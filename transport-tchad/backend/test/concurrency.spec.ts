import { auth, createApp, Ctx, GRAND_MARCHE, idem, inboundSms, offset, quickDriver, quickUser, reset } from './helpers';

let ctx: Ctx;
beforeAll(async () => { ctx = await createApp(); });
afterAll(async () => { await ctx.app.close(); });
beforeEach(async () => { await reset(ctx); });

describe('Anti double attribution (accès concurrents réels sur PostgreSQL)', () => {
  it('20 tours × 10 chauffeurs qui acceptent la même course au même instant : exactement 1 gagnant à chaque fois', async () => {
    ctx.config.matching.waveSizeApp = 10;
    try {
      const drivers: Awaited<ReturnType<typeof quickDriver>>[] = [];
      for (let i = 0; i < 10; i++) drivers.push(await quickDriver(ctx, { at: offset(GRAND_MARCHE, 50 * i) }));
      for (let round = 0; round < 20; round++) {
        const client = await quickUser(ctx);
        const r = await ctx.http().post('/v1/rides').set(auth(client.token)).set('Idempotency-Key', idem())
          .send({ transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
        const offers = await ctx.db.query('SELECT id, driver_id FROM ride_offers WHERE ride_id = $1', [r.body.id]);
        expect(offers).toHaveLength(10);

        const results = await Promise.all(offers.map((o) => {
          const d = drivers.find((x) => x.userId === String(o.driver_id))!;
          return ctx.http().post(`/v1/offers/${o.id}/accept`).set(auth(d.token));
        }));
        const codes = results.map((x) => x.status).sort();
        expect(codes.filter((c) => c === 200)).toHaveLength(1);
        expect(codes.filter((c) => c === 409)).toHaveLength(9);

        const [ride] = await ctx.db.query('SELECT status, driver_id FROM rides WHERE id = $1', [r.body.id]);
        expect(ride.status).toBe('ASSIGNED');
        const accepted = await ctx.db.query(`SELECT driver_id FROM ride_offers WHERE ride_id = $1 AND response = 'accepted'`, [r.body.id]);
        expect(accepted).toHaveLength(1);
        expect(String(accepted[0].driver_id)).toBe(String(ride.driver_id));
        const assignedEvents = await ctx.db.query(`SELECT 1 FROM ride_events WHERE ride_id = $1 AND to_status = 'ASSIGNED'`, [r.body.id]);
        expect(assignedEvents).toHaveLength(1);

        // Libère le chauffeur gagnant pour le tour suivant.
        await ctx.db.query(`UPDATE rides SET status = 'COMPLETED' WHERE id = $1`, [r.body.id]);
      }
    } finally {
      ctx.config.matching.waveSizeApp = 3;
    }
  });

  it('double appui du même chauffeur (2 requêtes simultanées) : 1 succès', async () => {
    const client = await quickUser(ctx);
    const d = await quickDriver(ctx);
    await ctx.http().post('/v1/rides').set(auth(client.token)).set('Idempotency-Key', idem())
      .send({ transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    const [o] = await ctx.db.query('SELECT id FROM ride_offers');
    const res = await Promise.all([1, 2, 3].map(() => ctx.http().post(`/v1/offers/${o.id}/accept`).set(auth(d.token))));
    expect(res.map((r) => r.status).sort()).toEqual([200, 409, 409]);
  });

  it('acceptation par SMS et par application en même temps : 1 seul gagnant', async () => {
    ctx.config.matching.waveSizeApp = 2;
    try {
      const appDriver = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 10) });
      const smsDriver = await quickDriver(ctx, { channel: 'sms', at: offset(GRAND_MARCHE, 20) });
      const client = await quickUser(ctx);
      const r = await ctx.http().post('/v1/rides').set(auth(client.token)).set('Idempotency-Key', idem())
        .send({ transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
      const [appOffer] = await ctx.db.query('SELECT id FROM ride_offers WHERE driver_id = $1', [appDriver.userId]);
      await Promise.all([
        ctx.http().post(`/v1/offers/${appOffer.id}/accept`).set(auth(appDriver.token)),
        inboundSms(ctx, smsDriver.phone, `OUI ${r.body.code}`),
      ]);
      const accepted = await ctx.db.query(`SELECT 1 FROM ride_offers WHERE response = 'accepted'`);
      expect(accepted).toHaveLength(1);
      expect((await ctx.db.query('SELECT status FROM rides'))[0].status).toBe('ASSIGNED');
    } finally {
      ctx.config.matching.waveSizeApp = 3;
    }
  });

  it('10 demandes simultanées du même client avec des clés différentes : 1 seule course créée', async () => {
    const client = await quickUser(ctx);
    const res = await Promise.all(Array.from({ length: 10 }, () =>
      ctx.http().post('/v1/rides').set(auth(client.token)).set('Idempotency-Key', idem()).send({ transportType: 'clando', pickupLandmarkId: 1 })));
    expect(res.filter((r) => r.status === 201)).toHaveLength(1);
    expect(res.filter((r) => r.status === 409)).toHaveLength(9);
  });

  it('expiration (tick) en concurrence avec des acceptations : jamais deux chauffeurs', async () => {
    for (let round = 0; round < 10; round++) {
      await reset(ctx);
      const drivers = [await quickDriver(ctx), await quickDriver(ctx), await quickDriver(ctx)];
      const client = await quickUser(ctx);
      await ctx.http().post('/v1/rides').set(auth(client.token)).set('Idempotency-Key', idem())
        .send({ transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
      const offers = await ctx.db.query('SELECT id, driver_id FROM ride_offers');
      ctx.clock.advance(19.99); // à la limite de l'expiration
      await Promise.all([
        ctx.rides.tick(),
        ...offers.map((o) => ctx.http().post(`/v1/offers/${o.id}/accept`).set(auth(drivers.find((d) => d.userId === String(o.driver_id))!.token))),
        ctx.rides.tick(),
      ]);
      const accepted = await ctx.db.query(`SELECT 1 FROM ride_offers WHERE response = 'accepted'`);
      expect(accepted.length).toBeLessThanOrEqual(1);
      const [ride] = await ctx.db.query('SELECT status, driver_id FROM rides');
      if (accepted.length) expect(ride.status).toBe('ASSIGNED');
    }
  });
});
