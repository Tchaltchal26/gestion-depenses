import { auth, createApp, Ctx, GRAND_MARCHE, idem, login, newPhone, offset, quickDriver, quickUser, reset } from './helpers';

let ctx: Ctx;
beforeAll(async () => { ctx = await createApp(); });
afterAll(async () => { await ctx.app.close(); });
beforeEach(async () => { await reset(ctx); });

const createRide = (token: string, body: Record<string, unknown>, key = idem()) =>
  ctx.http().post('/v1/rides').set(auth(token)).set('Idempotency-Key', key).send(body);

describe('Authentification OTP', () => {
  it('connexion par code SMS (le code n\'apparaît pas en clair dans le journal SMS)', async () => {
    const phone = newPhone();
    const { token } = await login(ctx, phone);
    expect(token).toBeTruthy();
    const logged = await ctx.db.query(`SELECT body FROM sms_messages WHERE kind = 'otp'`);
    expect(logged[0].body).toContain('******');
    expect(logged[0].body).not.toMatch(/\d{6}/);
  });

  it('refuse un mauvais code, puis bloque après 5 essais', async () => {
    const phone = newPhone();
    await ctx.http().post('/v1/auth/otp/request').send({ phone }).expect(204);
    for (let i = 0; i < 5; i++) await ctx.http().post('/v1/auth/otp/verify').send({ phone, code: '000000' }).expect(401);
    await ctx.http().post('/v1/auth/otp/verify').send({ phone, code: '000000' }).expect(429);
  });

  it('refuse un code expiré', async () => {
    const phone = newPhone();
    await ctx.http().post('/v1/auth/otp/request').send({ phone }).expect(204);
    const code = ctx.sms.sent.at(-1)!.body.match(/(\d{6})/)![1];
    ctx.clock.advance(301);
    await ctx.http().post('/v1/auth/otp/verify').send({ phone, code }).expect(401);
  });

  it('limite à 3 envois de code par 15 minutes', async () => {
    const phone = newPhone();
    for (let i = 0; i < 3; i++) await ctx.http().post('/v1/auth/otp/request').send({ phone }).expect(204);
    await ctx.http().post('/v1/auth/otp/request').send({ phone }).expect(429);
    ctx.clock.advance(901);
    await ctx.http().post('/v1/auth/otp/request').send({ phone }).expect(204);
  });

  it('rejette un numéro non tchadien et les requêtes sans jeton', async () => {
    await ctx.http().post('/v1/auth/otp/request').send({ phone: '+33612345678' }).expect(400);
    await ctx.http().get('/v1/landmarks/search?q=marche').expect(401);
  });

  it('un compte suspendu perd immédiatement l\'accès', async () => {
    const client = await quickUser(ctx);
    const admin = await quickUser(ctx, 'admin');
    await ctx.http().get('/v1/landmarks/search?q=marche').set(auth(client.token)).expect(200);
    await ctx.http().post(`/v1/admin/users/${client.userId}/suspend`).set(auth(admin.token)).expect(204);
    await ctx.http().get('/v1/landmarks/search?q=marche').set(auth(client.token)).expect(401);
  });
});

describe('Inscription et validation chauffeur', () => {
  it('un chauffeur non validé ne peut pas se rendre disponible ; après validation admin, si', async () => {
    const { token, userId } = await login(ctx, newPhone());
    await ctx.http().post('/v1/driver/register').set(auth(token)).send({ name: 'Ali', transportType: 'clando', plate: 'TD-55' }).expect(201);
    // Le rôle a changé : il faut un nouveau jeton (le rôle est relu en base par le garde).
    await ctx.http().put('/v1/driver/availability').set(auth(token)).send({ available: true }).expect(403);
    const admin = await quickUser(ctx, 'admin');
    await ctx.http().post(`/v1/admin/drivers/${userId}/approve`).set(auth(token)).expect(403); // pas admin
    await ctx.http().post(`/v1/admin/drivers/${userId}/approve`).set(auth(admin.token)).expect(204);
    await ctx.http().put('/v1/driver/availability').set(auth(token)).send({ available: true }).expect(204);
    await ctx.http().post('/v1/driver/location').set(auth(token)).send({ lat: 13.83, lng: 20.83, batteryPct: 80 }).expect(204);
  });
});

describe('Repères et tarifs', () => {
  it('recherche tolérante aux fautes et aux alias', async () => {
    const { token } = await quickUser(ctx);
    const r1 = await ctx.http().get('/v1/landmarks/search?q=grand marcher').set(auth(token)).expect(200);
    expect(r1.body[0].name).toBe('Grand marche');
    const r2 = await ctx.http().get('/v1/landmarks/search?q=mustashfa').set(auth(token)).expect(200);
    expect(r2.body[0].name).toBe('Hopital provincial');
  });

  it('donne un prix indicatif arrondi à 50 FCFA', async () => {
    const { token } = await quickUser(ctx);
    const r = await ctx.http().get('/v1/tariffs/quote?type=clando&from=1&to=4').set(auth(token)).expect(200);
    expect(r.body.priceXaf % 50).toBe(0);
    expect(r.body.priceXaf).toBeGreaterThanOrEqual(250);
  });
});

describe('Course complète par application', () => {
  it('attribue la course au chauffeur éligible et refuse les autres', async () => {
    const client = await quickUser(ctx);
    const near = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 300), name: 'Proche' });
    const far = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 1000), name: 'Loin' });
    const taxi = await quickDriver(ctx, { type: 'taxi' });                          // mauvais type
    const pending = await quickDriver(ctx, { approved: false });                    // non validé
    const stale = await quickDriver(ctx, { locationAgeSec: 600 });                  // position périmée
    const outside = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 8000) });     // hors rayon max

    const res = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1, dropoffLandmarkId: 3 }).expect(201);
    expect(res.body.status).toBe('OFFERED');
    expect(res.body.priceXaf).toBeGreaterThan(0);

    const offered = ctx.notifier.events.filter((e) => e.event === 'offer.new').map((e) => e.userId).sort();
    expect(offered).toEqual([near.userId, far.userId].sort());
    for (const d of [taxi, pending, stale, outside]) expect(offered).not.toContain(d.userId);

    const [offerNear] = (await ctx.http().get('/v1/driver/offers').set(auth(near.token)).expect(200)).body;
    const [offerFar] = (await ctx.http().get('/v1/driver/offers').set(auth(far.token)).expect(200)).body;

    const acc = await ctx.http().post(`/v1/offers/${offerNear.offerId}/accept`).set(auth(near.token)).expect(200);
    expect(acc.body.status).toBe('ASSIGNED');
    expect(acc.body.client.phone).toBe(client.phone);
    await ctx.http().post(`/v1/offers/${offerFar.offerId}/accept`).set(auth(far.token)).expect(409);
    expect(ctx.notifier.events.some((e) => e.userId === far.userId && e.event === 'offer.revoked')).toBe(true);

    const view = await ctx.http().get(`/v1/rides/${res.body.id}`).set(auth(client.token)).expect(200);
    expect(view.body.status).toBe('ASSIGNED');
    expect(view.body.driver).toMatchObject({ name: 'Proche', plate: 'TD-1234', phone: near.phone });
    const pin = view.body.pickupPin;

    await ctx.http().post(`/v1/rides/${res.body.id}/arrived`).set(auth(near.token)).expect(200);
    await ctx.http().post(`/v1/rides/${res.body.id}/start`).set(auth(near.token)).send({ pin: pin === '0000' ? '1111' : '0000' }).expect(422);
    await ctx.http().post(`/v1/rides/${res.body.id}/start`).set(auth(near.token)).send({ pin }).expect(200);
    const done = await ctx.http().post(`/v1/rides/${res.body.id}/complete`).set(auth(near.token)).expect(200);
    expect(done.body.status).toBe('COMPLETED');
    expect(done.body.client.phone).toBeUndefined(); // numéro masqué une fois la course terminée

    const events = await ctx.db.query('SELECT to_status FROM ride_events WHERE ride_id = $1 ORDER BY id', [res.body.id]);
    expect(events.map((e) => e.to_status)).toEqual(['SEARCHING', 'OFFERED', 'ASSIGNED', 'ARRIVED', 'IN_PROGRESS', 'COMPLETED']);
  });

  it('idempotence : même clé = même course ; une seule demande active par client', async () => {
    const client = await quickUser(ctx);
    const key = idem();
    const a = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }, key).expect(201);
    const b = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }, key).expect(201);
    expect(b.body.id).toBe(a.body.id);
    const c = await createRide(client.token, { transportType: 'taxi', pickupLandmarkId: 1 }).expect(409);
    expect(c.body.code).toBe('ALREADY_ACTIVE');
    await ctx.http().post('/v1/rides').set(auth(client.token)).send({ transportType: 'clando', pickupLandmarkId: 1 }).expect(400);
  });

  it('refuse un départ hors zone et les champs inconnus', async () => {
    const client = await quickUser(ctx);
    await createRide(client.token, { transportType: 'clando', pickupLat: 12.1, pickupLng: 15.0 }).expect(422);
    await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1, hack: true }).expect(400);
    await createRide(client.token, { transportType: 'avion', pickupLandmarkId: 1 }).expect(400);
  });

  it('limite le nombre de demandes par heure', async () => {
    const client = await quickUser(ctx);
    for (let i = 0; i < 3; i++) {
      const r = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
      await ctx.http().post(`/v1/rides/${r.body.id}/cancel`).set(auth(client.token)).expect(200);
    }
    await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(429);
  });

  it('confidentialité : un autre client ne voit pas la course, un client ne peut pas accepter d\'offre', async () => {
    const client = await quickUser(ctx);
    const other = await quickUser(ctx);
    await quickDriver(ctx);
    const r = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    await ctx.http().get(`/v1/rides/${r.body.id}`).set(auth(other.token)).expect(404);
    const [offer] = await ctx.db.query('SELECT id FROM ride_offers');
    await ctx.http().post(`/v1/offers/${offer.id}/accept`).set(auth(client.token)).expect(403);
  });
});

describe('Vagues, expirations et absence de chauffeur', () => {
  it('passe à la vague suivante après expiration, refuse une acceptation tardive, puis NO_DRIVER', async () => {
    const client = await quickUser(ctx);
    const drivers: Awaited<ReturnType<typeof quickDriver>>[] = [];
    for (let i = 0; i < 4; i++) drivers.push(await quickDriver(ctx, { at: offset(GRAND_MARCHE, 100 * (i + 1)) }));

    const r = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    let offers = await ctx.db.query('SELECT driver_id, wave FROM ride_offers ORDER BY id');
    expect(offers.map((o) => String(o.driver_id))).toEqual(drivers.slice(0, 3).map((d) => d.userId)); // les 3 plus proches

    ctx.clock.advance(21);
    await ctx.rides.tick();
    offers = await ctx.db.query('SELECT driver_id, wave, response FROM ride_offers ORDER BY id');
    expect(offers.slice(0, 3).every((o) => o.response === 'expired')).toBe(true);
    expect(offers[3]).toMatchObject({ driver_id: drivers[3].userId, wave: 2, response: null });

    const [late] = await ctx.db.query('SELECT id FROM ride_offers WHERE driver_id = $1', [drivers[0].userId]);
    await ctx.http().post(`/v1/offers/${late.id}/accept`).set(auth(drivers[0].token)).expect(409);

    ctx.clock.advance(300);
    await ctx.rides.tick();
    const view = await ctx.http().get(`/v1/rides/${r.body.id}`).set(auth(client.token)).expect(200);
    expect(view.body.status).toBe('NO_DRIVER');
    expect(ctx.notifier.events.some((e) => e.userId === client.userId && e.event === 'ride.no_driver')).toBe(true);
  });

  it('un refus déclenche immédiatement la vague suivante', async () => {
    ctx.config.matching.waveSizeApp = 1;
    try {
      const client = await quickUser(ctx);
      const d1 = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 100) });
      const d2 = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 200) });
      await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
      const [o1] = (await ctx.http().get('/v1/driver/offers').set(auth(d1.token))).body;
      await ctx.http().post(`/v1/offers/${o1.offerId}/decline`).set(auth(d1.token)).expect(204);
      const offers2 = (await ctx.http().get('/v1/driver/offers').set(auth(d2.token))).body;
      expect(offers2).toHaveLength(1);
    } finally {
      ctx.config.matching.waveSizeApp = 3;
    }
  });

  it('élargit le rayon quand personne n\'est proche', async () => {
    const client = await quickUser(ctx);
    const d = await quickDriver(ctx, { at: offset(GRAND_MARCHE, 4000) }); // hors 1,5 et 3 km, dans 5 km
    await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    const [ride] = await ctx.db.query('SELECT radius_idx, status FROM rides');
    expect(ride).toMatchObject({ radius_idx: 2, status: 'OFFERED' });
    expect((await ctx.http().get('/v1/driver/offers').set(auth(d.token))).body).toHaveLength(1);
  });

  it('un chauffeur déjà en course ne reçoit pas d\'offre ; batterie < 5 % le retire', async () => {
    const c1 = await quickUser(ctx);
    const c2 = await quickUser(ctx);
    const d = await quickDriver(ctx);
    await createRide(c1.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    const [o] = (await ctx.http().get('/v1/driver/offers').set(auth(d.token))).body;
    await ctx.http().post(`/v1/offers/${o.offerId}/accept`).set(auth(d.token)).expect(200);
    await createRide(c2.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    expect(await ctx.db.query('SELECT 1 FROM ride_offers WHERE driver_id = $1', [d.userId])).toHaveLength(1);

    const d2 = await quickDriver(ctx);
    await ctx.http().post('/v1/driver/location').set(auth(d2.token)).send({ ...GRAND_MARCHE, batteryPct: 3 }).expect(204);
    const [p] = await ctx.db.query('SELECT available FROM driver_presence WHERE driver_id = $1', [d2.userId]);
    expect(p.available).toBe(false);
  });

  it('annulation par le client : offres révoquées, chauffeur attribué prévenu', async () => {
    const client = await quickUser(ctx);
    const d = await quickDriver(ctx);
    const r = await createRide(client.token, { transportType: 'clando', pickupLandmarkId: 1 }).expect(201);
    const [o] = (await ctx.http().get('/v1/driver/offers').set(auth(d.token))).body;
    await ctx.http().post(`/v1/offers/${o.offerId}/accept`).set(auth(d.token)).expect(200);
    const c = await ctx.http().post(`/v1/rides/${r.body.id}/cancel`).set(auth(client.token)).expect(200);
    expect(c.body.status).toBe('CANCELLED_BY_CLIENT');
    expect(ctx.notifier.events.some((e) => e.userId === d.userId && e.event === 'ride.cancelled')).toBe(true);
    await ctx.http().post(`/v1/rides/${r.body.id}/cancel`).set(auth(client.token)).expect(409);
  });
});
