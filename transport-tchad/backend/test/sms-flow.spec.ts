import { createApp, Ctx, inboundSms, lastSmsTo, newPhone, quickDriver, reset, smsTo } from './helpers';

let ctx: Ctx;
beforeAll(async () => { ctx = await createApp(); });
afterAll(async () => { await ctx.app.close(); });
beforeEach(async () => { await reset(ctx); });

const codeIn = (s: string | undefined) => s!.match(/Course ([A-Z0-9]{5})/)![1];

describe('Webhook SMS', () => {
  it('rejette une signature absente ou fausse', async () => {
    await inboundSms(ctx, newPhone(), 'CLANDO grand marche', {}, 'mauvais-secret').expect(401);
    await ctx.http().post('/v1/sms/inbound/simgw').send({ messageId: 'x', from: newPhone(), text: 'aide' }).expect(401);
    expect(await ctx.db.query('SELECT 1 FROM rides')).toHaveLength(0);
  });

  it('ignore un message renvoyé deux fois par la passerelle (même identifiant)', async () => {
    const phone = newPhone();
    await quickDriver(ctx, { channel: 'sms' });
    const r1 = await inboundSms(ctx, phone, 'CLANDO grand marche', { messageId: 'dup-1' }).expect(200);
    const r2 = await inboundSms(ctx, phone, 'CLANDO grand marche', { messageId: 'dup-1' }).expect(200);
    expect(r1.body.status).toBe('processed');
    expect(r2.body.status).toBe('duplicate');
    expect(await ctx.db.query('SELECT 1 FROM rides')).toHaveLength(1);
  });
});

describe('Course complète par SMS (client et chauffeur sans Internet)', () => {
  it('demande → offre → OUI → confirmations avec numéros → FIN', async () => {
    const client = newPhone();
    const driver = await quickDriver(ctx, { channel: 'sms', at: null, name: 'Moussa' });

    // Le chauffeur se déclare disponible par SMS près d'un repère.
    await inboundSms(ctx, driver.phone, 'DISPO marche central').expect(200);
    expect(lastSmsTo(ctx, driver.phone)).toContain('DISPONIBLE pres de Grand marche');

    await inboundSms(ctx, client, 'Clando grand marché > hôpital').expect(200);
    const offer = lastSmsTo(ctx, driver.phone)!;
    expect(offer).toMatch(/^Course [A-Z0-9]{5}: CLANDO Grand marche > Hopital provincial, ~\d+F\. Client a \d+m\. Repondez OUI [A-Z0-9]{5} \(90s\)$/);
    const code = codeIn(offer);

    await inboundSms(ctx, driver.phone, `oui ${code.toLowerCase()}`).expect(200);
    const toClient = lastSmsTo(ctx, client)!;
    expect(toClient).toContain(`${code}: chauffeur Moussa, TD-1234 rouge. Tel ${driver.phone}`);
    expect(toClient).toMatch(/Code a donner: \d{4}$/);
    expect(lastSmsTo(ctx, driver.phone)).toContain(`OK ${code}. Client: Client ${client}`);

    await inboundSms(ctx, driver.phone, `FIN ${code}`).expect(200);
    expect(lastSmsTo(ctx, driver.phone)).toBe(`Course ${code} terminee. Merci.`);
    const [ride] = await ctx.db.query('SELECT status, channel FROM rides');
    expect(ride).toEqual({ status: 'COMPLETED', channel: 'sms' });

    // Économie : nombre de SMS sortants payés pour cette course (hors DISPO).
    const out = await ctx.db.query(`SELECT kind FROM sms_messages WHERE direction = 'out' AND ride_id IS NOT NULL`);
    expect(out.map((o) => o.kind).sort()).toEqual(['assigned', 'assigned', 'offer']);
  });

  it('format compact envoyé par l\'application sans data (position GPS)', async () => {
    const client = newPhone();
    const driver = await quickDriver(ctx, { channel: 'sms' });
    await inboundSms(ctx, client, 'M1 C 13.8292 20.8321 D 3').expect(200);
    expect(lastSmsTo(ctx, driver.phone)).toMatch(/CLANDO position GPS > Hopital provincial/);
  });

  it('lieu ambigu : propose des choix, le client répond par un chiffre', async () => {
    const client = newPhone();
    await quickDriver(ctx, { channel: 'sms' });
    await inboundSms(ctx, client, 'CLANDO marche').expect(200);
    const choices = lastSmsTo(ctx, client)!;
    expect(choices).toMatch(/^Plusieurs lieux: 1=.+ 2=.+ 3=.+\. Repondez 1 ou 2 ou 3$/);
    const second = choices.match(/2=(.+?) 3=/)![1];
    expect(await ctx.db.query('SELECT 1 FROM rides')).toHaveLength(0);

    await inboundSms(ctx, client, '2').expect(200);
    const [ride] = await ctx.db.query('SELECT pickup_text FROM rides');
    expect(ride.pickup_text).toBe(second);
  });

  it('lieu introuvable : message d\'aide, pas de course', async () => {
    const client = newPhone();
    await inboundSms(ctx, client, 'TAXI zzqxw').expect(200);
    expect(lastSmsTo(ctx, client)).toContain('non trouve');
    expect(await ctx.db.query('SELECT 1 FROM rides')).toHaveLength(0);
  });

  it('SMS arrivé en retard (réseau) : la demande n\'est pas servie', async () => {
    const client = newPhone();
    const receivedAt = new Date(ctx.clock.now().getTime() - 20 * 60_000).toISOString();
    await inboundSms(ctx, client, 'CLANDO grand marche', { receivedAt }).expect(200);
    expect(lastSmsTo(ctx, client)).toContain('arrivee en retard');
    expect(await ctx.db.query('SELECT 1 FROM rides')).toHaveLength(0);
  });

  it('client qui renvoie le même SMS : pas de seconde course, rappel du code', async () => {
    const client = newPhone();
    await quickDriver(ctx, { channel: 'sms' });
    await inboundSms(ctx, client, 'CLANDO grand marche').expect(200);
    await inboundSms(ctx, client, 'CLANDO grand marche').expect(200);
    await inboundSms(ctx, client, 'TAXI gare').expect(200);
    const rides = await ctx.db.query('SELECT public_code FROM rides');
    expect(rides).toHaveLength(1);
    expect(smsTo(ctx, client).filter((m) => m.includes(`deja une demande en cours (${rides[0].public_code})`))).toHaveLength(2);
  });

  it('réponse du chauffeur après expiration : refusée ; offre suivante à un autre chauffeur', async () => {
    const client = newPhone();
    const d1 = await quickDriver(ctx, { channel: 'sms', name: 'Un' });
    await inboundSms(ctx, client, 'CLANDO grand marche').expect(200);
    const code = codeIn(lastSmsTo(ctx, d1.phone));
    const d2 = await quickDriver(ctx, { channel: 'sms', name: 'Deux' });

    ctx.clock.advance(91);
    await ctx.rides.tick();
    expect(lastSmsTo(ctx, d2.phone)).toContain(`Course ${code}`);

    await inboundSms(ctx, d1.phone, `OUI ${code}`).expect(200);
    expect(lastSmsTo(ctx, d1.phone)).toBe(`${code} deja attribuee ou expiree.`);
    await inboundSms(ctx, d2.phone, `OUI ${code}`).expect(200);
    expect(lastSmsTo(ctx, client)).toContain('chauffeur Deux');
  });

  it('une vague ne contient jamais plus d\'un chauffeur SMS (coût)', async () => {
    for (let i = 0; i < 3; i++) await quickDriver(ctx, { channel: 'sms' });
    await inboundSms(ctx, newPhone(), 'CLANDO grand marche').expect(200);
    expect(await ctx.db.query(`SELECT 1 FROM sms_messages WHERE kind = 'offer'`)).toHaveLength(1);
  });

  it('aucun chauffeur : le client est prévenu par SMS à la fin du délai', async () => {
    const client = newPhone();
    await inboundSms(ctx, client, 'CLANDO grand marche').expect(200);
    ctx.clock.advance(481);
    await ctx.rides.tick();
    expect(lastSmsTo(ctx, client)).toMatch(/^Aucun chauffeur disponible pour [A-Z0-9]{5}/);
  });

  it('ANNULER par SMS : course annulée et chauffeur SMS prévenu', async () => {
    const client = newPhone();
    const d = await quickDriver(ctx, { channel: 'sms' });
    await inboundSms(ctx, client, 'CLANDO grand marche').expect(200);
    const code = codeIn(lastSmsTo(ctx, d.phone));
    await inboundSms(ctx, d.phone, `OUI ${code}`).expect(200);
    await inboundSms(ctx, client, 'annuler').expect(200);
    expect(lastSmsTo(ctx, client)).toBe(`Demande ${code} annulee.`);
    expect(lastSmsTo(ctx, d.phone)).toBe(`Course ${code} annulee par le client.`);
  });

  it('chauffeur non validé : refus ; numéro inconnu sans commande : aide', async () => {
    const d = await quickDriver(ctx, { channel: 'sms', approved: false });
    await inboundSms(ctx, d.phone, 'DISPO grand marche').expect(200);
    expect(lastSmsTo(ctx, d.phone)).toContain('non valide');
    const someone = newPhone();
    await inboundSms(ctx, someone, 'bonjour').expect(200);
    expect(lastSmsTo(ctx, someone)).toContain('Pour un transport envoyez');
  });

  it('tous les SMS envoyés pendant les tests tiennent en 1 SMS (<= 160 caractères, sans accents)', async () => {
    const client = newPhone();
    const d = await quickDriver(ctx, { channel: 'sms' });
    await inboundSms(ctx, client, 'CLANDO universite adam barka > marche a betail').expect(200);
    await inboundSms(ctx, d.phone, `OUI ${codeIn(lastSmsTo(ctx, d.phone))}`).expect(200);
    for (const m of ctx.sms.sent) {
      expect(m.body.length).toBeLessThanOrEqual(160);
      expect(m.body).toMatch(/^[\x20-\x7E\n]*$/);
    }
  });
});
