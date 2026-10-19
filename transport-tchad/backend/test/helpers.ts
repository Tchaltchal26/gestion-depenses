import { INestApplication, ValidationPipe } from '@nestjs/common';
import { JwtService } from '@nestjs/jwt';
import { Test } from '@nestjs/testing';
import { createHmac } from 'crypto';
import request from 'supertest';
import { AppModule } from '../src/app.module';
import { Clock } from '../src/common/clock';
import { CryptoService } from '../src/common/crypto.service';
import { AppConfig, CONFIG } from '../src/config';
import { DbService } from '../src/db/db.service';
import { Notifier } from '../src/notifier';
import { RidesService } from '../src/rides/rides.service';
import { FakeSmsProvider, SMS_PROVIDER } from '../src/sms/sms.provider';

// Points de test (coordonnées fictives, cohérentes avec sql/900_seed_dev.sql).
export const GRAND_MARCHE = { lat: 13.829, lng: 20.832 };
export const offset = (p: { lat: number; lng: number }, metersNorth: number) => ({ lat: p.lat + metersNorth / 111_320, lng: p.lng });

export interface Ctx {
  app: INestApplication;
  http: () => ReturnType<typeof request>;
  db: DbService;
  sms: FakeSmsProvider;
  notifier: Notifier;
  clock: Clock;
  rides: RidesService;
  crypto: CryptoService;
  jwt: JwtService;
  config: AppConfig;
}

export async function createApp(): Promise<Ctx> {
  const mod = await Test.createTestingModule({ imports: [AppModule] }).compile();
  const app = mod.createNestApplication({ rawBody: true });
  app.useGlobalPipes(new ValidationPipe({ whitelist: true, forbidNonWhitelisted: true, transform: true }));
  await app.init();
  return {
    app,
    http: () => request(app.getHttpServer()),
    db: mod.get(DbService),
    sms: mod.get(SMS_PROVIDER),
    notifier: mod.get(Notifier),
    clock: mod.get(Clock),
    rides: mod.get(RidesService),
    crypto: mod.get(CryptoService),
    jwt: mod.get(JwtService),
    config: mod.get(CONFIG),
  };
}

export async function reset(ctx: Ctx) {
  await ctx.db.query(
    'TRUNCATE ride_events, ride_offers, rides, driver_presence, drivers, otp_codes, sms_messages, sms_pending, users RESTART IDENTITY CASCADE');
  ctx.sms.sent.length = 0;
  ctx.notifier.events.length = 0;
  ctx.clock.reset();
}

let phoneSeq = 0;
/** Numéro tchadien de test unique (format +2356XXXXXXX). */
export const newPhone = () => '+2356' + String(1_000_000 + ++phoneSeq).slice(-7);

/** Connexion complète via OTP (le code est lu dans les SMS simulés). */
export async function login(ctx: Ctx, phone: string): Promise<{ token: string; userId: string }> {
  await ctx.http().post('/v1/auth/otp/request').send({ phone }).expect(204);
  const msg = [...ctx.sms.sent].reverse().find((m) => m.to === phone)!;
  const code = msg.body.match(/(\d{6})/)![1];
  const res = await ctx.http().post('/v1/auth/otp/verify').send({ phone, code }).expect(200);
  return { token: res.body.token, userId: res.body.userId };
}

/** Création rapide d'un utilisateur directement en base (sans OTP) + jeton. */
export async function quickUser(ctx: Ctx, role: 'client' | 'admin' = 'client', phone = newPhone()) {
  const [u] = await ctx.db.query(
    'INSERT INTO users (phone_hash, phone_enc, role, name) VALUES ($1, $2, $3, $4) RETURNING id',
    [ctx.crypto.hashPhone(phone), ctx.crypto.encrypt(phone), role, role === 'admin' ? 'Admin' : 'Client test']);
  return { userId: String(u.id), phone, token: ctx.jwt.sign({ sub: String(u.id), role }) };
}

export interface DriverOpts {
  type?: 'taxi' | 'rakcha' | 'clando';
  at?: { lat: number; lng: number } | null;
  channel?: 'app' | 'sms';
  approved?: boolean;
  locationAgeSec?: number;
  name?: string;
}

/** Chauffeur créé directement en base (validé, disponible, positionné) + jeton. */
export async function quickDriver(ctx: Ctx, o: DriverOpts = {}) {
  const phone = newPhone();
  const at = o.at === undefined ? GRAND_MARCHE : o.at;
  const [u] = await ctx.db.query(
    `INSERT INTO users (phone_hash, phone_enc, role, name) VALUES ($1, $2, 'driver', $3) RETURNING id`,
    [ctx.crypto.hashPhone(phone), ctx.crypto.encrypt(phone), o.name ?? 'Moussa']);
  const id = String(u.id);
  await ctx.db.query(
    `INSERT INTO drivers (user_id, transport_type, verification_status, vehicle_plate, vehicle_color) VALUES ($1, $2, $3, 'TD-1234', 'rouge')`,
    [id, o.type ?? 'clando', o.approved === false ? 'pending' : 'approved']);
  const locAt = new Date(ctx.clock.now().getTime() - (o.locationAgeSec ?? 0) * 1000);
  await ctx.db.query(
    `INSERT INTO driver_presence (driver_id, available, channel, location, location_source, location_at)
     VALUES ($1, true, $2, CASE WHEN $3::float IS NULL THEN NULL ELSE ST_SetSRID(ST_MakePoint($4, $3), 4326)::geography END,
             $5, $6)`,
    [id, o.channel ?? 'app', at?.lat ?? null, at?.lng ?? null, o.channel === 'sms' ? 'landmark' : 'gps', locAt]);
  return { userId: id, phone, token: ctx.jwt.sign({ sub: id, role: 'driver' }) };
}

export const auth = (token: string) => ({ Authorization: `Bearer ${token}` });

let idemSeq = 0;
export const idem = () => `test-${Date.now()}-${++idemSeq}`;

/** Envoie un SMS entrant signé au webhook, comme le ferait la passerelle. */
let smsSeq = 0;
export function inboundSms(ctx: Ctx, from: string, text: string, extra: Record<string, unknown> = {}, secret = 'test-webhook-secret') {
  const body = JSON.stringify({ messageId: `m-${Date.now()}-${++smsSeq}`, from, text, ...extra });
  const sig = createHmac('sha256', secret).update(body).digest('hex');
  return ctx.http().post('/v1/sms/inbound/simgw').set('Content-Type', 'application/json').set('X-Signature', sig).send(body);
}

export const smsTo = (ctx: Ctx, phone: string) => ctx.sms.sent.filter((m) => m.to === phone).map((m) => m.body);
export const lastSmsTo = (ctx: Ctx, phone: string) => smsTo(ctx, phone).at(-1);
