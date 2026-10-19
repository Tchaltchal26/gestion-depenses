function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Variable d'environnement manquante : ${name}`);
  return v;
}

function key32(name: string): Buffer {
  const buf = Buffer.from(required(name), 'base64');
  if (buf.length !== 32) throw new Error(`${name} doit faire 32 octets encodés en base64`);
  return buf;
}

export interface AppConfig {
  port: number;
  databaseUrl: string;
  phoneEncKey: Buffer;
  phoneHashKey: Buffer;
  jwtSecret: string;
  smsWebhookSecret: string;
  smsProvider: string;
  dispatchTickMs: number;
  matching: {
    radiiMeters: number[];
    waveSizeApp: number;
    offerTtlAppSec: number;
    offerTtlSmsSec: number;
    searchTimeoutAppSec: number;
    searchTimeoutSmsSec: number;
    gpsFreshSec: number;
    landmarkFreshSec: number;
  };
  maxRideRequestsPerHour: number;
  smsMaxAgeSec: number;
}

export function loadConfig(): AppConfig {
  return {
    port: Number(process.env.PORT ?? 3000),
    databaseUrl: required('DATABASE_URL'),
    phoneEncKey: key32('PHONE_ENC_KEY'),
    phoneHashKey: key32('PHONE_HASH_KEY'),
    jwtSecret: required('JWT_SECRET'),
    smsWebhookSecret: required('SMS_WEBHOOK_SECRET'),
    smsProvider: process.env.SMS_PROVIDER ?? 'fake',
    dispatchTickMs: Number(process.env.DISPATCH_TICK_MS ?? 2000),
    matching: {
      radiiMeters: [1500, 3000, 5000],
      waveSizeApp: 3,
      offerTtlAppSec: 20,
      offerTtlSmsSec: 90,
      searchTimeoutAppSec: 300,
      searchTimeoutSmsSec: 480,
      gpsFreshSec: 120,
      landmarkFreshSec: 1800,
    },
    maxRideRequestsPerHour: 3,
    smsMaxAgeSec: 900,
  };
}

export const CONFIG = Symbol('CONFIG');
