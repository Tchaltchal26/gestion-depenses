import { Body, Controller, Get, Header, HttpCode, Inject, NotFoundException, Post } from '@nestjs/common';
import { readFileSync } from 'fs';
import { join } from 'path';
import { AuthService } from '../auth/auth.service';
import { DbService } from '../db/db.service';
import { RidesService } from '../rides/rides.service';
import { SmsInboundService } from '../sms/sms.inbound.service';
import { FakeSmsProvider, SMS_PROVIDER } from '../sms/sms.provider';

/**
 * OUTILS DE DÉVELOPPEMENT — actifs uniquement si DEV_TOOLS=1 (jamais en production).
 * Simulateur de téléphones dans le navigateur : http://localhost:3000/dev
 */
@Controller('dev')
export class DevController {
  constructor(
    private readonly inbound: SmsInboundService,
    private readonly rides: RidesService,
    private readonly auth: AuthService,
    private readonly db: DbService,
    @Inject(SMS_PROVIDER) private readonly sms: FakeSmsProvider,
  ) {}

  @Get() @Header('Content-Type', 'text/html; charset=utf-8')
  page() {
    return readFileSync(join(__dirname, 'simulateur.html'), 'utf8');
  }

  /** Simule un SMS reçu par la passerelle (sans signature : outil local uniquement). */
  @Post('sms') @HttpCode(200)
  receive(@Body() b: { from: string; text: string }) {
    return this.inbound.handle('dev', { messageId: `dev-${Date.now()}-${Math.random()}`, from: b.from, text: b.text });
  }

  @Get('outbox')
  outbox() {
    return this.sms.sent.slice(-200);
  }

  @Post('tick') @HttpCode(200)
  async tick() {
    await this.rides.tick();
    return { ok: true };
  }

  /** Crée un chauffeur déjà validé (raccourci de test à la place du parcours d'inscription + validation admin). */
  @Post('driver') @HttpCode(200)
  async driver(@Body() b: { phone: string; name: string; type: 'taxi' | 'rakcha' | 'clando'; plate?: string }) {
    const e164 = this.auth.phoneOrThrow(b.phone);
    const u = await this.auth.findOrCreateUser(e164);
    await this.db.query(`UPDATE users SET role = 'driver', name = $2 WHERE id = $1`, [u.id, b.name]);
    await this.db.query(
      `INSERT INTO drivers (user_id, transport_type, verification_status, vehicle_plate, vehicle_color)
       VALUES ($1, $2, 'approved', $3, 'rouge') ON CONFLICT (user_id) DO UPDATE SET transport_type = EXCLUDED.transport_type, verification_status = 'approved'`,
      [u.id, b.type, b.plate ?? 'TD-0000']);
    await this.db.query('INSERT INTO driver_presence (driver_id) VALUES ($1) ON CONFLICT DO NOTHING', [u.id]);
    return { ok: true, phone: e164 };
  }

  @Get('rides')
  ridesList() {
    return this.db.query(
      `SELECT public_code, status, transport_type, channel, pickup_text, dropoff_text, price_quoted_xaf, created_at
         FROM rides ORDER BY id DESC LIMIT 20`);
  }
}

export function devToolsEnabled() {
  if (process.env.DEV_TOOLS !== '1') return false;
  if (process.env.NODE_ENV === 'production') throw new NotFoundException('DEV_TOOLS interdit en production');
  return true;
}
