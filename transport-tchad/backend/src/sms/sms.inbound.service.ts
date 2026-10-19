import { ConflictException, HttpException, Inject, Injectable, UnprocessableEntityException } from '@nestjs/common';
import { createHash } from 'crypto';
import { AuthService } from '../auth/auth.service';
import { Clock } from '../common/clock';
import { CryptoService } from '../common/crypto.service';
import { normalizeChadPhone } from '../common/phone';
import { normalizeText } from '../common/text';
import { AppConfig, CONFIG } from '../config';
import { DbService } from '../db/db.service';
import { GeoService } from '../geo/geo.service';
import { DriversService } from '../rides/drivers.service';
import { CreateRideInput, RidesService } from '../rides/rides.service';
import { parseSms, ParsedSms } from './sms.parser';
import { SmsSender } from './sms.sender';
import { Sms } from './sms.templates';

export interface InboundSms {
  messageId: string;
  from: string;
  text: string;
  receivedAt?: string;
}

const PENDING_TTL_SEC = 600;

@Injectable()
export class SmsInboundService {
  constructor(
    private readonly db: DbService,
    private readonly crypto: CryptoService,
    private readonly auth: AuthService,
    private readonly geo: GeoService,
    private readonly rides: RidesService,
    private readonly drivers: DriversService,
    private readonly sms: SmsSender,
    private readonly clock: Clock,
    @Inject(CONFIG) private readonly config: AppConfig,
  ) {}

  async handle(provider: string, msg: InboundSms): Promise<{ status: string }> {
    const phone = normalizeChadPhone(msg.from);
    if (!phone) return { status: 'ignored_invalid_number' };
    const ph = this.crypto.hashPhone(phone);

    // Le fournisseur peut renvoyer le même message (nouvelle tentative de webhook) : on l'ignore.
    const inserted = await this.db.query(
      `INSERT INTO sms_messages (direction, provider, provider_message_id, phone_hash, body, status)
       VALUES ('in', $1, $2, $3, $4, 'received') ON CONFLICT (provider, provider_message_id) DO NOTHING RETURNING id`,
      [provider, msg.messageId, ph, (msg.text ?? '').slice(0, 1000)]);
    if (!inserted.length) return { status: 'duplicate' };

    const user = await this.auth.findOrCreateUser(phone);
    if (user.status !== 'active') return { status: 'ignored_suspended' };
    const parsed = parseSms(msg.text);
    const reply = (body: string, kind = 'reply') => this.sms.sendToPhone(phone, body, kind);

    const [driver] = await this.db.query('SELECT user_id FROM drivers WHERE user_id = $1', [user.id]);
    if (driver) {
      await this.handleDriver(String(user.id), parsed, reply);
      return { status: 'processed' };
    }

    const receivedAt = msg.receivedAt ? new Date(msg.receivedAt) : this.clock.now();
    const tooOld = this.clock.now().getTime() - receivedAt.getTime() > this.config.smsMaxAgeSec * 1000;
    await this.handleClient(String(user.id), ph, msg.text, parsed, tooOld, reply);
    return { status: 'processed' };
  }

  private async handleClient(userId: string, ph: string, rawText: string, p: ParsedSms, tooOld: boolean, reply: (b: string, k?: string) => Promise<void>) {
    switch (p.kind) {
      case 'compact':
      case 'request': {
        if (tooOld) return reply(Sms.tooLate());
        const base: CreateRideInput = { transportType: p.type, channel: 'sms', idempotencyKey: this.smsIdempotencyKey(rawText) };
        if (p.kind === 'compact') {
          return this.createAndReport(userId, { ...base, pickupLat: p.lat, pickupLng: p.lng, pickupText: 'position GPS',
            dropoffLandmarkId: p.dropoffLandmarkId ?? undefined }, reply);
        }
        const res = await this.geo.resolve(p.pickupText);
        const dropoff = p.dropoffText ? await this.geo.resolve(p.dropoffText) : null;
        const dropoffFields = dropoff?.status === 'unique'
          ? { dropoffLandmarkId: dropoff.landmark.id }
          : { dropoffText: p.dropoffText ?? undefined }; // destination non bloquante : texte brut conservé
        if (res.status === 'none') return reply(Sms.notFound(p.pickupText));
        if (res.status === 'ambiguous') {
          await this.db.query(
            `INSERT INTO sms_pending (phone_hash, data, expires_at) VALUES ($1, $2, $3)
             ON CONFLICT (phone_hash) DO UPDATE SET data = EXCLUDED.data, expires_at = EXCLUDED.expires_at`,
            [ph, JSON.stringify({ ...base, ...dropoffFields, candidates: res.candidates.map((c) => c.id) }),
             new Date(this.clock.now().getTime() + PENDING_TTL_SEC * 1000)]);
          return reply(Sms.choices(res.candidates.map((c) => c.name)), 'choices');
        }
        return this.createAndReport(userId, { ...base, ...dropoffFields, pickupLandmarkId: res.landmark.id }, reply);
      }
      case 'choice': {
        const [pending] = await this.db.query(
          'DELETE FROM sms_pending WHERE phone_hash = $1 AND expires_at > $2 RETURNING data', [ph, this.clock.now()]);
        const id = pending?.data.candidates[p.n - 1];
        if (!id) return reply(Sms.clientHelp(), 'help');
        const { candidates, ...input } = pending.data;
        return this.createAndReport(userId, { ...input, pickupLandmarkId: id }, reply);
      }
      case 'cancel': {
        const active = await this.rides.findActiveRideForClient(userId);
        if (!active) return reply(Sms.nothingToCancel());
        const code = await this.rides.cancelByClient(userId, String(active.id), 'sms');
        return reply(Sms.cancelOk(code));
      }
      default:
        return reply(Sms.clientHelp(), 'help');
    }
  }

  private async createAndReport(userId: string, input: CreateRideInput, reply: (b: string, k?: string) => Promise<void>) {
    try {
      // Pas d'accusé de réception systématique (coût) : le client reçoit la confirmation du chauffeur
      // ou le message « aucun chauffeur ».
      const { ride, created } = await this.rides.createRide(userId, input);
      // Même SMS renvoyé (le client n'a pas encore de réponse) : on lui rappelle sa demande en cours.
      if (!created && ['SEARCHING', 'OFFERED'].includes(ride.status)) return reply(Sms.alreadyActive(ride.public_code));
    } catch (e) {
      if (e instanceof ConflictException) return reply(Sms.alreadyActive((e.getResponse() as any).rideCode ?? ''));
      if (e instanceof UnprocessableEntityException) return reply(Sms.outsideArea());
      if (e instanceof HttpException && e.getStatus() === 429) return reply(Sms.tooMany());
      throw e;
    }
  }

  private async handleDriver(driverId: string, p: ParsedSms, reply: (b: string, k?: string) => Promise<void>) {
    const d = await this.drivers.getDriver(driverId);
    if (d.verification_status !== 'approved') return reply(Sms.notApproved());
    switch (p.kind) {
      case 'available': {
        if (!p.placeText) return reply(Sms.driverHelp(), 'help');
        const res = await this.geo.resolve(p.placeText);
        if (res.status === 'none') return reply(Sms.notFound(p.placeText));
        const lm = res.status === 'unique' ? res.landmark : res.candidates[0];
        await this.drivers.setAvailableAtLandmark(driverId, lm.id);
        return reply(Sms.available(lm.name));
      }
      case 'stop':
        await this.drivers.setAvailability(driverId, false, 'sms');
        return reply(Sms.stopped());
      case 'accept': {
        const r = await this.rides.accept(driverId, { rideCode: p.code }, 'sms');
        if (!r.ok) return reply(r.reason === 'not_found' ? Sms.unknownCode(p.code) : Sms.alreadyTaken(p.code));
        return; // la confirmation détaillée est envoyée par RidesService.accept
      }
      case 'decline':
        await this.rides.decline(driverId, { rideCode: p.code });
        return;
      case 'finish': {
        const rideId = await this.rides.findRideByCodeForDriver(driverId, p.code);
        if (!rideId) return reply(Sms.unknownCode(p.code));
        try {
          await this.rides.complete(driverId, rideId, 'sms');
        } catch (e) {
          if (e instanceof ConflictException) return reply(Sms.alreadyTaken(p.code));
          throw e;
        }
        return reply(Sms.finished(p.code));
      }
      default:
        return reply(Sms.driverHelp(), 'help');
    }
  }

  /** Même texte renvoyé dans la même fenêtre de 5 minutes = même demande. */
  private smsIdempotencyKey(text: string) {
    const bucket = Math.floor(this.clock.now().getTime() / 300_000);
    return 'sms:' + createHash('sha256').update(normalizeText(text)).digest('hex').slice(0, 24) + ':' + bucket;
  }
}
