import { ConflictException, ForbiddenException, HttpException, HttpStatus, Inject, Injectable, NotFoundException, UnprocessableEntityException } from '@nestjs/common';
import { PoolClient } from 'pg';
import { Clock } from '../common/clock';
import { CryptoService } from '../common/crypto.service';
import { AppConfig, CONFIG } from '../config';
import { DbService } from '../db/db.service';
import { GeoService } from '../geo/geo.service';
import { Notifier } from '../notifier';
import { SmsSender } from '../sms/sms.sender';
import { Sms } from '../sms/sms.templates';

export type TransportType = 'taxi' | 'rakcha' | 'clando';
export type Channel = 'app' | 'sms' | 'ops';

export const ACTIVE_STATUSES = ['SEARCHING', 'OFFERED', 'ASSIGNED', 'ARRIVED', 'IN_PROGRESS'];

export interface CreateRideInput {
  transportType: TransportType;
  pickupLat?: number;
  pickupLng?: number;
  pickupLandmarkId?: number;
  pickupText?: string;
  dropoffLandmarkId?: number;
  dropoffText?: string;
  channel: Channel;
  idempotencyKey?: string;
}

export type AcceptResult =
  | { ok: true; rideId: string }
  | { ok: false; reason: 'not_found' | 'already_taken' | 'driver_busy' };

const isUniqueViolation = (e: any, constraint?: string) =>
  e?.code === '23505' && (!constraint || e?.constraint === constraint);

@Injectable()
export class RidesService {
  constructor(
    private readonly db: DbService,
    private readonly geo: GeoService,
    private readonly crypto: CryptoService,
    private readonly sms: SmsSender,
    private readonly notifier: Notifier,
    private readonly clock: Clock,
    @Inject(CONFIG) private readonly config: AppConfig,
  ) {}

  // ---------------------------------------------------------------- création

  async createRide(clientId: string, input: CreateRideInput) {
    if (input.idempotencyKey) {
      const [existing] = await this.db.query(
        'SELECT id FROM rides WHERE client_id = $1 AND idempotency_key = $2',
        [clientId, input.idempotencyKey],
      );
      if (existing) return { ride: await this.getRideRow(existing.id), created: false };
    }

    let lat = input.pickupLat;
    let lng = input.pickupLng;
    let pickupText = input.pickupText ?? null;
    if (input.pickupLandmarkId) {
      const lm = await this.geo.getLandmark(input.pickupLandmarkId);
      if (!lm) throw new UnprocessableEntityException('Repère de départ inconnu');
      lat = lm.lat;
      lng = lm.lng;
      pickupText = pickupText ?? lm.name;
    }
    if (lat == null || lng == null) throw new UnprocessableEntityException('Lieu de départ requis');
    if (!(await this.geo.inServiceArea(lat, lng))) throw new UnprocessableEntityException('Lieu de départ hors zone desservie');

    let dropoffText = input.dropoffText ?? null;
    let price: number | null = null;
    if (input.dropoffLandmarkId) {
      const d = await this.geo.getLandmark(input.dropoffLandmarkId);
      if (!d) throw new UnprocessableEntityException('Repère de destination inconnu');
      dropoffText = dropoffText ?? d.name;
      price = await this.geo.quote(input.transportType, { lat, lng }, d);
    }

    const now = this.clock.now();
    const [{ n }] = await this.db.query(
      `SELECT count(*)::int AS n FROM rides WHERE client_id = $1 AND created_at > $2`,
      [clientId, new Date(now.getTime() - 3600_000)],
    );
    if (n >= this.config.maxRideRequestsPerHour) {
      throw new HttpException({ message: 'Trop de demandes', code: 'TOO_MANY' }, HttpStatus.TOO_MANY_REQUESTS);
    }

    const m = this.config.matching;
    const timeout = input.channel === 'sms' ? m.searchTimeoutSmsSec : m.searchTimeoutAppSec;
    let rideId: string | undefined;
    for (let attempt = 0; attempt < 5 && !rideId; attempt++) {
      try {
        rideId = await this.db.tx(async (c) => {
          const { rows } = await c.query(
            `INSERT INTO rides (public_code, client_id, transport_type, status, channel, pickup, pickup_landmark_id,
                                pickup_text, dropoff_landmark_id, dropoff_text, price_quoted_xaf, pickup_pin,
                                idempotency_key, search_deadline, created_at)
             VALUES ($1, $2, $3, 'SEARCHING', $4, ST_SetSRID(ST_MakePoint($6, $5), 4326)::geography, $7, $8, $9, $10, $11, $12, $13, $14, $15)
             RETURNING id`,
            [this.crypto.randomCode(5), clientId, input.transportType, input.channel, lat, lng,
             input.pickupLandmarkId ?? null, pickupText, input.dropoffLandmarkId ?? null, dropoffText, price,
             this.crypto.randomDigits(4), input.idempotencyKey ?? null, new Date(now.getTime() + timeout * 1000), now],
          );
          await this.logEvent(c, rows[0].id, null, 'SEARCHING', clientId, input.channel);
          return rows[0].id as string;
        });
      } catch (e) {
        if (isUniqueViolation(e, 'rides_public_code_key')) continue; // collision de code : on réessaie
        if (isUniqueViolation(e, 'rides_one_searching_per_client')) {
          const [active] = await this.db.query(
            `SELECT public_code FROM rides WHERE client_id = $1 AND status IN ('SEARCHING','OFFERED')`, [clientId]);
          throw new ConflictException({ message: 'Une demande est déjà en cours', code: 'ALREADY_ACTIVE', rideCode: active?.public_code });
        }
        if (isUniqueViolation(e, 'rides_client_id_idempotency_key_key')) {
          const [existing] = await this.db.query(
            'SELECT id FROM rides WHERE client_id = $1 AND idempotency_key = $2', [clientId, input.idempotencyKey]);
          return { ride: await this.getRideRow(existing.id), created: false };
        }
        throw e;
      }
    }
    if (!rideId) throw new Error('Impossible de générer un code de course unique');

    await this.dispatch(rideId);
    return { ride: await this.getRideRow(rideId), created: true };
  }

  // ---------------------------------------------------------------- attribution

  /**
   * Envoie la vague d'offres suivante pour une course en recherche, ou la clôt si le délai
   * global est dépassé. Sans effet si des offres sont encore ouvertes.
   */
  async dispatch(rideId: string): Promise<void> {
    const m = this.config.matching;
    const now = this.clock.now();
    const outbox: (() => Promise<void>)[] = [];

    await this.db.tx(async (c) => {
      const { rows: [ride] } = await c.query(
        `SELECT r.*, ST_Y(r.pickup::geometry) AS lat, ST_X(r.pickup::geometry) AS lng FROM rides r WHERE id = $1 FOR UPDATE`, [rideId]);
      if (!ride || !['SEARCHING', 'OFFERED'].includes(ride.status)) return;
      const { rows: open } = await c.query('SELECT 1 FROM ride_offers WHERE ride_id = $1 AND response IS NULL LIMIT 1', [rideId]);
      if (open.length) return;

      if (new Date(ride.search_deadline) <= now) {
        await this.setStatus(c, ride, 'NO_DRIVER', null, 'system');
        outbox.push(() => this.tellClient(ride, Sms.noDriver(ride.public_code), 'ride.no_driver', 'no_driver'));
        return;
      }
      if (ride.status === 'OFFERED') await this.setStatus(c, ride, 'SEARCHING', null, 'system');

      let radiusIdx = ride.radius_idx;
      let candidates: any[] = [];
      for (; radiusIdx < m.radiiMeters.length; radiusIdx++) {
        candidates = await this.findCandidates(c, ride, m.radiiMeters[radiusIdx], now);
        if (candidates.length) break;
      }
      radiusIdx = Math.min(radiusIdx, m.radiiMeters.length - 1);
      if (!candidates.length) {
        await c.query('UPDATE rides SET radius_idx = $2 WHERE id = $1', [rideId, radiusIdx]);
        return; // personne pour l'instant : on réessaiera au prochain passage du répartiteur
      }

      // Vague : jusqu'à N chauffeurs « app », au plus 1 chauffeur « SMS » (chaque offre SMS coûte).
      const wave: any[] = [];
      let smsInWave = 0;
      for (const cand of candidates) {
        if (wave.length >= m.waveSizeApp) break;
        if (cand.channel === 'sms') {
          if (smsInWave) continue;
          smsInWave++;
        }
        wave.push(cand);
      }

      const waveNo = ride.wave + 1;
      const offered: any[] = [];
      for (const d of wave) {
        const ttl = d.channel === 'sms' ? m.offerTtlSmsSec : m.offerTtlAppSec;
        const { rows } = await c.query(
          `INSERT INTO ride_offers (ride_id, driver_id, wave, channel, sent_at, expires_at)
           VALUES ($1, $2, $3, $4, $5, $6)
           ON CONFLICT DO NOTHING RETURNING id`,
          [rideId, d.driver_id, waveNo, d.channel, now, new Date(now.getTime() + ttl * 1000)],
        );
        if (rows.length) offered.push({ ...d, offerId: rows[0].id, ttl });
      }
      if (!offered.length) return;

      await c.query('UPDATE drivers SET offered_count = offered_count + 1 WHERE user_id = ANY($1::bigint[])', [offered.map((o) => o.driver_id)]);
      await c.query('UPDATE rides SET wave = $2, radius_idx = $3 WHERE id = $1', [rideId, waveNo, radiusIdx]);
      await this.setStatus(c, { ...ride, status: 'SEARCHING' }, 'OFFERED', null, 'system', { wave: waveNo, drivers: offered.map((o) => o.driver_id) });

      for (const o of offered) {
        const data = { offerId: o.offerId, rideCode: ride.public_code, transportType: ride.transport_type, pickup: ride.pickup_text,
          dropoff: ride.dropoff_text, priceXaf: ride.price_quoted_xaf, distanceM: Math.round(o.dist), expiresInSec: o.ttl };
        if (o.channel === 'sms') {
          outbox.push(() => this.sms.sendToUser(o.driver_id, Sms.offer(ride.public_code, ride.transport_type, ride.pickup_text ?? 'position GPS',
            ride.dropoff_text, ride.price_quoted_xaf, o.dist, o.ttl), 'offer', rideId));
        } else {
          outbox.push(async () => this.notifier.notify(o.driver_id, 'offer.new', data));
        }
      }
    });

    for (const send of outbox) await send();
  }

  /**
   * Chauffeurs éligibles, classés par score croissant :
   *   0,7 × distance normalisée par le rayon + 0,3 × (1 − taux d'acceptation lissé).
   * L'ETA routier (OSRM) remplacera la distance quand sa fiabilité aura été validée sur le terrain.
   */
  private async findCandidates(c: PoolClient, ride: any, radius: number, now: Date) {
    const m = this.config.matching;
    const { rows } = await c.query(
      `SELECT d.user_id AS driver_id, p.channel,
              ST_Distance(p.location, r.pickup) AS dist,
              0.7 * ST_Distance(p.location, r.pickup) / $3
                + 0.3 * (1 - (d.accepted_count + 1)::float / (d.offered_count + 2)) AS score
         FROM rides r
         JOIN drivers d ON d.transport_type = r.transport_type AND d.verification_status = 'approved'
         JOIN users u ON u.id = d.user_id AND u.status = 'active'
         JOIN driver_presence p ON p.driver_id = d.user_id AND p.available
        WHERE r.id = $1
          AND ((p.location_source = 'gps' AND p.location_at > $2::timestamptz - make_interval(secs => $4))
            OR (p.location_source = 'landmark' AND p.location_at > $2::timestamptz - make_interval(secs => $5)))
          AND ST_DWithin(p.location, r.pickup, $3)
          AND NOT EXISTS (SELECT 1 FROM rides a WHERE a.driver_id = d.user_id AND a.status IN ('ASSIGNED','ARRIVED','IN_PROGRESS'))
          AND NOT EXISTS (SELECT 1 FROM ride_offers o WHERE o.driver_id = d.user_id AND (o.ride_id = r.id OR o.response IS NULL))
        ORDER BY score ASC
        LIMIT 10`,
      [ride.id, now, radius, m.gpsFreshSec, m.landmarkFreshSec],
    );
    return rows.map((r) => ({ ...r, dist: Number(r.dist) }));
  }

  /**
   * Acceptation d'une offre. Ordre de verrouillage constant (course puis offres) pour éviter
   * les interblocages. Une seule acceptation peut réussir par course ; l'index unique
   * « une course active par chauffeur » protège en plus contre un chauffeur sur deux courses.
   */
  async accept(driverId: string, ref: { offerId?: string; rideCode?: string }, channel: 'app' | 'sms'): Promise<AcceptResult> {
    const now = this.clock.now();
    const offer = await this.findOffer(driverId, ref);
    if (!offer) return { ok: false, reason: 'not_found' };

    let revoked: any[] = [];
    let ride: any;
    try {
      const ok = await this.db.tx(async (c) => {
        ride = (await c.query('SELECT * FROM rides WHERE id = $1 FOR UPDATE', [offer.ride_id])).rows[0];
        const o = (await c.query('SELECT * FROM ride_offers WHERE id = $1 FOR UPDATE', [offer.id])).rows[0];
        if (o.response !== null || ride.status !== 'OFFERED') return false;
        if (new Date(o.expires_at) <= now) {
          await c.query(`UPDATE ride_offers SET response = 'expired', responded_at = $2 WHERE id = $1`, [o.id, now]);
          return false;
        }
        await c.query(`UPDATE rides SET driver_id = $2, assigned_at = $3 WHERE id = $1`, [ride.id, driverId, now]);
        await this.setStatus(c, ride, 'ASSIGNED', driverId, channel);
        await c.query(`UPDATE ride_offers SET response = 'accepted', responded_at = $2 WHERE id = $1`, [o.id, now]);
        revoked = (await c.query(
          `UPDATE ride_offers SET response = 'revoked', responded_at = $2 WHERE ride_id = $1 AND response IS NULL RETURNING driver_id, channel`,
          [ride.id, now])).rows;
        await c.query('UPDATE drivers SET accepted_count = accepted_count + 1 WHERE user_id = $1', [driverId]);
        return true;
      });
      if (!ok) return { ok: false, reason: 'already_taken' };
    } catch (e) {
      if (isUniqueViolation(e, 'rides_one_active_per_driver')) return { ok: false, reason: 'driver_busy' };
      throw e;
    }

    // Notifications après validation de la transaction.
    for (const r of revoked) if (r.channel === 'app') this.notifier.notify(r.driver_id, 'offer.revoked', { rideCode: ride.public_code });
    const full = await this.getRideRow(ride.id);
    const driver = await this.driverPublicInfo(driverId);
    const client = await this.userContact(ride.client_id);
    await this.tellClient(full, Sms.clientConfirmed(full.public_code, driver.name, driver.vehicle_plate, driver.vehicle_color, driver.phone, full.pickup_pin),
      'ride.assigned', 'assigned', { driver });
    if (channel === 'sms') {
      await this.sms.sendToUser(driverId, Sms.driverConfirmed(full.public_code, client.name, client.phone), 'assigned', ride.id);
    }
    return { ok: true, rideId: String(ride.id) };
  }

  async decline(driverId: string, ref: { offerId?: string; rideCode?: string }): Promise<boolean> {
    const offer = await this.findOffer(driverId, ref);
    if (!offer) return false;
    const now = this.clock.now();
    const done = await this.db.tx(async (c) => {
      await c.query('SELECT 1 FROM rides WHERE id = $1 FOR UPDATE', [offer.ride_id]);
      const { rowCount } = await c.query(
        `UPDATE ride_offers SET response = 'declined', responded_at = $2 WHERE id = $1 AND response IS NULL`, [offer.id, now]);
      return rowCount === 1;
    });
    if (done) await this.dispatch(String(offer.ride_id));
    return done;
  }

  /** Passage périodique : expire les offres, lance les vagues suivantes, clôt les recherches trop longues. */
  async tick(): Promise<void> {
    const now = this.clock.now();
    await this.db.query(`UPDATE ride_offers SET response = 'expired', responded_at = $1 WHERE response IS NULL AND expires_at <= $1`, [now]);
    await this.db.query('DELETE FROM sms_pending WHERE expires_at <= $1', [now]);
    const rides = await this.db.query(`SELECT id FROM rides WHERE status IN ('SEARCHING','OFFERED') ORDER BY created_at`);
    for (const r of rides) await this.dispatch(String(r.id));
  }

  // ---------------------------------------------------------------- déroulement

  async arrived(driverId: string, rideId: string) {
    await this.driverTransition(driverId, rideId, ['ASSIGNED'], 'ARRIVED', 'app');
    const ride = await this.getRideRow(rideId);
    if (ride.channel !== 'sms') this.notifier.notify(ride.client_id, 'ride.arrived', { rideCode: ride.public_code });
  }

  async start(driverId: string, rideId: string, pin: string) {
    const ride = await this.getRideRow(rideId);
    if (String(ride.driver_id) !== driverId) throw new ForbiddenException();
    if (ride.pickup_pin !== pin) throw new UnprocessableEntityException('Code client incorrect');
    await this.driverTransition(driverId, rideId, ['ASSIGNED', 'ARRIVED'], 'IN_PROGRESS', 'app');
  }

  async complete(driverId: string, rideId: string, channel: 'app' | 'sms') {
    // Par SMS, le chauffeur ne peut pas saisir le code client : FIN est accepté dès l'attribution.
    const from = channel === 'sms' ? ['ASSIGNED', 'ARRIVED', 'IN_PROGRESS'] : ['IN_PROGRESS'];
    await this.driverTransition(driverId, rideId, from, 'COMPLETED', channel);
    await this.db.query('UPDATE rides SET completed_at = $2 WHERE id = $1', [rideId, this.clock.now()]);
  }

  async cancelByClient(clientId: string, rideId: string, channel: Channel) {
    const now = this.clock.now();
    let ride: any;
    let revoked: any[] = [];
    await this.db.tx(async (c) => {
      ride = (await c.query('SELECT * FROM rides WHERE id = $1 FOR UPDATE', [rideId])).rows[0];
      if (!ride) throw new NotFoundException('Course introuvable');
      if (String(ride.client_id) !== clientId) throw new ForbiddenException();
      if (!['SEARCHING', 'OFFERED', 'ASSIGNED', 'ARRIVED'].includes(ride.status)) {
        throw new ConflictException('Cette course ne peut plus être annulée');
      }
      await this.setStatus(c, ride, 'CANCELLED_BY_CLIENT', clientId, channel);
      await c.query('UPDATE rides SET cancelled_at = $2 WHERE id = $1', [rideId, now]);
      revoked = (await c.query(
        `UPDATE ride_offers SET response = 'revoked', responded_at = $2 WHERE ride_id = $1 AND response IS NULL RETURNING driver_id, channel`,
        [rideId, now])).rows;
    });
    for (const r of revoked) if (r.channel === 'app') this.notifier.notify(r.driver_id, 'offer.revoked', { rideCode: ride.public_code });
    if (ride.driver_id) {
      const [p] = await this.db.query('SELECT channel FROM driver_presence WHERE driver_id = $1', [ride.driver_id]);
      if (p?.channel === 'sms') await this.sms.sendToUser(ride.driver_id, Sms.cancelledToDriver(ride.public_code), 'cancelled', rideId);
      else this.notifier.notify(ride.driver_id, 'ride.cancelled', { rideCode: ride.public_code });
    }
    return ride.public_code as string;
  }

  async cancelByDriver(driverId: string, rideId: string) {
    await this.driverTransition(driverId, rideId, ['ASSIGNED', 'ARRIVED'], 'CANCELLED_BY_DRIVER', 'app');
    await this.db.query('UPDATE rides SET cancelled_at = $2 WHERE id = $1', [rideId, this.clock.now()]);
    const ride = await this.getRideRow(rideId);
    await this.tellClient(ride, Sms.cancelledToClient(ride.public_code), 'ride.cancelled', 'cancelled');
  }

  // ---------------------------------------------------------------- lecture

  async getRideForUser(user: { id: string; role: string }, rideId: string) {
    const r = await this.getRideRow(rideId);
    const isClient = String(r.client_id) === user.id;
    const isDriver = r.driver_id && String(r.driver_id) === user.id;
    if (!isClient && !isDriver && !['ops', 'admin'].includes(user.role)) throw new NotFoundException('Course introuvable');
    const view: any = {
      id: String(r.id), code: r.public_code, status: r.status, transportType: r.transport_type, channel: r.channel,
      pickup: { text: r.pickup_text, lat: r.lat, lng: r.lng }, dropoff: { text: r.dropoff_text }, priceXaf: r.price_quoted_xaf,
      createdAt: r.created_at, assignedAt: r.assigned_at,
    };
    const contactVisible = ['ASSIGNED', 'ARRIVED', 'IN_PROGRESS'].includes(r.status);
    if (isClient) {
      view.pickupPin = r.pickup_pin;
      if (r.driver_id) {
        const d = await this.driverPublicInfo(r.driver_id);
        view.driver = { name: d.name, plate: d.vehicle_plate, color: d.vehicle_color, phone: contactVisible ? d.phone : undefined };
      }
    }
    if (isDriver) {
      const c = await this.userContact(r.client_id);
      view.client = { name: c.name, phone: contactVisible ? c.phone : undefined };
    }
    return view;
  }

  async openOffersForDriver(driverId: string) {
    const rows = await this.db.query(
      `SELECT o.id, o.expires_at, r.public_code, r.transport_type, r.pickup_text, r.dropoff_text, r.price_quoted_xaf
         FROM ride_offers o JOIN rides r ON r.id = o.ride_id
        WHERE o.driver_id = $1 AND o.response IS NULL AND o.expires_at > $2`,
      [driverId, this.clock.now()],
    );
    return rows.map((o) => ({ offerId: String(o.id), rideCode: o.public_code, transportType: o.transport_type, pickup: o.pickup_text,
      dropoff: o.dropoff_text, priceXaf: o.price_quoted_xaf, expiresAt: o.expires_at }));
  }

  async findActiveRideForClient(clientId: string) {
    const [r] = await this.db.query(
      `SELECT id, public_code FROM rides WHERE client_id = $1 AND status = ANY($2) ORDER BY created_at DESC LIMIT 1`,
      [clientId, ['SEARCHING', 'OFFERED', 'ASSIGNED', 'ARRIVED']]);
    return r ?? null;
  }

  async findRideByCodeForDriver(driverId: string, code: string) {
    const [r] = await this.db.query('SELECT id FROM rides WHERE public_code = $1 AND driver_id = $2', [code, driverId]);
    return r ? String(r.id) : null;
  }

  async getRideRow(rideId: string | number) {
    const [r] = await this.db.query(
      `SELECT r.*, ST_Y(r.pickup::geometry) AS lat, ST_X(r.pickup::geometry) AS lng FROM rides r WHERE id = $1`, [rideId]);
    if (!r) throw new NotFoundException('Course introuvable');
    return r;
  }

  // ---------------------------------------------------------------- utilitaires internes

  private async findOffer(driverId: string, ref: { offerId?: string; rideCode?: string }) {
    const [o] = ref.offerId
      ? await this.db.query('SELECT id, ride_id FROM ride_offers WHERE id = $1 AND driver_id = $2', [ref.offerId, driverId])
      : await this.db.query(
          `SELECT o.id, o.ride_id FROM ride_offers o JOIN rides r ON r.id = o.ride_id WHERE r.public_code = $1 AND o.driver_id = $2`,
          [ref.rideCode, driverId]);
    return o ?? null;
  }

  private async driverTransition(driverId: string, rideId: string, from: string[], to: string, channel: string) {
    await this.db.tx(async (c) => {
      const ride = (await c.query('SELECT * FROM rides WHERE id = $1 FOR UPDATE', [rideId])).rows[0];
      if (!ride || String(ride.driver_id) !== driverId) throw new NotFoundException('Course introuvable');
      if (!from.includes(ride.status)) throw new ConflictException(`Transition impossible depuis ${ride.status}`);
      await this.setStatus(c, ride, to, driverId, channel);
    });
  }

  private async setStatus(c: PoolClient, ride: any, to: string, actorId: string | null, channel: string, meta?: unknown) {
    await c.query('UPDATE rides SET status = $2 WHERE id = $1', [ride.id, to]);
    await this.logEvent(c, ride.id, ride.status, to, actorId, channel, meta);
    ride.status = to;
  }

  private async logEvent(c: PoolClient, rideId: string, from: string | null, to: string, actorId: string | null, channel: string, meta?: unknown) {
    await c.query(
      `INSERT INTO ride_events (ride_id, from_status, to_status, actor_id, channel, at, meta) VALUES ($1, $2, $3, $4, $5, $6, $7)`,
      [rideId, from, to, actorId, channel, this.clock.now(), meta ? JSON.stringify(meta) : null]);
  }

  /** Client SMS → SMS ; client application → notification (simulée). */
  private async tellClient(ride: any, smsText: string, event: string, kind: string, data: Record<string, unknown> = {}) {
    if (ride.channel === 'sms') await this.sms.sendToUser(ride.client_id, smsText, kind, ride.id);
    else this.notifier.notify(ride.client_id, event, { rideCode: ride.public_code, ...data });
  }

  private async driverPublicInfo(driverId: string | number) {
    const [d] = await this.db.query(
      `SELECT u.name, u.phone_enc, d.vehicle_plate, d.vehicle_color FROM drivers d JOIN users u ON u.id = d.user_id WHERE d.user_id = $1`,
      [driverId]);
    return { name: d.name, vehicle_plate: d.vehicle_plate, vehicle_color: d.vehicle_color, phone: this.crypto.decrypt(d.phone_enc) };
  }

  private async userContact(userId: string | number) {
    const [u] = await this.db.query('SELECT name, phone_enc FROM users WHERE id = $1', [userId]);
    return { name: u.name as string | null, phone: this.crypto.decrypt(u.phone_enc) };
  }
}
