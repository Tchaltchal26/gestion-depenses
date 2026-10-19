import { ConflictException, ForbiddenException, Injectable, NotFoundException, UnprocessableEntityException } from '@nestjs/common';
import { Clock } from '../common/clock';
import { DbService } from '../db/db.service';
import { GeoService } from '../geo/geo.service';
import { TransportType } from './rides.service';

@Injectable()
export class DriversService {
  constructor(private readonly db: DbService, private readonly geo: GeoService, private readonly clock: Clock) {}

  async register(userId: string, input: { name: string; transportType: TransportType; plate?: string; color?: string }) {
    await this.db.tx(async (c) => {
      const { rows: [u] } = await c.query('SELECT role FROM users WHERE id = $1 FOR UPDATE', [userId]);
      if (u.role !== 'client' && u.role !== 'driver') throw new ForbiddenException();
      const { rowCount } = await c.query(
        `INSERT INTO drivers (user_id, transport_type, vehicle_plate, vehicle_color) VALUES ($1, $2, $3, $4) ON CONFLICT DO NOTHING`,
        [userId, input.transportType, input.plate ?? null, input.color ?? null]);
      if (!rowCount) throw new ConflictException('Déjà inscrit comme chauffeur');
      await c.query(`UPDATE users SET role = 'driver', name = $2 WHERE id = $1`, [userId, input.name]);
      await c.query('INSERT INTO driver_presence (driver_id) VALUES ($1)', [userId]);
    });
    return { status: 'pending' };
  }

  async getDriver(driverId: string) {
    const [d] = await this.db.query('SELECT * FROM drivers WHERE user_id = $1', [driverId]);
    if (!d) throw new NotFoundException('Chauffeur introuvable');
    return d;
  }

  private async approvedOrThrow(driverId: string) {
    const d = await this.getDriver(driverId);
    if (d.verification_status !== 'approved') throw new ForbiddenException('Compte chauffeur non encore validé');
  }

  async setAvailability(driverId: string, available: boolean, channel: 'app' | 'sms' = 'app') {
    if (available) await this.approvedOrThrow(driverId);
    await this.db.query('UPDATE driver_presence SET available = $2, channel = $3 WHERE driver_id = $1', [driverId, available, channel]);
  }

  async updateLocation(driverId: string, lat: number, lng: number, batteryPct?: number) {
    await this.db.query(
      `UPDATE driver_presence SET location = ST_SetSRID(ST_MakePoint($3, $2), 4326)::geography, location_source = 'gps',
              location_at = $4, battery_pct = $5::smallint,
              -- Batterie critique : retiré des candidats.
              available = CASE WHEN $5::int IS NOT NULL AND $5::int < 5 THEN false ELSE available END
        WHERE driver_id = $1`,
      [driverId, lat, lng, this.clock.now(), batteryPct ?? null]);
  }

  /** Disponibilité déclarée par SMS : la position est celle du repère, datée (elle « vieillit »). */
  async setAvailableAtLandmark(driverId: string, landmarkId: number) {
    await this.approvedOrThrow(driverId);
    const lm = await this.geo.getLandmark(landmarkId);
    if (!lm) throw new UnprocessableEntityException('Repère inconnu');
    await this.db.query(
      `UPDATE driver_presence SET available = true, channel = 'sms', location_source = 'landmark', location_at = $4,
              location = ST_SetSRID(ST_MakePoint($3, $2), 4326)::geography WHERE driver_id = $1`,
      [driverId, lm.lat, lm.lng, this.clock.now()]);
  }
}
