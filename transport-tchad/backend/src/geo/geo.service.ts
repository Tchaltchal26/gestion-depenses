import { Injectable } from '@nestjs/common';
import { normalizeText } from '../common/text';
import { DbService, Queryable } from '../db/db.service';

export interface LandmarkMatch {
  id: number;
  name: string;
  quartier: string | null;
  lat: number;
  lng: number;
  score: number;
}

export type Resolution =
  | { status: 'unique'; landmark: LandmarkMatch }
  | { status: 'ambiguous'; candidates: LandmarkMatch[] }
  | { status: 'none' };

@Injectable()
export class GeoService {
  constructor(private readonly db: DbService) {}

  /** Recherche floue (trigrammes) sur les noms et alias des repères, tolérante aux fautes. */
  async searchLandmarks(q: string, limit = 5): Promise<LandmarkMatch[]> {
    const norm = normalizeText(q);
    if (norm.length < 2) return [];
    const rows = await this.db.query(
      `SELECT l.id, l.name, l.quartier, ST_Y(l.location::geometry) AS lat, ST_X(l.location::geometry) AS lng,
              max(greatest(similarity(n.norm, $1), word_similarity($1, n.norm))) AS score
         FROM landmark_names n JOIN landmarks l ON l.id = n.landmark_id
        WHERE greatest(similarity(n.norm, $1), word_similarity($1, n.norm)) >= 0.3
        GROUP BY l.id
        ORDER BY score DESC, l.name
        LIMIT $2`,
      [norm, limit],
    );
    return rows.map((r) => ({ ...r, score: Number(r.score) }));
  }

  /** Décide si un texte désigne un repère sans ambiguïté. */
  async resolve(q: string): Promise<Resolution> {
    const found = await this.searchLandmarks(q, 5);
    if (!found.length) return { status: 'none' };
    const [top, second] = found;
    if (!second || (top.score >= 0.5 && second.score <= top.score - 0.15)) return { status: 'unique', landmark: top };
    return { status: 'ambiguous', candidates: found.slice(0, 3) };
  }

  async getLandmark(id: number): Promise<LandmarkMatch | null> {
    const [r] = await this.db.query(
      `SELECT id, name, quartier, ST_Y(location::geometry) AS lat, ST_X(location::geometry) AS lng, 1 AS score
         FROM landmarks WHERE id = $1`,
      [id],
    );
    return r ?? null;
  }

  async inServiceArea(lat: number, lng: number, c: Queryable = this.db.pool): Promise<boolean> {
    const { rows } = await c.query(
      `SELECT 1 FROM zones WHERE active AND ST_Covers(area, ST_SetSRID(ST_MakePoint($2, $1), 4326)::geography) LIMIT 1`,
      [lat, lng],
    );
    return rows.length > 0;
  }

  /**
   * Prix indicatif : base + prix/km sur la distance à vol d'oiseau majorée de 30 % (approximation
   * du trajet routier tant qu'OSRM n'est pas branché), arrondi aux 50 FCFA supérieurs.
   */
  async quote(type: string, from: { lat: number; lng: number }, to: { lat: number; lng: number }): Promise<number | null> {
    const [r] = await this.db.query(
      `SELECT t.base_xaf, t.per_km_xaf, t.min_xaf,
              ST_Distance(ST_SetSRID(ST_MakePoint($2, $1), 4326)::geography,
                          ST_SetSRID(ST_MakePoint($4, $3), 4326)::geography) AS meters
         FROM tariffs t WHERE t.transport_type = $5`,
      [from.lat, from.lng, to.lat, to.lng, type],
    );
    if (!r) return null;
    const km = (Number(r.meters) * 1.3) / 1000;
    const raw = Math.max(r.min_xaf, r.base_xaf + r.per_km_xaf * km);
    return Math.ceil(raw / 50) * 50;
  }
}
