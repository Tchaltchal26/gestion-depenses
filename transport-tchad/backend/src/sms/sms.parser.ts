import { normalizeText } from '../common/text';

export type TransportType = 'taxi' | 'rakcha' | 'clando';

export type ParsedSms =
  | { kind: 'compact'; type: TransportType; lat: number; lng: number; dropoffLandmarkId: number | null }
  | { kind: 'request'; type: TransportType; pickupText: string; dropoffText: string | null }
  | { kind: 'accept' | 'decline' | 'finish'; code: string }
  | { kind: 'cancel'; code: string | null }
  | { kind: 'available'; placeText: string | null }
  | { kind: 'stop' }
  | { kind: 'help' }
  | { kind: 'choice'; n: number }
  | { kind: 'unknown' };

const TYPE_WORDS: Record<string, TransportType> = {
  taxi: 'taxi', t: 'taxi',
  rakcha: 'rakcha', raksha: 'rakcha', rikcha: 'rakcha', rickshaw: 'rakcha', r: 'rakcha',
  clando: 'clando', klando: 'clando', moto: 'clando', c: 'clando',
};
const COMPACT_TYPES: Record<string, TransportType> = { t: 'taxi', r: 'rakcha', c: 'clando' };
const CODE = '([a-z0-9]{4,6})';

/**
 * Interprète un SMS entrant. Formats acceptés (insensibles à la casse et aux accents) :
 *   CLANDO grand marche > hopital   |  TAXI gare vers universite  |  RAKCHA mosquee
 *   M1 C 13.8291 20.8324 D 412       (format compact généré par l'application)
 *   OUI K7Q4P | NON K7Q4P | FIN K7Q4P | ANNULER [K7Q4P] | DISPO grand marche | STOP | AIDE | 1..9
 */
export function parseSms(raw: string): ParsedSms {
  const t = normalizeText(raw);
  if (t === '' || /^(aide|help|info)$/.test(t)) return { kind: 'help' };

  let m = t.match(/^m1 ([trc]) (-?\d{1,2}\.\d+) (-?\d{1,3}\.\d+)(?: d (\d+))?$/);
  if (m) {
    const id = m[4] ? Number(m[4]) : 0;
    return { kind: 'compact', type: COMPACT_TYPES[m[1]], lat: Number(m[2]), lng: Number(m[3]), dropoffLandmarkId: id > 0 ? id : null };
  }
  if ((m = t.match(new RegExp(`^(oui|ok|yes) ${CODE}$`)))) return { kind: 'accept', code: m[2].toUpperCase() };
  if ((m = t.match(new RegExp(`^(non|no) ${CODE}$`)))) return { kind: 'decline', code: m[2].toUpperCase() };
  if ((m = t.match(new RegExp(`^fin ${CODE}$`)))) return { kind: 'finish', code: m[1].toUpperCase() };
  if ((m = t.match(new RegExp(`^annul(?:er|e)?(?: ${CODE})?$`)))) return { kind: 'cancel', code: m[1] ? m[1].toUpperCase() : null };
  if ((m = t.match(/^dispo(?:nible)?(?: (.+))?$/))) return { kind: 'available', placeText: m[1]?.trim() || null };
  if (/^(stop|pause)$/.test(t)) return { kind: 'stop' };
  if ((m = t.match(/^([1-9])$/))) return { kind: 'choice', n: Number(m[1]) };

  const [first, ...rest] = t.split(' ');
  const type = TYPE_WORDS[first];
  if (type && rest.length) {
    const [pickup, dropoff] = rest.join(' ').split(/\s*>\s*|\s+vers\s+/);
    if (pickup && pickup.trim()) {
      return { kind: 'request', type, pickupText: pickup.trim(), dropoffText: dropoff?.trim() || null };
    }
  }
  return { kind: 'unknown' };
}
