/**
 * Normalise un numéro tchadien en E.164 (+235XXXXXXXX).
 * Accepte : 66123456, 66 12 34 56, +23566123456, 0023566123456, 23566123456.
 * Retourne null si le numéro n'est pas valide.
 */
export function normalizeChadPhone(input: string): string | null {
  let d = (input ?? '').replace(/[\s.\-()]/g, '');
  if (d.startsWith('+')) d = d.slice(1);
  else if (d.startsWith('00')) d = d.slice(2);
  if (d.length === 8) d = '235' + d;
  if (!/^235\d{8}$/.test(d)) return null;
  return '+' + d;
}
