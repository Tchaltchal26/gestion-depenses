/** Minuscules, sans accents, espaces normalisés. */
export function normalizeText(s: string): string {
  return (s ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9>\s.\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Convertit en alphabet GSM-7 « sûr » (ASCII imprimable) pour rester à 160 caractères par SMS. */
export function toGsmSafe(s: string): string {
  return (s ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^\x20-\x7E\n]/g, '?');
}
