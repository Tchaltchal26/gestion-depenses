// Modèles de SMS : alphabet latin sans accents, <= 160 caractères (1 SMS facturé).
const cut = (s: string | null | undefined, n: number) => {
  const v = (s ?? '').trim();
  return v.length > n ? v.slice(0, n - 1) + '.' : v;
};
const TYPE_LABEL: Record<string, string> = { taxi: 'TAXI', rakcha: 'RAKCHA', clando: 'CLANDO' };

export const Sms = {
  clientHelp: () =>
    'Pour un transport envoyez: CLANDO lieu > destination (ou TAXI, RAKCHA). Ex: CLANDO grand marche > hopital. ANNULER pour annuler.',
  driverHelp: () => 'DISPO lieu = disponible. STOP = pause. OUI code = accepter. NON code = refuser. FIN code = course terminee.',
  choices: (names: string[]) =>
    `Plusieurs lieux: ${names.map((n, i) => `${i + 1}=${cut(n, 25)}`).join(' ')}. Repondez ${names.map((_, i) => i + 1).join(' ou ')}`,
  notFound: (text: string) => `Lieu "${cut(text, 30)}" non trouve. Indiquez un repere connu (marche, mosquee, ecole...).`,
  outsideArea: () => 'Desole, ce lieu est hors de la zone desservie.',
  tooLate: () => 'Votre demande est arrivee en retard (reseau). Renvoyez-la si vous avez encore besoin.',
  alreadyActive: (code: string) => `Vous avez deja une demande en cours (${code}). Envoyez ANNULER pour l'annuler.`,
  tooMany: () => 'Trop de demandes en peu de temps. Reessayez plus tard.',
  offer: (code: string, type: string, pickup: string, dropoff: string | null, price: number | null, distM: number, ttl: number) =>
    `Course ${code}: ${TYPE_LABEL[type]} ${cut(pickup, 30)}${dropoff ? ' > ' + cut(dropoff, 30) : ''}${price ? `, ~${price}F` : ''}. Client a ${Math.round(distM / 100) * 100}m. Repondez OUI ${code} (${ttl}s)`,
  driverConfirmed: (code: string, clientName: string | null, clientPhone: string) =>
    `OK ${code}. Client: ${cut(clientName || 'Client', 20)} ${clientPhone}. Appelez-le. A la fin envoyez FIN ${code}`,
  clientConfirmed: (code: string, driverName: string | null, plate: string | null, color: string | null, phone: string, pin: string) =>
    `${code}: chauffeur ${cut(driverName || '', 20)}, ${cut(plate || '', 12)} ${cut(color || '', 10)}. Tel ${phone}. Il vous appelle. Code a donner: ${pin}`,
  alreadyTaken: (code: string) => `${code} deja attribuee ou expiree.`,
  noDriver: (code: string) => `Aucun chauffeur disponible pour ${code}. Reessayez dans quelques minutes.`,
  cancelledToDriver: (code: string) => `Course ${code} annulee par le client.`,
  cancelledToClient: (code: string) => `Course ${code} annulee par le chauffeur. Renvoyez une demande si besoin.`,
  cancelOk: (code: string) => `Demande ${code} annulee.`,
  nothingToCancel: () => 'Aucune demande en cours.',
  available: (place: string) => `Vous etes DISPONIBLE pres de ${cut(place, 40)} (30 min). STOP pour arreter.`,
  stopped: () => 'Vous etes en PAUSE. DISPO lieu pour reprendre.',
  notApproved: () => 'Compte chauffeur non valide. Contactez la plateforme.',
  finished: (code: string) => `Course ${code} terminee. Merci.`,
  unknownCode: (code: string) => `Code ${code} inconnu.`,
  otp: (code: string) => `Votre code de connexion: ${code}. Ne le donnez a personne.`,
};
