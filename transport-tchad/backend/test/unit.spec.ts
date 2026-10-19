import { normalizeChadPhone } from '../src/common/phone';
import { parseSms } from '../src/sms/sms.parser';
import { Sms } from '../src/sms/sms.templates';

describe('normalizeChadPhone', () => {
  it.each([
    ['66123456', '+23566123456'],
    ['66 12 34 56', '+23566123456'],
    ['+235 66-12-34-56', '+23566123456'],
    ['0023566123456', '+23566123456'],
    ['23566123456', '+23566123456'],
  ])('%s → %s', (input, out) => expect(normalizeChadPhone(input)).toBe(out));

  it.each(['', '1234', '+33612345678', '6612345', 'abc'])('rejette %p', (input) => expect(normalizeChadPhone(input)).toBeNull());
});

describe('parseSms', () => {
  it('demande en texte libre avec destination', () => {
    expect(parseSms('CLANDO Grand Marché > Hôpital')).toEqual({ kind: 'request', type: 'clando', pickupText: 'grand marche', dropoffText: 'hopital' });
  });
  it('accepte « vers » et les abréviations', () => {
    expect(parseSms('t gare vers universite')).toEqual({ kind: 'request', type: 'taxi', pickupText: 'gare', dropoffText: 'universite' });
    expect(parseSms('Raksha mosquee')).toEqual({ kind: 'request', type: 'rakcha', pickupText: 'mosquee', dropoffText: null });
    expect(parseSms('moto marche')).toMatchObject({ type: 'clando' });
  });
  it('format compact de l\'application', () => {
    expect(parseSms('M1 C 13.8291 20.8324 D 4')).toEqual({ kind: 'compact', type: 'clando', lat: 13.8291, lng: 20.8324, dropoffLandmarkId: 4 });
    expect(parseSms('m1 t 13.8291 20.8324')).toEqual({ kind: 'compact', type: 'taxi', lat: 13.8291, lng: 20.8324, dropoffLandmarkId: null });
  });
  it('commandes chauffeur et client', () => {
    expect(parseSms('oui k7q4p')).toEqual({ kind: 'accept', code: 'K7Q4P' });
    expect(parseSms('NON K7Q4P')).toEqual({ kind: 'decline', code: 'K7Q4P' });
    expect(parseSms('Fin K7Q4P')).toEqual({ kind: 'finish', code: 'K7Q4P' });
    expect(parseSms('annuler')).toEqual({ kind: 'cancel', code: null });
    expect(parseSms('Dispo grand marche')).toEqual({ kind: 'available', placeText: 'grand marche' });
    expect(parseSms('stop')).toEqual({ kind: 'stop' });
    expect(parseSms('2')).toEqual({ kind: 'choice', n: 2 });
    expect(parseSms('aide')).toEqual({ kind: 'help' });
    expect(parseSms('?')).toEqual({ kind: 'help' });
  });
  it('texte non reconnu', () => {
    expect(parseSms('bonjour je veux aller au marche')).toEqual({ kind: 'unknown' });
    expect(parseSms('taxi')).toEqual({ kind: 'unknown' });
  });
});

describe('modèles de SMS', () => {
  it('restent sous 160 caractères même avec des valeurs longues', () => {
    const long = 'x'.repeat(200);
    const msgs = [
      Sms.clientHelp(), Sms.driverHelp(), Sms.choices([long, long, long]), Sms.notFound(long),
      Sms.offer('K7Q4P', 'clando', long, long, 99999, 4999, 90),
      Sms.driverConfirmed('K7Q4P', long, '+23566123456'),
      Sms.clientConfirmed('K7Q4P', long, long, long, '+23566123456', '1234'),
      Sms.available(long), Sms.alreadyActive('K7Q4P'), Sms.tooLate(),
    ];
    for (const m of msgs) expect(m.length).toBeLessThanOrEqual(160);
  });
});
