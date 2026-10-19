import { Inject, Injectable } from '@nestjs/common';
import { createCipheriv, createDecipheriv, createHmac, randomBytes, randomInt, timingSafeEqual } from 'crypto';
import { AppConfig, CONFIG } from '../config';

// Alphabet sans caractères ambigus (pas de O/0, I/1, L) pour les codes lus et tapés par SMS.
const CODE_ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

@Injectable()
export class CryptoService {
  constructor(@Inject(CONFIG) private readonly config: AppConfig) {}

  /** Empreinte déterministe du numéro : permet la recherche sans stocker le numéro en clair. */
  hashPhone(e164: string): string {
    return createHmac('sha256', this.config.phoneHashKey).update(e164).digest('hex');
  }

  encrypt(plain: string): string {
    const iv = randomBytes(12);
    const cipher = createCipheriv('aes-256-gcm', this.config.phoneEncKey, iv);
    const enc = Buffer.concat([cipher.update(plain, 'utf8'), cipher.final()]);
    return [iv, cipher.getAuthTag(), enc].map((b) => b.toString('base64')).join('.');
  }

  decrypt(payload: string): string {
    const [iv, tag, enc] = payload.split('.').map((s) => Buffer.from(s, 'base64'));
    const decipher = createDecipheriv('aes-256-gcm', this.config.phoneEncKey, iv);
    decipher.setAuthTag(tag);
    return Buffer.concat([decipher.update(enc), decipher.final()]).toString('utf8');
  }

  hmac(value: string, secret: string): string {
    return createHmac('sha256', secret).update(value).digest('hex');
  }

  safeEqualHex(a: string, b: string): boolean {
    const ba = Buffer.from(a, 'hex');
    const bb = Buffer.from(b, 'hex');
    return ba.length === bb.length && ba.length > 0 && timingSafeEqual(ba, bb);
  }

  randomCode(length: number): string {
    let s = '';
    for (let i = 0; i < length; i++) s += CODE_ALPHABET[randomInt(CODE_ALPHABET.length)];
    return s;
  }

  randomDigits(length: number): string {
    let s = '';
    for (let i = 0; i < length; i++) s += String(randomInt(10));
    return s;
  }
}
