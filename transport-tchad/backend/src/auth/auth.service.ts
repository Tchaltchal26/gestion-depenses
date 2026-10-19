import { HttpException, HttpStatus, Inject, Injectable, UnauthorizedException, BadRequestException } from '@nestjs/common';
import { JwtService } from '@nestjs/jwt';
import { Clock } from '../common/clock';
import { CryptoService } from '../common/crypto.service';
import { normalizeChadPhone } from '../common/phone';
import { AppConfig, CONFIG } from '../config';
import { DbService } from '../db/db.service';
import { SmsSender } from '../sms/sms.sender';
import { Sms } from '../sms/sms.templates';

const OTP_TTL_SEC = 300;
const OTP_MAX_ATTEMPTS = 5;
const OTP_WINDOW_SEC = 900;
const OTP_MAX_SENDS_PER_WINDOW = 3;

@Injectable()
export class AuthService {
  constructor(
    private readonly db: DbService,
    private readonly crypto: CryptoService,
    private readonly sms: SmsSender,
    private readonly jwt: JwtService,
    private readonly clock: Clock,
    @Inject(CONFIG) private readonly config: AppConfig,
  ) {}

  phoneOrThrow(raw: string): string {
    const e164 = normalizeChadPhone(raw);
    if (!e164) throw new BadRequestException('Numéro de téléphone tchadien invalide');
    return e164;
  }

  private codeHash(phoneHash: string, code: string) {
    return this.crypto.hmac(`${phoneHash}:${code}`, this.config.jwtSecret);
  }

  async requestOtp(rawPhone: string): Promise<void> {
    const e164 = this.phoneOrThrow(rawPhone);
    const ph = this.crypto.hashPhone(e164);
    const now = this.clock.now();
    const code = this.crypto.randomDigits(6);
    const expires = new Date(now.getTime() + OTP_TTL_SEC * 1000);
    const windowStart = new Date(now.getTime() - OTP_WINDOW_SEC * 1000);

    // Insertion ou mise à jour atomique avec limitation du nombre d'envois par fenêtre.
    const rows = await this.db.query(
      `INSERT INTO otp_codes (phone_hash, code_hash, expires_at, attempts, window_start, sent_in_window)
       VALUES ($1, $2, $3, 0, $4, 1)
       ON CONFLICT (phone_hash) DO UPDATE SET
         code_hash = EXCLUDED.code_hash, expires_at = EXCLUDED.expires_at, attempts = 0,
         window_start = CASE WHEN otp_codes.window_start < $5 THEN EXCLUDED.window_start ELSE otp_codes.window_start END,
         sent_in_window = CASE WHEN otp_codes.window_start < $5 THEN 1 ELSE otp_codes.sent_in_window + 1 END
       WHERE otp_codes.window_start < $5 OR otp_codes.sent_in_window < $6
       RETURNING phone_hash`,
      [ph, this.codeHash(ph, code), expires, now, windowStart, OTP_MAX_SENDS_PER_WINDOW],
    );
    if (!rows.length) throw new HttpException('Trop de codes demandés. Réessayez dans 15 minutes.', HttpStatus.TOO_MANY_REQUESTS);
    await this.sms.sendToPhone(e164, Sms.otp(code), 'otp', undefined, Sms.otp('******'));
  }

  async verifyOtp(rawPhone: string, code: string): Promise<{ token: string; userId: string; role: string }> {
    const e164 = this.phoneOrThrow(rawPhone);
    const ph = this.crypto.hashPhone(e164);
    const now = this.clock.now();
    const [row] = await this.db.query(
      `UPDATE otp_codes SET attempts = attempts + 1 WHERE phone_hash = $1 RETURNING code_hash, expires_at, attempts`,
      [ph],
    );
    if (!row || new Date(row.expires_at) <= now) throw new UnauthorizedException('Code expiré ou inexistant');
    if (row.attempts > OTP_MAX_ATTEMPTS) throw new HttpException('Trop d\'essais. Demandez un nouveau code.', HttpStatus.TOO_MANY_REQUESTS);
    if (!this.crypto.safeEqualHex(row.code_hash, this.codeHash(ph, String(code ?? '')))) {
      throw new UnauthorizedException('Code incorrect');
    }
    await this.db.query('DELETE FROM otp_codes WHERE phone_hash = $1', [ph]);
    const user = await this.findOrCreateUser(e164);
    if (user.status !== 'active') throw new UnauthorizedException('Compte suspendu');
    return { token: await this.jwt.signAsync({ sub: user.id, role: user.role }), userId: user.id, role: user.role };
  }

  async findOrCreateUser(e164: string): Promise<{ id: string; role: string; status: string }> {
    const ph = this.crypto.hashPhone(e164);
    const [u] = await this.db.query(
      `INSERT INTO users (phone_hash, phone_enc) VALUES ($1, $2)
       ON CONFLICT (phone_hash) DO UPDATE SET phone_hash = EXCLUDED.phone_hash
       RETURNING id, role, status`,
      [ph, this.crypto.encrypt(e164)],
    );
    return u;
  }
}
