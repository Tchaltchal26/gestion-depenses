import { Inject, Injectable, Logger } from '@nestjs/common';
import { CryptoService } from '../common/crypto.service';
import { toGsmSafe } from '../common/text';
import { DbService } from '../db/db.service';
import { SMS_PROVIDER, SmsProvider } from './sms.provider';

export function smsSegments(body: string): number {
  return body.length <= 160 ? 1 : Math.ceil(body.length / 153);
}

@Injectable()
export class SmsSender {
  private readonly logger = new Logger('SmsSender');

  constructor(
    @Inject(SMS_PROVIDER) private readonly provider: SmsProvider,
    private readonly db: DbService,
    private readonly crypto: CryptoService,
  ) {}

  async sendToUser(userId: string | number, body: string, kind: string, rideId?: string | number) {
    const [u] = await this.db.query<{ phone_enc: string }>('SELECT phone_enc FROM users WHERE id = $1', [userId]);
    if (!u) return;
    await this.sendToPhone(this.crypto.decrypt(u.phone_enc), body, kind, rideId);
  }

  /** `logBody` permet de stocker une version masquée (ex. code OTP) dans le journal. */
  async sendToPhone(e164: string, body: string, kind: string, rideId?: string | number, logBody?: string) {
    const text = toGsmSafe(body);
    const [row] = await this.db.query<{ id: string }>(
      `INSERT INTO sms_messages (direction, provider, phone_hash, body, kind, ride_id, status, segments)
       VALUES ('out', $1, $2, $3, $4, $5, 'queued', $6) RETURNING id`,
      [this.provider.name, this.crypto.hashPhone(e164), toGsmSafe(logBody ?? body), kind, rideId ?? null, smsSegments(text)],
    );
    try {
      const { providerMessageId } = await this.provider.send(e164, text);
      await this.db.query(`UPDATE sms_messages SET status = 'sent', provider_message_id = $2 WHERE id = $1`, [row.id, providerMessageId]);
    } catch (e) {
      this.logger.error(`Échec d'envoi SMS ${row.id} : ${(e as Error).message}`);
      await this.db.query(`UPDATE sms_messages SET status = 'failed' WHERE id = $1`, [row.id]);
    }
  }
}
