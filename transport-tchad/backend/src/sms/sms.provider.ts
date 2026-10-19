import { Injectable, Logger } from '@nestjs/common';

export interface SmsProvider {
  readonly name: string;
  send(toE164: string, body: string): Promise<{ providerMessageId: string }>;
}

export const SMS_PROVIDER = Symbol('SMS_PROVIDER');

/**
 * Fournisseur SIMULÉ : aucun SMS réel n'est envoyé. Les messages sont gardés en mémoire
 * (lisibles par les tests) et écrits dans le journal. À remplacer par un adaptateur réel
 * (passerelle SIM, agrégateur HTTP ou SMPP opérateur) implémentant la même interface.
 */
@Injectable()
export class FakeSmsProvider implements SmsProvider {
  readonly name = 'fake';
  readonly sent: { to: string; body: string }[] = [];
  private readonly logger = new Logger('FakeSms');
  private seq = 0;

  async send(toE164: string, body: string) {
    this.sent.push({ to: toE164, body });
    this.logger.debug(`[SIMULÉ] → ${toE164.slice(0, 7)}***** : ${body}`);
    return { providerMessageId: `fake-${++this.seq}` };
  }
}
