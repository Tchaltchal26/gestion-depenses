import { Injectable, Logger } from '@nestjs/common';

/**
 * Notifications vers les applications : SIMULÉES pour l'instant (journal + mémoire).
 * Socket.IO et Firebase Cloud Messaging ne sont pas encore branchés ; les applications
 * peuvent en attendant interroger GET /v1/rides/:id et GET /v1/driver/offers.
 */
@Injectable()
export class Notifier {
  readonly events: { userId: string; event: string; data: unknown }[] = [];
  private readonly logger = new Logger('Notifier');

  notify(userId: string | number, event: string, data: unknown) {
    this.events.push({ userId: String(userId), event, data });
    this.logger.debug(`[SIMULÉ] ${event} → utilisateur ${userId}`);
  }
}
