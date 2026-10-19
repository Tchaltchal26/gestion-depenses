import { Inject, Injectable, Logger, OnApplicationBootstrap, OnApplicationShutdown } from '@nestjs/common';
import { AppConfig, CONFIG } from '../config';
import { RidesService } from './rides.service';

/** Lance RidesService.tick() à intervalle régulier (désactivé si DISPATCH_TICK_MS = 0, ex. tests). */
@Injectable()
export class Dispatcher implements OnApplicationBootstrap, OnApplicationShutdown {
  private timer?: NodeJS.Timeout;
  private running = false;
  private readonly logger = new Logger('Dispatcher');

  constructor(private readonly rides: RidesService, @Inject(CONFIG) private readonly config: AppConfig) {}

  onApplicationBootstrap() {
    if (this.config.dispatchTickMs <= 0) return;
    this.timer = setInterval(async () => {
      if (this.running) return;
      this.running = true;
      try {
        await this.rides.tick();
      } catch (e) {
        this.logger.error(`tick : ${(e as Error).message}`);
      } finally {
        this.running = false;
      }
    }, this.config.dispatchTickMs);
  }

  onApplicationShutdown() {
    if (this.timer) clearInterval(this.timer);
  }
}
