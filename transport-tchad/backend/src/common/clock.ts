import { Injectable } from '@nestjs/common';

/** Horloge injectable : les tests peuvent avancer le temps pour simuler les expirations. */
@Injectable()
export class Clock {
  private offsetMs = 0;
  now(): Date {
    return new Date(Date.now() + this.offsetMs);
  }
  reset(): void {
    this.offsetMs = 0;
  }
  advance(seconds: number): void {
    this.offsetMs += seconds * 1000;
  }
}
