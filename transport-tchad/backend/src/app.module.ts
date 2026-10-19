import { Module } from '@nestjs/common';
import { JwtModule } from '@nestjs/jwt';
import { AuthGuard } from './auth/auth.guard';
import { AuthService } from './auth/auth.service';
import { Clock } from './common/clock';
import { CryptoService } from './common/crypto.service';
import { CONFIG, loadConfig } from './config';
import { AdminController, AuthController, DriverController, GeoController, RidesController, SmsController } from './controllers';
import { DbService } from './db/db.service';
import { DevController, devToolsEnabled } from './dev/dev.controller';
import { GeoService } from './geo/geo.service';
import { Notifier } from './notifier';
import { Dispatcher } from './rides/dispatcher';
import { DriversService } from './rides/drivers.service';
import { RidesService } from './rides/rides.service';
import { SmsInboundService } from './sms/sms.inbound.service';
import { FakeSmsProvider, SMS_PROVIDER } from './sms/sms.provider';
import { SmsSender } from './sms/sms.sender';

const config = loadConfig();
if (config.smsProvider !== 'fake') {
  throw new Error(`SMS_PROVIDER=${config.smsProvider} non disponible : seul "fake" (simulé) est implémenté pour l'instant`);
}

@Module({
  imports: [JwtModule.register({ secret: config.jwtSecret, signOptions: { expiresIn: '15m' } })],
  controllers: [AuthController, GeoController, RidesController, DriverController, AdminController, SmsController,
    ...(devToolsEnabled() ? [DevController] : [])],
  providers: [
    { provide: CONFIG, useValue: config },
    { provide: SMS_PROVIDER, useClass: FakeSmsProvider },
    Clock, CryptoService, DbService, GeoService, Notifier, SmsSender, AuthService, AuthGuard,
    RidesService, DriversService, SmsInboundService, Dispatcher,
  ],
})
export class AppModule {}
