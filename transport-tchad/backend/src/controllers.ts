import { Body, Controller, Get, Headers, HttpCode, NotFoundException, Param, Post, Put, Query, Req, UnauthorizedException, UseGuards, BadRequestException, ConflictException } from '@nestjs/common';
import { Inject } from '@nestjs/common';
import { AuthGuard, AuthUser, CurrentUser, Roles } from './auth/auth.guard';
import { AuthService } from './auth/auth.service';
import { CryptoService } from './common/crypto.service';
import { AppConfig, CONFIG } from './config';
import { DbService } from './db/db.service';
import { AvailabilityDto, CreateRideDto, InboundSmsDto, LocationDto, OtpRequestDto, OtpVerifyDto, QuoteQueryDto, RegisterDriverDto, SearchQueryDto, StartRideDto } from './dto';
import { GeoService } from './geo/geo.service';
import { DriversService } from './rides/drivers.service';
import { RidesService } from './rides/rides.service';
import { SmsInboundService } from './sms/sms.inbound.service';

@Controller('v1/auth')
export class AuthController {
  constructor(private readonly auth: AuthService) {}

  @Post('otp/request') @HttpCode(204)
  async request(@Body() dto: OtpRequestDto) {
    await this.auth.requestOtp(dto.phone);
  }

  @Post('otp/verify') @HttpCode(200)
  verify(@Body() dto: OtpVerifyDto) {
    return this.auth.verifyOtp(dto.phone, dto.code);
  }
}

@Controller('v1')
@UseGuards(AuthGuard)
export class GeoController {
  constructor(private readonly geo: GeoService) {}

  @Get('landmarks/search')
  search(@Query() q: SearchQueryDto) {
    return this.geo.searchLandmarks(q.q, 10);
  }

  @Get('tariffs/quote')
  async quote(@Query() q: QuoteQueryDto) {
    const [from, to] = await Promise.all([this.geo.getLandmark(q.from), this.geo.getLandmark(q.to)]);
    if (!from || !to) throw new NotFoundException('Repère inconnu');
    return { priceXaf: await this.geo.quote(q.type, from, to), indicative: true };
  }
}

@Controller('v1/rides')
@UseGuards(AuthGuard)
export class RidesController {
  constructor(private readonly rides: RidesService) {}

  @Post() @Roles('client')
  async create(@CurrentUser() u: AuthUser, @Body() dto: CreateRideDto, @Headers('idempotency-key') key?: string) {
    if (!key || key.length > 100) throw new BadRequestException('En-tête Idempotency-Key requis');
    const { ride } = await this.rides.createRide(u.id, { ...dto, channel: 'app', idempotencyKey: key });
    return this.rides.getRideForUser(u, ride.id);
  }

  @Get(':id')
  get(@CurrentUser() u: AuthUser, @Param('id') id: string) {
    return this.rides.getRideForUser(u, id);
  }

  @Post(':id/cancel') @HttpCode(200)
  async cancel(@CurrentUser() u: AuthUser, @Param('id') id: string) {
    if (u.role === 'driver') await this.rides.cancelByDriver(u.id, id);
    else await this.rides.cancelByClient(u.id, id, 'app');
    return this.rides.getRideForUser(u, id);
  }

  @Post(':id/arrived') @Roles('driver') @HttpCode(200)
  async arrived(@CurrentUser() u: AuthUser, @Param('id') id: string) {
    await this.rides.arrived(u.id, id);
    return this.rides.getRideForUser(u, id);
  }

  @Post(':id/start') @Roles('driver') @HttpCode(200)
  async start(@CurrentUser() u: AuthUser, @Param('id') id: string, @Body() dto: StartRideDto) {
    await this.rides.start(u.id, id, dto.pin);
    return this.rides.getRideForUser(u, id);
  }

  @Post(':id/complete') @Roles('driver') @HttpCode(200)
  async complete(@CurrentUser() u: AuthUser, @Param('id') id: string) {
    await this.rides.complete(u.id, id, 'app');
    return this.rides.getRideForUser(u, id);
  }
}

@Controller('v1')
@UseGuards(AuthGuard)
export class DriverController {
  constructor(private readonly drivers: DriversService, private readonly rides: RidesService) {}

  @Post('driver/register') @Roles('client')
  register(@CurrentUser() u: AuthUser, @Body() dto: RegisterDriverDto) {
    return this.drivers.register(u.id, dto);
  }

  @Put('driver/availability') @Roles('driver') @HttpCode(204)
  async availability(@CurrentUser() u: AuthUser, @Body() dto: AvailabilityDto) {
    await this.drivers.setAvailability(u.id, dto.available, 'app');
  }

  @Post('driver/location') @Roles('driver') @HttpCode(204)
  async location(@CurrentUser() u: AuthUser, @Body() dto: LocationDto) {
    await this.drivers.updateLocation(u.id, dto.lat, dto.lng, dto.batteryPct);
  }

  @Get('driver/offers') @Roles('driver')
  offers(@CurrentUser() u: AuthUser) {
    return this.rides.openOffersForDriver(u.id);
  }

  @Post('offers/:id/accept') @Roles('driver') @HttpCode(200)
  async accept(@CurrentUser() u: AuthUser, @Param('id') id: string) {
    const r = await this.rides.accept(u.id, { offerId: id }, 'app');
    if (!r.ok) {
      if (r.reason === 'not_found') throw new NotFoundException('Offre introuvable');
      throw new ConflictException({ message: 'Course déjà attribuée ou offre expirée', code: r.reason });
    }
    return this.rides.getRideForUser(u, r.rideId);
  }

  @Post('offers/:id/decline') @Roles('driver') @HttpCode(204)
  async decline(@CurrentUser() u: AuthUser, @Param('id') id: string) {
    await this.rides.decline(u.id, { offerId: id });
  }
}

@Controller('v1/admin')
@UseGuards(AuthGuard)
@Roles('admin')
export class AdminController {
  constructor(private readonly db: DbService) {}

  @Post('drivers/:id/approve') @HttpCode(204)
  async approve(@Param('id') id: string) {
    const r = await this.db.query(`UPDATE drivers SET verification_status = 'approved' WHERE user_id = $1 RETURNING user_id`, [id]);
    if (!r.length) throw new NotFoundException('Chauffeur introuvable');
  }

  @Post('users/:id/suspend') @HttpCode(204)
  async suspend(@Param('id') id: string) {
    const r = await this.db.query(`UPDATE users SET status = 'suspended' WHERE id = $1 AND role <> 'admin' RETURNING id`, [id]);
    if (!r.length) throw new NotFoundException('Utilisateur introuvable');
    await this.db.query('UPDATE driver_presence SET available = false WHERE driver_id = $1', [id]);
  }
}

/** Webhook de la passerelle SMS. La signature HMAC-SHA256 du corps brut est obligatoire. */
@Controller('v1/sms')
export class SmsController {
  constructor(private readonly inbound: SmsInboundService, private readonly crypto: CryptoService, @Inject(CONFIG) private readonly config: AppConfig) {}

  @Post('inbound/:provider') @HttpCode(200)
  handle(@Param('provider') provider: string, @Req() req: any, @Headers('x-signature') sig: string, @Body() dto: InboundSmsDto) {
    const raw: Buffer | undefined = req.rawBody;
    if (!raw || !sig || !/^[a-f0-9]{64}$/i.test(sig) ||
        !this.crypto.safeEqualHex(sig.toLowerCase(), this.crypto.hmac(raw.toString('utf8'), this.config.smsWebhookSecret))) {
      throw new UnauthorizedException('Signature invalide');
    }
    if (!/^[a-z0-9_-]{1,30}$/.test(provider)) throw new BadRequestException();
    return this.inbound.handle(provider, dto);
  }
}
