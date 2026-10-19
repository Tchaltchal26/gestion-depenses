import { Type } from 'class-transformer';
import { IsBoolean, IsIn, IsInt, IsLatitude, IsLongitude, IsNotEmpty, IsOptional, IsString, Length, Matches, Max, MaxLength, Min } from 'class-validator';

const TYPES = ['taxi', 'rakcha', 'clando'];

export class OtpRequestDto {
  @IsString() @MaxLength(20) phone: string;
}
export class OtpVerifyDto {
  @IsString() @MaxLength(20) phone: string;
  @Matches(/^\d{6}$/) code: string;
}
export class CreateRideDto {
  @IsIn(TYPES) transportType: 'taxi' | 'rakcha' | 'clando';
  @IsOptional() @IsLatitude() pickupLat?: number;
  @IsOptional() @IsLongitude() pickupLng?: number;
  @IsOptional() @IsInt() @Min(1) pickupLandmarkId?: number;
  @IsOptional() @IsString() @MaxLength(120) pickupText?: string;
  @IsOptional() @IsInt() @Min(1) dropoffLandmarkId?: number;
  @IsOptional() @IsString() @MaxLength(120) dropoffText?: string;
}
export class RegisterDriverDto {
  @IsString() @IsNotEmpty() @MaxLength(80) name: string;
  @IsIn(TYPES) transportType: 'taxi' | 'rakcha' | 'clando';
  @IsOptional() @IsString() @MaxLength(20) plate?: string;
  @IsOptional() @IsString() @MaxLength(30) color?: string;
}
export class AvailabilityDto {
  @IsBoolean() available: boolean;
}
export class LocationDto {
  @IsLatitude() lat: number;
  @IsLongitude() lng: number;
  @IsOptional() @IsInt() @Min(0) @Max(100) batteryPct?: number;
}
export class StartRideDto {
  @Matches(/^\d{4}$/) pin: string;
}
export class SearchQueryDto {
  @IsString() @Length(2, 80) q: string;
}
export class QuoteQueryDto {
  @IsIn(TYPES) type: string;
  @Type(() => Number) @IsInt() from: number;
  @Type(() => Number) @IsInt() to: number;
}
export class InboundSmsDto {
  @IsString() @IsNotEmpty() @MaxLength(100) messageId: string;
  @IsString() @MaxLength(20) from: string;
  @IsString() @MaxLength(1000) text: string;
  @IsOptional() @IsString() receivedAt?: string;
}
