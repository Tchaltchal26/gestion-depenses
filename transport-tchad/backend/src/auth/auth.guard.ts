import { CanActivate, ExecutionContext, ForbiddenException, Injectable, SetMetadata, UnauthorizedException, createParamDecorator } from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { JwtService } from '@nestjs/jwt';
import { DbService } from '../db/db.service';

export type Role = 'client' | 'driver' | 'ops' | 'admin';
export interface AuthUser { id: string; role: Role }

export const Roles = (...roles: Role[]) => SetMetadata('roles', roles);
export const CurrentUser = createParamDecorator((_: unknown, ctx: ExecutionContext): AuthUser => ctx.switchToHttp().getRequest().user);

/** Vérifie le jeton, recharge l'utilisateur (statut et rôle à jour) et contrôle le rôle exigé. */
@Injectable()
export class AuthGuard implements CanActivate {
  constructor(private readonly jwt: JwtService, private readonly db: DbService, private readonly reflector: Reflector) {}

  async canActivate(ctx: ExecutionContext): Promise<boolean> {
    const req = ctx.switchToHttp().getRequest();
    const header: string = req.headers.authorization ?? '';
    if (!header.startsWith('Bearer ')) throw new UnauthorizedException();
    let payload: { sub: string };
    try {
      payload = await this.jwt.verifyAsync(header.slice(7));
    } catch {
      throw new UnauthorizedException();
    }
    const [u] = await this.db.query('SELECT id, role, status FROM users WHERE id = $1', [payload.sub]);
    if (!u || u.status !== 'active') throw new UnauthorizedException();
    req.user = { id: String(u.id), role: u.role } as AuthUser;
    const roles = this.reflector.getAllAndOverride<Role[]>('roles', [ctx.getHandler(), ctx.getClass()]);
    if (roles?.length && !roles.includes(u.role)) throw new ForbiddenException();
    return true;
  }
}
