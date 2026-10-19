import 'reflect-metadata';
import { existsSync } from 'fs';
if (existsSync('.env')) process.loadEnvFile('.env');
import { ValidationPipe } from '@nestjs/common';
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { DbService } from './db/db.service';

async function bootstrap() {
  const app = await NestFactory.create(AppModule, { rawBody: true });
  // Vérifie la base avant de démarrer : message clair au lieu d'erreurs répétées toutes les 2 secondes.
  try {
    await app.get(DbService).query('SELECT 1');
  } catch (e) {
    console.error(`\nImpossible de se connecter à la base de données : ${(e as Error).message}`);
    console.error(`DATABASE_URL utilisée : ${(process.env.DATABASE_URL ?? '').replace(/:[^:@/]+@/, ':****@')}`);
    console.error('Vérifiez que Docker Desktop est ouvert, que "npm run db:up" a réussi, et que le port dans .env');
    console.error('correspond à celui de docker-compose.yml (5433). Voir README, section « Problèmes fréquents ».\n');
    await app.close();
    process.exit(1);
  }
  app.useGlobalPipes(new ValidationPipe({ whitelist: true, forbidNonWhitelisted: true, transform: true }));
  app.enableShutdownHooks();
  const port = Number(process.env.PORT ?? 3000);
  await app.listen(port);
  console.log(`API démarrée sur http://localhost:${port}/v1`);
}
bootstrap();
