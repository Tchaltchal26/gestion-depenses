import { Client } from 'pg';
import { migrate } from '../src/db/migrate';

/** Recrée entièrement le schéma de la base de test puis applique migrations + données de dev. */
export default async function () {
  require('./env');
  const url = process.env.DATABASE_URL!;
  const c = new Client({ connectionString: url });
  await c.connect();
  await c.query('DROP SCHEMA IF EXISTS public CASCADE; CREATE SCHEMA public;');
  await c.end();
  await migrate(url, true);
}
