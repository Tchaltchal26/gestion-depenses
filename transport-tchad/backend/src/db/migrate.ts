import { existsSync } from 'fs';
if (require.main === module && existsSync('.env')) process.loadEnvFile('.env');
import { readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import { Client } from 'pg';

/**
 * Applique les fichiers sql/0xx_*.sql non encore appliqués.
 * Avec --seed-dev, applique aussi sql/9xx_seed_dev.sql (données fictives de développement).
 */
export async function migrate(databaseUrl: string, seedDev = false): Promise<string[]> {
  const dir = join(__dirname, '..', '..', 'sql');
  const files = readdirSync(dir)
    .filter((f) => f.endsWith('.sql') && (seedDev || !f.startsWith('9')))
    .sort();
  const client = new Client({ connectionString: databaseUrl });
  await client.connect();
  const applied: string[] = [];
  try {
    await client.query('CREATE TABLE IF NOT EXISTS schema_migrations (name TEXT PRIMARY KEY, applied_at TIMESTAMPTZ NOT NULL DEFAULT now())');
    const done = new Set((await client.query('SELECT name FROM schema_migrations')).rows.map((r) => r.name));
    for (const f of files) {
      if (done.has(f)) continue;
      await client.query('BEGIN');
      try {
        await client.query(readFileSync(join(dir, f), 'utf8'));
        await client.query('INSERT INTO schema_migrations (name) VALUES ($1)', [f]);
        await client.query('COMMIT');
        applied.push(f);
      } catch (e) {
        await client.query('ROLLBACK');
        throw new Error(`Échec de la migration ${f} : ${(e as Error).message}`);
      }
    }
  } finally {
    await client.end();
  }
  return applied;
}

if (require.main === module) {
  const url = process.env.DATABASE_URL;
  if (!url) throw new Error('DATABASE_URL manquante');
  migrate(url, process.argv.includes('--seed-dev'))
    .then((a) => console.log(a.length ? `Appliqué : ${a.join(', ')}` : 'Base déjà à jour'))
    .catch((e) => {
      console.error(e.message);
      process.exit(1);
    });
}
