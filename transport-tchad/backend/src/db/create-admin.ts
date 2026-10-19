import { existsSync } from 'fs';
if (require.main === module && existsSync('.env')) process.loadEnvFile('.env');
import { Client } from 'pg';
import { CryptoService } from '../common/crypto.service';
import { loadConfig } from '../config';
import { normalizeChadPhone } from '../common/phone';

/** Donne le rôle administrateur au numéro indiqué : npm run admin:create -- 66123456 */
async function main() {
  const config = loadConfig();
  const phone = normalizeChadPhone(process.argv[2] ?? '');
  if (!phone) throw new Error('Usage : npm run admin:create -- <numéro tchadien>');
  const crypto = new CryptoService(config);
  const c = new Client({ connectionString: config.databaseUrl });
  await c.connect();
  await c.query(
    `INSERT INTO users (phone_hash, phone_enc, role, name) VALUES ($1, $2, 'admin', 'Administrateur')
     ON CONFLICT (phone_hash) DO UPDATE SET role = 'admin'`,
    [crypto.hashPhone(phone), crypto.encrypt(phone)]);
  await c.end();
  console.log(`${phone} est administrateur. Connexion par code SMS via /v1/auth/otp/request.`);
}
main().catch((e) => { console.error(e.message); process.exit(1); });
