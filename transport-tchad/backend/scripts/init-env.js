// Crée un fichier .env de développement avec des clés aléatoires (ne remplace pas un .env existant).
const fs = require('fs');
const path = require('path');
const { randomBytes } = require('crypto');
const target = path.join(__dirname, '..', '.env');
if (fs.existsSync(target)) {
  console.log('.env existe déjà : rien à faire.');
  process.exit(0);
}
const key = () => randomBytes(32).toString('base64');
fs.writeFileSync(target, [
  'PORT=3000',
  'DATABASE_URL=postgres://transport:transport@localhost:5433/transport',
  `PHONE_ENC_KEY=${key()}`,
  `PHONE_HASH_KEY=${key()}`,
  `JWT_SECRET=${randomBytes(32).toString('hex')}`,
  `SMS_WEBHOOK_SECRET=${randomBytes(32).toString('hex')}`,
  'SMS_PROVIDER=fake',
  'DISPATCH_TICK_MS=2000',
  'DEV_TOOLS=1',
  '',
].join('\n'));
console.log('.env créé avec des clés aléatoires (développement uniquement).');
