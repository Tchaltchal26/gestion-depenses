// Environnement de test : base dédiée, clés de test (jamais utilisées en production).
process.env.DATABASE_URL = process.env.TEST_DATABASE_URL ?? 'postgres://transport:transport@localhost:5432/transport_test';
process.env.PHONE_ENC_KEY = Buffer.alloc(32, 1).toString('base64');
process.env.PHONE_HASH_KEY = Buffer.alloc(32, 2).toString('base64');
process.env.JWT_SECRET = 'test-jwt-secret';
process.env.SMS_WEBHOOK_SECRET = 'test-webhook-secret';
process.env.SMS_PROVIDER = 'fake';
process.env.DISPATCH_TICK_MS = '0'; // les tests appellent tick() eux-mêmes
