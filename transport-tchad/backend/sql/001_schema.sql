-- Schéma initial. Exécuté par src/db/migrate.ts.
CREATE EXTENSION IF NOT EXISTS postgis;
CREATE EXTENSION IF NOT EXISTS pg_trgm;

CREATE TABLE users (
  id          BIGSERIAL PRIMARY KEY,
  phone_hash  TEXT NOT NULL UNIQUE,           -- HMAC du numéro E.164 (recherche)
  phone_enc   TEXT NOT NULL,                  -- numéro chiffré AES-256-GCM
  name        TEXT,
  role        TEXT NOT NULL DEFAULT 'client' CHECK (role IN ('client','driver','ops','admin')),
  status      TEXT NOT NULL DEFAULT 'active' CHECK (status IN ('active','suspended')),
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE drivers (
  user_id             BIGINT PRIMARY KEY REFERENCES users(id),
  transport_type      TEXT NOT NULL CHECK (transport_type IN ('taxi','rakcha','clando')),
  verification_status TEXT NOT NULL DEFAULT 'pending' CHECK (verification_status IN ('pending','approved','rejected')),
  vehicle_plate       TEXT,
  vehicle_color       TEXT,
  offered_count       INT NOT NULL DEFAULT 0,
  accepted_count      INT NOT NULL DEFAULT 0
);

CREATE TABLE driver_presence (
  driver_id        BIGINT PRIMARY KEY REFERENCES drivers(user_id),
  available        BOOLEAN NOT NULL DEFAULT false,
  channel          TEXT NOT NULL DEFAULT 'app' CHECK (channel IN ('app','sms')),
  location         GEOGRAPHY(Point, 4326),
  location_source  TEXT CHECK (location_source IN ('gps','landmark')),
  location_at      TIMESTAMPTZ,
  battery_pct      SMALLINT
);
CREATE INDEX driver_presence_location_idx ON driver_presence USING GIST (location);

CREATE TABLE zones (
  id      SERIAL PRIMARY KEY,
  name    TEXT NOT NULL,
  area    GEOGRAPHY(Polygon, 4326) NOT NULL,
  active  BOOLEAN NOT NULL DEFAULT true
);

CREATE TABLE landmarks (
  id        SERIAL PRIMARY KEY,
  name      TEXT NOT NULL,
  quartier  TEXT,
  location  GEOGRAPHY(Point, 4326) NOT NULL,
  verified  BOOLEAN NOT NULL DEFAULT false
);
-- Un nom principal + des alias (français, arabe translittéré...), normalisés sans accents.
CREATE TABLE landmark_names (
  landmark_id INT NOT NULL REFERENCES landmarks(id) ON DELETE CASCADE,
  norm        TEXT NOT NULL,
  PRIMARY KEY (landmark_id, norm)
);
CREATE INDEX landmark_names_trgm_idx ON landmark_names USING GIN (norm gin_trgm_ops);

CREATE TABLE tariffs (
  transport_type TEXT PRIMARY KEY CHECK (transport_type IN ('taxi','rakcha','clando')),
  base_xaf       INT NOT NULL,
  per_km_xaf     INT NOT NULL,
  min_xaf        INT NOT NULL
);

CREATE TABLE rides (
  id                  BIGSERIAL PRIMARY KEY,
  public_code         TEXT NOT NULL UNIQUE,
  client_id           BIGINT NOT NULL REFERENCES users(id),
  driver_id           BIGINT REFERENCES drivers(user_id),
  transport_type      TEXT NOT NULL CHECK (transport_type IN ('taxi','rakcha','clando')),
  status              TEXT NOT NULL CHECK (status IN ('SEARCHING','OFFERED','ASSIGNED','ARRIVED','IN_PROGRESS','COMPLETED','NO_DRIVER','CANCELLED_BY_CLIENT','CANCELLED_BY_DRIVER','CANCELLED_BY_OPS')),
  channel             TEXT NOT NULL CHECK (channel IN ('app','sms','ops')),
  pickup              GEOGRAPHY(Point, 4326) NOT NULL,
  pickup_landmark_id  INT REFERENCES landmarks(id),
  pickup_text         TEXT,
  dropoff_landmark_id INT REFERENCES landmarks(id),
  dropoff_text        TEXT,
  price_quoted_xaf    INT,
  pickup_pin          TEXT NOT NULL,
  idempotency_key     TEXT,
  radius_idx          SMALLINT NOT NULL DEFAULT 0,
  wave                SMALLINT NOT NULL DEFAULT 0,
  search_deadline     TIMESTAMPTZ NOT NULL,
  created_at          TIMESTAMPTZ NOT NULL,
  assigned_at         TIMESTAMPTZ,
  completed_at        TIMESTAMPTZ,
  cancelled_at        TIMESTAMPTZ,
  UNIQUE (client_id, idempotency_key)
);
-- Un chauffeur ne peut avoir qu'une seule course active.
CREATE UNIQUE INDEX rides_one_active_per_driver ON rides (driver_id)
  WHERE status IN ('ASSIGNED','ARRIVED','IN_PROGRESS');
-- Un client ne peut avoir qu'une seule demande en recherche.
CREATE UNIQUE INDEX rides_one_searching_per_client ON rides (client_id)
  WHERE status IN ('SEARCHING','OFFERED');
CREATE INDEX rides_status_idx ON rides (status);

CREATE TABLE ride_offers (
  id           BIGSERIAL PRIMARY KEY,
  ride_id      BIGINT NOT NULL REFERENCES rides(id),
  driver_id    BIGINT NOT NULL REFERENCES drivers(user_id),
  wave         SMALLINT NOT NULL,
  channel      TEXT NOT NULL CHECK (channel IN ('app','sms')),
  sent_at      TIMESTAMPTZ NOT NULL,
  expires_at   TIMESTAMPTZ NOT NULL,
  response     TEXT CHECK (response IN ('accepted','declined','expired','revoked')),
  responded_at TIMESTAMPTZ,
  UNIQUE (ride_id, driver_id)
);
-- Un chauffeur ne reçoit qu'une offre ouverte à la fois.
CREATE UNIQUE INDEX ride_offers_one_open_per_driver ON ride_offers (driver_id) WHERE response IS NULL;

CREATE TABLE ride_events (
  id          BIGSERIAL PRIMARY KEY,
  ride_id     BIGINT NOT NULL REFERENCES rides(id),
  from_status TEXT,
  to_status   TEXT NOT NULL,
  actor_id    BIGINT,
  channel     TEXT,
  at          TIMESTAMPTZ NOT NULL,
  meta        JSONB
);

CREATE TABLE otp_codes (
  phone_hash    TEXT PRIMARY KEY,
  code_hash     TEXT NOT NULL,
  expires_at    TIMESTAMPTZ NOT NULL,
  attempts      SMALLINT NOT NULL DEFAULT 0,
  window_start  TIMESTAMPTZ NOT NULL,
  sent_in_window SMALLINT NOT NULL DEFAULT 1
);

CREATE TABLE sms_messages (
  id                  BIGSERIAL PRIMARY KEY,
  direction           TEXT NOT NULL CHECK (direction IN ('in','out')),
  provider            TEXT NOT NULL,
  provider_message_id TEXT,
  phone_hash          TEXT NOT NULL,
  body                TEXT NOT NULL,          -- les codes OTP sont masqués avant stockage
  kind                TEXT,
  ride_id             BIGINT REFERENCES rides(id),
  status              TEXT NOT NULL,
  segments            SMALLINT NOT NULL DEFAULT 1,
  created_at          TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE (provider, provider_message_id)
);
CREATE INDEX sms_messages_phone_idx ON sms_messages (phone_hash, created_at);

-- Demande SMS en attente d'un choix (désambiguïsation du lieu de départ).
CREATE TABLE sms_pending (
  phone_hash TEXT PRIMARY KEY,
  data       JSONB NOT NULL,
  expires_at TIMESTAMPTZ NOT NULL
);
