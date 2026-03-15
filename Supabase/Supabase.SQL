-- =============================================================
-- OptiSheets — Full Supabase Schema (Reconstructed)
-- Run this in the Supabase SQL Editor in one go.
-- =============================================================

-- ---------------------------------------------------------------
-- Extensions
-- ---------------------------------------------------------------
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ---------------------------------------------------------------
-- 1. customers
-- ---------------------------------------------------------------
CREATE TABLE IF NOT EXISTS customers (
  id          UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  username    TEXT        NOT NULL,
  email       TEXT        NOT NULL UNIQUE,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE customers ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON customers FOR ALL USING (false);

-- ---------------------------------------------------------------
-- 2. templates
-- (Referenced by purchases, api/templates.js, admin/templates.js)
-- ---------------------------------------------------------------
CREATE TABLE IF NOT EXISTS templates (
  id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  slug         TEXT        NOT NULL UNIQUE,
  name         TEXT        NOT NULL,
  description  TEXT,
  price_cents  INTEGER     NOT NULL CHECK (price_cents >= 0),
  category     TEXT,
  preview_url  TEXT,
  download_url TEXT,
  published    BOOLEAN     NOT NULL DEFAULT false,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE templates ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON templates FOR ALL USING (false);

-- ---------------------------------------------------------------
-- 3. purchases
-- (Referenced by api/purchase.js, api/generate.js, admin/purchases.js)
-- ---------------------------------------------------------------
CREATE TABLE IF NOT EXISTS purchases (
  id                 UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  template_id        UUID        NOT NULL REFERENCES templates (id) ON DELETE RESTRICT,
  user_email         TEXT        NOT NULL,
  payment_reference  TEXT        NOT NULL,
  purchased_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE purchases ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON purchases FOR ALL USING (false);

-- ---------------------------------------------------------------
-- 4. licenses
-- ---------------------------------------------------------------
CREATE TABLE IF NOT EXISTS licenses (
  id            UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  customer_id   UUID        NOT NULL REFERENCES customers (id) ON DELETE CASCADE,
  template_id   TEXT        NOT NULL,
  private_key   TEXT        UNIQUE,  -- NULL for standard (non-AI) templates
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE licenses ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON licenses FOR ALL USING (false);

-- ---------------------------------------------------------------
-- 5. credit_wallets
-- ---------------------------------------------------------------
CREATE TABLE IF NOT EXISTS credit_wallets (
  id          UUID    PRIMARY KEY DEFAULT gen_random_uuid(),
  license_id  UUID    NOT NULL UNIQUE REFERENCES licenses (id) ON DELETE CASCADE,
  balance     INTEGER NOT NULL DEFAULT 0 CHECK (balance >= 0)
);

ALTER TABLE credit_wallets ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON credit_wallets FOR ALL USING (false);

-- ---------------------------------------------------------------
-- 6. ledger
-- ---------------------------------------------------------------
CREATE TYPE ledger_event_type AS ENUM ('purchase', 'usage', 'refund', 'adjustment');

CREATE TABLE IF NOT EXISTS ledger (
  id          UUID              PRIMARY KEY DEFAULT gen_random_uuid(),
  license_id  UUID              NOT NULL REFERENCES licenses (id) ON DELETE CASCADE,
  event_type  ledger_event_type NOT NULL,
  amount      INTEGER           NOT NULL,  -- positive = credit, negative = debit
  note        TEXT,
  created_at  TIMESTAMPTZ       NOT NULL DEFAULT NOW()
);

ALTER TABLE templates ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON ledger FOR ALL USING (false);

-- ---------------------------------------------------------------
-- 7. ai_cache
-- ---------------------------------------------------------------
CREATE TABLE IF NOT EXISTS ai_cache (
  id            UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  request_hash  TEXT        NOT NULL UNIQUE,  -- SHA-256 of (template_id + sorted inputs)
  template_id   TEXT        NOT NULL,
  output        TEXT        NOT NULL,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE ai_cache ENABLE ROW LEVEL SECURITY;

CREATE POLICY "deny all to non-service roles"
  ON ai_cache FOR ALL USING (false);

-- ---------------------------------------------------------------
-- Indexes
-- ---------------------------------------------------------------
CREATE INDEX IF NOT EXISTS idx_templates_slug          ON templates     (slug);
CREATE INDEX IF NOT EXISTS idx_purchases_template_id   ON purchases     (template_id);
CREATE INDEX IF NOT EXISTS idx_purchases_user_email    ON purchases     (user_email);
CREATE INDEX IF NOT EXISTS idx_licenses_customer_id    ON licenses      (customer_id);
CREATE INDEX IF NOT EXISTS idx_licenses_template_id    ON licenses      (template_id);
CREATE INDEX IF NOT EXISTS idx_ledger_license_id       ON ledger        (license_id);
CREATE INDEX IF NOT EXISTS idx_ledger_event_type       ON ledger        (event_type);
CREATE INDEX IF NOT EXISTS idx_ai_cache_template_id    ON ai_cache      (template_id);