'use strict';

// ---------------------------------------------------------------------------
// OptiSheets — End-to-end integration test
//
// Prerequisites:
//   1. Copy .env.example → .env and fill in real credentials
//   2. Start the local dev server:  node api/server.js
//   3. Run this script:             node test/testFlow.js
//
// What it does:
//   • Seeds a mock customer + license + wallet (balance = 3) directly in Supabase
//   • Runs the full API flow against the local server (validate-key, get-recommendations)
//   • Verifies credit deduction and cache behaviour
//   • Cleans up ALL seeded rows (including ai_cache + ledger) after the run
//
// WARNING: Step 3 calls OpenAI (gpt-4o-mini). It consumes one real API credit
//          and may take several seconds. Cost is negligible (<$0.001).
//
// Environment variables (loaded from .env):
//   SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY — direct DB access for seed/cleanup
//   OPENAI_API_KEY                          — used by the server, not directly here
//   TEST_BASE_URL                           — optional, defaults to http://localhost:3000
// ---------------------------------------------------------------------------

require('dotenv').config();

const crypto = require('crypto');
const { v4: uuidv4 } = require('uuid');
const { createClient } = require('@supabase/supabase-js');
const { SYSTEM_PROMPT } = require('../templates/internship-tracker/prompt');

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const BASE_URL   = (process.env.TEST_BASE_URL || 'http://localhost:3000').replace(/\/+$/, '');
const TEMPLATE_ID = 'internship-tracker';

// ---------------------------------------------------------------------------
// Supabase client (service role — bypasses RLS for seeding and cleanup)
// ---------------------------------------------------------------------------

if (!process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) {
  console.error('ERROR: SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set in .env');
  process.exit(1);
}

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY,
  { auth: { persistSession: false, autoRefreshToken: false } }
);

// ---------------------------------------------------------------------------
// ANSI colours
// ---------------------------------------------------------------------------

const C = {
  green:  '\x1b[32m',
  red:    '\x1b[31m',
  yellow: '\x1b[33m',
  cyan:   '\x1b[36m',
  bold:   '\x1b[1m',
  reset:  '\x1b[0m',
};

// ---------------------------------------------------------------------------
// Test reporter
// ---------------------------------------------------------------------------

let passed = 0;
let failed = 0;

function pass(label) {
  passed++;
  console.log(`${C.green}  ✓ PASS${C.reset}  ${label}`);
}

function fail(label, detail) {
  failed++;
  console.log(`${C.red}  ✗ FAIL${C.reset}  ${label}`);
  if (detail !== undefined) {
    console.log(`         ${C.red}${detail}${C.reset}`);
  }
}

function section(title) {
  console.log(`\n${C.cyan}${C.bold}── ${title}${C.reset}`);
}

function info(msg) {
  console.log(`  ${C.yellow}${msg}${C.reset}`);
}

// ---------------------------------------------------------------------------
// HTTP helpers
// ---------------------------------------------------------------------------

/**
 * POST JSON to the local server and return { status, body }.
 * Uses a 60-second AbortController timeout (enough for an OpenAI round-trip).
 */
async function post(path, body, extraHeaders) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 60_000);
  try {
    const res = await fetch(`${BASE_URL}${path}`, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json', ...extraHeaders },
      body:    JSON.stringify(body),
      signal:  controller.signal,
    });
    const json = await res.json();
    return { status: res.status, body: json };
  } finally {
    clearTimeout(timer);
  }
}

// ---------------------------------------------------------------------------
// Mirrors lib/aiRequestHandler.js — needed to predict the cache key for cleanup
// ---------------------------------------------------------------------------

function sortObjectKeys(obj) {
  if (typeof obj !== 'object' || obj === null || Array.isArray(obj)) return obj;
  return Object.keys(obj)
    .sort()
    .reduce((acc, k) => { acc[k] = sortObjectKeys(obj[k]); return acc; }, {});
}

function computeRequestHash(templateId, userData) {
  const canonical = JSON.stringify({
    template_id: templateId,
    user_data:   sortObjectKeys(userData),
  });
  return crypto.createHash('sha256').update(canonical).digest('hex');
}

// ---------------------------------------------------------------------------
// Seed helpers
// ---------------------------------------------------------------------------

async function seedCustomer(id, email) {
  const { error } = await supabase.from('customers').insert({
    id,
    username: 'test-user',
    email,
  });
  if (error) throw new Error(`Customer seed failed: ${error.message}`);
  pass('Seeded customer');
}

async function seedLicense(id, customerId, privateKey) {
  const { error } = await supabase.from('licenses').insert({
    id,
    customer_id: customerId,
    template_id: TEMPLATE_ID,
    private_key: privateKey,
  });
  if (error) throw new Error(`License seed failed: ${error.message}`);
  pass('Seeded license');
}

async function seedWallet(id, licenseId, balance) {
  const { error } = await supabase.from('credit_wallets').insert({
    id,
    license_id: licenseId,
    balance,
  });
  if (error) throw new Error(`Wallet seed failed (balance ${balance}): ${error.message}`);
  pass(`Seeded credit wallet (balance = ${balance})`);
}

// ---------------------------------------------------------------------------
// Cleanup
// ---------------------------------------------------------------------------

async function cleanup({ customerId, licenseId, walletId, requestHash }) {
  section('Cleanup');

  const steps = [
    ['ai_cache entry',    () => supabase.from('ai_cache').delete().eq('request_hash', requestHash)],
    ['ledger entries',    () => supabase.from('ledger').delete().eq('license_id', licenseId)],
    ['credit_wallet',     () => supabase.from('credit_wallets').delete().eq('id', walletId)],
    ['license',           () => supabase.from('licenses').delete().eq('id', licenseId)],
    ['customer',          () => supabase.from('customers').delete().eq('id', customerId)],
  ];

  for (const [label, fn] of steps) {
    const { error } = await fn();
    if (error) {
      console.log(`  ${C.yellow}⚠ cleanup warning — ${label}: ${error.message}${C.reset}`);
    } else {
      console.log(`  cleaned up ${label}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------

function printSummary() {
  const total = passed + failed;
  console.log(`\n${'─'.repeat(50)}`);
  if (failed === 0) {
    console.log(`${C.green}${C.bold}All ${total} tests passed.${C.reset}`);
  } else {
    console.log(
      `${C.bold}Results: ${C.green}${passed} passed${C.reset}${C.bold}, ` +
      `${C.red}${failed} failed${C.reset}${C.bold} / ${total} total${C.reset}`
    );
  }
  console.log('─'.repeat(50));
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  // Check Node version (fetch requires ≥ 18)
  const [major] = process.versions.node.split('.').map(Number);
  if (major < 18) {
    console.error(`ERROR: Node.js 18+ required (found ${process.version}). Upgrade and retry.`);
    process.exit(1);
  }

  console.log(`${C.bold}OptiSheets — Integration Test${C.reset}`);
  console.log(`Server: ${BASE_URL}`);

  // --- server health check ---
  try {
    const res = await fetch(`${BASE_URL}/`, { signal: AbortSignal.timeout(4_000) });
    // Any response (including 404) means the server is up.
    void res;
  } catch {
    console.error(
      `\nERROR: Cannot reach the local server at ${BASE_URL}.\n` +
      `Start it first with:  node api/server.js\n`
    );
    process.exit(1);
  }

  // Unique IDs for this test run
  const customerId  = uuidv4();
  const licenseId   = uuidv4();
  const walletId    = uuidv4();
  const privateKey  = `test-${uuidv4()}`;
  const testEmail   = `test-${Date.now()}@optisheets-test.invalid`;

  // Mock internship data — kept constant so both API calls hash identically
  const userData = {
    applications: [
      {
        company:     'Stripe',
        role:        'Software Engineering Intern',
        status:      'Phone Screen',
        appliedDate: '2026-01-10',
        notes:       'Referred by alum',
      },
      {
        company:     'Notion',
        role:        'Product Management Intern',
        status:      'Applied',
        appliedDate: '2026-01-15',
      },
      {
        company: 'Figma',
        role:    'Design Technologist Intern',
        status:  'Rejected',
      },
      {
        company:     'Vercel',
        role:        'Software Engineering Intern',
        status:      'Technical Screen',
        appliedDate: '2026-01-20',
        notes:       'Take-home due Friday',
      },
    ],
    targetRole:   'Software Engineering Intern',
    targetSeason: 'Summer 2026',
    targetCount:  2,
  };

  const requestHash = computeRequestHash(TEMPLATE_ID, userData);
  let seeded = false;

  try {
    // ── Seed ───────────────────────────────────────────────────────────────
    section('Seeding mock data into Supabase');
    await seedCustomer(customerId, testEmail);
    await seedLicense(licenseId, customerId, privateKey);
    await seedWallet(walletId, licenseId, 3);
    seeded = true;

    // ── Step 1: validate-key ───────────────────────────────────────────────
    section('Step 1 — POST /api/validate-key');

    const vk1 = await post('/api/validate-key', { private_key: privateKey });

    vk1.body.valid === true
      ? pass('valid = true')
      : fail('valid = true', `Got: ${JSON.stringify(vk1.body)}`);

    vk1.status === 200
      ? pass('HTTP 200')
      : fail('HTTP 200', `Got: ${vk1.status}`);

    vk1.body.balance === 3
      ? pass('balance = 3')
      : fail('balance = 3', `Got: ${vk1.body.balance}`);

    vk1.body.template_id === TEMPLATE_ID
      ? pass(`template_id = "${TEMPLATE_ID}"`)
      : fail(`template_id = "${TEMPLATE_ID}"`, `Got: "${vk1.body.template_id}"`);

    // ── Step 2: get-recommendations — cache miss ───────────────────────────
    section('Step 2 — POST /api/get-recommendations (first call — expects cache miss)');
    info('Calling OpenAI via the server; this may take a few seconds…');

    const rec1 = await post('/api/get-recommendations', {
      private_key:   privateKey,
      template_id:   TEMPLATE_ID,
      user_data:     userData,
      system_prompt: SYSTEM_PROMPT,
    });

    rec1.body.success === true
      ? pass('success = true')
      : fail('success = true', `Got: ${JSON.stringify(rec1.body)}`);

    rec1.status === 200
      ? pass('HTTP 200')
      : fail('HTTP 200', `Got: ${rec1.status}`);

    rec1.body.cached === false
      ? pass('cached = false  (cache miss as expected)')
      : fail('cached = false', `Got cached = ${rec1.body.cached}`);

    typeof rec1.body.output === 'string' && rec1.body.output.length > 0
      ? pass(`output is non-empty string  (${rec1.body.output.length} chars)`)
      : fail('output is non-empty string', `Got: ${JSON.stringify(rec1.body.output)}`);

    rec1.body.remaining_credits === 2
      ? pass('remaining_credits = 2  (1 credit consumed)')
      : fail('remaining_credits = 2', `Got: ${rec1.body.remaining_credits}`);

    // ── Step 3: confirm balance via validate-key ───────────────────────────
    section('Step 3 — Confirm balance decreased to 2 via /api/validate-key');

    const vk2 = await post('/api/validate-key', { private_key: privateKey });

    vk2.body.balance === 2
      ? pass('balance = 2  (credit was deducted)')
      : fail('balance = 2', `Got: ${vk2.body.balance}`);

    // ── Step 4: get-recommendations — cache hit ────────────────────────────
    section('Step 4 — POST /api/get-recommendations (identical data — expects cache hit)');

    const rec2 = await post('/api/get-recommendations', {
      private_key:   privateKey,
      template_id:   TEMPLATE_ID,
      user_data:     userData,       // identical to step 2 → same hash
      system_prompt: SYSTEM_PROMPT,
    });

    rec2.body.success === true
      ? pass('success = true')
      : fail('success = true', `Got: ${JSON.stringify(rec2.body)}`);

    rec2.status === 200
      ? pass('HTTP 200')
      : fail('HTTP 200', `Got: ${rec2.status}`);

    rec2.body.cached === true
      ? pass('cached = true  (cache hit as expected)')
      : fail('cached = true', `Got cached = ${rec2.body.cached}`);

    rec2.body.remaining_credits === 2
      ? pass('remaining_credits still = 2  (no credit deducted for cache hit)')
      : fail('remaining_credits still = 2', `Got: ${rec2.body.remaining_credits}`);

    rec2.body.output === rec1.body.output
      ? pass('output matches first response  (served from cache)')
      : fail('output matches first response', 'Outputs differ — possible cache miss or data mismatch');

    // ── Step 5: error cases ────────────────────────────────────────────────
    section('Step 5 — Error cases');

    const badKey = await post('/api/validate-key', { private_key: 'invalid-key-xyz' });
    badKey.body.valid === false && badKey.status === 200
      ? pass('validate-key with unknown key → { valid: false } HTTP 200')
      : fail('validate-key with unknown key', `Got status=${badKey.status} body=${JSON.stringify(badKey.body)}`);

    const noKey = await post('/api/get-recommendations', {
      private_key:   'invalid-key-xyz',
      template_id:   TEMPLATE_ID,
      user_data:     userData,
      system_prompt: SYSTEM_PROMPT,
    });
    noKey.status === 401
      ? pass('get-recommendations with invalid key → HTTP 401')
      : fail('get-recommendations with invalid key → HTTP 401', `Got: ${noKey.status}`);

  } catch (err) {
    console.error(`\n${C.red}${C.bold}Fatal error:${C.reset} ${err.message}`);
    if (err.cause?.code === 'ECONNREFUSED') {
      console.error(`Is the server running at ${BASE_URL}?  Run: node api/server.js`);
    }
    failed++;
  } finally {
    if (seeded) {
      await cleanup({ customerId, licenseId, walletId, requestHash });
    }
    printSummary();
    process.exit(failed > 0 ? 1 : 0);
  }
}

main();
