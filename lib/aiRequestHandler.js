require('dotenv').config();
const crypto = require('crypto');
const { v4: uuidv4 } = require('uuid');
const OpenAI = require('openai');
const supabase = require('./supabaseClient');

// ---------------------------------------------------------------------------
// OpenAI client — instantiated lazily so missing key only throws at call time
// ---------------------------------------------------------------------------
function getOpenAIClient() {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('Missing OPENAI_API_KEY environment variable');
  return new OpenAI({ apiKey });
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const MAX_TOKENS = 500;

// GPT-4o-mini pricing (as of 2024): $0.00015 / 1K input tokens
// Using a conservative ceiling to stay safely under the $1.00 guard.
const COST_PER_1K_INPUT_TOKENS = 0.00015;
const MAX_ALLOWED_COST_USD = 1.00;

// Rough character-to-token ratio (OpenAI ~4 chars per token)
const CHARS_PER_TOKEN = 4;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Compute SHA-256 hash of the canonical request payload.
 * Sorting user_data keys ensures identical inputs always produce the same hash.
 */
function computeRequestHash(templateId, userData) {
  const canonical = JSON.stringify({
    template_id: templateId,
    user_data: sortObjectKeys(userData),
  });
  return crypto.createHash('sha256').update(canonical).digest('hex');
}

function sortObjectKeys(obj) {
  if (typeof obj !== 'object' || obj === null || Array.isArray(obj)) return obj;
  return Object.keys(obj)
    .sort()
    .reduce((acc, k) => {
      acc[k] = sortObjectKeys(obj[k]);
      return acc;
    }, {});
}

/**
 * Estimate the USD cost of the input payload before hitting OpenAI.
 * Conservative: counts prompt chars only (system + user), ignores output cost.
 */
function estimateInputCostUSD(systemPrompt, userContent) {
  const totalChars = (systemPrompt?.length ?? 0) + (userContent?.length ?? 0);
  const estimatedTokens = Math.ceil(totalChars / CHARS_PER_TOKEN);
  return (estimatedTokens / 1000) * COST_PER_1K_INPUT_TOKENS;
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

/**
 * Execute the full AI credit flow.
 *
 * @param {object} params
 * @param {string} params.private_key   - License private key from the request
 * @param {string} params.template_id   - Template slug / ID
 * @param {object} params.user_data     - User-supplied inputs for the prompt
 * @param {string} params.system_prompt - System prompt string for the template
 * @returns {Promise<{ output: string, cached: boolean }>}
 * @throws {Error} with a human-readable `.message` on any validation failure
 */
async function handleAIRequest({ private_key, template_id, user_data, system_prompt }) {
  // -------------------------------------------------------------------------
  // Step 1 — Validate private_key against licenses table
  // -------------------------------------------------------------------------
  const { data: license, error: licenseError } = await supabase
    .from('licenses')
    .select('id, template_id')
    .eq('private_key', private_key)
    .maybeSingle();

  if (licenseError) throw new Error(`License lookup failed: ${licenseError.message}`);
  if (!license) throw new Error('Invalid license key');

  // -------------------------------------------------------------------------
  // Step 2 — Compute SHA-256 hash of the canonical payload
  // -------------------------------------------------------------------------
  const requestHash = computeRequestHash(template_id, user_data);

  // -------------------------------------------------------------------------
  // Step 3 — Check ai_cache for a cache hit
  // -------------------------------------------------------------------------
  const { data: cached, error: cacheError } = await supabase
    .from('ai_cache')
    .select('output')
    .eq('request_hash', requestHash)
    .maybeSingle();

  if (cacheError) throw new Error(`Cache lookup failed: ${cacheError.message}`);

  if (cached) {
    return { output: cached.output, cached: true };
  }

  // -------------------------------------------------------------------------
  // Step 4 — Check credit balance
  // -------------------------------------------------------------------------
  const { data: wallet, error: walletError } = await supabase
    .from('credit_wallets')
    .select('id, balance')
    .eq('license_id', license.id)
    .maybeSingle();

  if (walletError) throw new Error(`Wallet lookup failed: ${walletError.message}`);
  if (!wallet) throw new Error('No credit wallet found for this license');
  if (wallet.balance < 1) throw new Error('Insufficient AI Credits');

  // -------------------------------------------------------------------------
  // Step 5 — Estimate token cost; reject oversized requests
  // -------------------------------------------------------------------------
  const userContent = JSON.stringify(user_data);
  const estimatedCost = estimateInputCostUSD(system_prompt, userContent);

  if (estimatedCost > MAX_ALLOWED_COST_USD) {
    throw new Error(
      `Request too large (estimated cost $${estimatedCost.toFixed(4)} exceeds $${MAX_ALLOWED_COST_USD.toFixed(2)} limit)`
    );
  }

  // -------------------------------------------------------------------------
  // Step 6 — Reserve 1 credit (decrement balance + write 'usage' ledger entry)
  // -------------------------------------------------------------------------
  const ledgerEntryId = uuidv4();

  const { error: debitError } = await supabase
    .from('credit_wallets')
    .update({ balance: wallet.balance - 1 })
    .eq('id', wallet.id);

  if (debitError) throw new Error(`Credit reservation failed: ${debitError.message}`);

  const { error: usageLedgerError } = await supabase.from('ledger').insert({
    id: ledgerEntryId,
    license_id: license.id,
    event_type: 'usage',
    amount: -1,
    note: `AI request for template "${template_id}" (hash: ${requestHash.slice(0, 12)}…)`,
  });

  if (usageLedgerError) {
    // Best-effort rollback of the balance decrement before throwing
    await supabase
      .from('credit_wallets')
      .update({ balance: wallet.balance })
      .eq('id', wallet.id);
    throw new Error(`Ledger write failed: ${usageLedgerError.message}`);
  }

  // -------------------------------------------------------------------------
  // Step 7 — Call OpenAI
  // -------------------------------------------------------------------------
  let aiOutput;
  try {
    const openai = getOpenAIClient();
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      max_tokens: MAX_TOKENS,
      temperature: 0.7,
      messages: [
        { role: 'system', content: system_prompt },
        { role: 'user', content: userContent },
      ],
    });

    aiOutput = completion.choices[0]?.message?.content ?? '';
    if (!aiOutput) throw new Error('OpenAI returned an empty response');
  } catch (openAIError) {
    // -----------------------------------------------------------------------
    // Step 9 — OpenAI failed: refund the credit
    // -----------------------------------------------------------------------
    await supabase
      .from('credit_wallets')
      .update({ balance: wallet.balance }) // restore original balance
      .eq('id', wallet.id);

    await supabase.from('ledger').insert({
      id: uuidv4(),
      license_id: license.id,
      event_type: 'refund',
      amount: 1,
      note: `Refund for failed AI request (usage ledger: ${ledgerEntryId}): ${openAIError.message}`,
    });

    throw new Error(`AI generation failed: ${openAIError.message}`);
  }

  // -------------------------------------------------------------------------
  // Step 8 — Success: persist to ai_cache
  // -------------------------------------------------------------------------
  const { error: cacheInsertError } = await supabase.from('ai_cache').insert({
    id: uuidv4(),
    request_hash: requestHash,
    template_id,
    output: aiOutput,
  });

  // Cache write failure is non-fatal — log it but don't fail the request
  if (cacheInsertError) {
    console.warn('ai_cache insert failed (non-fatal):', cacheInsertError.message);
  }

  return { output: aiOutput, cached: false };
}

module.exports = { handleAIRequest };
