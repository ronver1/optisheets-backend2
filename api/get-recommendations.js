require('dotenv').config();
const supabase = require('../lib/supabaseClient');
const { handleAIRequest } = require('../lib/aiRequestHandler');

// Known validation-layer error messages that map to 4xx responses rather than 500.
const CLIENT_ERRORS = new Set([
  'Invalid license key',
  'No credit wallet found for this license',
  'Insufficient AI Credits',
]);
const CLIENT_ERROR_PREFIX_REQUEST_TOO_LARGE = 'Request too large';

/**
 * Classify an aiRequestHandler error into an appropriate HTTP status code.
 * @param {string} message
 * @returns {number}
 */
function httpStatusForError(message) {
  if (message === 'Invalid license key') return 401;
  if (message === 'Insufficient AI Credits') return 402;
  if (CLIENT_ERRORS.has(message)) return 400;
  if (message.startsWith(CLIENT_ERROR_PREFIX_REQUEST_TOO_LARGE)) return 413;
  return 500;
}

/**
 * POST /api/get-recommendations
 * Body: { private_key, template_id, user_data, system_prompt }
 * Returns: { success: true, output, remaining_credits, cached }
 *        | { success: false, error }
 */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  const { private_key, template_id, user_data, system_prompt } = req.body ?? {};

  // ------------------------------------------------------------------
  // Input validation
  // ------------------------------------------------------------------
  if (!private_key || typeof private_key !== 'string') {
    return res.status(400).json({ success: false, error: 'private_key is required' });
  }
  if (!template_id || typeof template_id !== 'string') {
    return res.status(400).json({ success: false, error: 'template_id is required' });
  }
  if (!user_data || typeof user_data !== 'object' || Array.isArray(user_data)) {
    return res.status(400).json({ success: false, error: 'user_data must be a non-array object' });
  }
  if (!system_prompt || typeof system_prompt !== 'string') {
    return res.status(400).json({ success: false, error: 'system_prompt is required' });
  }

  // ------------------------------------------------------------------
  // Run the full AI credit flow
  // ------------------------------------------------------------------
  let output, cached;
  try {
    ({ output, cached } = await handleAIRequest({
      private_key,
      template_id,
      user_data,
      system_prompt,
    }));
  } catch (err) {
    const status = httpStatusForError(err.message);
    console.error(`get-recommendations [${status}]:`, err.message);
    return res.status(status).json({ success: false, error: err.message });
  }

  // ------------------------------------------------------------------
  // Fetch updated balance for the response
  // ------------------------------------------------------------------
  const { data: license } = await supabase
    .from('licenses')
    .select('id')
    .eq('private_key', private_key)
    .maybeSingle();

  let remaining_credits = null;
  if (license) {
    const { data: wallet } = await supabase
      .from('credit_wallets')
      .select('balance')
      .eq('license_id', license.id)
      .maybeSingle();
    remaining_credits = wallet?.balance ?? null;
  }

  return res.status(200).json({
    success: true,
    output,
    cached,
    remaining_credits,
  });
};
