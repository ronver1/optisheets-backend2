require('dotenv').config();
const supabase = require('../lib/supabaseClient');

/**
 * POST /api/validate-key
 * Body: { private_key: string }
 * Returns: { valid: boolean, balance?: number, template_id?: string }
 */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { private_key } = req.body ?? {};

  if (!private_key || typeof private_key !== 'string') {
    return res.status(400).json({ error: 'private_key is required' });
  }

  // ------------------------------------------------------------------
  // Look up license by private key
  // ------------------------------------------------------------------
  const { data: license, error: licenseError } = await supabase
    .from('licenses')
    .select('id, template_id')
    .eq('private_key', private_key)
    .maybeSingle();

  if (licenseError) {
    console.error('validate-key: license lookup error:', licenseError.message);
    return res.status(500).json({ error: 'Internal server error' });
  }

  if (!license) {
    // Return 200 with valid:false — not an error, just an invalid key
    return res.status(200).json({ valid: false });
  }

  // ------------------------------------------------------------------
  // Fetch wallet balance
  // ------------------------------------------------------------------
  const { data: wallet, error: walletError } = await supabase
    .from('credit_wallets')
    .select('balance')
    .eq('license_id', license.id)
    .maybeSingle();

  if (walletError) {
    console.error('validate-key: wallet lookup error:', walletError.message);
    return res.status(500).json({ error: 'Internal server error' });
  }

  return res.status(200).json({
    valid: true,
    template_id: license.template_id,
    balance: wallet?.balance ?? 0,
  });
};
