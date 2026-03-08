require('dotenv').config();
const { v4: uuidv4 } = require('uuid');
const supabase = require('../lib/supabaseClient');

const MAX_CREDIT_ADD = 10_000; // sanity ceiling per single request

/**
 * POST /api/add-credits
 * Admin/webhook use only — requires Authorization: Bearer <ADMIN_SECRET>
 * Body: { private_key: string, amount: number, note?: string }
 * Returns: { success: true, new_balance: number }
 *        | { success: false, error: string }
 */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  // ------------------------------------------------------------------
  // Admin auth — Bearer token must match ADMIN_SECRET
  // ------------------------------------------------------------------
  const adminSecret = process.env.ADMIN_SECRET;
  if (!adminSecret) {
    console.error('add-credits: ADMIN_SECRET is not set');
    return res.status(500).json({ success: false, error: 'Server misconfiguration' });
  }

  const authHeader = req.headers['authorization'] ?? '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (!token || token !== adminSecret) {
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }

  // ------------------------------------------------------------------
  // Input validation
  // ------------------------------------------------------------------
  const { private_key, amount, note } = req.body ?? {};

  if (!private_key || typeof private_key !== 'string') {
    return res.status(400).json({ success: false, error: 'private_key is required' });
  }
  if (amount === undefined || amount === null) {
    return res.status(400).json({ success: false, error: 'amount is required' });
  }

  const parsedAmount = Number(amount);
  if (!Number.isInteger(parsedAmount) || parsedAmount < 1) {
    return res.status(400).json({ success: false, error: 'amount must be a positive integer' });
  }
  if (parsedAmount > MAX_CREDIT_ADD) {
    return res.status(400).json({
      success: false,
      error: `amount exceeds the maximum allowed per request (${MAX_CREDIT_ADD})`,
    });
  }

  // ------------------------------------------------------------------
  // Resolve license → wallet
  // ------------------------------------------------------------------
  const { data: license, error: licenseError } = await supabase
    .from('licenses')
    .select('id')
    .eq('private_key', private_key)
    .maybeSingle();

  if (licenseError) {
    console.error('add-credits: license lookup error:', licenseError.message);
    return res.status(500).json({ success: false, error: 'Internal server error' });
  }
  if (!license) {
    return res.status(404).json({ success: false, error: 'Invalid license key' });
  }

  const { data: wallet, error: walletError } = await supabase
    .from('credit_wallets')
    .select('id, balance')
    .eq('license_id', license.id)
    .maybeSingle();

  if (walletError) {
    console.error('add-credits: wallet lookup error:', walletError.message);
    return res.status(500).json({ success: false, error: 'Internal server error' });
  }
  if (!wallet) {
    return res.status(404).json({ success: false, error: 'No credit wallet found for this license' });
  }

  // ------------------------------------------------------------------
  // Increment balance
  // ------------------------------------------------------------------
  const newBalance = wallet.balance + parsedAmount;

  const { error: updateError } = await supabase
    .from('credit_wallets')
    .update({ balance: newBalance })
    .eq('id', wallet.id);

  if (updateError) {
    console.error('add-credits: balance update error:', updateError.message);
    return res.status(500).json({ success: false, error: 'Failed to update credit balance' });
  }

  // ------------------------------------------------------------------
  // Write 'purchase' ledger entry
  // ------------------------------------------------------------------
  const { error: ledgerError } = await supabase.from('ledger').insert({
    id: uuidv4(),
    license_id: license.id,
    event_type: 'purchase',
    amount: parsedAmount,
    note: note?.trim() || `${parsedAmount} credit(s) added via API`,
  });

  if (ledgerError) {
    // Balance was already updated — log loudly but don't fail the response.
    // A missing ledger entry is an audit gap, not a functional failure.
    console.error('add-credits: ledger insert error (balance already updated):', ledgerError.message);
  }

  return res.status(200).json({ success: true, new_balance: newBalance });
};
