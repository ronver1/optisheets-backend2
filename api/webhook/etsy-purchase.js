'use strict';

require('dotenv').config();

const supabase = require('../../lib/supabaseClient');
const { generatePrivateKey } = require('../../lib/keyGenerator');
const {
  sendAITemplatePurchaseEmail,
  sendStandardTemplatePurchaseEmail,
  sendCreditPurchaseEmail,
} = require('../../lib/emailService');
const { TEMPLATES, CREDIT_PACKS } = require('../../config/templates');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function matchListing(listingTitle) {
  const lower = listingTitle.toLowerCase();

  for (const template of Object.values(TEMPLATES)) {
    if (template.etsy_title_keywords.some((kw) => lower.includes(kw))) {
      return { type: template.isAI ? 'ai_template' : 'standard_template', item: template };
    }
  }

  for (const pack of Object.values(CREDIT_PACKS)) {
    if (pack.etsy_title_keywords.some((kw) => lower.includes(kw))) {
      return { type: 'credits', item: pack };
    }
  }

  return null;
}

/** Returns existing customer id or creates a new row and returns the id. */
async function upsertCustomer(email, name) {
  const { data: existing, error: lookupError } = await supabase
    .from('customers')
    .select('id')
    .eq('email', email)
    .maybeSingle();

  if (lookupError) throw new Error(`Customer lookup failed: ${lookupError.message}`);

  if (existing) return existing.id;

  const { data: created, error: createError } = await supabase
    .from('customers')
    .insert({ email, username: name })
    .select('id')
    .single();

  if (createError) throw new Error(`Customer creation failed: ${createError.message}`);
  return created.id;
}

// ---------------------------------------------------------------------------
// Purchase handlers
// ---------------------------------------------------------------------------

async function handleAITemplate({ buyerEmail, buyerName, orderId, template }) {
  const customerId = await upsertCustomer(buyerEmail, buyerName);
  const privateKey = generatePrivateKey(template.id);

  const { data: license, error: licenseError } = await supabase
    .from('licenses')
    .insert({ customer_id: customerId, template_id: template.id, private_key: privateKey })
    .select('id')
    .single();

  if (licenseError) throw new Error(`License insert failed: ${licenseError.message}`);

  const { error: walletError } = await supabase
    .from('credit_wallets')
    .insert({ license_id: license.id, balance: 0 });

  if (walletError) throw new Error(`Wallet insert failed: ${walletError.message}`);

  const { error: ledgerError } = await supabase
    .from('ledger')
    .insert({
      license_id: license.id,
      event_type: 'purchase',
      amount: 0,
      note: `New license created - order ${orderId}`,
    });

  if (ledgerError) {
    // Non-fatal — log and continue
    console.error(`[etsy-webhook] Ledger insert failed (non-fatal): ${ledgerError.message}`);
  }

  await sendAITemplatePurchaseEmail({
    to: buyerEmail,
    customerName: buyerName,
    templateName: template.name,
    privateKey,
    sheetUrl: template.sheetUrl,
    pdfUrl: template.pdfUrl,
  });

  console.log(`[etsy-webhook] AI template purchase complete — order ${orderId}, template ${template.id}, key ${privateKey}`);
}

async function handleStandardTemplate({ buyerEmail, buyerName, orderId, template }) {
  const customerId = await upsertCustomer(buyerEmail, buyerName);

  const { data: license, error: licenseError } = await supabase
    .from('licenses')
    .insert({ customer_id: customerId, template_id: template.id, private_key: null })
    .select('id')
    .single();

  if (licenseError) throw new Error(`License insert failed: ${licenseError.message}`);

  const { error: ledgerError } = await supabase
    .from('ledger')
    .insert({
      license_id: license.id,
      event_type: 'purchase',
      amount: 0,
      note: `Standard template purchase - order ${orderId}`,
    });

  if (ledgerError) {
    console.error(`[etsy-webhook] Ledger insert failed (non-fatal): ${ledgerError.message}`);
  }

  await sendStandardTemplatePurchaseEmail({
    to: buyerEmail,
    customerName: buyerName,
    templateName: template.name,
    sheetUrl: template.sheetUrl,
    pdfUrl: template.pdfUrl,
  });

  console.log(`[etsy-webhook] Standard template purchase complete — order ${orderId}, template ${template.id}`);
}

async function handleCreditPack({ buyerEmail, buyerName, orderId, pack, privateKey: targetKey }) {
  // Find the customer
  const { data: customer, error: customerError } = await supabase
    .from('customers')
    .select('id')
    .eq('email', buyerEmail)
    .maybeSingle();

  if (customerError) throw new Error(`Customer lookup failed: ${customerError.message}`);
  if (!customer) throw new Error(`No customer found for email ${buyerEmail}`);

  // Resolve the target license
  let licenseId;

  if (targetKey) {
    const { data: license, error: licenseError } = await supabase
      .from('licenses')
      .select('id')
      .eq('private_key', targetKey)
      .eq('customer_id', customer.id)
      .maybeSingle();

    if (licenseError) throw new Error(`License lookup failed: ${licenseError.message}`);
    if (!license) throw new Error(`No license found for provided private_key`);
    licenseId = license.id;
  } else {
    // Fall back to their most recent license
    const { data: license, error: licenseError } = await supabase
      .from('licenses')
      .select('id')
      .eq('customer_id', customer.id)
      .order('created_at', { ascending: false })
      .limit(1)
      .maybeSingle();

    if (licenseError) throw new Error(`License lookup failed: ${licenseError.message}`);
    if (!license) throw new Error(`No license found for customer ${buyerEmail}`);
    licenseId = license.id;
  }

  // Fetch current balance
  const { data: wallet, error: walletError } = await supabase
    .from('credit_wallets')
    .select('id, balance')
    .eq('license_id', licenseId)
    .maybeSingle();

  if (walletError) throw new Error(`Wallet lookup failed: ${walletError.message}`);
  if (!wallet) throw new Error(`No wallet found for license ${licenseId}`);

  const newBalance = wallet.balance + pack.amount;

  const { error: updateError } = await supabase
    .from('credit_wallets')
    .update({ balance: newBalance })
    .eq('id', wallet.id);

  if (updateError) throw new Error(`Wallet update failed: ${updateError.message}`);

  const { error: ledgerError } = await supabase
    .from('ledger')
    .insert({
      license_id: licenseId,
      event_type: 'purchase',
      amount: pack.amount,
      note: `Credit pack purchase - order ${orderId}`,
    });

  if (ledgerError) {
    console.error(`[etsy-webhook] Ledger insert failed (non-fatal): ${ledgerError.message}`);
  }

  await sendCreditPurchaseEmail({
    to: buyerEmail,
    customerName: buyerName,
    amount: pack.amount,
    newBalance,
  });

  console.log(`[etsy-webhook] Credit pack purchase complete — order ${orderId}, +${pack.amount} credits, new balance ${newBalance}`);
}

// ---------------------------------------------------------------------------
// Handler
// ---------------------------------------------------------------------------

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // 1. Verify webhook secret
  if (req.headers['x-webhook-secret'] !== process.env.WEBHOOK_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const { buyer_email, buyer_name, listing_title, order_id, private_key } = req.body ?? {};

  if (!buyer_email || !buyer_name || !listing_title || !order_id) {
    return res.status(400).json({ error: 'Missing required fields: buyer_email, buyer_name, listing_title, order_id' });
  }

  console.log(`[etsy-webhook] Received order ${order_id} — "${listing_title}" for ${buyer_email}`);

  // 2. Identify purchase type
  const match = matchListing(listing_title);

  if (!match) {
    console.log(`[etsy-webhook] Unrecognized listing title: "${listing_title}" (order ${order_id})`);
    return res.status(200).json({ success: false, reason: 'unrecognized listing' });
  }

  // 3. Dispatch to the appropriate handler
  try {
    if (match.type === 'ai_template') {
      await handleAITemplate({ buyerEmail: buyer_email, buyerName: buyer_name, orderId: order_id, template: match.item });
    } else if (match.type === 'standard_template') {
      await handleStandardTemplate({ buyerEmail: buyer_email, buyerName: buyer_name, orderId: order_id, template: match.item });
    } else {
      await handleCreditPack({ buyerEmail: buyer_email, buyerName: buyer_name, orderId: order_id, pack: match.item, privateKey: private_key ?? null });
    }

    return res.status(200).json({ success: true, type: match.type });
  } catch (err) {
    console.error(`[etsy-webhook] Error processing order ${order_id}:`, err.message);
    return res.status(500).json({ success: false, error: err.message });
  }
};
