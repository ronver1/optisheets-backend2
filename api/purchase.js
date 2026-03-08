require('dotenv').config();
const express = require('express');
const { v4: uuidv4 } = require('uuid');
const supabase = require('../lib/supabase');

const app = express();
app.use(express.json());

// POST /api/purchase - record a completed purchase and return download link
app.post('/api/purchase', async (req, res) => {
  const { template_slug, user_email, payment_reference } = req.body;

  if (!template_slug || !user_email || !payment_reference) {
    return res.status(400).json({ error: 'template_slug, user_email, and payment_reference are required' });
  }

  try {
    // Fetch template
    const { data: template, error: tplError } = await supabase
      .from('templates')
      .select('id, name, download_url')
      .eq('slug', template_slug)
      .eq('published', true)
      .single();

    if (tplError || !template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    // Record purchase
    const purchaseId = uuidv4();
    const { error: insertError } = await supabase.from('purchases').insert({
      id: purchaseId,
      template_id: template.id,
      user_email,
      payment_reference,
      purchased_at: new Date().toISOString(),
    });

    if (insertError) throw insertError;

    res.json({
      purchase_id: purchaseId,
      template_name: template.name,
      download_url: template.download_url,
    });
  } catch (err) {
    console.error('POST /api/purchase error:', err.message);
    res.status(500).json({ error: 'Purchase recording failed' });
  }
});

module.exports = app;
