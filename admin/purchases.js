require('dotenv').config();
const express = require('express');
const supabase = require('../lib/supabase');
const requireAdmin = require('./middleware/requireAdmin');

const app = express();
app.use(express.json());
app.use(requireAdmin);

// GET /admin/purchases - list all purchases with optional email filter
app.get('/admin/purchases', async (req, res) => {
  const { email } = req.query;

  try {
    let query = supabase
      .from('purchases')
      .select('*, templates(name, slug)')
      .order('purchased_at', { ascending: false });

    if (email) query = query.eq('user_email', email);

    const { data, error } = await query;
    if (error) throw error;
    res.json({ purchases: data });
  } catch (err) {
    console.error('GET /admin/purchases error:', err.message);
    res.status(500).json({ error: 'Failed to fetch purchases' });
  }
});

module.exports = app;
