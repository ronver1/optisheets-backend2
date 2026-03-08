require('dotenv').config();
const express = require('express');
const supabase = require('../lib/supabase');
const requireAdmin = require('./middleware/requireAdmin');

const app = express();
app.use(express.json());
app.use(requireAdmin);

// GET /admin/templates - list all templates (including unpublished)
app.get('/admin/templates', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('templates')
      .select('*')
      .order('created_at', { ascending: false });

    if (error) throw error;
    res.json({ templates: data });
  } catch (err) {
    console.error('GET /admin/templates error:', err.message);
    res.status(500).json({ error: 'Failed to fetch templates' });
  }
});

// POST /admin/templates - create a new template
app.post('/admin/templates', async (req, res) => {
  const { slug, name, description, price_cents, category, preview_url, download_url, published } = req.body;

  if (!slug || !name || price_cents === undefined) {
    return res.status(400).json({ error: 'slug, name, and price_cents are required' });
  }

  try {
    const { data, error } = await supabase
      .from('templates')
      .insert({ slug, name, description, price_cents, category, preview_url, download_url, published: published ?? false })
      .select()
      .single();

    if (error) throw error;
    res.status(201).json({ template: data });
  } catch (err) {
    console.error('POST /admin/templates error:', err.message);
    res.status(500).json({ error: 'Failed to create template' });
  }
});

// PATCH /admin/templates/:id - update a template
app.patch('/admin/templates/:id', async (req, res) => {
  const { id } = req.params;
  const updates = req.body;

  try {
    const { data, error } = await supabase
      .from('templates')
      .update({ ...updates, updated_at: new Date().toISOString() })
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;
    res.json({ template: data });
  } catch (err) {
    console.error(`PATCH /admin/templates/${id} error:`, err.message);
    res.status(500).json({ error: 'Failed to update template' });
  }
});

// DELETE /admin/templates/:id - delete a template
app.delete('/admin/templates/:id', async (req, res) => {
  const { id } = req.params;

  try {
    const { error } = await supabase.from('templates').delete().eq('id', id);
    if (error) throw error;
    res.json({ success: true });
  } catch (err) {
    console.error(`DELETE /admin/templates/${id} error:`, err.message);
    res.status(500).json({ error: 'Failed to delete template' });
  }
});

module.exports = app;
