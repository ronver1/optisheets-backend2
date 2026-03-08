require('dotenv').config();
const express = require('express');
const supabase = require('../lib/supabase');
const { getOrFetch } = require('../lib/cache');

const app = express();
app.use(express.json());

// GET /api/templates - list all published templates
app.get('/api/templates', async (req, res) => {
  try {
    const templates = await getOrFetch('all_templates', async () => {
      const { data, error } = await supabase
        .from('templates')
        .select('id, slug, name, description, price_cents, category, preview_url')
        .eq('published', true)
        .order('created_at', { ascending: false });

      if (error) throw error;
      return data;
    }, 5 * 60_000); // cache 5 minutes

    res.json({ templates });
  } catch (err) {
    console.error('GET /api/templates error:', err.message);
    res.status(500).json({ error: 'Failed to fetch templates' });
  }
});

// GET /api/templates/:slug - get a single template by slug
app.get('/api/templates/:slug', async (req, res) => {
  const { slug } = req.params;
  try {
    const template = await getOrFetch(`template:${slug}`, async () => {
      const { data, error } = await supabase
        .from('templates')
        .select('*')
        .eq('slug', slug)
        .eq('published', true)
        .single();

      if (error) throw error;
      return data;
    }, 5 * 60_000);

    if (!template) return res.status(404).json({ error: 'Template not found' });
    res.json({ template });
  } catch (err) {
    console.error(`GET /api/templates/${slug} error:`, err.message);
    res.status(500).json({ error: 'Failed to fetch template' });
  }
});

module.exports = app;
