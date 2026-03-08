require('dotenv').config();
const express = require('express');
const openai = require('../lib/openai');
const supabase = require('../lib/supabase');

const app = express();
app.use(express.json());

// POST /api/generate - generate AI-powered content for a template
app.post('/api/generate', async (req, res) => {
  const { template_slug, user_email, prompt_inputs } = req.body;

  if (!template_slug || !user_email || !prompt_inputs) {
    return res.status(400).json({ error: 'template_slug, user_email, and prompt_inputs are required' });
  }

  try {
    // Verify user has purchased this template
    const { data: purchase, error: purchaseError } = await supabase
      .from('purchases')
      .select('id')
      .eq('user_email', user_email)
      .eq('template_id', (
        await supabase
          .from('templates')
          .select('id')
          .eq('slug', template_slug)
          .single()
      ).data?.id)
      .maybeSingle();

    if (purchaseError) throw purchaseError;
    if (!purchase) {
      return res.status(403).json({ error: 'No valid purchase found for this template' });
    }

    // Load template-specific system prompt
    let systemPrompt;
    try {
      systemPrompt = require(`../templates/${template_slug}/prompt.js`);
    } catch {
      return res.status(404).json({ error: `No prompt config found for template: ${template_slug}` });
    }

    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: JSON.stringify(prompt_inputs) },
      ],
      temperature: 0.7,
    });

    const result = completion.choices[0].message.content;
    res.json({ result });
  } catch (err) {
    console.error('POST /api/generate error:', err.message);
    res.status(500).json({ error: 'AI generation failed' });
  }
});

module.exports = app;
