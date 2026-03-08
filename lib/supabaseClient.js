require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');

const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!supabaseUrl || !supabaseKey) {
  throw new Error(
    'Missing required environment variables: SUPABASE_URL and/or SUPABASE_SERVICE_ROLE_KEY'
  );
}

/**
 * Supabase client initialized with the service role key.
 * This bypasses Row Level Security and must only be used server-side.
 * Never expose this client or its key to the browser.
 */
const supabase = createClient(supabaseUrl, supabaseKey, {
  auth: {
    // Disable automatic session persistence — not needed in a serverless context.
    persistSession: false,
    autoRefreshToken: false,
  },
});

module.exports = supabase;
