require('dotenv').config();
const supabase = require('../../lib/supabaseClient');

// ------------------------------------------------------------------
// CSV helpers
// ------------------------------------------------------------------

/**
 * Escape a single CSV cell value.
 * Wraps in double-quotes and escapes any internal double-quotes.
 */
function csvCell(value) {
  if (value === null || value === undefined) return '';
  const str = String(value);
  // If the value contains a comma, newline, or double-quote it must be quoted.
  if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

/** Convert an array of objects to a CSV string. */
function toCsv(rows, columns) {
  const header = columns.join(',');
  const body = rows
    .map(row => columns.map(col => csvCell(row[col])).join(','))
    .join('\n');
  return `${header}\n${body}`;
}

// ------------------------------------------------------------------
// Handler
// ------------------------------------------------------------------

/**
 * GET /api/admin/export-csv
 * Headers: Authorization: Bearer <ADMIN_SECRET>
 * Query params:
 *   ?table=customers|licenses|wallets|ledger|full (default: full)
 * Returns a downloadable CSV file.
 */
module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // ------------------------------------------------------------------
  // Admin auth
  // ------------------------------------------------------------------
  const adminSecret = process.env.ADMIN_SECRET;
  if (!adminSecret) {
    console.error('export-csv: ADMIN_SECRET is not set');
    return res.status(500).json({ error: 'Server misconfiguration' });
  }

  const authHeader = req.headers['authorization'] ?? '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (!token || token !== adminSecret) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  // ------------------------------------------------------------------
  // Fetch all data in parallel
  // ------------------------------------------------------------------
  const [
    { data: customers, error: custErr },
    { data: licenses,  error: licErr  },
    { data: wallets,   error: walErr  },
    { data: ledger,    error: ledErr  },
  ] = await Promise.all([
    supabase
      .from('customers')
      .select('id, username, email, created_at')
      .order('created_at', { ascending: true }),

    supabase
      .from('licenses')
      .select('id, customer_id, template_id, private_key, created_at')
      .order('created_at', { ascending: true }),

    supabase
      .from('credit_wallets')
      .select('id, license_id, balance'),

    supabase
      .from('ledger')
      .select('id, license_id, event_type, amount, note, created_at')
      .order('created_at', { ascending: true }),
  ]);

  const errors = [
    custErr && `customers: ${custErr.message}`,
    licErr  && `licenses: ${licErr.message}`,
    walErr  && `wallets: ${walErr.message}`,
    ledErr  && `ledger: ${ledErr.message}`,
  ].filter(Boolean);

  if (errors.length) {
    console.error('export-csv: query errors:', errors);
    return res.status(500).json({ error: 'Data fetch failed', details: errors });
  }

  // ------------------------------------------------------------------
  // Build a denormalized flat view joining all four tables
  // ------------------------------------------------------------------

  // Index wallets by license_id for O(1) lookup
  const walletByLicenseId = Object.fromEntries(
    (wallets ?? []).map(w => [w.license_id, w])
  );

  // Index customers by id
  const customerById = Object.fromEntries(
    (customers ?? []).map(c => [c.id, c])
  );

  // Build one row per license
  const licenseRows = (licenses ?? []).map(lic => {
    const customer = customerById[lic.customer_id] ?? {};
    const wallet   = walletByLicenseId[lic.id] ?? {};
    return {
      license_id:        lic.id,
      license_created_at: lic.created_at,
      template_id:       lic.template_id,
      private_key:       lic.private_key,
      customer_id:       lic.customer_id,
      customer_username: customer.username ?? '',
      customer_email:    customer.email    ?? '',
      customer_created_at: customer.created_at ?? '',
      wallet_id:         wallet.id      ?? '',
      credit_balance:    wallet.balance  ?? 0,
    };
  });

  // Build ledger rows (flat, already have all columns needed)
  const ledgerRows = (ledger ?? []).map(entry => ({
    ledger_id:   entry.id,
    license_id:  entry.license_id,
    event_type:  entry.event_type,
    amount:      entry.amount,
    note:        entry.note ?? '',
    created_at:  entry.created_at,
  }));

  // ------------------------------------------------------------------
  // Determine which table to export
  // ------------------------------------------------------------------
  const table = req.query?.table ?? 'full';

  let csvContent, filename;

  switch (table) {
    case 'customers':
      csvContent = toCsv(customers ?? [], ['id', 'username', 'email', 'created_at']);
      filename   = 'customers.csv';
      break;

    case 'licenses':
      csvContent = toCsv(licenses ?? [], ['id', 'customer_id', 'template_id', 'private_key', 'created_at']);
      filename   = 'licenses.csv';
      break;

    case 'wallets':
      csvContent = toCsv(wallets ?? [], ['id', 'license_id', 'balance']);
      filename   = 'wallets.csv';
      break;

    case 'ledger':
      csvContent = toCsv(ledgerRows, ['ledger_id', 'license_id', 'event_type', 'amount', 'note', 'created_at']);
      filename   = 'ledger.csv';
      break;

    case 'full':
    default: {
      // Two sections separated by a blank line in one file
      const licensesCsv = toCsv(licenseRows, [
        'license_id', 'license_created_at', 'template_id', 'private_key',
        'customer_id', 'customer_username', 'customer_email', 'customer_created_at',
        'wallet_id', 'credit_balance',
      ]);
      const ledgerCsv = toCsv(ledgerRows, [
        'ledger_id', 'license_id', 'event_type', 'amount', 'note', 'created_at',
      ]);
      csvContent = `# LICENSES + CUSTOMERS + WALLETS\n${licensesCsv}\n\n# LEDGER\n${ledgerCsv}`;
      filename   = 'optisheets-export.csv';
      break;
    }
  }

  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  return res.status(200).send(csvContent);
};
