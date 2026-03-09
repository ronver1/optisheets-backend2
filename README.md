# OptiSheets Backend

Node.js/Express backend for OptiSheets — AI-powered Google Sheets templates for college students. Deployed as Vercel serverless functions.

## Stack

- **Runtime**: Node.js (CommonJS)
- **Framework**: Express (one app instance per serverless function)
- **Database**: Supabase (Postgres + service role client)
- **AI**: OpenAI `gpt-4o-mini`
- **Email**: Nodemailer (Gmail SMTP)
- **Deployment**: Vercel

## Project Structure

```
api/
  index.js                  # GET  /api/             — health check
  validate-key.js           # POST /api/validate-key — check license key + balance
  get-recommendations.js    # POST /api/get-recommendations — AI credit flow
  add-credits.js            # POST /api/add-credits  — add credits (admin)
  webhook/
    etsy-purchase.js        # POST /api/webhook/etsy-purchase — Zapier/Etsy webhook
admin/
  middleware/requireAdmin.js
  templates.js              # CRUD /admin/templates
  purchases.js              # GET  /admin/purchases
  export-csv.js             # GET  /admin/export-csv
config/
  templates.js              # TEMPLATES and CREDIT_PACKS config objects
lib/
  supabaseClient.js         # Supabase singleton (service role)
  openai.js                 # OpenAI singleton
  aiRequestHandler.js       # AI credit flow logic
  emailService.js           # Transactional email functions
  keyGenerator.js           # License key generation + validation
  cache.js                  # In-memory cache helper
templates/
  <slug>/
    prompt.js               # System prompt for OpenAI
    meta.json               # Template metadata
```

## Environment Variables

Copy `.env.example` to `.env` and fill in each value before running locally. All variables are also required in your Vercel project settings for production.

### Supabase

| Variable | Description |
|---|---|
| `SUPABASE_URL` | Your Supabase project URL (e.g. `https://xyz.supabase.co`) |
| `SUPABASE_SERVICE_ROLE_KEY` | Service role key — bypasses RLS. **Never expose to the browser.** |

### OpenAI

| Variable | Description |
|---|---|
| `OPENAI_API_KEY` | OpenAI API key used for AI recommendations (`gpt-4o-mini`) |

### Admin

| Variable | Description |
|---|---|
| `ADMIN_SECRET` | Bearer token required for all `/admin/*` routes and `POST /api/add-credits` |

### Webhook

| Variable | Description |
|---|---|
| `WEBHOOK_SECRET` | Shared secret sent by Zapier in the `x-webhook-secret` header. Set the same value in both Vercel and your Zapier webhook action. Generate with `openssl rand -hex 32`. |

### Email (Nodemailer / Gmail SMTP)

| Variable | Description |
|---|---|
| `EMAIL_HOST` | SMTP hostname (e.g. `smtp.gmail.com`) |
| `EMAIL_PORT` | SMTP port — `587` for STARTTLS, `465` for SSL |
| `EMAIL_USER` | SMTP login username (your Gmail address) |
| `EMAIL_PASS` | SMTP password — for Gmail, use an **App Password**, not your account password. Generate one at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) (requires 2FA). |
| `EMAIL_FROM` | The `From` header on outgoing emails (e.g. `OptiSheets <you@gmail.com>`) |

## Local Development

```bash
npm install
cp .env.example .env   # then fill in your values
node server.js
```

The local Express server mounts all API handlers and listens on `http://localhost:3000`.

## Deployment

Push to `main` — Vercel deploys automatically. Each file in `api/` becomes an independent serverless function. Set all environment variables in the Vercel dashboard under **Settings → Environment Variables**.

## Etsy → Zapier Webhook Flow

1. A customer places an order on Etsy.
2. Zapier triggers on the new order and sends a `POST` to `/api/webhook/etsy-purchase` with:
   ```json
   {
     "buyer_email": "customer@example.com",
     "buyer_name": "Jane Smith",
     "listing_title": "Internship Tracker AI — Google Sheets Template",
     "order_id": "1234567890",
     "private_key": "OS-INT-XXXX-XXXX-XXXX"
   }
   ```
   `private_key` is optional — only needed for credit pack purchases targeting a specific license.
3. The webhook identifies the purchase type from `listing_title` keywords defined in `config/templates.js`, provisions the customer/license/wallet in Supabase, and sends a transactional email.
4. Returns `{ "success": true, "type": "ai_template" | "standard_template" | "credits" }`.

> The `x-webhook-secret` header must match `WEBHOOK_SECRET` or the request is rejected with 401.
