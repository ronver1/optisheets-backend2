'use strict';

const nodemailer = require('nodemailer');

const SUPPORT_EMAIL = 'support@optisheets.com';

function createTransporter() {
  return nodemailer.createTransport({
    host: process.env.EMAIL_HOST,
    port: Number(process.env.EMAIL_PORT) || 587,
    secure: Number(process.env.EMAIL_PORT) === 465,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });
}

const baseStyles = `
  body { margin: 0; padding: 0; background: #f4f4f5; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #18181b; }
  .wrapper { padding: 40px 16px; }
  .card { background: #ffffff; border-radius: 8px; max-width: 560px; margin: 0 auto; padding: 40px 36px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
  .logo { font-size: 20px; font-weight: 700; color: #16a34a; margin-bottom: 28px; }
  h1 { font-size: 22px; font-weight: 700; margin: 0 0 8px; color: #18181b; }
  p { font-size: 15px; line-height: 1.6; margin: 0 0 16px; color: #3f3f46; }
  .key-box { background: #f0fdf4; border: 1.5px solid #86efac; border-radius: 6px; padding: 18px 20px; margin: 24px 0; text-align: center; }
  .key-label { font-size: 12px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: #16a34a; margin-bottom: 8px; }
  .key-value { font-family: 'Courier New', Courier, monospace; font-size: 18px; font-weight: 700; color: #14532d; letter-spacing: 0.04em; word-break: break-all; }
  .balance-box { background: #eff6ff; border: 1.5px solid #93c5fd; border-radius: 6px; padding: 18px 20px; margin: 24px 0; text-align: center; }
  .balance-label { font-size: 12px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: #1d4ed8; margin-bottom: 8px; }
  .balance-value { font-size: 28px; font-weight: 700; color: #1e3a8a; }
  .steps { margin: 24px 0; padding: 0; }
  .step { display: flex; align-items: flex-start; margin-bottom: 14px; }
  .step-num { flex-shrink: 0; width: 26px; height: 26px; background: #16a34a; color: #fff; border-radius: 50%; font-size: 13px; font-weight: 700; display: flex; align-items: center; justify-content: center; margin-right: 12px; margin-top: 1px; }
  .step-text { font-size: 15px; color: #3f3f46; line-height: 1.5; }
  .btn { display: inline-block; background: #16a34a; color: #ffffff !important; text-decoration: none; font-size: 14px; font-weight: 600; padding: 11px 22px; border-radius: 6px; margin: 4px 6px 4px 0; }
  .btn-outline { display: inline-block; background: #ffffff; color: #16a34a !important; text-decoration: none; font-size: 14px; font-weight: 600; padding: 10px 22px; border-radius: 6px; border: 1.5px solid #16a34a; margin: 4px 0; }
  .note { background: #fafafa; border-left: 3px solid #d4d4d8; border-radius: 0 4px 4px 0; padding: 12px 16px; margin: 20px 0; font-size: 14px; color: #52525b; }
  .divider { border: none; border-top: 1px solid #e4e4e7; margin: 28px 0; }
  .footer { font-size: 13px; color: #a1a1aa; text-align: center; margin-top: 28px; }
  .footer a { color: #71717a; }
`;

function renderStep(num, html) {
  return `
    <tr>
      <td valign="top" style="padding-bottom:14px;">
        <table cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td valign="top" style="padding-right:12px; padding-top:1px;">
              <div style="width:26px;height:26px;background:#16a34a;border-radius:50%;text-align:center;line-height:26px;font-size:13px;font-weight:700;color:#ffffff;">${num}</div>
            </td>
            <td valign="top" style="font-size:15px;color:#3f3f46;line-height:1.5;">${html}</td>
          </tr>
        </table>
      </td>
    </tr>`;
}

function layout(bodyContent) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <style>${baseStyles}</style>
</head>
<body>
  <div class="wrapper">
    <div class="card">
      <div class="logo">OptiSheets</div>
      ${bodyContent}
      <hr class="divider" />
      <p class="footer">
        Questions? Reply to this email or contact us at
        <a href="mailto:${SUPPORT_EMAIL}">${SUPPORT_EMAIL}</a>
      </p>
    </div>
  </div>
</body>
</html>`;
}

async function sendAITemplatePurchaseEmail({ to, customerName, templateName, privateKey, sheetUrl, pdfUrl }) {
  const html = layout(`
    <h1>You're all set, ${customerName}!</h1>
    <p>Thanks for purchasing <strong>${templateName}</strong>. Below is your unique license key — keep it safe, you'll need it to activate AI features in your sheet.</p>

    <div class="key-box">
      <div class="key-label">Your License Key</div>
      <div class="key-value">${privateKey}</div>
    </div>

    <p style="font-weight:600;margin-bottom:10px;">Getting started:</p>
    <table class="steps" cellpadding="0" cellspacing="0" border="0" width="100%">
      ${renderStep(1, `Click the template link below to open the view-only sheet.`)}
      ${renderStep(2, `Go to <strong>File → Make a Copy</strong> to get your own editable version.`)}
      ${renderStep(3, `Open the <strong>Settings</strong> sheet and paste your license key into cell <strong>B2</strong>.`)}
      ${renderStep(4, `Click <strong>OptiSheets AI → Get AI Recommendations</strong> to get started.`)}
    </table>

    <p>
      <a href="${sheetUrl}" class="btn">Open Template</a>
      <a href="${pdfUrl}" class="btn-outline">View PDF Instructions</a>
    </p>

    <div class="note">
      <strong>Note:</strong> AI features require AI Credits, which are sold separately.
      Purchase a credit pack from our Etsy shop to start using AI recommendations.
    </div>
  `);

  await createTransporter().sendMail({
    from: process.env.EMAIL_FROM,
    to,
    subject: `Your OptiSheets ${templateName} — License Key & Template Access`,
    html,
  });
}

async function sendStandardTemplatePurchaseEmail({ to, customerName, templateName, sheetUrl, pdfUrl }) {
  const html = layout(`
    <h1>You're all set, ${customerName}!</h1>
    <p>Thanks for purchasing <strong>${templateName}</strong>. Follow the steps below to start using your template.</p>

    <p style="font-weight:600;margin-bottom:10px;">Getting started:</p>
    <table class="steps" cellpadding="0" cellspacing="0" border="0" width="100%">
      ${renderStep(1, `Click the template link below to open the view-only sheet.`)}
      ${renderStep(2, `Go to <strong>File → Make a Copy</strong> to get your own editable version.`)}
      ${renderStep(3, `Follow the PDF instructions to get started.`)}
    </table>

    <p>
      <a href="${sheetUrl}" class="btn">Open Template</a>
      <a href="${pdfUrl}" class="btn-outline">View PDF Instructions</a>
    </p>
  `);

  await createTransporter().sendMail({
    from: process.env.EMAIL_FROM,
    to,
    subject: `Your OptiSheets ${templateName} — Template Access`,
    html,
  });
}

async function sendCreditPurchaseEmail({ to, customerName, amount, newBalance }) {
  const html = layout(`
    <h1>Your credits have been added!</h1>
    <p>Hi ${customerName}, <strong>${amount} AI credit${amount !== 1 ? 's' : ''}</strong> have been added to your account.</p>

    <div class="balance-box">
      <div class="balance-label">New Credit Balance</div>
      <div class="balance-value">${newBalance} credit${newBalance !== 1 ? 's' : ''}</div>
    </div>

    <p>
      To check your balance at any time, open your OptiSheets template and look at the
      <strong>Settings</strong> sheet — your current credit balance is displayed in cell <strong>B3</strong>.
    </p>

    <div class="note">
      Once credits are loaded, click <strong>OptiSheets AI → Get AI Recommendations</strong> in your sheet to use them.
    </div>
  `);

  await createTransporter().sendMail({
    from: process.env.EMAIL_FROM,
    to,
    subject: 'Your OptiSheets AI Credits Have Been Added',
    html,
  });
}

module.exports = {
  sendAITemplatePurchaseEmail,
  sendStandardTemplatePurchaseEmail,
  sendCreditPurchaseEmail,
};
