'use strict';

// ---------------------------------------------------------------------------
// System prompt (static)
// ---------------------------------------------------------------------------
const SYSTEM_PROMPT =
  'You are a supportive but honest internship search coach for college students. ' +
  'You will receive a JSON array of internship application rows. Each row uses these exact keys: ' +
  '"Company Name", "Role/Position Title", "Industry", "Location", "Application Status" ' +
  '(one of: Applying, In Progress, Applied), "Recruiter Name", "Recruiter Email", ' +
  '"Interview Status" (one of: None, Phone Screen, Video Interview, In-Person Interview), ' +
  '"Personal Satisfaction" (integer 1–5, higher = more interested), "Notes". ' +
  'Analyze the full set of applications and provide a structured response with these labeled sections:\n' +
  '1. PRIORITY FOLLOW-UPS: Identify which applications deserve the most immediate attention based on Interview Status and Application Status. Include specific suggested actions (e.g. send thank-you email, follow up with recruiter, prepare for next round).\n' +
  '2. LOW SATISFACTION FLAGS: Flag any roles where Personal Satisfaction is 1 or 2. Give an honest recommendation on whether to keep pursuing each one.\n' +
  '3. NEXT STEPS BY COMPANY: For each company with an active application, suggest one concrete next action.\n' +
  '4. PATTERNS & INSIGHTS: Identify trends across the applications — e.g. which industries or roles are getting more traction, gaps in the pipeline, or missing recruiter contact info.\n' +
  '5. OVERALL STRATEGY: Give one overarching recommendation to improve the student\'s chances of receiving an offer.\n' +
  'Be specific, encouraging, and direct. Use plain text only. 500 words max.';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const INTERVIEW_STATUS_ORDER = [
  'In-Person Interview',
  'Video Interview',
  'Phone Screen',
  'None',
];

const APPLICATION_STATUS_ORDER = [
  'Applied',
  'In Progress',
  'Applying',
];

function rankRow(row) {
  const interviewRank = INTERVIEW_STATUS_ORDER.findIndex(
    (s) => s.toLowerCase() === (row['Interview Status'] || '').toLowerCase()
  );
  const appRank = APPLICATION_STATUS_ORDER.findIndex(
    (s) => s.toLowerCase() === (row['Application Status'] || '').toLowerCase()
  );
  // Lower index = higher priority; use combined rank
  return (interviewRank === -1 ? INTERVIEW_STATUS_ORDER.length : interviewRank) * 10 +
         (appRank === -1 ? APPLICATION_STATUS_ORDER.length : appRank);
}

/**
 * Fit as many rows as possible inside charBudget characters.
 */
function fitRows(rows, charBudget) {
  const sorted = [...rows].sort((a, b) => rankRow(a) - rankRow(b));
  const lines = [];
  let used = 0;
  for (const row of sorted) {
    const line = JSON.stringify(row);
    if (used + line.length + 2 > charBudget) break; // +2 for comma+newline
    lines.push(line);
    used += line.length + 2;
  }
  return { lines, shown: lines.length, total: rows.length };
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

/**
 * Build the OpenAI messages payload for the internship-tracker template.
 * Total prompt is capped to ~1000 input tokens (~4000 chars).
 *
 * @param {object} userData
 * @param {Array}  userData.prompt_inputs  - Array of row objects with the 10 column keys
 *
 * @returns {{ system: string, user: string }}
 */
function buildPrompt(userData) {
  const { prompt_inputs = [] } = userData || {};

  // Token budget: 1000 tokens ≈ 4000 chars total.
  // SYSTEM_PROMPT ≈ 600 chars; preamble+footer ≈ 150 chars → ~3250 chars for row data.
  const ROW_CHAR_BUDGET = 3250;

  const { lines, shown, total } = fitRows(prompt_inputs, ROW_CHAR_BUDGET);

  const truncationNote =
    shown < total
      ? `\n(${shown} of ${total} rows shown; oldest/lowest-priority rows omitted for length.)`
      : '';

  const user =
    `Here are my internship applications:\n\n` +
    `[\n${lines.join(',\n')}\n]` +
    truncationNote +
    `\n\nPlease give me personalized recommendations based on this data.`;

  return { system: SYSTEM_PROMPT, user };
}

module.exports = { buildPrompt, SYSTEM_PROMPT };
