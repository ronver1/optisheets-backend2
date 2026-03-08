'use strict';

// ---------------------------------------------------------------------------
// System prompt (static)
// ---------------------------------------------------------------------------
const SYSTEM_PROMPT =
  'You are a recruiting coach helping college students land internships. ' +
  'Given a student\'s application tracker data, give concise actionable recommendations. ' +
  'Cover: which applications to follow up on urgently, interview prep for active stages, ' +
  'pipeline weak spots (low response rate, stalled stages), and concrete next steps. ' +
  'Be specific. Use short labeled sections. Plain text only. 400 words max.';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Sort applications so active/late-stage ones come first.
 * Unknown statuses go to the end.
 */
const STATUS_ORDER = [
  'Offer',
  'Final Round',
  'Onsite',
  'Technical Screen',
  'Phone Screen',
  'Applied',
  'Rejected',
  'Withdrawn',
];

function rankStatus(status = '') {
  const idx = STATUS_ORDER.findIndex(
    (s) => s.toLowerCase() === status.toLowerCase()
  );
  return idx === -1 ? STATUS_ORDER.length : idx;
}

/**
 * Render one application row as a single line.
 * Keeps notes capped at 80 chars to contain token bloat.
 */
function renderApp(app) {
  const parts = [`- ${app.company || '?'} | ${app.role || '?'} | ${app.status || '?'}`];
  if (app.appliedDate) parts.push(`| Applied: ${app.appliedDate}`);
  if (app.notes) parts.push(`| Notes: ${String(app.notes).slice(0, 80)}`);
  return parts.join(' ');
}

/**
 * Fit as many application rows as possible inside charBudget characters.
 */
function fitApplications(apps, charBudget) {
  const sorted = [...apps].sort((a, b) => rankStatus(a.status) - rankStatus(b.status));
  const lines = [];
  let used = 0;
  for (const app of sorted) {
    const line = renderApp(app);
    if (used + line.length + 1 > charBudget) break; // +1 for newline
    lines.push(line);
    used += line.length + 1;
  }
  return { text: lines.join('\n'), shown: lines.length, total: apps.length };
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

/**
 * Build the OpenAI messages payload for the internship-tracker template.
 * Total prompt is capped to ~1000 input tokens (~4000 chars).
 *
 * @param {object}   userData
 * @param {Array}    userData.applications          - List of application objects
 * @param {string}   userData.applications[].company
 * @param {string}   userData.applications[].role
 * @param {string}   userData.applications[].status  - e.g. "Applied" | "Phone Screen" |
 *                                                      "Technical Screen" | "Onsite" |
 *                                                      "Final Round" | "Offer" |
 *                                                      "Rejected" | "Withdrawn"
 * @param {string}   [userData.applications[].appliedDate] - ISO date string (YYYY-MM-DD)
 * @param {string}   [userData.applications[].notes]       - Free-text notes (truncated at 80 chars)
 * @param {string}   [userData.targetRole]    - e.g. "Software Engineering Intern"
 * @param {string}   [userData.targetSeason]  - e.g. "Summer 2026"
 * @param {number}   [userData.targetCount]   - How many offers the student is aiming for
 *
 * @returns {{ system: string, user: string }}
 */
function buildPrompt(userData) {
  const {
    applications = [],
    targetRole = 'internship',
    targetSeason = '',
    targetCount = null,
  } = userData || {};

  // Token budget: 1000 tokens ≈ 4000 chars total.
  // SYSTEM_PROMPT ≈ 280 chars; preamble+footer ≈ 200 chars → ~3520 chars for app rows.
  const APP_CHAR_BUDGET = 3520;

  const { text: appBlock, shown, total } = fitApplications(applications, APP_CHAR_BUDGET);

  const truncationNote =
    shown < total ? `\n(${shown} of ${total} applications shown; remainder omitted for length.)` : '';

  const header = [
    `I am tracking my ${targetRole} search`,
    targetSeason ? ` for ${targetSeason}` : '',
    targetCount ? ` (targeting ${targetCount} offer${targetCount !== 1 ? 's' : ''})` : '',
    '.',
  ].join('');

  const user =
    `${header}\n\n` +
    `Applications (active stages first):\n${appBlock}${truncationNote}\n\n` +
    `Give me specific, actionable recruiting recommendations based on this data.`;

  return { system: SYSTEM_PROMPT, user };
}

module.exports = { buildPrompt, SYSTEM_PROMPT };
