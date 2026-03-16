'use strict';

// ---------------------------------------------------------------------------
// System prompt (static)
// ---------------------------------------------------------------------------
const SYSTEM_PROMPT =
  'You are an expert academic advisor helping a college student build a personalized 4-year academic plan. ' +
  'You will receive a JSON object with two keys: "profile" (student settings) and "classes" (list of required courses). ' +
  'The profile has these keys: university, major, minor, start_year, target_gpa, max_credits, min_credits, preferred_load ' +
  '(Light, Moderate, or Heavy), summer_available (Yes or No). ' +
  'Each class has: class_name, credit_hours (integer), difficulty (1-10), prerequisites (comma-separated class names or blank), ' +
  'already_completed (Yes or No). ' +
  'Produce a structured plain-text response with these exact labeled sections:\n' +
  '1. 4-YEAR SEMESTER PLAN: Lay out all 8 semesters (Fall Year 1 through Spring Year 4) plus Summer slots if summer_available is Yes. ' +
  'For each semester list the exact classes scheduled, total credit hours, total difficulty score, and a load rating (Light / Moderate / Heavy). ' +
  'Never exceed max_credits or go below min_credits per semester. Never schedule a class before its prerequisites. ' +
  'Do not reschedule already-completed classes.\n' +
  '2. SEMESTER LOAD ANALYSIS: Evaluate each semester\'s difficulty score. Flag any semester where the combined difficulty score exceeds 35 as High Risk. ' +
  'Recommend one class to move to a different semester for each flagged semester.\n' +
  '3. PREREQUISITE CHAIN: Identify the 3 most critical prerequisite chains in the student\'s class list (the sequences that unlock the most future classes). ' +
  'Explain why these must be taken early.\n' +
  '4. STRATEGIC ADVICE: Recommend when to take the hardest classes, which semesters to keep lighter for GPA protection, ' +
  'and which classes pair well or poorly together based on difficulty and workload.\n' +
  '5. GRADUATION RISK FLAGS: Calculate total credit hours across all classes excluding already-completed ones. ' +
  'If the total does not fit within the available semesters given the credit constraints, flag it clearly and suggest specific adjustments.\n' +
  '6. FLEXIBILITY BUFFER: Recommend leaving 1-2 elective or free slots open per year and explain why this benefits the student long term.\n' +
  'Be specific, practical, and direct. Use plain text only. 700 words max.';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Priority rank for a class row.
 * Lower number = higher scheduling importance:
 *   0 — not yet completed AND has prerequisites (most critical to schedule early)
 *   1 — not yet completed, no prerequisites
 *   2 — already completed (lowest priority; included last)
 */
function rankClass(cls) {
  const notDone = (cls.already_completed || '').toLowerCase() !== 'yes';
  const hasPrereqs = !!(cls.prerequisites && String(cls.prerequisites).trim());
  if (notDone && hasPrereqs) return 0;
  if (notDone && !hasPrereqs) return 1;
  return 2;
}

/**
 * Fit as many class rows as possible inside charBudget characters,
 * prioritising classes not yet completed and those with prerequisites so the AI
 * always sees the most scheduling-critical courses within the token budget.
 */
function fitRows(rows, charBudget) {
  const sorted = [...rows].sort((a, b) => rankClass(a) - rankClass(b));
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
 * Build the OpenAI messages payload for the four-year-planner template.
 * Total prompt is capped to approximately 1200 input tokens (~4800 chars).
 *
 * @param {object}   userData
 * @param {object}   userData.profile  - Student profile settings object
 * @param {Array}    userData.classes  - Array of class objects
 *
 * @returns {{ system: string, user: string }}
 */
function buildPrompt(userData) {
  const { profile = {}, classes = [] } = userData || {};

  // Token budget: ~1200 tokens ≈ 4800 chars total.
  // SYSTEM_PROMPT ≈ 800 chars; profile block + preamble/footer ≈ 200 chars → ~3800 chars for class rows.
  const ROW_CHAR_BUDGET = 3800;

  const { lines, shown, total } = fitRows(classes, ROW_CHAR_BUDGET);

  const truncationNote =
    shown < total
      ? `\n(${shown} of ${total} classes shown; already-completed and lower-priority classes omitted for length.)`
      : '';

  const profileText = JSON.stringify(profile);

  const user =
    `Here is my student profile:\n\n${profileText}\n\n` +
    `Here are my classes:\n\n` +
    `[\n${lines.join(',\n')}\n]` +
    truncationNote +
    `\n\nPlease build my personalized 4-year academic plan based on this data.`;

  return { system: SYSTEM_PROMPT, user };
}

module.exports = { buildPrompt, SYSTEM_PROMPT };
