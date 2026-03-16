'use strict';

// ---------------------------------------------------------------------------
// System prompt (static)
// ---------------------------------------------------------------------------
const SYSTEM_PROMPT =
  'You are a supportive but honest academic coach for college students. ' +
  'You will receive a JSON object with three keys: ' +
  '"gpa_tracker" (array of class rows with Semester, Class Name, Credit Hours, Letter Grade, Grade Points, Weighted Points, Notes), ' +
  '"grade_tracker" (object with category breakdowns: assignments, quizzes, exams, projects, participation — each with count, weight, and an array of score entries), ' +
  'and "settings" (target_gpa, current_cumulative_gpa, prior_credits_completed, graduation_credits_required, current_semester_gpa, credits_this_semester). ' +
  'Analyze all three data sources together and provide a structured response with these exact labeled sections:\n' +
  '1. GPA SNAPSHOT: State the student\'s current semester GPA, cumulative GPA, and gap to their target GPA. If they are on track, acknowledge it. If not, be specific about what needs to change.\n' +
  '2. GRADE RESCUE ALERTS: Flag any class currently trending toward a C or below. Give specific, actionable advice for each flagged class including what to prioritize studying and what minimum scores are needed on remaining work.\n' +
  '3. PRIORITY CLASS STRATEGY: Using credit-hour weighting, tell the student exactly which classes deserve the most energy this semester and why. Recommend whether to aim for an A, A-, or B+ in each class based on current standing and GPA targets.\n' +
  '4. FINAL GRADE SIMULATIONS: Based on the Grade Tracker data, tell the student what scores they need on remaining assignments, quizzes, and exams to hit their target grade in the currently tracked class. Be specific with numbers.\n' +
  '5. BURNOUT CHECK: If the student is performing well across all classes, acknowledge it and advise on maintaining performance without over-studying. If they are struggling broadly, provide motivational but realistic advice.\n' +
  '6. SEMESTER REFLECTION & NEXT STEPS: Give 3 to 5 concrete action items the student should do in the next 7 days based on their data. End with a one-sentence motivational close.\n' +
  'Be specific, encouraging, and direct. Use plain text only. 600 words max.';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Rank a GPA tracker row for prioritization.
 * Higher credit hours + lower/missing grade = higher priority (lower rank number).
 */
const GPA_POINTS_MAP = {
  'A': 4.0, 'A-': 3.67, 'B+': 3.33, 'B': 3.0, 'B-': 2.67,
  'C+': 2.33, 'C': 2.0, 'C-': 1.67, 'D+': 1.33, 'D': 1.0, 'F': 0.0,
};

function rankGpaRow(row) {
  const credits = Number(row['Credit Hours']) || 0;
  const grade   = (row['Letter Grade'] || '').trim();
  const points  = grade === '' ? -1 : (GPA_POINTS_MAP[grade] !== undefined ? GPA_POINTS_MAP[grade] : 2.0);
  // Lower GPA points + more credit hours = higher priority (smaller rank)
  return points * 10 - credits;
}

/**
 * Fit as many GPA tracker rows as possible inside charBudget characters.
 * Rows with lower grades and more credits come first.
 */
function fitRows(rows, charBudget) {
  const sorted = [...rows].sort((a, b) => rankGpaRow(a) - rankGpaRow(b));
  const lines  = [];
  let used = 0;
  for (const row of sorted) {
    const line = JSON.stringify(row);
    if (used + line.length + 2 > charBudget) break;
    lines.push(line);
    used += line.length + 2;
  }
  return { lines, shown: lines.length, total: rows.length };
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

/**
 * Build the OpenAI messages payload for the gpa-and-grade-tracker template.
 * Total prompt is capped to ~1200 input tokens (~4800 chars).
 *
 * @param {object} userData
 * @param {Array}  userData.gpa_tracker    - Array of class rows from GPA Tracker sheet
 * @param {object} userData.grade_tracker  - Current class grade breakdown from Grade Tracker sheet
 * @param {object} userData.settings       - Cumulative settings and computed GPA values
 *
 * @returns {{ system: string, user: string }}
 */
function buildPrompt(userData) {
  const {
    gpa_tracker   = [],
    grade_tracker = {},
    settings      = {},
  } = userData || {};

  // Token budget: 1200 tokens ≈ 4800 chars total.
  // SYSTEM_PROMPT ≈ 900 chars; settings+grade_tracker+preamble ≈ 650 chars
  // → ~3250 chars for gpa_tracker row data.
  const ROW_CHAR_BUDGET = 3250;

  const { lines, shown, total } = fitRows(gpa_tracker, ROW_CHAR_BUDGET);

  const truncationNote =
    shown < total
      ? `\n(${shown} of ${total} GPA tracker rows shown; lowest-priority rows omitted for length.)`
      : '';

  const settingsStr   = JSON.stringify(settings);
  const gradeTrackerStr = JSON.stringify(grade_tracker);

  const user =
    `Here is my academic data:\n\n` +
    `SETTINGS: ${settingsStr}\n\n` +
    `GPA TRACKER (${shown} classes):\n[\n${lines.join(',\n')}\n]` +
    truncationNote +
    `\n\nGRADE TRACKER (current class breakdown):\n${gradeTrackerStr}` +
    `\n\nPlease give me personalized academic recommendations based on this data.`;

  return { system: SYSTEM_PROMPT, user };
}

module.exports = { buildPrompt, SYSTEM_PROMPT };
