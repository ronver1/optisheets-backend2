'use strict';

// ---------------------------------------------------------------------------
// System prompt (static)
// ---------------------------------------------------------------------------
const SYSTEM_PROMPT =
  'You are a knowledgeable, encouraging personal fitness coach. ' +
  'You will receive a JSON array of workout log rows. Each row uses these exact keys: ' +
  '"Date", "Workout Type" (one of: Strength, Cardio, HIIT, Flexibility, Sports, Other), ' +
  '"Exercise Name", "Sets" (number), "Reps" (number), "Weight (lbs)" (number), ' +
  '"Duration (min)" (number), "Cardio Distance (miles)" (number), ' +
  '"Perceived Effort" (integer 1\u201310, higher = harder), "Notes". ' +
  'Analyze the full workout log and produce a structured response with these exact labeled sections:\n' +
  '1. PROGRESS HIGHLIGHTS: Identify measurable improvements across the log such as increased weight, more reps, faster pace, or longer duration. Call out specific exercises where the student is clearly getting stronger or fitter.\n' +
  '2. WORKOUT BALANCE ANALYSIS: Analyze the mix of workout types. Flag if the student is overtraining one type (e.g. all strength, no cardio) or skipping recovery. Recommend adjustments to their weekly split.\n' +
  '3. EFFORT & RECOVERY FLAGS: Look at Perceived Effort scores. If multiple sessions show effort 9\u201310 with no low-effort or rest days between them, flag potential overtraining. If effort scores are consistently low (1\u20133), suggest increasing intensity.\n' +
  '4. EXERCISE-SPECIFIC RECOMMENDATIONS: For the top 3 most frequently logged exercises, give one specific tip to improve form, increase progressive overload, or avoid a common plateau.\n' +
  '5. NEXT WORKOUT SUGGESTION: Based on what was last logged and the overall pattern, suggest exactly what the student should do in their next workout session \u2014 specific exercises, approximate sets/reps/weight, and type.\n' +
  '6. WEEKLY GOALS: Give 2\u20133 specific, measurable goals for the coming week based on current performance (e.g. "increase squat weight by 5 lbs", "add one 20-minute cardio session", "hit 8 hours of sleep between sessions").\n' +
  'Be specific, motivating, and data-driven. Use plain text only. 500 words max.';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Fit as many rows as possible inside charBudget characters.
 * Sorts by date descending so the AI always sees the newest workouts first
 * when the log must be trimmed.
 */
function fitRows(rows, charBudget) {
  const sorted = [...rows].sort((a, b) => {
    const dateA = a['Date'] ? new Date(a['Date']) : new Date(0);
    const dateB = b['Date'] ? new Date(b['Date']) : new Date(0);
    return dateB - dateA; // most recent first
  });
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
 * Build the OpenAI messages payload for the workout-tracker template.
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
      ? `\n(${shown} of ${total} rows shown; oldest rows omitted for length.)`
      : '';

  const user =
    'Here are my workout log entries:\n\n' +
    `[\n${lines.join(',\n')}\n]` +
    truncationNote +
    '\n\nPlease give me personalized fitness recommendations based on this data.';

  return { system: SYSTEM_PROMPT, user };
}

module.exports = { buildPrompt, SYSTEM_PROMPT };
