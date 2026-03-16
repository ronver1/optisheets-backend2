// =============================================================================
// OptiSheets — Workout & Fitness Tracker
// Google Apps Script
// =============================================================================
//
// SETUP INSTRUCTIONS
// ------------------
// 1. In the Apps Script editor, go to Project Settings → Script Properties and add:
//      OPTISHEETS_BASE_URL  →  https://your-backend.vercel.app
//
// 2. Make sure your spreadsheet has these sheets (exact names):
//      "Settings"      — license key and user profile (see layout below)
//      "Workout Log"   — your workout rows (see layout below)
//
// 3. Run "OptiSheets AI → Setup All Sheets" to configure headers and
//    data validation automatically.
//
// 4. "AI Recommendations" will be created automatically on first run.
//
// SETTINGS SHEET LAYOUT
// ---------------------
//   A2: "License Key"                B2: <your private key>   ← PRIVATE_KEY_CELL
//   A3: "Your Name"                  B3: <name>
//   A4: "Fitness Goal"               B4: <dropdown>
//   A5: "Training Days Per Week"     B5: <number 1–7>
//   A6: "Current Fitness Level"      B6: <dropdown>
//   A7: "Any Injuries or Limitations" B7: <text>
//
// WORKOUT LOG SHEET LAYOUT
// ------------------------
//   Row 1 = banner, Row 2 = headers (set by "Setup Workout Log Sheet")
//   Col A: Date
//   Col B: Workout Type   (Strength | Cardio | HIIT | Flexibility | Sports | Other)
//   Col C: Exercise Name
//   Col D: Sets
//   Col E: Reps
//   Col F: Weight (lbs)
//   Col G: Duration (min)
//   Col H: Cardio Distance (miles)
//   Col I: Perceived Effort  (integer 1–10)
//   Col J: Notes
// =============================================================================

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var TEMPLATE_ID      = 'workout-tracker';
var SETTINGS_SHEET   = 'Settings';
var WORKOUT_SHEET    = 'Workout Log';
var OUTPUT_SHEET     = 'AI Recommendations';
var PRIVATE_KEY_CELL = 'B2';

/** Exact column headers — must match meta.json column names */
var WORKOUT_HEADERS = [
  'Date',
  'Workout Type',
  'Exercise Name',
  'Sets',
  'Reps',
  'Weight (lbs)',
  'Duration (min)',
  'Cardio Distance (miles)',
  'Perceived Effort',
  'Notes'
];

var WORKOUT_TYPE_OPTIONS  = ['Strength', 'Cardio', 'HIIT', 'Flexibility', 'Sports', 'Other'];
var FITNESS_GOAL_OPTIONS  = ['Lose Weight', 'Build Muscle', 'Improve Endurance', 'General Fitness', 'Sport Performance'];
var FITNESS_LEVEL_OPTIONS = ['Beginner', 'Intermediate', 'Advanced'];

// Column indices (1-based) matching WORKOUT_HEADERS order
var COL_DATE             = 1;
var COL_WORKOUT_TYPE     = 2;
var COL_EXERCISE_NAME    = 3;
var COL_SETS             = 4;
var COL_REPS             = 5;
var COL_WEIGHT_LBS       = 6;
var COL_DURATION_MIN     = 7;
var COL_CARDIO_DIST      = 8;
var COL_PERCEIVED_EFFORT = 9;
var COL_NOTES            = 10;

var NUM_COLUMNS = 10;

// Brand colors
var COLOR_DARK_GREEN       = '#1B5E20';
var COLOR_GREEN            = '#388E3C';
var COLOR_LIGHT_GREEN      = '#A5D6A7';
var COLOR_VERY_LIGHT_GREEN = '#F1F8E9';
var COLOR_WHITE            = '#FFFFFF';
var COLOR_BLACK            = '#212121';

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu('OptiSheets AI', [
      { name: 'Get AI Recommendations',  functionName: 'getAIRecommendations' },
      { name: 'Setup All Sheets',        functionName: 'setupAllSheets' },
      null, // separator
      { name: 'Setup Settings Sheet',    functionName: 'setupSettingsSheet' },
      { name: 'Setup Workout Log Sheet', functionName: 'setupWorkoutLogSheet' },
      { name: 'About / Setup Help',      functionName: 'showSetupHelp' },
    ]);
}

// ---------------------------------------------------------------------------
// Setup: configure all sheets at once
// ---------------------------------------------------------------------------

function setupAllSheets() {
  setupSettingsSheet();
  setupWorkoutLogSheet();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('All sheets configured! Fill in your profile and log workouts starting at row 3.', 'OptiSheets AI', 7);
}

// ---------------------------------------------------------------------------
// Setup: Settings sheet
// ---------------------------------------------------------------------------

/**
 * Creates or resets the Settings sheet with profile fields and license key cell.
 * Saves existing values in B2 and B3:B7 before clearing, then restores them
 * after rebuilding so user data is never lost on re-setup.
 */
function setupSettingsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
  }

  // ── Save existing values before clearing ──────────────────────────────────
  var savedKey     = sheet.getRange('B2').getValue();
  var savedProfile = sheet.getRange('B3:B7').getValues(); // 5-row × 1-col array

  // ── Clear content and formats ─────────────────────────────────────────────
  sheet.clearContents();
  sheet.clearFormats();

  // ── Row 1: Title banner — merge A1:D1 ────────────────────────────────────
  sheet.getRange('A1:D1').merge()
    .setValue('OptiSheets \u2014 Workout Tracker Settings')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── Row 2: License Key ────────────────────────────────────────────────────
  sheet.getRange('A2')
    .setValue('License Key')
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);
  // B2 is PRIVATE_KEY_CELL — left blank for user to paste key

  // ── Row 3: Your Name ──────────────────────────────────────────────────────
  sheet.getRange('A3')
    .setValue('Your Name')
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);

  // ── Row 4: Fitness Goal (dropdown) ────────────────────────────────────────
  sheet.getRange('A4')
    .setValue('Fitness Goal')
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);
  var fitnessGoalRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FITNESS_GOAL_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B4').setDataValidation(fitnessGoalRule);

  // ── Row 5: Training Days Per Week (number 1–7) ────────────────────────────
  sheet.getRange('A5')
    .setValue('Training Days Per Week')
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);
  var trainingDaysRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 7)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B5').setDataValidation(trainingDaysRule);

  // ── Row 6: Current Fitness Level (dropdown) ───────────────────────────────
  sheet.getRange('A6')
    .setValue('Current Fitness Level')
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);
  var fitnessLevelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FITNESS_LEVEL_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B6').setDataValidation(fitnessLevelRule);

  // ── Row 7: Injuries / Limitations ────────────────────────────────────────
  sheet.getRange('A7')
    .setValue('Any Injuries or Limitations')
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);

  // ── Row 8: Instruction hint ───────────────────────────────────────────────
  sheet.getRange('A8')
    .setValue('Fill in your profile above, then run Setup All Sheets to configure your Workout Log.')
    .setFontColor('#757575')
    .setFontStyle('italic');

  // ── Restore saved values ──────────────────────────────────────────────────
  sheet.getRange('B2').setValue(savedKey);
  sheet.getRange('B3:B7').setValues(savedProfile);

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 200); // col A
  sheet.setColumnWidth(2, 260); // col B

  // ── Tab color ─────────────────────────────────────────────────────────────
  sheet.setTabColor(COLOR_DARK_GREEN);
}

// ---------------------------------------------------------------------------
// Setup: Workout Log sheet
// ---------------------------------------------------------------------------

/**
 * Creates or resets the Workout Log sheet with headers, data validation,
 * alternating row colors, and correct column widths.
 */
function setupWorkoutLogSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WORKOUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WORKOUT_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();

  // ── Row 1: Title banner — merge A1:J1 ────────────────────────────────────
  sheet.getRange(1, 1, 1, NUM_COLUMNS).merge()
    .setValue('OptiSheets \u2014 Workout Log')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── Row 2: Column headers ─────────────────────────────────────────────────
  sheet.getRange(2, 1, 1, NUM_COLUMNS).setValues([WORKOUT_HEADERS]);
  sheet.getRange(2, 1, 1, NUM_COLUMNS)
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 28);

  // ── Data rows 3–202 (200 rows): alternating colors ────────────────────────
  for (var r = 3; r <= 202; r++) {
    var rowBg = (r % 2 === 1) ? COLOR_WHITE : COLOR_VERY_LIGHT_GREEN; // odd=white, even=very light green
    sheet.getRange(r, 1, 1, NUM_COLUMNS).setBackground(rowBg);
    sheet.setRowHeight(r, 24);
  }

  // ── Data validation: Column B (Workout Type dropdown) ────────────────────
  var workoutTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(WORKOUT_TYPE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(3, COL_WORKOUT_TYPE, 200, 1).setDataValidation(workoutTypeRule);

  // ── Data validation: Column I (Perceived Effort 1–10) ────────────────────
  // Applied as a single range call; no help text per spec
  var effortRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 10)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(3, COL_PERCEIVED_EFFORT, 200, 1).setDataValidation(effortRule);

  // ── Data validation: Column A (Date) ─────────────────────────────────────
  // Applied as a single range call; no help text per spec
  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(3, COL_DATE, 200, 1).setDataValidation(dateRule);

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(COL_DATE,             100);
  sheet.setColumnWidth(COL_WORKOUT_TYPE,     110);
  sheet.setColumnWidth(COL_EXERCISE_NAME,    180);
  sheet.setColumnWidth(COL_SETS,              60);
  sheet.setColumnWidth(COL_REPS,              60);
  sheet.setColumnWidth(COL_WEIGHT_LBS,       100);
  sheet.setColumnWidth(COL_DURATION_MIN,     110);
  sheet.setColumnWidth(COL_CARDIO_DIST,      160);
  sheet.setColumnWidth(COL_PERCEIVED_EFFORT, 120);
  sheet.setColumnWidth(COL_NOTES,            200);

  // ── Freeze first 2 rows (banner + headers) ────────────────────────────────
  sheet.setFrozenRows(2);

  // ── Tab color ─────────────────────────────────────────────────────────────
  sheet.setTabColor(COLOR_GREEN);

  ss.toast('Workout Log sheet is ready! Log your workouts starting at row 3.', 'OptiSheets AI', 6);
}

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------

function getAIRecommendations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Preparing your data\u2026', 'OptiSheets AI', 5);

  // ── 1. Read license key from Settings B2 ──────────────────────────────────
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!settingsSheet) {
    showError(
      'Missing sheet: "' + SETTINGS_SHEET + '"',
      'Please run OptiSheets AI \u2192 Setup All Sheets first to configure your spreadsheet.'
    );
    return;
  }

  var privateKey = String(settingsSheet.getRange(PRIVATE_KEY_CELL).getValue()).trim();
  if (!privateKey) {
    showError(
      'License key not found',
      'Cell ' + PRIVATE_KEY_CELL + ' on the "' + SETTINGS_SHEET + '" sheet is empty.\n\n' +
      'Paste your OptiSheets license key there and try again.'
    );
    return;
  }

  // ── 2. Read user profile from Settings B3:B7 (single call) ───────────────
  var profileValues = settingsSheet.getRange('B3:B7').getValues();
  var profile = {
    name:          String(profileValues[0][0]).trim(),
    goal:          String(profileValues[1][0]).trim(),
    training_days: profileValues[2][0] !== '' ? Number(profileValues[2][0]) : '',
    fitness_level: String(profileValues[3][0]).trim(),
    injuries:      String(profileValues[4][0]).trim(),
  };

  // ── 3. Read backend URL from Script Properties ────────────────────────────
  var baseUrl = getScriptProperty('OPTISHEETS_BASE_URL');
  if (!baseUrl) {
    showError(
      'Backend URL not configured',
      'Go to Extensions \u2192 Apps Script \u2192 Project Settings \u2192 Script Properties\nand add:\n\n' +
      '  Key:   OPTISHEETS_BASE_URL\n  Value: https://your-backend.vercel.app'
    );
    return;
  }
  baseUrl = baseUrl.replace(/\/+$/, ''); // strip trailing slash

  // ── 4. Read workout rows ──────────────────────────────────────────────────
  var rows = readWorkoutRows(ss);
  if (rows === null) return; // readWorkoutRows already showed an alert
  if (rows.length === 0) {
    showError(
      'No workout data found',
      'Your Workout Log sheet appears to be empty.\n\nAdd some workout rows starting at row 3 and try again.'
    );
    return;
  }

  // ── 5. Build request payload ──────────────────────────────────────────────
  var payload = {
    private_key:   privateKey,
    template_id:   TEMPLATE_ID,
    user_data:     { prompt_inputs: rows, profile: profile },
    system_prompt: buildSystemPromptString(),
  };

  // ── 6. Call backend ───────────────────────────────────────────────────────
  ss.toast('Calling OptiSheets AI (' + rows.length + ' workout entries)\u2026', 'OptiSheets AI', 30);

  var result = callBackend(baseUrl, payload);
  if (!result) return; // callBackend already showed an alert

  // ── 7. Write output ───────────────────────────────────────────────────────
  writeOutput(ss, result, rows.length);
  ss.toast('Done! ' + result.remaining_credits + ' credit(s) remaining.', 'OptiSheets AI', 8);
}

// ---------------------------------------------------------------------------
// Build system prompt (mirrors prompt.js SYSTEM_PROMPT exactly)
// ---------------------------------------------------------------------------

function buildSystemPromptString() {
  return (
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
    'Be specific, motivating, and data-driven. Use plain text only. 500 words max.'
  );
}

// ---------------------------------------------------------------------------
// Read workout rows
// ---------------------------------------------------------------------------

/**
 * Returns an array of row objects using exact column names from meta.json,
 * or null if the sheet is missing (alert already shown).
 * Skips rows where both Date (col A) and Exercise Name (col C) are blank.
 *
 * Expected columns A–J:
 *   Date | Workout Type | Exercise Name | Sets | Reps | Weight (lbs) |
 *   Duration (min) | Cardio Distance (miles) | Perceived Effort | Notes
 */
function readWorkoutRows(ss) {
  var sheet = ss.getSheetByName(WORKOUT_SHEET);
  if (!sheet) {
    showError(
      'Missing sheet: "' + WORKOUT_SHEET + '"',
      'Please run OptiSheets AI \u2192 Setup All Sheets to create the Workout Log sheet.'
    );
    return null;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // no data rows yet

  // Read all data rows at once (columns A–J), starting at row 3
  var data = sheet.getRange(3, 1, lastRow - 2, NUM_COLUMNS).getValues();
  var rows = [];
  var tz   = ss.getSpreadsheetTimeZone();

  for (var i = 0; i < data.length; i++) {
    var row          = data[i];
    var dateVal      = row[COL_DATE - 1];
    var exerciseName = String(row[COL_EXERCISE_NAME - 1]).trim();

    // Skip rows where both Date and Exercise Name are blank
    if (!dateVal && !exerciseName) continue;

    // Format date value
    var dateStr = '';
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else if (dateVal) {
      dateStr = String(dateVal).trim();
    }

    // Helper: parse numeric cells, return '' if blank/non-numeric
    function numOrBlank(raw) {
      if (raw === '' || raw === null || raw === undefined) return '';
      var n = Number(raw);
      return isNaN(n) ? '' : n;
    }

    var rowObj = {};
    rowObj['Date']                    = dateStr;
    rowObj['Workout Type']            = String(row[COL_WORKOUT_TYPE - 1]).trim()  || '';
    rowObj['Exercise Name']           = exerciseName;
    rowObj['Sets']                    = numOrBlank(row[COL_SETS - 1]);
    rowObj['Reps']                    = numOrBlank(row[COL_REPS - 1]);
    rowObj['Weight (lbs)']            = numOrBlank(row[COL_WEIGHT_LBS - 1]);
    rowObj['Duration (min)']          = numOrBlank(row[COL_DURATION_MIN - 1]);
    rowObj['Cardio Distance (miles)'] = numOrBlank(row[COL_CARDIO_DIST - 1]);
    rowObj['Perceived Effort']        = (function(raw) {
      if (raw === '' || raw === null || raw === undefined) return '';
      var n = Number(raw);
      return isNaN(n) ? '' : parseInt(n, 10);
    }(row[COL_PERCEIVED_EFFORT - 1]));
    rowObj['Notes']                   = String(row[COL_NOTES - 1]).trim()         || '';

    rows.push(rowObj);
  }

  return rows;
}

// ---------------------------------------------------------------------------
// HTTP call
// ---------------------------------------------------------------------------

/**
 * POSTs to /api/get-recommendations and returns the parsed JSON body,
 * or null if anything went wrong (friendly alert already shown).
 */
function callBackend(baseUrl, payload) {
  var url     = baseUrl + '/api/get-recommendations';
  var options = {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true, // we handle all status codes ourselves
  };

  var response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    showError(
      'Network error',
      'Could not reach the OptiSheets backend.\n\nDetails: ' + e.message +
      '\n\nCheck your internet connection and verify the OPTISHEETS_BASE_URL script property.'
    );
    return null;
  }

  var statusCode = response.getResponseCode();
  var bodyText   = response.getContentText();
  var body;

  try {
    body = JSON.parse(bodyText);
  } catch (_) {
    showError(
      'Unexpected server response (HTTP ' + statusCode + ')',
      'The server returned a non-JSON response. This usually means the backend URL is wrong.\n\n' +
      'Raw response:\n' + bodyText.slice(0, 300)
    );
    return null;
  }

  if (body.success) return body;

  var friendlyMessage = friendlyErrorMessage(statusCode, body.error || 'Unknown error');
  showError('OptiSheets AI error (HTTP ' + statusCode + ')', friendlyMessage);
  return null;
}

/**
 * Returns a human-friendly error string based on status code and raw error message.
 */
function friendlyErrorMessage(statusCode, rawError) {
  switch (statusCode) {
    case 401:
      return (
        'Your license key was not recognised.\n\n' +
        'Double-check the key in cell ' + PRIVATE_KEY_CELL + ' of the "' + SETTINGS_SHEET + '" sheet.\n\n' +
        'If you just purchased, make sure you copied the full key.'
      );
    case 402:
      return (
        'You\'ve run out of AI Credits.\n\n' +
        'Visit optisheets.com to top up your balance, then try again.'
      );
    case 413:
      return (
        'Your workout log has too many entries for a single request.\n\n' +
        'Try removing older rows and re-run, or split your log into smaller batches.'
      );
    case 400:
      return 'Bad request: ' + rawError + '\n\nThis is likely a bug \u2014 please contact support.';
    case 500:
      return (
        'The OptiSheets server encountered an internal error.\n\n' +
        'Please try again in a moment. If this keeps happening, contact support.\n\n' +
        'Details: ' + rawError
      );
    default:
      return rawError;
  }
}

// ---------------------------------------------------------------------------
// Write output
// ---------------------------------------------------------------------------

/**
 * Writes AI recommendations to the "AI Recommendations" sheet.
 * Creates the sheet if it doesn't exist.
 */
function writeOutput(ss, result, rowCount) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();

  // ── Tab color ─────────────────────────────────────────────────────────────
  sheet.setTabColor(COLOR_LIGHT_GREEN);

  // ── Row 1: Branded banner ─────────────────────────────────────────────────
  sheet.getRange('A1:A1').merge()
    .setValue('OptiSheets \u2014 AI Recommendations')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM d, yyyy 'at' h:mm a");
  var cacheNote = result.cached ? ' (cached \u2014 no credit used)' : '';

  // ── Row 2: Timestamp ──────────────────────────────────────────────────────
  sheet.getRange(2, 1)
    .setValue('Generated: ' + timestamp + cacheNote)
    .setFontColor('#555555')
    .setFontStyle('italic');

  // ── Row 3: Workout count + credits remaining ──────────────────────────────
  sheet.getRange(3, 1)
    .setValue('Workouts analysed: ' + rowCount + '   |   Credits remaining: ' + result.remaining_credits)
    .setFontColor('#555555');

  // ── Row 4: Blank spacer ───────────────────────────────────────────────────
  sheet.getRange(4, 1).setValue('');

  // ── Row 5: AI output text ─────────────────────────────────────────────────
  var outputCell = sheet.getRange(5, 1);
  outputCell
    .setValue(result.output)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(11)
    .setBackground(COLOR_VERY_LIGHT_GREEN)
    .setBorder(null, true, null, null, null, null, COLOR_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ── Column width so text is readable ─────────────────────────────────────
  sheet.setColumnWidth(1, 720);

  // ── Row height: give the output cell plenty of space ─────────────────────
  sheet.setRowHeight(5, 500);

  // Navigate the user to the output sheet
  ss.setActiveSheet(sheet);
  sheet.setActiveRange(outputCell);
}

// ---------------------------------------------------------------------------
// Setup help dialog
// ---------------------------------------------------------------------------

function showSetupHelp() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:sans-serif;font-size:13px;padding:12px;line-height:1.6}' +
    'h2{margin-top:0;color:#1B5E20}' +
    'code{background:#f4f4f4;padding:2px 5px;border-radius:3px}' +
    'table{border-collapse:collapse;width:100%}' +
    'th,td{border:1px solid #ccc;padding:5px 8px;text-align:left}' +
    'th{background:#E8F5E9}' +
    'ol{padding-left:18px}' +
    'li{margin-bottom:8px}' +
    '</style>' +
    '<h2>OptiSheets Workout Tracker \u2014 Setup</h2>' +

    '<ol>' +
    '<li><b>Run Setup All Sheets</b> from the <b>OptiSheets AI</b> menu to configure the Settings and Workout Log sheets automatically.</li>' +
    '<li><b>Fill in your profile</b> on the Settings sheet: name, fitness goal, training days per week, fitness level, and any injuries or limitations.</li>' +
    '<li><b>Paste your license key</b> in cell <code>B2</code> of the Settings sheet.</li>' +
    '<li><b>Log your workouts</b> in the Workout Log sheet starting at row 3 \u2014 one row per exercise set or session.</li>' +
    '<li><b>Run Get AI Recommendations</b> from the OptiSheets AI menu to receive personalised fitness coaching.</li>' +
    '</ol>' +

    '<p><b>One-time script property:</b><br>' +
    'Extensions \u2192 Apps Script \u2192 Project Settings \u2192 Script Properties<br>' +
    'Add <code>OPTISHEETS_BASE_URL</code> \u2192 your backend URL</p>' +

    '<b>Workout Log columns (A\u2013J):</b><br>' +
    '<table>' +
    '<tr><th>Column</th><th>Field</th><th>What to enter</th></tr>' +
    '<tr><td>A</td><td>Date</td><td>Date of the workout session</td></tr>' +
    '<tr><td>B</td><td>Workout Type</td><td>Strength, Cardio, HIIT, Flexibility, Sports, or Other</td></tr>' +
    '<tr><td>C</td><td>Exercise Name</td><td>Name of the exercise (e.g. Squat, Treadmill)</td></tr>' +
    '<tr><td>D</td><td>Sets</td><td>Number of sets performed</td></tr>' +
    '<tr><td>E</td><td>Reps</td><td>Reps per set (leave blank for cardio)</td></tr>' +
    '<tr><td>F</td><td>Weight (lbs)</td><td>Weight used in lbs (leave blank if bodyweight)</td></tr>' +
    '<tr><td>G</td><td>Duration (min)</td><td>Duration of exercise or session in minutes</td></tr>' +
    '<tr><td>H</td><td>Cardio Distance (miles)</td><td>Distance covered (cardio only)</td></tr>' +
    '<tr><td>I</td><td>Perceived Effort</td><td>How hard it felt: 1 (easy) to 10 (max effort)</td></tr>' +
    '<tr><td>J</td><td>Notes</td><td>Any notes, form cues, or observations</td></tr>' +
    '</table>' +
    '<br><small>Rows 1\u20132 are the banner and headers \u2014 log your data from row 3 onwards.</small>'
  )
    .setTitle('OptiSheets Setup Help')
    .setWidth(560)
    .setHeight(480);

  SpreadsheetApp.getUi().showModalDialog(html, 'OptiSheets Setup Help');
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function showError(title, message) {
  SpreadsheetApp.getUi().alert('\u26a0\ufe0f ' + title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function getScriptProperty(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? String(value).trim() : '';
}
