// =============================================================================
// OptiSheets — 4-Year College Planner
// Google Apps Script
// =============================================================================
//
// SETUP INSTRUCTIONS
// ------------------
// 1. Open the Apps Script editor (Extensions → Apps Script).
// 2. From the menu bar choose OptiSheets AI → Setup All Sheets to create and
//    configure all three sheets automatically.
// 3. On the Settings sheet, paste your license key into cell B2 and fill in
//    the rest of your academic profile (rows 3–11).
// 4. On the Class List sheet, add every class you need to take.
// 5. Run OptiSheets AI → Get AI Recommendations to receive your personalized
//    4-year academic plan.
//
// SHEET LAYOUT
// ------------
//   Settings          — academic profile + license key (row 2)
//   Class List        — one row per course (columns A–E)
//   AI Recommendations — created automatically on first run
// =============================================================================

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var TEMPLATE_ID      = 'four-year-planner';
var SETTINGS_SHEET   = 'Settings';
var CLASS_LIST_SHEET = 'Class List';
var OUTPUT_SHEET     = 'AI Recommendations';
var PRIVATE_KEY_CELL = 'B2';

// Brand palette
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
      { name: 'Get AI Recommendations', functionName: 'getAIRecommendations' },
      { name: 'Setup All Sheets',        functionName: 'setupAllSheets' },
      null, // separator
      { name: 'Setup Settings Sheet',    functionName: 'setupSettingsSheet' },
      { name: 'Setup Class List Sheet',  functionName: 'setupClassListSheet' },
      { name: 'About / Setup Help',      functionName: 'showSetupHelp' },
    ]);
}

// ---------------------------------------------------------------------------
// Setup: Settings sheet
// ---------------------------------------------------------------------------

/**
 * Creates or rebuilds the Settings sheet with branded layout and data
 * validation. Saves and restores any values already in B2:B11 so that
 * re-running setup never wipes the user's profile or license key.
 */
function setupSettingsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
  }

  // ── 1. Save B2:B11 values before clearing ────────────────────────────────
  var savedValues = sheet.getRange('B2:B11').getValues();

  // ── 2. Clear everything ───────────────────────────────────────────────────
  sheet.clear();

  // ── 3. Row 1: branded banner across A1:C1 ────────────────────────────────
  sheet.getRange('A1:C1').merge()
    .setValue('OptiSheets — 4-Year Planner Settings')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── 4. Column A labels (rows 2–11) ───────────────────────────────────────
  var labels = [
    ['License Key'],
    ['University Name'],
    ['Major'],
    ['Minor (optional)'],
    ['Start Year'],
    ['Target GPA'],
    ['Max Credits Per Semester'],
    ['Min Credits Per Semester'],
    ['Preferred Semester Load'],
    ['Summer Semesters Available'],
  ];
  sheet.getRange(2, 1, 10, 1).setValues(labels)
    .setFontWeight('bold')
    .setFontColor(COLOR_DARK_GREEN);

  // ── 5. Data validation for B6–B11 ────────────────────────────────────────

  // B6 — Start Year: integer 2020–2040
  sheet.getRange('B6').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(2020, 2040)
      .setAllowInvalid(false)
      .build()
  );

  // B7 — Target GPA: 0.0–4.0
  sheet.getRange('B7').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0.0, 4.0)
      .setAllowInvalid(false)
      .build()
  );

  // B8 — Max Credits Per Semester: 12–22
  sheet.getRange('B8').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(12, 22)
      .setAllowInvalid(false)
      .build()
  );

  // B9 — Min Credits Per Semester: 9–18
  sheet.getRange('B9').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(9, 18)
      .setAllowInvalid(false)
      .build()
  );

  // B10 — Preferred Semester Load: dropdown
  sheet.getRange('B10').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Light', 'Moderate', 'Heavy'], true)
      .setAllowInvalid(false)
      .build()
  );

  // B11 — Summer Semesters Available: dropdown
  sheet.getRange('B11').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .build()
  );

  // ── 6. Row 13: instructional hint ────────────────────────────────────────
  sheet.getRange('A13')
    .setValue('Fill in your profile above, then add your classes to the Class List sheet.')
    .setFontColor('#757575')
    .setFontStyle('italic');

  // ── 7. Column widths ──────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 240);

  // ── 8. Tab color ──────────────────────────────────────────────────────────
  sheet.setTabColor(COLOR_DARK_GREEN);

  // ── 9. Restore saved B2:B11 values ───────────────────────────────────────
  sheet.getRange('B2:B11').setValues(savedValues);
}

// ---------------------------------------------------------------------------
// Setup: Class List sheet
// ---------------------------------------------------------------------------

/**
 * Creates or rebuilds the Class List sheet. Saves all existing class data
 * (rows 3+) before clearing and restores it after rebuilding so that
 * re-running setup never loses the user's class entries.
 */
function setupClassListSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLASS_LIST_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CLASS_LIST_SHEET);
  }

  // ── 1. Save existing class data (rows 3+) before clearing ────────────────
  var lastRow = sheet.getLastRow();
  var savedClassData = null;
  if (lastRow >= 3) {
    savedClassData = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  }

  // ── 2. Clear everything ───────────────────────────────────────────────────
  sheet.clear();

  // ── 3. Row 1: branded banner across A1:E1 ────────────────────────────────
  sheet.getRange(1, 1, 1, 5).merge()
    .setValue('OptiSheets — Class List')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── 4. Row 2: column headers ──────────────────────────────────────────────
  sheet.getRange(2, 1, 1, 5)
    .setValues([['Class Name', 'Credit Hours', 'Difficulty (1-10)', 'Prerequisites', 'Already Completed']])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 28);

  // ── 5. Data rows 3–102: alternating row colors ────────────────────────────
  for (var r = 3; r <= 102; r++) {
    var rowBg = (r % 2 === 1) ? COLOR_WHITE : COLOR_VERY_LIGHT_GREEN;
    sheet.getRange(r, 1, 1, 5).setBackground(rowBg);
    sheet.setRowHeight(r, 24);
  }

  // ── 6. Data validation (single range calls, no help text) ────────────────

  // Column B (Credit Hours): 1–6
  sheet.getRange(3, 2, 100, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 6)
      .setAllowInvalid(false)
      .build()
  );

  // Column C (Difficulty): 1–10
  sheet.getRange(3, 3, 100, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 10)
      .setAllowInvalid(false)
      .build()
  );

  // Column E (Already Completed): Yes / No
  sheet.getRange(3, 5, 100, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .build()
  );

  // ── 7. Column widths ──────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 140);

  // ── 8. Freeze header rows ─────────────────────────────────────────────────
  sheet.setFrozenRows(2);

  // ── 9. Tab color ──────────────────────────────────────────────────────────
  sheet.setTabColor(COLOR_GREEN);

  // ── 10. Restore saved class data ─────────────────────────────────────────
  if (savedClassData && savedClassData.length > 0) {
    sheet.getRange(3, 1, savedClassData.length, 5).setValues(savedClassData);
  }
}

// ---------------------------------------------------------------------------
// Setup: all sheets at once
// ---------------------------------------------------------------------------

function setupAllSheets() {
  setupSettingsSheet();
  setupClassListSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'All sheets configured! Fill in Settings and add your classes to get started.',
    'OptiSheets AI',
    8
  );
}

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------

function getAIRecommendations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Preparing your data…', 'OptiSheets AI', 5);

  // ── 1. Read license key from Settings ────────────────────────────────────
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!settingsSheet) {
    showError(
      'Missing sheet: "' + SETTINGS_SHEET + '"',
      'Please run OptiSheets AI → Setup All Sheets first to create and configure all sheets.'
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

  // ── 2. Read profile from Settings B3:B11 (single getValues call) ─────────
  var profileValues = settingsSheet.getRange('B3:B11').getValues();
  var profile = {
    university:       String(profileValues[0][0]).trim(),
    major:            String(profileValues[1][0]).trim(),
    minor:            String(profileValues[2][0]).trim(),
    start_year:       Number(profileValues[3][0]) || 0,
    target_gpa:       Number(profileValues[4][0]) || 0,
    max_credits:      Number(profileValues[5][0]) || 0,
    min_credits:      Number(profileValues[6][0]) || 0,
    preferred_load:   String(profileValues[7][0]).trim(),
    summer_available: String(profileValues[8][0]).trim(),
  };

  // ── 3. Read backend URL from Script Property ──────────────────────────────
  var baseUrl = getScriptProperty('OPTISHEETS_BASE_URL');
  if (!baseUrl) {
    showError(
      'Backend URL not configured',
      'Go to Extensions → Apps Script → Project Settings → Script Properties\n' +
      'and add:\n\n  Key:   OPTISHEETS_BASE_URL\n  Value: https://your-backend.vercel.app'
    );
    return;
  }
  baseUrl = baseUrl.replace(/\/+$/, '');

  // ── 4. Read class list ────────────────────────────────────────────────────
  var classList = readClassList(ss);
  if (classList === null) return; // readClassList already showed an alert

  if (classList.length === 0) {
    showError(
      'No classes found',
      'Your Class List sheet appears to be empty.\n\n' +
      'Add your required courses starting at row 3 and try again.'
    );
    return;
  }

  // ── 5. Build request payload ──────────────────────────────────────────────
  var userData = {
    profile: profile,
    classes: classList,
  };

  var payload = {
    private_key:   privateKey,
    template_id:   TEMPLATE_ID,
    user_data:     userData,
    system_prompt: buildSystemPromptString(),
  };

  // ── 6. Call backend ───────────────────────────────────────────────────────
  ss.toast('Calling OptiSheets AI (' + classList.length + ' classes)…', 'OptiSheets AI', 30);

  var result = callBackend(baseUrl, payload);
  if (!result) return; // callBackend already showed an alert

  // ── 7. Write output ───────────────────────────────────────────────────────
  writeOutput(ss, result, classList.length);
  ss.toast('Done! ' + result.remaining_credits + ' credit(s) remaining.', 'OptiSheets AI', 8);
}

// ---------------------------------------------------------------------------
// Build system prompt (mirrors templates/four-year-planner/prompt.js)
// ---------------------------------------------------------------------------

function buildSystemPromptString() {
  return (
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
    'Be specific, practical, and direct. Use plain text only. 700 words max.'
  );
}

// ---------------------------------------------------------------------------
// Read class list rows
// ---------------------------------------------------------------------------

/**
 * Returns an array of class objects from the Class List sheet, or null if
 * the sheet is missing (alert already shown).
 * Skips rows where both Class Name (col A) and Credit Hours (col B) are blank.
 */
function readClassList(ss) {
  var sheet = ss.getSheetByName(CLASS_LIST_SHEET);
  if (!sheet) {
    showError(
      'Missing sheet: "' + CLASS_LIST_SHEET + '"',
      'Please run OptiSheets AI → Setup All Sheets first to create and configure all sheets.'
    );
    return null;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Read all data rows at once (columns A–E), starting at row 3
  var data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  var classes = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var className      = String(row[0]).trim();
    var creditHoursRaw = row[1];
    var creditEmpty    = (creditHoursRaw === '' || creditHoursRaw === null || String(creditHoursRaw).trim() === '');

    // Skip rows where both Class Name and Credit Hours are blank
    if (!className && creditEmpty) continue;

    classes.push({
      class_name:        className,
      credit_hours:      creditEmpty ? '' : Number(creditHoursRaw),
      difficulty:        (row[2] === '' || row[2] === null) ? '' : Number(row[2]),
      prerequisites:     String(row[3]).trim(),
      already_completed: String(row[4]).trim(),
    });
  }

  return classes;
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
    muteHttpExceptions: true,
  };

  var response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    showError(
      'Network error',
      'Could not reach the OptiSheets backend.\n\nDetails: ' + e.message +
      '\n\nCheck your internet connection and try again.'
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
 * Maps HTTP status codes and raw error messages to user-friendly text.
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
        'Your class list is too large for a single request.\n\n' +
        'Try removing classes that are already completed and re-run, or split the list into smaller batches.'
      );
    case 400:
      return 'Bad request: ' + rawError + '\n\nThis is likely a bug — please contact support.';
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
 * Creates the sheet if it does not yet exist.
 */
function writeOutput(ss, result, classCount) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();

  // ── Tab color ─────────────────────────────────────────────────────────────
  sheet.setTabColor(COLOR_LIGHT_GREEN);

  // ── Row 1: branded banner ────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, 3).merge()
    .setValue('OptiSheets — 4-Year Planner Recommendations')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── Rows 2–4: metadata ────────────────────────────────────────────────────
  var now       = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'MMMM d, yyyy \'at\' h:mm a');
  var cacheNote = result.cached ? ' (cached — no credit used)' : '';

  sheet.getRange(2, 1)
    .setValue('Generated: ' + timestamp + cacheNote)
    .setFontColor('#555555')
    .setFontStyle('italic');

  sheet.getRange(3, 1)
    .setValue('Classes analysed: ' + classCount)
    .setFontColor('#555555');

  sheet.getRange(4, 1)
    .setValue('Credits remaining: ' + result.remaining_credits)
    .setFontColor('#555555');

  // ── Row 5+: AI output text ────────────────────────────────────────────────
  var outputCell = sheet.getRange(5, 1);
  outputCell
    .setValue(result.output)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(11)
    .setBackground(COLOR_VERY_LIGHT_GREEN)
    .setBorder(null, true, null, null, null, null, COLOR_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ── Column width and row height ───────────────────────────────────────────
  sheet.setColumnWidth(1, 720);
  sheet.setRowHeight(5, 600);

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
    'body{font-family:sans-serif;font-size:13px;padding:14px 16px;line-height:1.65;color:#212121}' +
    'h2{margin-top:0;color:#1B5E20;font-size:16px}' +
    'ol{padding-left:18px}li{margin-bottom:7px}' +
    'b{color:#1B5E20}' +
    'table{width:100%;border-collapse:collapse;margin-top:8px}' +
    'th{background:#1B5E20;color:#FFFFFF;text-align:left;padding:6px 8px;font-weight:bold}' +
    'td{padding:5px 8px;border-bottom:1px solid #E0E0E0;vertical-align:top}' +
    'tr:nth-child(even) td{background:#F1F8E9}' +
    '</style>' +
    '<h2>OptiSheets — 4-Year College Planner</h2>' +
    '<b>Getting started — follow these steps:</b>' +
    '<ol>' +
    '<li>Run <b>OptiSheets AI → Setup All Sheets</b> from the menu to create and configure all three sheets automatically.</li>' +
    '<li>Fill in your academic profile on the <b>Settings</b> sheet — university, major, credits per semester, target GPA, and so on.</li>' +
    '<li>Paste your license key into cell <b>B2</b> of the Settings sheet.</li>' +
    '<li>Add all your required courses to the <b>Class List</b> sheet — include credit hours, difficulty, prerequisites, and whether each class is already completed.</li>' +
    '<li>Run <b>OptiSheets AI → Get AI Recommendations</b> — the AI will build your personalized 4-year academic plan.</li>' +
    '</ol>' +
    '<br>' +
    '<b>Class List columns explained:</b>' +
    '<table>' +
    '<tr><th>Column</th><th>What to enter</th></tr>' +
    '<tr><td><b>Class Name</b></td><td>Full course name (e.g. Calculus I, Organic Chemistry)</td></tr>' +
    '<tr><td><b>Credit Hours</b></td><td>Number of credits the course is worth (1–6)</td></tr>' +
    '<tr><td><b>Difficulty (1-10)</b></td><td>Your estimated difficulty — 1 = very easy, 10 = extremely hard</td></tr>' +
    '<tr><td><b>Prerequisites</b></td><td>Comma-separated names of courses that must be completed first, or leave blank if none</td></tr>' +
    '<tr><td><b>Already Completed</b></td><td><b>Yes</b> if you have already passed this course, <b>No</b> if you still need to take it</td></tr>' +
    '</table>'
  )
    .setTitle('OptiSheets Setup Help')
    .setWidth(580)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'OptiSheets Setup Help');
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function showError(title, message) {
  SpreadsheetApp.getUi().alert('⚠️ ' + title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function getScriptProperty(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? String(value).trim() : '';
}
