// =============================================================================
// OptiSheets — Internship Tracker
// Google Apps Script
// =============================================================================
//
// SETUP INSTRUCTIONS
// ------------------
// 1. In the Apps Script editor, go to Project Settings → Script Properties and add:
//      OPTISHEETS_BASE_URL  →  https://your-backend.vercel.app
//
// 2. Make sure your spreadsheet has these sheets (exact names):
//      "Settings"  — configuration cells (see SETTINGS SHEET LAYOUT below)
//      "Tracker"   — your internship application rows (see TRACKER SHEET LAYOUT below)
//
// 3. Run "Setup Tracker Sheet" from the OptiSheets AI menu to create headers
//    and configure data validation automatically.
//
// 4. "AI Recommendations" will be created automatically on first run.
//
// SETTINGS SHEET LAYOUT
// ---------------------
//   A2: "License Key"   B2: <your private key>
//
// TRACKER SHEET LAYOUT
// --------------------
//   Row 1 = headers (set automatically by "Setup Tracker Sheet")
//   Col A: Company Name
//   Col B: Role/Position Title
//   Col C: Industry
//   Col D: Location
//   Col E: Application Status   (Applying | In Progress | Applied)
//   Col F: Recruiter Name
//   Col G: Recruiter Email
//   Col H: Interview Status     (None | Phone Screen | Video Interview | In-Person Interview)
//   Col I: Personal Satisfaction (integer 1–5)
//   Col J: Notes
// =============================================================================

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var TEMPLATE_ID       = "internship-tracker";
var SETTINGS_SHEET    = "Settings";
var TRACKER_SHEET     = "Tracker";
var OUTPUT_SHEET      = "AI Recommendations";
var PRIVATE_KEY_CELL  = "B2";

/** Exact column headers — must match meta.json column names */
var TRACKER_HEADERS = [
  "Company Name",
  "Role/Position Title",
  "Industry",
  "Location",
  "Application Status",
  "Recruiter Name",
  "Recruiter Email",
  "Interview Status",
  "Personal Satisfaction",
  "Notes"
];

var APPLICATION_STATUS_OPTIONS = ["Applying", "In Progress", "Applied"];
var INTERVIEW_STATUS_OPTIONS   = ["None", "Phone Screen", "Video Interview", "In-Person Interview"];

// Column indices (1-based) matching TRACKER_HEADERS order
var COL_COMPANY_NAME          = 1;
var COL_ROLE_TITLE            = 2;
var COL_INDUSTRY              = 3;
var COL_LOCATION              = 4;
var COL_APPLICATION_STATUS    = 5;
var COL_RECRUITER_NAME        = 6;
var COL_RECRUITER_EMAIL       = 7;
var COL_INTERVIEW_STATUS      = 8;
var COL_PERSONAL_SATISFACTION = 9;
var COL_NOTES                 = 10;

var NUM_COLUMNS = 10;

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("OptiSheets AI", [
      { name: "Get AI Recommendations",  functionName: "getAIRecommendations" },
      { name: "Setup Tracker Sheet",     functionName: "setupTrackerSheet" },
      null, // separator
      { name: "About / Setup Help",      functionName: "showSetupHelp" },
    ]);
  try { setupSettingsSheet(); } catch (e) { /* silent — branding only */ }
}

// ---------------------------------------------------------------------------
// Setup: create headers and data validation
// ---------------------------------------------------------------------------

/**
 * Creates or resets the Tracker sheet headers and configures data validation
 * for Application Status, Interview Status, and Personal Satisfaction columns.
 */
function setupTrackerSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRACKER_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(TRACKER_SHEET);
  }

  // ── Insert title banner row at row 1 (shifts existing content down) ───────
  sheet.insertRowBefore(1);
  var bannerRange = sheet.getRange(1, 1, 1, NUM_COLUMNS);
  bannerRange.merge()
    .setValue("OptiSheets — Internship Tracker")
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);

  // ── Write headers to row 2 ────────────────────────────────────────────────
  sheet.getRange(2, 1, 1, NUM_COLUMNS).setValues([TRACKER_HEADERS]);
  sheet.getRange(2, 1, 1, NUM_COLUMNS)
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF")
    .setFontSize(12)
    .setFontWeight("bold")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 32);

  // ── Freeze first 2 rows (banner + header) ────────────────────────────────
  sheet.setFrozenRows(2);

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(COL_COMPANY_NAME,          160);
  sheet.setColumnWidth(COL_ROLE_TITLE,            180);
  sheet.setColumnWidth(COL_INDUSTRY,              130);
  sheet.setColumnWidth(COL_LOCATION,              130);
  sheet.setColumnWidth(COL_APPLICATION_STATUS,    150);
  sheet.setColumnWidth(COL_RECRUITER_NAME,        140);
  sheet.setColumnWidth(COL_RECRUITER_EMAIL,       180);
  sheet.setColumnWidth(COL_INTERVIEW_STATUS,      180);
  sheet.setColumnWidth(COL_PERSONAL_SATISFACTION, 160);
  sheet.setColumnWidth(COL_NOTES,                 220);

  // ── Alternating row colors (rows 3–101, data rows) ───────────────────────
  for (var r = 3; r <= 101; r++) {
    var rowBg = (r % 2 === 1) ? "#FFFFFF" : "#F1F8E9"; // odd = white, even = very light green
    sheet.getRange(r, 1, 1, NUM_COLUMNS).setBackground(rowBg);
    sheet.setRowHeight(r, 26);
  }

  // ── Column alignment ─────────────────────────────────────────────────────
  // Left-align: A, B, C, D, F, G, J (text columns)
  sheet.getRange(3, COL_COMPANY_NAME,       99, 1).setHorizontalAlignment("left");
  sheet.getRange(3, COL_ROLE_TITLE,         99, 1).setHorizontalAlignment("left");
  sheet.getRange(3, COL_INDUSTRY,           99, 1).setHorizontalAlignment("left");
  sheet.getRange(3, COL_LOCATION,           99, 1).setHorizontalAlignment("left");
  sheet.getRange(3, COL_RECRUITER_NAME,     99, 1).setHorizontalAlignment("left");
  sheet.getRange(3, COL_RECRUITER_EMAIL,    99, 1).setHorizontalAlignment("left");
  sheet.getRange(3, COL_NOTES,              99, 1).setHorizontalAlignment("left");
  // Center-align: E, H, I (status and satisfaction columns)
  sheet.getRange(3, COL_APPLICATION_STATUS,    99, 1).setHorizontalAlignment("center");
  sheet.getRange(3, COL_INTERVIEW_STATUS,      99, 1).setHorizontalAlignment("center");
  sheet.getRange(3, COL_PERSONAL_SATISFACTION, 99, 1).setHorizontalAlignment("center");

  var lastDataRow = Math.max(sheet.getMaxRows(), 100);

  // ── Conditional formatting: Application Status (col E) ───────────────────
  var appStatusCfRange = sheet.getRange(3, COL_APPLICATION_STATUS, lastDataRow - 2, 1);
  var cfRules = [];
  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Applying")
    .setBackground("#FFF9C4")
    .setRanges([appStatusCfRange])
    .build());
  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("In Progress")
    .setBackground("#BBDEFB")
    .setRanges([appStatusCfRange])
    .build());
  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Applied")
    .setBackground("#C8E6C9")
    .setRanges([appStatusCfRange])
    .build());

  // ── Conditional formatting: Personal Satisfaction (col I) ────────────────
  var satisfactionCfRange = sheet.getRange(3, COL_PERSONAL_SATISFACTION, lastDataRow - 2, 1);
  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 2)
    .setBackground("#FFCDD2")
    .setRanges([satisfactionCfRange])
    .build());
  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(3)
    .setBackground("#FFF9C4")
    .setRanges([satisfactionCfRange])
    .build());
  cfRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 5)
    .setBackground("#C8E6C9")
    .setRanges([satisfactionCfRange])
    .build());
  sheet.setConditionalFormatRules(cfRules);

  // ── Application Status dropdown (col E) ───────────────────────────────────
  var appStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(APPLICATION_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .setHelpText("Choose: " + APPLICATION_STATUS_OPTIONS.join(", "))
    .build();
  sheet.getRange(3, COL_APPLICATION_STATUS, lastDataRow - 2, 1)
    .setDataValidation(appStatusRule);

  // ── Interview Status dropdown (col H) ────────────────────────────────────
  var interviewStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(INTERVIEW_STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .setHelpText("Choose: " + INTERVIEW_STATUS_OPTIONS.join(", "))
    .build();
  sheet.getRange(3, COL_INTERVIEW_STATUS, lastDataRow - 2, 1)
    .setDataValidation(interviewStatusRule);

  // ── Personal Satisfaction integer 1–5 (col I) ────────────────────────────
  var satisfactionRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 5)
    .setAllowInvalid(false)
    .setHelpText("Enter a whole number from 1 (low) to 5 (high)")
    .build();
  sheet.getRange(3, COL_PERSONAL_SATISFACTION, lastDataRow - 2, 1)
    .setDataValidation(satisfactionRule);

  // ── Tab color ─────────────────────────────────────────────────────────────
  sheet.setTabColor("#1B5E20");

  // ── Brand the Settings sheet too ─────────────────────────────────────────
  setupSettingsSheet();

  ss.toast("Tracker sheet is ready! Fill in your applications starting at row 3.", "OptiSheets AI", 6);
}

// ---------------------------------------------------------------------------
// Setup: brand the Settings sheet
// ---------------------------------------------------------------------------

/**
 * Applies OptiSheets branding to the Settings sheet.
 * Safe to call on every open — only touches visual/formatting properties.
 */
function setupSettingsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) return; // Settings sheet doesn't exist yet — nothing to brand

  // ── Tab color ─────────────────────────────────────────────────────────────
  sheet.setTabColor("#388E3C");

  // ── Title banner: merge A1:B1 ─────────────────────────────────────────────
  var bannerRange = sheet.getRange("A1:B1");
  bannerRange.merge()
    .setValue("OptiSheets — Settings")
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  // ── A2: label styling ─────────────────────────────────────────────────────
  sheet.getRange("A2")
    .setFontWeight("bold")
    .setFontColor("#1B5E20");

  // ── A3: instructional hint ────────────────────────────────────────────────
  sheet.getRange("A3")
    .setValue("Enter your license key in cell B2 to get started.")
    .setFontColor("#757575")
    .setFontStyle("italic");

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 160); // col A
  sheet.setColumnWidth(2, 280); // col B
}

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------

function getAIRecommendations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Preparing your data…", "OptiSheets AI", 5);

  // ── 1. Read settings ──────────────────────────────────────────────────────
  var settings = readSettings(ss);
  if (!settings) return; // readSettings already showed an alert

  // ── 2. Read tracker rows ──────────────────────────────────────────────────
  var rows = readTrackerRows(ss);
  if (rows === null) return; // readTrackerRows already showed an alert
  if (rows.length === 0) {
    showError(
      "No applications found",
      "Your Tracker sheet appears to be empty. Add some rows and try again."
    );
    return;
  }

  // ── 3. Build request payload ──────────────────────────────────────────────
  var payload = {
    private_key:   settings.privateKey,
    template_id:   TEMPLATE_ID,
    user_data:     { prompt_inputs: rows },
    system_prompt: buildSystemPrompt(),
  };

  // ── 4. Call backend ───────────────────────────────────────────────────────
  ss.toast("Calling OptiSheets AI (" + rows.length + " applications)…", "OptiSheets AI", 30);

  var result = callBackend(settings.baseUrl, payload);
  if (!result) return; // callBackend already showed an alert

  // ── 5. Write output ───────────────────────────────────────────────────────
  writeOutput(ss, result, rows.length);
  ss.toast("Done! " + result.remaining_credits + " credit(s) remaining.", "OptiSheets AI", 8);
}

// ---------------------------------------------------------------------------
// Build system prompt (mirrors prompt.js SYSTEM_PROMPT)
// ---------------------------------------------------------------------------

function buildSystemPrompt() {
  return (
    "You are a supportive but honest internship search coach for college students. " +
    "You will receive a JSON array of internship application rows. Each row uses these exact keys: " +
    "\"Company Name\", \"Role/Position Title\", \"Industry\", \"Location\", \"Application Status\" " +
    "(one of: Applying, In Progress, Applied), \"Recruiter Name\", \"Recruiter Email\", " +
    "\"Interview Status\" (one of: None, Phone Screen, Video Interview, In-Person Interview), " +
    "\"Personal Satisfaction\" (integer 1-5, higher = more interested), \"Notes\". " +
    "Analyze the full set of applications and provide a structured response with these labeled sections:\n" +
    "1. PRIORITY FOLLOW-UPS: Identify which applications deserve the most immediate attention based on Interview Status and Application Status. Include specific suggested actions (e.g. send thank-you email, follow up with recruiter, prepare for next round).\n" +
    "2. LOW SATISFACTION FLAGS: Flag any roles where Personal Satisfaction is 1 or 2. Give an honest recommendation on whether to keep pursuing each one.\n" +
    "3. NEXT STEPS BY COMPANY: For each company with an active application, suggest one concrete next action.\n" +
    "4. PATTERNS & INSIGHTS: Identify trends across the applications -- e.g. which industries or roles are getting more traction, gaps in the pipeline, or missing recruiter contact info.\n" +
    "5. OVERALL STRATEGY: Give one overarching recommendation to improve the student's chances of receiving an offer.\n" +
    "Be specific, encouraging, and direct. Use plain text only. 500 words max."
  );
}

// ---------------------------------------------------------------------------
// Read settings sheet
// ---------------------------------------------------------------------------

/**
 * Returns { privateKey, baseUrl } or null if validation fails (alert already shown).
 */
function readSettings(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: \"" + SETTINGS_SHEET + "\"",
      "Please create a sheet named \"" + SETTINGS_SHEET + "\" with your license key in cell " + PRIVATE_KEY_CELL + ".\n\nSee the README for the full layout."
    );
    return null;
  }

  var privateKey = String(sheet.getRange(PRIVATE_KEY_CELL).getValue()).trim();
  if (!privateKey) {
    showError(
      "License key not found",
      "Cell " + PRIVATE_KEY_CELL + " on the \"" + SETTINGS_SHEET + "\" sheet is empty.\n\nPaste your OptiSheets license key there and try again."
    );
    return null;
  }

  var baseUrl = getScriptProperty("OPTISHEETS_BASE_URL");
  if (!baseUrl) {
    showError(
      "Backend URL not configured",
      "Go to Extensions → Apps Script → Project Settings → Script Properties\nand add:\n\n  Key:   OPTISHEETS_BASE_URL\n  Value: https://your-backend.vercel.app"
    );
    return null;
  }

  // Strip trailing slash so we can always append paths cleanly
  baseUrl = baseUrl.replace(/\/+$/, "");

  return { privateKey: privateKey, baseUrl: baseUrl };
}

// ---------------------------------------------------------------------------
// Read tracker rows
// ---------------------------------------------------------------------------

/**
 * Returns an array of row objects using exact column names from meta.json,
 * or null if the sheet is missing.
 * Skips the header row (row 1) and blank rows (no Company Name).
 *
 * Expected columns (A–J):
 *   Company Name | Role/Position Title | Industry | Location |
 *   Application Status | Recruiter Name | Recruiter Email |
 *   Interview Status | Personal Satisfaction | Notes
 */
function readTrackerRows(ss) {
  var sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: \"" + TRACKER_SHEET + "\"",
      "Please create a sheet named \"" + TRACKER_SHEET + "\" with your application rows.\n\nRun \"Setup Tracker Sheet\" from the OptiSheets AI menu to configure it automatically."
    );
    return null;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // header only or empty

  // Read all data rows at once (columns A–J)
  var data = sheet.getRange(2, 1, lastRow - 1, NUM_COLUMNS).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var companyName = String(row[COL_COMPANY_NAME - 1]).trim();
    if (!companyName) continue; // skip blank rows

    var satisfactionRaw = row[COL_PERSONAL_SATISFACTION - 1];
    var satisfaction = "";
    if (satisfactionRaw !== "" && !isNaN(Number(satisfactionRaw))) {
      satisfaction = parseInt(Number(satisfactionRaw), 10);
    }

    var rowObj = {};
    rowObj["Company Name"]          = companyName;
    rowObj["Role/Position Title"]   = String(row[COL_ROLE_TITLE - 1]).trim()          || "";
    rowObj["Industry"]              = String(row[COL_INDUSTRY - 1]).trim()             || "";
    rowObj["Location"]              = String(row[COL_LOCATION - 1]).trim()             || "";
    rowObj["Application Status"]    = String(row[COL_APPLICATION_STATUS - 1]).trim()   || "";
    rowObj["Recruiter Name"]        = String(row[COL_RECRUITER_NAME - 1]).trim()       || "";
    rowObj["Recruiter Email"]       = String(row[COL_RECRUITER_EMAIL - 1]).trim()      || "";
    rowObj["Interview Status"]      = String(row[COL_INTERVIEW_STATUS - 1]).trim()     || "";
    rowObj["Personal Satisfaction"] = satisfaction;
    rowObj["Notes"]                 = String(row[COL_NOTES - 1]).trim()                || "";

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
  var url = baseUrl + "/api/get-recommendations";
  var options = {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true, // we handle all status codes ourselves
  };

  var response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    showError(
      "Network error",
      "Could not reach the OptiSheets backend.\n\nDetails: " + e.message +
      "\n\nCheck your internet connection and verify the OPTISHEETS_BASE_URL script property."
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
      "Unexpected server response (HTTP " + statusCode + ")",
      "The server returned a non-JSON response. This usually means the backend URL is wrong.\n\n" +
      "Raw response:\n" + bodyText.slice(0, 300)
    );
    return null;
  }

  if (body.success) return body;

  // Map status codes / error messages to friendly text
  var friendlyMessage = friendlyErrorMessage(statusCode, body.error || "Unknown error");
  showError("OptiSheets AI error (HTTP " + statusCode + ")", friendlyMessage);
  return null;
}

/**
 * Returns a human-friendly error string based on status code and raw error message.
 */
function friendlyErrorMessage(statusCode, rawError) {
  switch (statusCode) {
    case 401:
      return (
        "Your license key was not recognised.\n\n" +
        "Double-check the key in cell " + PRIVATE_KEY_CELL + " of the \"" + SETTINGS_SHEET + "\" sheet.\n\n" +
        "If you just purchased, make sure you copied the full key."
      );
    case 402:
      return (
        "You've run out of AI Credits.\n\n" +
        "Visit optisheets.com to top up your balance, then try again."
      );
    case 413:
      return (
        "Your tracker has too many applications for a single request.\n\n" +
        "Try removing old rows and re-run, or split your tracker into smaller batches."
      );
    case 400:
      return "Bad request: " + rawError + "\n\nThis is likely a bug — please contact support.";
    case 500:
      return (
        "The OptiSheets server encountered an internal error.\n\n" +
        "Please try again in a moment. If this keeps happening, contact support.\n\n" +
        "Details: " + rawError
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
  sheet.setTabColor("#A5D6A7");

  // ── Branded banner in row 1 ───────────────────────────────────────────────
  var bannerCell = sheet.getRange(1, 1, 1, 3);
  bannerCell.merge()
    .setValue("OptiSheets — AI Recommendations")
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM d, yyyy 'at' h:mm a");
  var cacheNote = result.cached ? " (cached — no credit used)" : "";

  // ── Header block (shifted down by 1 row to accommodate banner) ───────────
  var headerData = [
    ["OptiSheets AI — Internship Tracker Recommendations"],
    ["Generated: " + timestamp + cacheNote],
    ["Applications analysed: " + rowCount + "   |   Credits remaining: " + result.remaining_credits],
    [""], // blank row spacer
  ];

  sheet.getRange(2, 1, headerData.length, 1).setValues(headerData);

  // Title style
  sheet.getRange(2, 1).setFontSize(14).setFontWeight("bold");
  // Meta line style
  sheet.getRange(3, 1).setFontColor("#555555").setFontStyle("italic");
  sheet.getRange(4, 1).setFontColor("#555555");

  // ── Recommendations text (shifted down by 1 row) ──────────────────────────
  var outputRow  = 1 + headerData.length + 1; // row 6
  var outputCell = sheet.getRange(outputRow, 1);
  outputCell
    .setValue(result.output)
    .setWrap(true)
    .setVerticalAlignment("top")
    .setFontSize(11)
    .setBackground("#F1F8E9")
    .setBorder(null, true, null, null, null, null, "#388E3C", SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ── Column width so text is readable ─────────────────────────────────────
  sheet.setColumnWidth(1, 720);

  // ── Row height: auto-expand the output cell ───────────────────────────────
  sheet.setRowHeight(outputRow, 400);

  // Navigate the user to the output sheet
  ss.setActiveSheet(sheet);
  sheet.setActiveRange(outputCell);
}

// ---------------------------------------------------------------------------
// Setup help dialog
// ---------------------------------------------------------------------------

function showSetupHelp() {
  var html = HtmlService.createHtmlOutput(
    "<style>body{font-family:sans-serif;font-size:13px;padding:12px;line-height:1.6}" +
    "h2{margin-top:0}code{background:#f4f4f4;padding:2px 5px;border-radius:3px}</style>" +
    "<h2>OptiSheets Internship Tracker — Setup</h2>" +
    "<b>1. Script property (one-time)</b><br>" +
    "Extensions → Apps Script → Project Settings → Script Properties<br>" +
    "<code>OPTISHEETS_BASE_URL</code> → your backend URL<br><br>" +
    "<b>2. Settings sheet</b><br>" +
    "<table border='1' cellpadding='4' style='border-collapse:collapse'>" +
    "<tr><th>Cell</th><th>Value</th></tr>" +
    "<tr><td>B2</td><td>Your license key <b>(required)</b></td></tr>" +
    "</table><br>" +
    "<b>3. Tracker sheet columns</b><br>" +
    "Run <b>OptiSheets AI → Setup Tracker Sheet</b> to auto-create headers and dropdowns.<br><br>" +
    "<table border='1' cellpadding='4' style='border-collapse:collapse'>" +
    "<tr><th>Column</th><th>Contents</th></tr>" +
    "<tr><td>A</td><td>Company Name</td></tr>" +
    "<tr><td>B</td><td>Role/Position Title</td></tr>" +
    "<tr><td>C</td><td>Industry</td></tr>" +
    "<tr><td>D</td><td>Location</td></tr>" +
    "<tr><td>E</td><td>Application Status (Applying / In Progress / Applied)</td></tr>" +
    "<tr><td>F</td><td>Recruiter Name</td></tr>" +
    "<tr><td>G</td><td>Recruiter Email</td></tr>" +
    "<tr><td>H</td><td>Interview Status (None / Phone Screen / Video Interview / In-Person Interview)</td></tr>" +
    "<tr><td>I</td><td>Personal Satisfaction (1 = low interest, 5 = high interest)</td></tr>" +
    "<tr><td>J</td><td>Notes</td></tr>" +
    "</table>" +
    "<br><small>Row 1 is treated as a header and skipped automatically.</small>"
  )
    .setTitle("OptiSheets Setup Help")
    .setWidth(560)
    .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, "OptiSheets Setup Help");
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function showError(title, message) {
  SpreadsheetApp.getUi().alert("⚠️ " + title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function getScriptProperty(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? String(value).trim() : "";
}
