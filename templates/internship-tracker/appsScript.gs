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
// 3. "AI Recommendations" will be created automatically on first run.
//
// SETTINGS SHEET LAYOUT
// ---------------------
//   A2: "License Key"       B2: <your private key>
//   A3: "Target Role"       B3: Software Engineering Intern   (optional)
//   A4: "Target Season"     B4: Summer 2026                   (optional)
//   A5: "Target Count"      B5: 3                             (optional, number)
//
// TRACKER SHEET LAYOUT
// --------------------
//   Row 1 = headers (skipped automatically)
//   Col A: Company
//   Col B: Role
//   Col C: Status   (Applied | Phone Screen | Technical Screen | Onsite |
//                    Final Round | Offer | Rejected | Withdrawn)
//   Col D: Applied Date  (any date format; stored as YYYY-MM-DD)
//   Col E: Notes         (free text, truncated server-side at 80 chars)
// =============================================================================

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Must match templates/internship-tracker/prompt.js → SYSTEM_PROMPT */
var SYSTEM_PROMPT =
  "You are a recruiting coach helping college students land internships. " +
  "Given a student's application tracker data, give concise actionable recommendations. " +
  "Cover: which applications to follow up on urgently, interview prep for active stages, " +
  "pipeline weak spots (low response rate, stalled stages), and concrete next steps. " +
  "Be specific. Use short labeled sections. Plain text only. 400 words max.";

var TEMPLATE_ID       = "internship-tracker";
var SETTINGS_SHEET    = "Settings";
var TRACKER_SHEET     = "Tracker";
var OUTPUT_SHEET      = "AI Recommendations";
var PRIVATE_KEY_CELL  = "B2";
var TARGET_ROLE_CELL  = "B3";
var TARGET_SEASON_CELL = "B4";
var TARGET_COUNT_CELL = "B5";

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("OptiSheets AI", [
      { name: "Get AI Recommendations", functionName: "getAIRecommendations" },
      null, // separator
      { name: "About / Setup Help",     functionName: "showSetupHelp" },
    ]);
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
  var applications = readTrackerRows(ss);
  if (applications === null) return; // readTrackerRows already showed an alert
  if (applications.length === 0) {
    showError(
      "No applications found",
      "Your Tracker sheet appears to be empty. Add some rows and try again."
    );
    return;
  }

  // ── 3. Build request payload ──────────────────────────────────────────────
  var userData = {
    applications:  applications,
    targetRole:    settings.targetRole   || "internship",
    targetSeason:  settings.targetSeason || "",
    targetCount:   settings.targetCount  || null,
  };

  var payload = {
    private_key:   settings.privateKey,
    template_id:   TEMPLATE_ID,
    user_data:     userData,
    system_prompt: SYSTEM_PROMPT,
  };

  // ── 4. Call backend ───────────────────────────────────────────────────────
  ss.toast("Calling OptiSheets AI (" + applications.length + " applications)…", "OptiSheets AI", 30);

  var result = callBackend(settings.baseUrl, payload);
  if (!result) return; // callBackend already showed an alert

  // ── 5. Write output ───────────────────────────────────────────────────────
  writeOutput(ss, result, applications.length);
  ss.toast("Done! " + result.remaining_credits + " credit(s) remaining.", "OptiSheets AI", 8);
}

// ---------------------------------------------------------------------------
// Read settings sheet
// ---------------------------------------------------------------------------

/**
 * Returns { privateKey, baseUrl, targetRole, targetSeason, targetCount }
 * or null if validation fails (alert already shown).
 */
function readSettings(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: "" + SETTINGS_SHEET + """,
      "Please create a sheet named "" + SETTINGS_SHEET + "" with your license key in cell " + PRIVATE_KEY_CELL + ".\n\nSee the README for the full layout."
    );
    return null;
  }

  var privateKey = String(sheet.getRange(PRIVATE_KEY_CELL).getValue()).trim();
  if (!privateKey) {
    showError(
      "License key not found",
      "Cell " + PRIVATE_KEY_CELL + " on the "" + SETTINGS_SHEET + "" sheet is empty.\n\nPaste your OptiSheets license key there and try again."
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

  var targetRole    = String(sheet.getRange(TARGET_ROLE_CELL).getValue()).trim()   || "";
  var targetSeason  = String(sheet.getRange(TARGET_SEASON_CELL).getValue()).trim() || "";
  var targetCountRaw = sheet.getRange(TARGET_COUNT_CELL).getValue();
  var targetCount   = (targetCountRaw && !isNaN(Number(targetCountRaw)))
    ? parseInt(Number(targetCountRaw), 10)
    : null;

  return {
    privateKey:   privateKey,
    baseUrl:      baseUrl,
    targetRole:   targetRole,
    targetSeason: targetSeason,
    targetCount:  targetCount,
  };
}

// ---------------------------------------------------------------------------
// Read tracker rows
// ---------------------------------------------------------------------------

/**
 * Returns an array of application objects, or null if the sheet is missing.
 * Skips the header row (row 1) and blank rows (no company name).
 *
 * Expected columns:
 *   A: Company | B: Role | C: Status | D: Applied Date | E: Notes
 */
function readTrackerRows(ss) {
  var sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: "" + TRACKER_SHEET + """,
      "Please create a sheet named "" + TRACKER_SHEET + "" with your application rows.\n\nSee the README for the column layout."
    );
    return null;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // header only or empty

  // Read all data rows at once (A2:E to end)
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var applications = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var company = String(row[0]).trim();
    if (!company) continue; // skip blank rows

    var appliedDateRaw = row[3];
    var appliedDate    = formatDateCell(appliedDateRaw);

    var notes = String(row[4]).trim() || undefined;

    applications.push({
      company:     company,
      role:        String(row[1]).trim() || "",
      status:      String(row[2]).trim() || "",
      appliedDate: appliedDate || undefined,
      notes:       notes,
    });
  }

  return applications;
}

/**
 * Converts a spreadsheet date cell value to "YYYY-MM-DD" string.
 * Returns "" if the value isn't a recognisable date.
 */
function formatDateCell(value) {
  if (!value) return "";
  var d;
  if (value instanceof Date) {
    d = value;
  } else {
    d = new Date(value);
  }
  if (isNaN(d.getTime())) return "";
  var yyyy = d.getUTCFullYear();
  var mm   = String(d.getUTCMonth() + 1).padStart(2, "0");
  var dd   = String(d.getUTCDate()).padStart(2, "0");
  return yyyy + "-" + mm + "-" + dd;
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
        "Double-check the key in cell " + PRIVATE_KEY_CELL + " of the "" + SETTINGS_SHEET + "" sheet.\n\n" +
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
        "Try removing old Rejected/Withdrawn rows and re-run, or split your tracker into smaller batches."
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
function writeOutput(ss, result, appCount) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM d, yyyy 'at' h:mm a");
  var cacheNote = result.cached ? " (cached — no credit used)" : "";

  // ── Header block ──────────────────────────────────────────────────────────
  var headerData = [
    ["OptiSheets AI — Internship Tracker Recommendations"],
    ["Generated: " + timestamp + cacheNote],
    ["Applications analysed: " + appCount + "   |   Credits remaining: " + result.remaining_credits],
    [""], // blank row spacer
  ];

  sheet.getRange(1, 1, headerData.length, 1).setValues(headerData);

  // Title style
  sheet.getRange(1, 1).setFontSize(14).setFontWeight("bold");
  // Meta line style
  sheet.getRange(2, 1).setFontColor("#555555").setFontStyle("italic");
  sheet.getRange(3, 1).setFontColor("#555555");

  // ── Recommendations text ──────────────────────────────────────────────────
  var outputRow = headerData.length + 1; // row 5
  var outputCell = sheet.getRange(outputRow, 1);
  outputCell
    .setValue(result.output)
    .setWrap(true)
    .setVerticalAlignment("top")
    .setFontSize(11);

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
    "<tr><td>B3</td><td>Target role (e.g. <i>Software Engineering Intern</i>)</td></tr>" +
    "<tr><td>B4</td><td>Target season (e.g. <i>Summer 2026</i>)</td></tr>" +
    "<tr><td>B5</td><td>Number of offers you want (e.g. <i>3</i>)</td></tr>" +
    "</table><br>" +
    "<b>3. Tracker sheet columns</b><br>" +
    "<table border='1' cellpadding='4' style='border-collapse:collapse'>" +
    "<tr><th>Column</th><th>Contents</th></tr>" +
    "<tr><td>A</td><td>Company</td></tr>" +
    "<tr><td>B</td><td>Role</td></tr>" +
    "<tr><td>C</td><td>Status (Applied / Phone Screen / Technical Screen / Onsite / Final Round / Offer / Rejected / Withdrawn)</td></tr>" +
    "<tr><td>D</td><td>Applied Date</td></tr>" +
    "<tr><td>E</td><td>Notes</td></tr>" +
    "</table>" +
    "<br><small>Row 1 is treated as a header and skipped automatically.</small>"
  )
    .setTitle("OptiSheets Setup Help")
    .setWidth(540)
    .setHeight(480);

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
