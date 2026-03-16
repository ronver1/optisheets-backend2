// =============================================================================
// OptiSheets — GPA & Grade Tracker
// Google Apps Script
// =============================================================================
//
// SETUP INSTRUCTIONS
// ------------------
// 1. In the Apps Script editor, go to Project Settings → Script Properties and add:
//      OPTISHEETS_BASE_URL  →  https://your-backend.vercel.app
//
// 2. Make sure your spreadsheet has these four sheets (exact names):
//      "Settings"          — configuration, license key, GPA conversion table,
//                            grade breakdown config
//      "GPA Tracker"       — one row per class across all semesters
//      "Grade Tracker"     — detailed score breakdown for the current class
//      "AI Recommendations"— auto-created on first AI run
//
// 3. Run "OptiSheets AI → Setup All Sheets" to create headers, dropdowns,
//    formulas, and formatting automatically.
//
// 4. Fill in Settings first (license key, GPA targets, grade breakdown config),
//    then enter class data in GPA Tracker, and scores in Grade Tracker.
// =============================================================================

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var TEMPLATE_ID      = "gpa-and-grade-tracker";
var SETTINGS_SHEET   = "Settings";
var GPA_SHEET        = "GPA Tracker";
var GRADE_SHEET      = "Grade Tracker";
var OUTPUT_SHEET     = "AI Recommendations";
var PRIVATE_KEY_CELL = "B3"; // Settings B3

// Brand colors
var COLOR_DARK_GREEN   = "#1B5E20";
var COLOR_GREEN        = "#388E3C";
var COLOR_LIGHT_GREEN  = "#A5D6A7";
var COLOR_VLIGHT_GREEN = "#F1F8E9";
var COLOR_WHITE        = "#FFFFFF";
var COLOR_BLACK        = "#212121";

// GPA conversion table entries (letter → points)
var GPA_TABLE = [
  ["A",  4.00],
  ["A-", 3.67],
  ["B+", 3.33],
  ["B",  3.00],
  ["B-", 2.67],
  ["C+", 2.33],
  ["C",  2.00],
  ["C-", 1.67],
  ["D+", 1.33],
  ["D",  1.00],
  ["F",  0.00],
];

var LETTER_GRADES = GPA_TABLE.map(function(r) { return r[0]; });

// Grade breakdown category rows (Settings D11:H15)
// Index: 0=Assignments, 1=Quizzes, 2=Exams, 3=Projects, 4=Participation
var GRADE_CATEGORIES = ["Assignments", "Quizzes", "Exams", "Projects", "Participation"];

// GPA Tracker column indices (1-based)
var GPA_COL_SEMESTER      = 1; // A
var GPA_COL_CLASS_NAME    = 2; // B
var GPA_COL_CREDITS       = 3; // C
var GPA_COL_LETTER_GRADE  = 4; // D
var GPA_COL_GRADE_POINTS  = 5; // E
var GPA_COL_WEIGHTED_PTS  = 6; // F
var GPA_COL_NOTES         = 7; // G
var GPA_NUM_COLS          = 7;

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("OptiSheets AI", [
      { name: "Get AI Recommendations",      functionName: "getAIRecommendations" },
      { name: "Setup All Sheets",            functionName: "setupAllSheets" },
      null, // separator
      { name: "Setup Settings Sheet",        functionName: "setupSettingsSheet" },
      { name: "Setup GPA Tracker Sheet",     functionName: "setupGPATrackerSheet" },
      { name: "Setup Grade Tracker Sheet",   functionName: "setupGradeTrackerSheet" },
      { name: "About / Setup Help",          functionName: "showSetupHelp" },
    ]);
}

function setupAllSheets() {
  setupSettingsSheet();
  setupGPATrackerSheet();
  setupGradeTrackerSheet();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast("All sheets configured! Fill in Settings first, then add your class data.", "OptiSheets AI", 8);
}

// ---------------------------------------------------------------------------
// Setup: Settings Sheet
// ---------------------------------------------------------------------------

function setupSettingsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOR_DARK_GREEN);

  // ── Block 1: Title banner A1:F1 ──────────────────────────────────────────
  sheet.getRange("A1:F1").merge()
    .setValue("OptiSheets — GPA & Grade Tracker Settings")
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  // ── Block 2: License key and cumulative config (rows 3–8) ────────────────
  var labels = [
    ["License Key",                ""],
    ["Target GPA",                 ""],
    ["Prior Credits Completed",    ""],
    ["Prior Cumulative GPA",       ""],
    ["Graduation Credits Required",""],
    ["Grading Scale",              ""],
  ];

  sheet.getRange(3, 1, 6, 2).setValues(labels);

  // Style label column
  var labelRange = sheet.getRange(3, 1, 6, 1);
  labelRange.setFontWeight("bold").setFontColor(COLOR_DARK_GREEN);

  // Data validation
  sheet.getRange("B4").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0.0, 4.0).setAllowInvalid(false)
      .setHelpText("Enter your target GPA (0.0 – 4.0)").build()
  );
  sheet.getRange("B5").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireWholeNumber().greaterThanOrEqualTo(0).setAllowInvalid(false)
      .setHelpText("Total credits completed before this semester").build()
  );
  sheet.getRange("B6").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0.0, 4.0).setAllowInvalid(false)
      .setHelpText("Your cumulative GPA before this semester (0.0 – 4.0)").build()
  );
  sheet.getRange("B7").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireWholeNumber().greaterThanOrEqualTo(0).setAllowInvalid(false)
      .setHelpText("Total credits needed to graduate").build()
  );
  sheet.getRange("B8").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["10-point scale", "7-point scale"], true)
      .setAllowInvalid(false).setHelpText("Select your school's grading scale").build()
  );

  // ── Block 3: GPA Conversion Table (A10:B21) ───────────────────────────────
  sheet.getRange("A10:B10")
    .setValues([["Letter Grade", "GPA Points"]])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  for (var i = 0; i < GPA_TABLE.length; i++) {
    var r  = 11 + i;
    var bg = (i % 2 === 0) ? COLOR_WHITE : COLOR_VLIGHT_GREEN;
    sheet.getRange(r, 1).setValue(GPA_TABLE[i][0]).setHorizontalAlignment("center");
    sheet.getRange(r, 2).setValue(GPA_TABLE[i][1]).setHorizontalAlignment("center");
    sheet.getRange(r, 1, 1, 2).setBackground(bg);
  }

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);

  // ── Block 4: Grade Breakdown Config Table (D10:H15+) ─────────────────────
  var gbHeaders = ["Category", "Count", "Total Weight (%)", "Drop Lowest?", "Extra Credit?"];
  sheet.getRange(10, 4, 1, 5).setValues([gbHeaders])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  for (var j = 0; j < GRADE_CATEGORIES.length; j++) {
    var gr = 11 + j;
    sheet.getRange(gr, 4).setValue(GRADE_CATEGORIES[j]).setFontWeight("bold");

    // Count (col E) — whole number 0–30, except Participation locked to 1
    if (GRADE_CATEGORIES[j] === "Participation") {
      sheet.getRange(gr, 5).setValue(1).setFontColor("#757575").setFontStyle("italic");
    } else {
      sheet.getRange(gr, 5).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireWholeNumber().between(0, 30).setAllowInvalid(false)
          .setHelpText("How many " + GRADE_CATEGORIES[j] + "?").build()
      );
    }

    // Total Weight % (col F)
    sheet.getRange(gr, 6).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberBetween(0, 100).setAllowInvalid(false)
        .setHelpText("Total weight % for " + GRADE_CATEGORIES[j]).build()
    );

    // Drop Lowest? (col G) — no dropdown for Participation
    if (GRADE_CATEGORIES[j] === "Participation") {
      sheet.getRange(gr, 7).setValue("N/A").setFontColor("#757575").setFontStyle("italic");
    } else {
      sheet.getRange(gr, 7).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(["Yes", "No"], true).setAllowInvalid(false).build()
      );
    }

    // Extra Credit? (col H)
    sheet.getRange(gr, 8).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Yes", "No"], true).setAllowInvalid(false).build()
    );
  }

  // Total Weight validation row (row 16)
  sheet.getRange(16, 4).setValue("Total Weight").setFontWeight("bold");
  sheet.getRange(16, 6).setFormula("=SUM(F11:F15)").setFontWeight("bold");

  // Conditional format F16: green if 100, red otherwise
  var f16Range = sheet.getRange("F16");
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(100).setBackground("#C8E6C9").setRanges([f16Range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(100).setBackground("#FFCDD2").setRanges([f16Range]).build(),
  ]);

  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);

  // Instructional hint in A2
  sheet.getRange("A2")
    .setValue("Paste your license key in B3. Fill Settings before running Setup.")
    .setFontColor("#757575")
    .setFontStyle("italic");
}

// ---------------------------------------------------------------------------
// Setup: GPA Tracker Sheet
// ---------------------------------------------------------------------------

function setupGPATrackerSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GPA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GPA_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOR_GREEN);

  // ── Row 1: Title banner A1:H1 ─────────────────────────────────────────────
  sheet.getRange(1, 1, 1, 8).merge()
    .setValue("OptiSheets — GPA Tracker")
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  // ── Row 2: Cumulative summary bar ────────────────────────────────────────
  // Semester GPA (cols A-B), Cumulative GPA (cols C-D), Credits (cols E-F), Target (cols G-H)
  sheet.getRange("A2:B2").merge()
    .setFormula(
      '=IF(SUMPRODUCT(IF(ISNUMBER(E4:E53),C4:C53,0))=0,"Semester GPA: —",' +
      '"Semester GPA: "&TEXT(IFERROR(SUMPRODUCT((C4:C53)*(IF(ISNUMBER(E4:E53),E4:E53,0)))' +
      '/SUMPRODUCT(IF(ISNUMBER(E4:E53),C4:C53,0)),0),"0.00"))'
    )
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("C2:D2").merge()
    .setFormula(
      '=IF(AND(Settings!B5="",Settings!B6="",SUMPRODUCT(IF(ISNUMBER(E4:E53),C4:C53,0))=0),' +
      '"Cumulative GPA: —",' +
      '"Cumulative GPA: "&TEXT(IFERROR((IFERROR(VALUE(Settings!B6),0)*IFERROR(VALUE(Settings!B5),0)' +
      '+SUMPRODUCT((C4:C53)*(IF(ISNUMBER(E4:E53),E4:E53,0))))' +
      '/(IFERROR(VALUE(Settings!B5),0)+SUMPRODUCT(IF(ISNUMBER(E4:E53),C4:C53,0))),0),"0.00"))'
    )
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("E2:F2").merge()
    .setFormula(
      '="Credits: "&TEXT(IFERROR(VALUE(Settings!B5),0)+SUMPRODUCT(IF(ISNUMBER(E4:E53),C4:C53,0)),"0")' +
      '&" / "&IF(Settings!B7="","—",Settings!B7)'
    )
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("G2:H2").merge()
    .setFormula('="Target GPA: "&IF(Settings!B4="","—",Settings!B4)')
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.setRowHeight(2, 30);

  // ── Row 3: Column headers ─────────────────────────────────────────────────
  var headers = ["Semester", "Class Name", "Credit Hours", "Letter Grade",
                 "Grade Points", "Weighted Points", "Notes"];
  sheet.getRange(3, 1, 1, GPA_NUM_COLS).setValues([headers])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(3, 30);

  // ── Data rows 4–53 ────────────────────────────────────────────────────────
  for (var r = 4; r <= 53; r++) {
    var bg = (r % 2 === 0) ? COLOR_WHITE : COLOR_VLIGHT_GREEN;
    sheet.getRange(r, 1, 1, GPA_NUM_COLS).setBackground(bg);
    sheet.setRowHeight(r, 26);

    // Grade Points formula (col E)
    sheet.getRange(r, GPA_COL_GRADE_POINTS)
      .setFormula('=IF(D' + r + '="","",IFERROR(VLOOKUP(D' + r + ',Settings!$A$11:$B$21,2,FALSE),"?"))')
      .setHorizontalAlignment("center");

    // Weighted Points formula (col F)
    sheet.getRange(r, GPA_COL_WEIGHTED_PTS)
      .setFormula('=IF(E' + r + '="","",E' + r + '*C' + r + ')')
      .setHorizontalAlignment("center");
  }

  // Protect formula columns (advisory — set background to indicate locked)
  sheet.getRange(4, GPA_COL_GRADE_POINTS, 50, 1).setBackground("#F9FBE7").setFontStyle("italic");
  sheet.getRange(4, GPA_COL_WEIGHTED_PTS, 50, 1).setBackground("#F9FBE7").setFontStyle("italic");

  // ── Data validation: Credit Hours (col C) ────────────────────────────────
  sheet.getRange(4, GPA_COL_CREDITS, 50, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireWholeNumber().between(1, 6).setAllowInvalid(false)
      .setHelpText("Enter credit hours (1–6)").build()
  );

  // ── Data validation: Letter Grade dropdown (col D) ───────────────────────
  sheet.getRange(4, GPA_COL_LETTER_GRADE, 50, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(LETTER_GRADES, true).setAllowInvalid(false)
      .setHelpText("Select a letter grade").build()
  );

  // ── Conditional formatting: Letter Grade (col D) ─────────────────────────
  var gradeCfRange = sheet.getRange(4, GPA_COL_LETTER_GRADE, 50, 1);
  var cfRules = [];

  // A, A- → soft green
  ["A", "A-"].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(g).setBackground("#C8E6C9").setRanges([gradeCfRange]).build());
  });
  // B+, B, B- → soft yellow
  ["B+", "B", "B-"].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(g).setBackground("#FFF9C4").setRanges([gradeCfRange]).build());
  });
  // C+, C, C- → soft orange
  ["C+", "C", "C-"].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(g).setBackground("#FFCCBC").setRanges([gradeCfRange]).build());
  });
  // D+, D, F → soft red
  ["D+", "D", "F"].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(g).setBackground("#FFCDD2").setRanges([gradeCfRange]).build());
  });
  sheet.setConditionalFormatRules(cfRules);

  // ── What-If GPA Simulator (rows 56–62) ───────────────────────────────────
  sheet.getRange(55, 1, 1, GPA_NUM_COLS).merge()
    .setValue("What-If GPA Simulator")
    .setBackground(COLOR_LIGHT_GREEN)
    .setFontColor(COLOR_DARK_GREEN)
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center");

  var simHeaders = ["Class Name", "Current Grade", "Hypothetical Grade", "Credit Hours",
                    "Impact on Cumulative GPA", "", ""];
  sheet.getRange(56, 1, 1, GPA_NUM_COLS).setValues([simHeaders])
    .setBackground(COLOR_VLIGHT_GREEN).setFontWeight("bold");

  for (var s = 57; s <= 59; s++) {
    sheet.getRange(s, 1).setFontStyle("italic").setFontColor("#757575")
      .setValue("Class " + (s - 56));
    sheet.getRange(s, 4).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireWholeNumber().between(1, 6).setAllowInvalid(false).build()
    );
    // Hypothetical grade dropdown
    sheet.getRange(s, 3).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(LETTER_GRADES, true).setAllowInvalid(false).build()
    );
    // Simulated cumulative GPA formula
    sheet.getRange(s, 5).setFormula(
      '=IFERROR(LET(priorC,IFERROR(VALUE(Settings!B5),0),priorG,IFERROR(VALUE(Settings!B6),0),' +
      'semC,SUMPRODUCT(IF(ISNUMBER(E4:E53),C4:C53,0)),' +
      'semWP,SUMPRODUCT((C4:C53)*(IF(ISNUMBER(E4:E53),E4:E53,0))),' +
      'adjC,IF(D' + s + '="",0,IFERROR(VALUE(D' + s + '),0)),' +
      'adjOldGP,IF(B' + s + '="",0,IFERROR(VLOOKUP(B' + s + ',Settings!$A$11:$B$21,2,FALSE),0)),' +
      'adjNewGP,IF(C' + s + '="",0,IFERROR(VLOOKUP(C' + s + ',Settings!$A$11:$B$21,2,FALSE),0)),' +
      '(priorC*priorG+semWP-adjC*adjOldGP+adjC*adjNewGP)/(priorC+semC)),"—")'
    ).setFontWeight("bold").setFontColor(COLOR_DARK_GREEN);
  }

  sheet.getRange(61, 1, 1, 3).merge()
    .setValue("Enter hypothetical grades above to see how your cumulative GPA would change.")
    .setFontColor("#757575").setFontStyle("italic");

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(GPA_COL_SEMESTER,     110);
  sheet.setColumnWidth(GPA_COL_CLASS_NAME,   200);
  sheet.setColumnWidth(GPA_COL_CREDITS,      100);
  sheet.setColumnWidth(GPA_COL_LETTER_GRADE, 120);
  sheet.setColumnWidth(GPA_COL_GRADE_POINTS, 110);
  sheet.setColumnWidth(GPA_COL_WEIGHTED_PTS, 130);
  sheet.setColumnWidth(GPA_COL_NOTES,        220);

  // ── Freeze rows 1–3 ───────────────────────────────────────────────────────
  sheet.setFrozenRows(3);

  ss.toast("GPA Tracker sheet is ready! Add class rows starting at row 4.", "OptiSheets AI", 6);
}

// ---------------------------------------------------------------------------
// Setup: Grade Tracker Sheet
// ---------------------------------------------------------------------------

function setupGradeTrackerSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GRADE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GRADE_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();
  sheet.clearNotes();
  sheet.setTabColor(COLOR_LIGHT_GREEN);

  // ── Row 1: Title banner A1:F1 ─────────────────────────────────────────────
  sheet.getRange(1, 1, 1, 6).merge()
    .setValue("OptiSheets — Grade Tracker")
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  // ── Row 2: Class selector ─────────────────────────────────────────────────
  sheet.getRange("A2").setValue("Select Class:").setFontWeight("bold").setFontColor(COLOR_DARK_GREEN);
  sheet.getRange("B2").setValue("").setFontStyle("italic"); // student fills in
  sheet.getRange("C2").setValue("Grading Scale:").setFontWeight("bold").setFontColor(COLOR_DARK_GREEN);
  sheet.getRange("D2").setFormula("=Settings!B8"); // pulls from Settings

  // ── Row 3: Live grade display bar ─────────────────────────────────────────
  // Weighted grade formula requires score data which is built dynamically below;
  // we use placeholder formulas referencing the data range that will be built.
  // Actual scores will live in col D starting at row 6+.
  // We'll compute weighted grade from the category summary rows.
  // For now set static labels; formulas updated after category rows are written.
  sheet.getRange("A3").setValue("Current Grade:").setFontWeight("bold");
  sheet.getRange("B3").setValue("").setFontWeight("bold"); // formula set after rows built
  sheet.getRange("C3").setValue("Letter Grade:").setFontWeight("bold");
  sheet.getRange("D3").setValue("").setFontWeight("bold"); // formula set after rows built
  sheet.getRange("E3").setValue("Graded So Far:").setFontWeight("bold");
  sheet.getRange("F3").setValue("").setFontWeight("bold"); // formula set after rows built
  sheet.getRange(3, 1, 1, 6)
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE)
    .setVerticalAlignment("middle");
  sheet.setRowHeight(3, 28);

  // ── Row 4: Grade Needed bar ───────────────────────────────────────────────
  sheet.getRange("A4").setValue("Grade Needed for Target GPA:").setFontWeight("bold")
    .setFontColor(COLOR_DARK_GREEN);
  sheet.getRange("B4").setValue("Fill in scores to calculate").setFontStyle("italic")
    .setFontColor("#757575");
  sheet.getRange(4, 1, 1, 6).setBackground(COLOR_VLIGHT_GREEN);
  sheet.setRowHeight(4, 26);

  // ── Row 5: Column headers ─────────────────────────────────────────────────
  var colHeaders = ["Category", "Item Name", "Weight per Item (%)", "Your Score (%)",
                    "Weighted Score", "Status"];
  sheet.getRange(5, 1, 1, 6).setValues([colHeaders])
    .setBackground(COLOR_DARK_GREEN).setFontColor(COLOR_WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(5, 28);

  // ── Read category config from Settings ───────────────────────────────────
  // Settings!E11:H15 = Count, Weight, Drop Lowest?, Extra Credit? per category
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  var categoryConfig = {};

  if (settingsSheet) {
    var configData = settingsSheet.getRange(11, 5, 5, 4).getValues();
    for (var ci = 0; ci < GRADE_CATEGORIES.length; ci++) {
      var cat = GRADE_CATEGORIES[ci];
      categoryConfig[cat] = {
        count:       cat === "Participation" ? 1 : (Number(configData[ci][0]) || 3),
        weight:      Number(configData[ci][1]) || 0,
        dropLowest:  String(configData[ci][2]).toLowerCase() === "yes",
        extraCredit: String(configData[ci][3]).toLowerCase() === "yes",
      };
    }
  } else {
    // Defaults if Settings not yet configured
    var defaults = {
      "Assignments":   { count: 5,  weight: 25, dropLowest: false, extraCredit: false },
      "Quizzes":       { count: 5,  weight: 20, dropLowest: false, extraCredit: false },
      "Exams":         { count: 3,  weight: 35, dropLowest: false, extraCredit: false },
      "Projects":      { count: 2,  weight: 15, dropLowest: false, extraCredit: false },
      "Participation": { count: 1,  weight: 5,  dropLowest: false, extraCredit: false },
    };
    categoryConfig = defaults;
  }

  // ── Build category sections dynamically ──────────────────────────────────
  var currentRow = 6;
  var allScoreCells = []; // for conditional formatting
  var categorySummaryRows = []; // track { category, summaryRow, weight, count, dropLowest }

  for (var catIdx = 0; catIdx < GRADE_CATEGORIES.length; catIdx++) {
    var catName = GRADE_CATEGORIES[catIdx];
    var cfg     = categoryConfig[catName] || { count: 1, weight: 0, dropLowest: false, extraCredit: false };
    var count   = Math.max(1, cfg.count);

    // Category sub-header row
    sheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue(catName)
      .setBackground(COLOR_GREEN)
      .setFontColor(COLOR_WHITE)
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(currentRow, 24);
    currentRow++;

    var itemStartRow = currentRow;

    // Per-item rows
    for (var itemIdx = 0; itemIdx < count; itemIdx++) {
      var itemLabel = catName === "Participation"
        ? "Participation"
        : catName.replace(/s$/, "") + " " + (itemIdx + 1);

      sheet.getRange(currentRow, 1).setValue(catName).setFontColor("#555555").setFontStyle("italic");
      sheet.getRange(currentRow, 2).setValue(itemLabel);

      // Weight per item = total weight / count (formula referencing Settings)
      var settingsWeightRef  = "Settings!$F$" + (11 + catIdx);
      var settingsCountRef   = catName === "Participation" ? "1" : "Settings!$E$" + (11 + catIdx);
      sheet.getRange(currentRow, 3)
        .setFormula("=IFERROR(" + settingsWeightRef + "/" + settingsCountRef + ",0)")
        .setHorizontalAlignment("center")
        .setFontColor("#555555");

      // Score input (col D)
      sheet.getRange(currentRow, 4).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, 150).setAllowInvalid(false)
          .setHelpText("Enter your score %").build()
      );

      // Weighted score (col E)
      sheet.getRange(currentRow, 5)
        .setFormula(
          '=IF(D' + currentRow + '="","",ROUND((C' + currentRow + '/100)*D' + currentRow + ',4))'
        )
        .setHorizontalAlignment("center")
        .setFontColor("#555555");

      // Status (col F)
      sheet.getRange(currentRow, 6)
        .setFormula(
          '=IF(D' + currentRow + '="","⏳ Pending","✅ Done")'
        )
        .setHorizontalAlignment("center");

      // Drop-lowest flag: mark last item in range if dropLowest
      if (cfg.dropLowest && itemIdx === count - 1) {
        sheet.getRange(currentRow, 1, 1, 6)
          .setFontStyle("italic").setFontColor("#9E9E9E");
        sheet.getRange(currentRow, 2)
          .setValue(itemLabel + " (drop lowest)");
      }

      // Alternating row background
      var itemBg = (itemIdx % 2 === 0) ? COLOR_WHITE : COLOR_VLIGHT_GREEN;
      sheet.getRange(currentRow, 1, 1, 6).setBackground(itemBg);
      sheet.setRowHeight(currentRow, 24);

      allScoreCells.push(currentRow);
      currentRow++;
    }

    var itemEndRow = currentRow - 1;

    // Extra credit row
    if (cfg.extraCredit) {
      sheet.getRange(currentRow, 1).setValue(catName).setFontColor("#555555").setFontStyle("italic");
      sheet.getRange(currentRow, 2).setValue("Extra Credit").setFontStyle("italic")
        .setFontColor(COLOR_DARK_GREEN).setFontWeight("bold");
      sheet.getRange(currentRow, 3).setValue(0).setHorizontalAlignment("center");
      sheet.getRange(currentRow, 4).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, 150).setAllowInvalid(false).build()
      );
      sheet.getRange(currentRow, 5)
        .setFormula('=IF(D' + currentRow + '="","",ROUND((C' + currentRow + '/100)*D' + currentRow + ',4))')
        .setHorizontalAlignment("center");
      sheet.getRange(currentRow, 6)
        .setFormula('=IF(D' + currentRow + '="","—","✅ EC")')
        .setHorizontalAlignment("center");
      sheet.getRange(currentRow, 1, 1, 6).setBackground("#E8F5E9");
      sheet.setRowHeight(currentRow, 24);
      allScoreCells.push(currentRow);
      itemEndRow = currentRow;
      currentRow++;
    }

    // Category summary row
    var avgFormula, dropNote;
    if (cfg.dropLowest && (itemEndRow - itemStartRow) >= 1) {
      // Exclude minimum score from average
      avgFormula =
        '=IFERROR(IF(COUNTA(D' + itemStartRow + ':D' + (itemEndRow - (cfg.extraCredit ? 1 : 0)) + ')=0,"—",' +
        'ROUND((SUM(D' + itemStartRow + ':D' + (itemEndRow - (cfg.extraCredit ? 1 : 0)) + ')' +
        '-MIN(IF(D' + itemStartRow + ':D' + (itemEndRow - (cfg.extraCredit ? 1 : 0)) + '<>"",D' +
        itemStartRow + ':D' + (itemEndRow - (cfg.extraCredit ? 1 : 0)) + ')))' +
        '/(MAX(COUNTA(D' + itemStartRow + ':D' + (itemEndRow - (cfg.extraCredit ? 1 : 0)) + ')-1,1)),1)),"—")';
      dropNote = " (drop lowest applied)";
    } else {
      avgFormula =
        '=IFERROR(IF(COUNTA(D' + itemStartRow + ':D' + itemEndRow + ')=0,"—",' +
        'ROUND(AVERAGEIF(D' + itemStartRow + ':D' + itemEndRow + ',"<>"),1)),"—")';
      dropNote = "";
    }

    sheet.getRange(currentRow, 1).setValue(catName);
    sheet.getRange(currentRow, 2).setValue(catName + " Average" + dropNote)
      .setFontWeight("bold").setFontColor(COLOR_DARK_GREEN);
    sheet.getRange(currentRow, 3).setFormula("=" + "Settings!$F$" + (11 + catIdx))
      .setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(currentRow, 4).setFormula(avgFormula)
      .setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(currentRow, 5)
      .setFormula(
        '=IFERROR(IF(D' + currentRow + '="—","—",ROUND((C' + currentRow + '/100)*D' + currentRow + ',4)),"—")'
      )
      .setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(currentRow, 6).setValue("").setFontWeight("bold");
    sheet.getRange(currentRow, 1, 1, 6)
      .setBackground(COLOR_LIGHT_GREEN).setFontColor(COLOR_DARK_GREEN);
    sheet.setRowHeight(currentRow, 26);

    categorySummaryRows.push({
      category: catName,
      summaryRow: currentRow,
      weightRef: "Settings!$F$" + (11 + catIdx),
    });

    currentRow++;
    // Spacer
    sheet.setRowHeight(currentRow, 8);
    currentRow++;
  }

  // ── Live grade formulas in row 3 (now that data rows exist) ──────────────
  // Weighted grade = sum of (category_avg * weight/100) over all categories
  var weightedGradeFormulaParts = categorySummaryRows.map(function(s) {
    return 'IF(D' + s.summaryRow + '="—",0,D' + s.summaryRow + '*(' + s.weightRef + '/100))';
  });
  var totalWeightedScoreFormula =
    '=IFERROR(ROUND((' + weightedGradeFormulaParts.join('+') + '),1),"—")';

  sheet.getRange("B3").setFormula(totalWeightedScoreFormula)
    .setFontWeight("bold").setFontColor(COLOR_WHITE);

  // Letter grade from weighted score using grading scale
  sheet.getRange("D3")
    .setFormula(
      '=IFERROR(IF(B3="—","—",IF(Settings!B8="7-point scale",' +
      'IF(B3>=93,"A",IF(B3>=90,"A-",IF(B3>=87,"B+",IF(B3>=83,"B",IF(B3>=80,"B-",' +
      'IF(B3>=77,"C+",IF(B3>=73,"C",IF(B3>=70,"C-",IF(B3>=67,"D+",IF(B3>=60,"D","F")))))))))),' +
      'IF(B3>=93,"A",IF(B3>=90,"A-",IF(B3>=87,"B+",IF(B3>=83,"B",IF(B3>=80,"B-",' +
      'IF(B3>=73,"C+",IF(B3>=70,"C",IF(B3>=67,"C-",IF(B3>=60,"D+",IF(B3>=57,"D","F")))))))))))),"—")'
    )
    .setFontWeight("bold").setFontColor(COLOR_WHITE);

  // Graded so far = sum of weights where a score has been entered
  var gradedSoFarParts = categorySummaryRows.map(function(s) {
    return 'IF(D' + s.summaryRow + '="—",0,' + s.weightRef + ')';
  });
  sheet.getRange("F3")
    .setFormula('=IFERROR(TEXT((' + gradedSoFarParts.join('+') + ')/100,"0%"),"—")')
    .setFontWeight("bold").setFontColor(COLOR_WHITE);

  // ── Conditional formatting: score input column D ──────────────────────────
  var cfRules = [];
  allScoreCells.forEach(function(rowNum) {
    var range = sheet.getRange(rowNum, 4);
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(90, 150).setBackground("#C8E6C9").setRanges([range]).build());
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(80, 89.99).setBackground("#FFF9C4").setRanges([range]).build());
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(70, 79.99).setBackground("#FFCCBC").setRanges([range]).build());
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(70).setBackground("#FFCDD2").setRanges([range]).build());
  });
  if (cfRules.length > 0) {
    sheet.setConditionalFormatRules(cfRules);
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 110);

  // ── Freeze rows 1–5 ───────────────────────────────────────────────────────
  sheet.setFrozenRows(5);

  ss.toast("Grade Tracker sheet is ready! Enter scores in column D.", "OptiSheets AI", 6);
}

// ---------------------------------------------------------------------------
// Main entry point: Get AI Recommendations
// ---------------------------------------------------------------------------

function getAIRecommendations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Reading your academic data…", "OptiSheets AI", 5);

  // ── 1. Read settings ──────────────────────────────────────────────────────
  var settings = readSettings(ss);
  if (!settings) return;

  // ── 2. Read GPA Tracker rows ──────────────────────────────────────────────
  var gpaRows = readGPATrackerRows(ss);
  if (gpaRows === null) return;

  // ── 3. Read Grade Tracker data ────────────────────────────────────────────
  var gradeData = readGradeTrackerData(ss);
  if (gradeData === null) return;

  if (gpaRows.length === 0 && Object.keys(gradeData).length === 0) {
    showError(
      "No data found",
      "Both GPA Tracker and Grade Tracker appear to be empty. " +
      "Add some class data and try again."
    );
    return;
  }

  // ── 4. Build settings payload ─────────────────────────────────────────────
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  var targetGpa     = settingsSheet ? Number(settingsSheet.getRange("B4").getValue()) || 0 : 0;
  var priorCredits  = settingsSheet ? Number(settingsSheet.getRange("B5").getValue()) || 0 : 0;
  var priorGpa      = settingsSheet ? Number(settingsSheet.getRange("B6").getValue()) || 0 : 0;
  var gradCredits   = settingsSheet ? Number(settingsSheet.getRange("B7").getValue()) || 0 : 0;

  // Compute semester GPA from GPA tracker rows
  var semWeightedSum = 0, semCreditSum = 0;
  gpaRows.forEach(function(row) {
    if (row["Grade Points"] !== "") {
      semWeightedSum += (Number(row["Grade Points"]) || 0) * (Number(row["Credit Hours"]) || 0);
      semCreditSum   += Number(row["Credit Hours"]) || 0;
    }
  });
  var semGpa = semCreditSum > 0 ? Math.round((semWeightedSum / semCreditSum) * 100) / 100 : 0;
  var cumGpa = (priorCredits + semCreditSum) > 0
    ? Math.round(((priorGpa * priorCredits + semWeightedSum) / (priorCredits + semCreditSum)) * 100) / 100
    : 0;

  var settingsPayload = {
    target_gpa:                 targetGpa,
    current_cumulative_gpa:     cumGpa,
    prior_credits_completed:    priorCredits,
    graduation_credits_required: gradCredits,
    current_semester_gpa:       semGpa,
    credits_this_semester:      semCreditSum,
  };

  // ── 5. Build full request payload ─────────────────────────────────────────
  var userData = {
    gpa_tracker:   gpaRows,
    grade_tracker: gradeData,
    settings:      settingsPayload,
  };

  var payload = {
    private_key:   settings.privateKey,
    template_id:   TEMPLATE_ID,
    user_data:     userData,
    system_prompt: buildSystemPromptString(),
  };

  // ── 6. Call backend ───────────────────────────────────────────────────────
  ss.toast("Calling OptiSheets AI (" + gpaRows.length + " classes)…", "OptiSheets AI", 30);
  var result = callBackend(settings.baseUrl, payload);
  if (!result) return;

  // ── 7. Write output ───────────────────────────────────────────────────────
  writeOutput(ss, result, gpaRows.length);
  ss.toast("Done! " + result.remaining_credits + " credit(s) remaining.", "OptiSheets AI", 8);
}

// ---------------------------------------------------------------------------
// System prompt string (mirrors prompt.js SYSTEM_PROMPT)
// ---------------------------------------------------------------------------

function buildSystemPromptString() {
  return (
    "You are a supportive but honest academic coach for college students. " +
    "You will receive a JSON object with three keys: " +
    "\"gpa_tracker\" (array of class rows with Semester, Class Name, Credit Hours, Letter Grade, Grade Points, Weighted Points, Notes), " +
    "\"grade_tracker\" (object with category breakdowns: assignments, quizzes, exams, projects, participation — each with count, weight, and an array of score entries), " +
    "and \"settings\" (target_gpa, current_cumulative_gpa, prior_credits_completed, graduation_credits_required, current_semester_gpa, credits_this_semester). " +
    "Analyze all three data sources together and provide a structured response with these exact labeled sections:\n" +
    "1. GPA SNAPSHOT: State the student's current semester GPA, cumulative GPA, and gap to their target GPA. If they are on track, acknowledge it. If not, be specific about what needs to change.\n" +
    "2. GRADE RESCUE ALERTS: Flag any class currently trending toward a C or below. Give specific, actionable advice for each flagged class including what to prioritize studying and what minimum scores are needed on remaining work.\n" +
    "3. PRIORITY CLASS STRATEGY: Using credit-hour weighting, tell the student exactly which classes deserve the most energy this semester and why. Recommend whether to aim for an A, A-, or B+ in each class based on current standing and GPA targets.\n" +
    "4. FINAL GRADE SIMULATIONS: Based on the Grade Tracker data, tell the student what scores they need on remaining assignments, quizzes, and exams to hit their target grade in the currently tracked class. Be specific with numbers.\n" +
    "5. BURNOUT CHECK: If the student is performing well across all classes, acknowledge it and advise on maintaining performance without over-studying. If they are struggling broadly, provide motivational but realistic advice.\n" +
    "6. SEMESTER REFLECTION & NEXT STEPS: Give 3 to 5 concrete action items the student should do in the next 7 days based on their data. End with a one-sentence motivational close.\n" +
    "Be specific, encouraging, and direct. Use plain text only. 600 words max."
  );
}

// ---------------------------------------------------------------------------
// Read settings
// ---------------------------------------------------------------------------

function readSettings(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: \"" + SETTINGS_SHEET + "\"",
      "Please create a sheet named \"" + SETTINGS_SHEET + "\" and run Setup All Sheets first."
    );
    return null;
  }

  var privateKey = String(sheet.getRange(PRIVATE_KEY_CELL).getValue()).trim();
  if (!privateKey) {
    showError(
      "License key not found",
      "Cell " + PRIVATE_KEY_CELL + " on the \"" + SETTINGS_SHEET + "\" sheet is empty.\n\n" +
      "Paste your OptiSheets license key there and try again."
    );
    return null;
  }

  var baseUrl = getScriptProperty("OPTISHEETS_BASE_URL");
  if (!baseUrl) {
    showError(
      "Backend URL not configured",
      "Go to Extensions → Apps Script → Project Settings → Script Properties\nand add:\n\n" +
      "  Key:   OPTISHEETS_BASE_URL\n  Value: https://your-backend.vercel.app"
    );
    return null;
  }

  baseUrl = baseUrl.replace(/\/+$/, "");
  return { privateKey: privateKey, baseUrl: baseUrl };
}

// ---------------------------------------------------------------------------
// Read GPA Tracker rows
// ---------------------------------------------------------------------------

function readGPATrackerRows(ss) {
  var sheet = ss.getSheetByName(GPA_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: \"" + GPA_SHEET + "\"",
      "Run \"OptiSheets AI → Setup All Sheets\" to create and configure all required sheets."
    );
    return null;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return []; // no data rows

  // Data starts at row 4; columns A–G (7 cols)
  var data = sheet.getRange(4, 1, lastRow - 3, GPA_NUM_COLS).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var row       = data[i];
    var className = String(row[GPA_COL_CLASS_NAME - 1]).trim();
    if (!className) continue;

    var obj = {};
    obj["Semester"]        = String(row[GPA_COL_SEMESTER - 1]).trim()     || "";
    obj["Class Name"]      = className;
    obj["Credit Hours"]    = Number(row[GPA_COL_CREDITS - 1]) || "";
    obj["Letter Grade"]    = String(row[GPA_COL_LETTER_GRADE - 1]).trim() || "";
    obj["Grade Points"]    = row[GPA_COL_GRADE_POINTS - 1] !== "" ? Number(row[GPA_COL_GRADE_POINTS - 1]) : "";
    obj["Weighted Points"] = row[GPA_COL_WEIGHTED_PTS - 1]  !== "" ? Number(row[GPA_COL_WEIGHTED_PTS - 1])  : "";
    obj["Notes"]           = String(row[GPA_COL_NOTES - 1]).trim()        || "";
    rows.push(obj);
  }

  return rows;
}

// ---------------------------------------------------------------------------
// Read Grade Tracker data
// ---------------------------------------------------------------------------

function readGradeTrackerData(ss) {
  var sheet = ss.getSheetByName(GRADE_SHEET);
  if (!sheet) {
    showError(
      "Missing sheet: \"" + GRADE_SHEET + "\"",
      "Run \"OptiSheets AI → Setup All Sheets\" to create and configure all required sheets."
    );
    return null;
  }

  var className    = String(sheet.getRange("B2").getValue()).trim() || "Unknown Class";
  var currentGrade = sheet.getRange("B3").getValue();
  var letterGrade  = String(sheet.getRange("D3").getValue()).trim() || "";
  var gradedSoFar  = String(sheet.getRange("F3").getValue()).trim() || "";

  // Read all visible data rows (rows 6 through last row)
  var lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    return {
      class_name:    className,
      current_grade: currentGrade,
      letter_grade:  letterGrade,
      graded_so_far: gradedSoFar,
      categories:    {},
    };
  }

  var rawData = sheet.getRange(6, 1, lastRow - 5, 6).getValues();
  var categories = {};

  rawData.forEach(function(row) {
    var cat    = String(row[0]).trim();
    var item   = String(row[1]).trim();
    var weight = row[2];
    var score  = row[3];
    var wScore = row[4];

    if (!cat || !item) return;

    // Skip sub-header merged rows (value = category name only, no item name)
    if (item === "") return;

    // Skip summary rows (contain "Average")
    if (item.indexOf("Average") !== -1) return;

    if (!categories[cat]) categories[cat] = { items: [] };
    categories[cat].items.push({
      name:           item,
      weight:         typeof weight === "number" ? Math.round(weight * 100) / 100 : 0,
      score:          score !== "" && score !== "—" ? Number(score) : null,
      weighted_score: wScore !== "" && wScore !== "—" ? Number(wScore) : null,
    });
  });

  return {
    class_name:    className,
    current_grade: typeof currentGrade === "number" ? Math.round(currentGrade * 10) / 10 : currentGrade,
    letter_grade:  letterGrade,
    graded_so_far: gradedSoFar,
    categories:    categories,
  };
}

// ---------------------------------------------------------------------------
// HTTP call
// ---------------------------------------------------------------------------

function callBackend(baseUrl, payload) {
  var url     = baseUrl + "/api/get-recommendations";
  var options = {
    method:             "post",
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
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

  var friendlyMessage = friendlyErrorMessage(statusCode, body.error || "Unknown error");
  showError("OptiSheets AI error (HTTP " + statusCode + ")", friendlyMessage);
  return null;
}

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
        "Your tracker has too many classes for a single request.\n\n" +
        "Try archiving older semester rows and re-run."
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

function writeOutput(ss, result, classCount) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOR_WHITE);

  // ── Title banner row 1 ────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, 3).merge()
    .setValue("OptiSheets — AI Recommendations")
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 36);

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM d, yyyy 'at' h:mm a");
  var cacheNote = result.cached ? " (cached — no credit used)" : "";

  var headerData = [
    ["OptiSheets AI — GPA & Grade Tracker Recommendations"],
    ["Generated: " + timestamp + cacheNote],
    ["Classes analysed: " + classCount + "   |   Credits remaining: " + result.remaining_credits],
    [""],
  ];

  sheet.getRange(2, 1, headerData.length, 1).setValues(headerData);
  sheet.getRange(2, 1).setFontSize(14).setFontWeight("bold");
  sheet.getRange(3, 1).setFontColor("#555555").setFontStyle("italic");
  sheet.getRange(4, 1).setFontColor("#555555");

  var outputRow  = 2 + headerData.length; // row 6
  var outputCell = sheet.getRange(outputRow, 1);
  outputCell
    .setValue(result.output)
    .setWrap(true)
    .setVerticalAlignment("top")
    .setFontSize(11)
    .setBackground(COLOR_VLIGHT_GREEN)
    .setBorder(null, true, null, null, null, null, COLOR_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);

  sheet.setColumnWidth(1, 720);
  sheet.setRowHeight(outputRow, 500);

  ss.setActiveSheet(sheet);
  sheet.setActiveRange(outputCell);
}

// ---------------------------------------------------------------------------
// Setup help dialog
// ---------------------------------------------------------------------------

function showSetupHelp() {
  var html = HtmlService.createHtmlOutput(
    "<style>" +
    "body{font-family:sans-serif;font-size:13px;padding:14px;line-height:1.6;color:#212121}" +
    "h2{margin-top:0;color:#1B5E20}h3{color:#388E3C;margin-top:16px;margin-bottom:4px}" +
    "code{background:#f4f4f4;padding:2px 5px;border-radius:3px}" +
    "table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:4px 8px}" +
    "th{background:#1B5E20;color:#fff}.badge{display:inline-block;background:#388E3C;color:#fff;" +
    "border-radius:3px;padding:1px 6px;font-size:11px;margin-right:4px}" +
    "</style>" +
    "<h2>OptiSheets GPA &amp; Grade Tracker — Setup Guide</h2>" +

    "<h3>Step 1 — Script property (one-time)</h3>" +
    "Extensions → Apps Script → Project Settings → Script Properties<br>" +
    "<code>OPTISHEETS_BASE_URL</code> → your backend URL (e.g. <code>https://your-app.vercel.app</code>)<br><br>" +

    "<h3>Step 2 — Fill in the Settings sheet</h3>" +
    "<table><tr><th>Cell</th><th>Value</th></tr>" +
    "<tr><td>B3</td><td>Your license key <b>(required)</b></td></tr>" +
    "<tr><td>B4</td><td>Target GPA (e.g. 3.5)</td></tr>" +
    "<tr><td>B5</td><td>Credits completed before this semester</td></tr>" +
    "<tr><td>B6</td><td>Cumulative GPA before this semester</td></tr>" +
    "<tr><td>B7</td><td>Credits required to graduate</td></tr>" +
    "<tr><td>B8</td><td>Grading scale (10-point or 7-point)</td></tr>" +
    "</table>" +
    "<br>Then fill in the <b>Grade Breakdown Config</b> table (columns D–H, rows 11–15) " +
    "to set how many assignments, quizzes, exams, projects, and participation grades you have, " +
    "plus the weight of each category (must total 100%).<br><br>" +

    "<h3>Step 3 — Run Setup All Sheets</h3>" +
    "Click <b>OptiSheets AI → Setup All Sheets</b>. This creates all 4 sheets with " +
    "headers, dropdowns, formulas, and formatting.<br><br>" +

    "<h3>Step 4 — Enter your class data</h3>" +
    "<span class='badge'>GPA Tracker</span> One row per class. Enter Semester, Class Name, " +
    "Credit Hours, and Letter Grade. Grade Points and Weighted Points calculate automatically.<br><br>" +
    "<span class='badge'>Grade Tracker</span> Enter your class name in B2, then type your " +
    "score percentage (0–100) in column D for each completed item. The current grade and " +
    "letter grade update in real time in row 3.<br><br>" +

    "<h3>Step 5 — Get AI Recommendations</h3>" +
    "Click <b>OptiSheets AI → Get AI Recommendations</b>. Results appear in the " +
    "<b>AI Recommendations</b> sheet with 6 labeled sections.<br><br>" +

    "<h3>Sheet overview</h3>" +
    "<table><tr><th>Sheet</th><th>Purpose</th></tr>" +
    "<tr><td>Settings</td><td>License key, GPA targets, conversion table, grade breakdown config</td></tr>" +
    "<tr><td>GPA Tracker</td><td>All classes across all semesters; cumulative GPA auto-calculated</td></tr>" +
    "<tr><td>Grade Tracker</td><td>Detailed score breakdown for one class at a time</td></tr>" +
    "<tr><td>AI Recommendations</td><td>Auto-created; shows the AI analysis</td></tr>" +
    "</table>"
  )
    .setTitle("OptiSheets GPA Tracker — Setup Help")
    .setWidth(600)
    .setHeight(580);

  SpreadsheetApp.getUi().showModalDialog(html, "OptiSheets GPA Tracker — Setup Help");
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function showError(title, message) {
  SpreadsheetApp.getUi().alert(
    "⚠️ " + title,
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getScriptProperty(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? String(value).trim() : "";
}
