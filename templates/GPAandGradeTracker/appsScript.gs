// =============================================================================
// OptiSheets — GPA & Grade Tracker
// Google Apps Script
// =============================================================================

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var TEMPLATE_ID      = 'gpa-and-grade-tracker';
var SETTINGS_SHEET   = 'Settings';
var GPA_SHEET        = 'GPA Tracker';
var GRADE_SHEET      = 'Grade Tracker';
var OUTPUT_SHEET     = 'AI Recommendations';
var PRIVATE_KEY_CELL = 'B3';

var COLOR_DARK_GREEN   = '#1B5E20';
var COLOR_GREEN        = '#388E3C';
var COLOR_LIGHT_GREEN  = '#A5D6A7';
var COLOR_VLIGHT_GREEN = '#F1F8E9';
var COLOR_WHITE        = '#FFFFFF';
var COLOR_BLACK        = '#212121';
var COLOR_GRAY         = '#757575';

var GPA_TABLE = [
  ['A',  4.00],
  ['A-', 3.67],
  ['B+', 3.33],
  ['B',  3.00],
  ['B-', 2.67],
  ['C+', 2.33],
  ['C',  2.00],
  ['C-', 1.67],
  ['D+', 1.33],
  ['D',  1.00],
  ['F',  0.00]
];

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu('OptiSheets AI', [
      { name: 'Get AI Recommendations',    functionName: 'getAIRecommendations' },
      { name: 'Setup All Sheets',          functionName: 'setupAllSheets' },
      null,
      { name: 'Setup Settings Sheet',      functionName: 'setupSettingsSheet' },
      { name: 'Setup GPA Tracker Sheet',   functionName: 'setupGPATrackerSheet' },
      { name: 'Setup Grade Tracker Sheet', functionName: 'setupGradeTrackerSheet' }
    ]);
}

// ---------------------------------------------------------------------------
// setupAllSheets
// ---------------------------------------------------------------------------

function setupAllSheets() {
  setupSettingsSheet();
  setupGPATrackerSheet();
  setupGradeTrackerSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('All sheets are set up and ready!', 'OptiSheets AI', 5);
}

// ---------------------------------------------------------------------------
// setupSettingsSheet
// ---------------------------------------------------------------------------

function setupSettingsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
  }

  // ── Save existing values before clearing ──────────────────────────────────
  var savedKey         = sheet.getRange('B3').getValue();
  var savedTargetGPA   = sheet.getRange('B4').getValue();
  var savedPriorCred   = sheet.getRange('B5').getValue();
  var savedPriorGPA    = sheet.getRange('B6').getValue();
  var savedGradCred    = sheet.getRange('B7').getValue();
  var savedScale       = sheet.getRange('B8').getValue();
  var savedGradeConfig = sheet.getRange('E11:H15').getValues(); // 5×4

  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOR_DARK_GREEN);

  // ── Row 1: title banner ───────────────────────────────────────────────────
  sheet.getRange('A1:F1').merge()
    .setValue('OptiSheets — GPA & Grade Tracker Settings')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── Row 2: hint ───────────────────────────────────────────────────────────
  sheet.getRange('A2')
    .setValue('Paste your license key in B3. Fill in grade config, then run Setup Grade Tracker Sheet.')
    .setFontStyle('italic')
    .setFontColor(COLOR_GRAY);

  // ── Rows 3–8: labels in column A ─────────────────────────────────────────
  var labels = [
    'License Key',
    'Target GPA',
    'Prior Credits Completed',
    'Prior Cumulative GPA',
    'Graduation Credits Required',
    'Grading Scale'
  ];
  for (var i = 0; i < labels.length; i++) {
    sheet.getRange(3 + i, 1)
      .setValue(labels[i])
      .setFontWeight('bold')
      .setFontColor(COLOR_DARK_GREEN);
  }

  // ── Restore B3:B8 ─────────────────────────────────────────────────────────
  sheet.getRange('B3').setValue(savedKey);
  sheet.getRange('B4').setValue(savedTargetGPA);
  sheet.getRange('B5').setValue(savedPriorCred);
  sheet.getRange('B6').setValue(savedPriorGPA);
  sheet.getRange('B7').setValue(savedGradCred);
  sheet.getRange('B8').setValue(savedScale);

  // ── Data validation B4:B8 — one statement per cell, no setHelpText ────────
  sheet.getRange('B4').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0.0, 4.0).setAllowInvalid(false).build()
  );
  sheet.getRange('B5').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 999).setAllowInvalid(false).build()
  );
  sheet.getRange('B6').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0.0, 4.0).setAllowInvalid(false).build()
  );
  sheet.getRange('B7').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 999).setAllowInvalid(false).build()
  );
  sheet.getRange('B8').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['10-point scale', '7-point scale'], true).setAllowInvalid(false).build()
  );

  // ── GPA Conversion Table rows 10–21, columns A–B ─────────────────────────
  sheet.getRange('A10:B10')
    .setValues([['Letter Grade', 'GPA Points']])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  for (var g = 0; g < GPA_TABLE.length; g++) {
    var gRow = 11 + g;
    var gBg  = (g % 2 === 0) ? COLOR_WHITE : COLOR_VLIGHT_GREEN;
    sheet.getRange(gRow, 1, 1, 2)
      .setValues([[GPA_TABLE[g][0], GPA_TABLE[g][1]]])
      .setBackground(gBg)
      .setHorizontalAlignment('center');
  }

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);

  // ── Grade Breakdown Config Table rows 10–16, columns D–H ─────────────────
  sheet.getRange('D10:H10')
    .setValues([['Category', 'Count', 'Total Weight (%)', 'Drop Lowest?', 'Extra Credit?']])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  var cats = ['Assignments', 'Quizzes', 'Exams', 'Projects', 'Participation'];
  for (var c = 0; c < cats.length; c++) {
    sheet.getRange(11 + c, 4).setValue(cats[c]);
  }

  // Participation count E15 = 1, gray italic, no validation
  sheet.getRange('E15').setValue(1).setFontColor(COLOR_GRAY).setFontStyle('italic');

  // Count validation E11:E14 — no setHelpText
  sheet.getRange('E11:E14').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 30).setAllowInvalid(false).build()
  );

  // Weight validation F11:F15 — no setHelpText
  sheet.getRange('F11:F15').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 100).setAllowInvalid(false).build()
  );

  // Drop Lowest dropdowns G11:G14 — no setHelpText
  sheet.getRange('G11:G14').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build()
  );
  // G15 = N/A gray italic
  sheet.getRange('G15').setValue('N/A').setFontColor(COLOR_GRAY).setFontStyle('italic');

  // Extra Credit dropdowns H11:H15 — no setHelpText
  sheet.getRange('H11:H15').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build()
  );

  // Row 16: total weight
  sheet.getRange('D16').setValue('Total Weight').setFontWeight('bold');
  sheet.getRange('F16').setFormula('=SUM(F11:F15)').setFontWeight('bold');

  // Conditional formatting on F16 only
  var f16Range = sheet.getRange('F16');
  var cfSettings = [];
  cfSettings.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(100).setBackground('#C8E6C9').setRanges([f16Range]).build());
  cfSettings.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(100).setBackground('#FFCDD2').setRanges([f16Range]).build());
  sheet.setConditionalFormatRules(cfSettings);

  // ── Restore saved grade config values E11:H15 ─────────────────────────────
  for (var ri = 0; ri < 5; ri++) {
    for (var ci = 0; ci < 4; ci++) {
      var val = savedGradeConfig[ri][ci];
      if (val !== '' && val !== null && val !== undefined) {
        // E15 (ri=4, ci=0) always stays 1
        if (!(ri === 4 && ci === 0)) {
          sheet.getRange(11 + ri, 5 + ci).setValue(val);
        }
      }
    }
  }

  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);
}

// ---------------------------------------------------------------------------
// setupGPATrackerSheet
// ---------------------------------------------------------------------------

function setupGPATrackerSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GPA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GPA_SHEET);
  }

  sheet.setTabColor(COLOR_GREEN);
  sheet.clearContents();
  sheet.clearFormats();

  // ── Row 1: title banner ───────────────────────────────────────────────────
  sheet.getRange('A1:H1').merge()
    .setValue('OptiSheets — GPA Tracker')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── Row 2: summary bar — four merged pairs ────────────────────────────────
  sheet.getRange('A2:B2').merge()
    .setValue('Semester GPA: —')
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('C2:D2').merge()
    .setValue('Cumulative GPA: —')
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E2:F2').merge()
    .setValue('Credits: — / —')
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('G2:H2').merge()
    .setValue('Target GPA: —')
    .setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 30);

  // ── Row 3: column headers ─────────────────────────────────────────────────
  sheet.getRange(3, 1, 1, 7)
    .setValues([['Semester', 'Class Name', 'Credit Hours', 'Letter Grade', 'Grade Points', 'Weighted Points', 'Notes']])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);

  // ── Data rows 4–53 — build E and F formula arrays in memory ──────────────
  var eFormulas = [];
  var fFormulas = [];
  for (var r = 4; r <= 53; r++) {
    var bg = (r % 2 === 0) ? COLOR_VLIGHT_GREEN : COLOR_WHITE;
    sheet.getRange(r, 1, 1, 7).setBackground(bg);
    sheet.setRowHeight(r, 26);
    eFormulas.push(['=IF(D' + r + '="","",IFERROR(VLOOKUP(D' + r + ',Settings!$A$11:$B$21,2,FALSE),"?"))']);
    fFormulas.push(['=IF(E' + r + '="","",E' + r + '*C' + r + ')']);
  }
  sheet.getRange(4, 5, 50, 1).setFormulas(eFormulas);
  sheet.getRange(4, 6, 50, 1).setFormulas(fFormulas);

  // Credit hours C4:C53 — single range call, no setHelpText
  sheet.getRange('C4:C53').setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(1, 6).setAllowInvalid(false).build()
  );

  // Letter grade D4:D53 — single range call
  var gradeLetters = GPA_TABLE.map(function(g) { return g[0]; });
  sheet.getRange('D4:D53').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(gradeLetters, true).setAllowInvalid(false).build()
  );

  // ── Conditional formatting D4:D53 — one setConditionalFormatRules call ────
  var dRange  = sheet.getRange('D4:D53');
  var cfRules = [];
  ['A', 'A-'].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(g).setBackground('#C8E6C9').setRanges([dRange]).build());
  });
  ['B+', 'B', 'B-'].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(g).setBackground('#FFF9C4').setRanges([dRange]).build());
  });
  ['C+', 'C', 'C-'].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(g).setBackground('#FFCCBC').setRanges([dRange]).build());
  });
  ['D+', 'D', 'F'].forEach(function(g) {
    cfRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(g).setBackground('#FFCDD2').setRanges([dRange]).build());
  });
  sheet.setConditionalFormatRules(cfRules);

  // ── What-If Simulator rows 55–63 ─────────────────────────────────────────
  sheet.getRange('A55:G55').merge()
    .setValue('What-If GPA Simulator')
    .setBackground(COLOR_LIGHT_GREEN)
    .setFontColor(COLOR_DARK_GREEN)
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(56, 1, 1, 7)
    .setValues([['Class Name', 'Current Letter Grade', 'Hypothetical Grade', 'Credit Hours', 'Simulated Cumulative GPA', '', '']])
    .setBackground(COLOR_VLIGHT_GREEN)
    .setFontWeight('bold');

  for (var sim = 0; sim < 6; sim++) {
    var simRow = 57 + sim;
    sheet.getRange(simRow, 1).setValue('Class ' + (sim + 1)).setFontColor(COLOR_GRAY).setFontStyle('italic');

    // Col C: hypothetical grade dropdown — no setHelpText
    sheet.getRange(simRow, 3).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(gradeLetters, true).setAllowInvalid(false).build()
    );

    // Col D: credit hours — no setHelpText
    sheet.getRange(simRow, 4).setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(1, 6).setAllowInvalid(false).build()
    );

    // Col E: simulated cumulative GPA — no LET(), nested IFERROR(VALUE(...))
    // Accounts for: prior credits×GPA, semester weighted points,
    // and swapping this class's current grade (col B) for hypothetical (col C)
    var sf =
      '=IF(OR(C' + simRow + '="",D' + simRow + '=""),"",IFERROR(' +
        '(' +
          'IFERROR(VALUE(Settings!$B$6),0)*IFERROR(VALUE(Settings!$B$5),0)' +
          '+SUMPRODUCT(($D$4:$D$53<>"")*IFERROR(VALUE($F$4:$F$53),0))' +
          '-IFERROR(VLOOKUP(B' + simRow + ',Settings!$A$11:$B$21,2,FALSE),0)*IFERROR(VALUE(D' + simRow + '),0)' +
          '+IFERROR(VLOOKUP(C' + simRow + ',Settings!$A$11:$B$21,2,FALSE),0)*IFERROR(VALUE(D' + simRow + '),0)' +
        ')/' +
        'IF(' +
          'IFERROR(VALUE(Settings!$B$5),0)+SUMPRODUCT(($D$4:$D$53<>"")*IFERROR(VALUE($C$4:$C$53),0))=0,' +
          '1,' +
          'IFERROR(VALUE(Settings!$B$5),0)+SUMPRODUCT(($D$4:$D$53<>"")*IFERROR(VALUE($C$4:$C$53),0))' +
        '),' +
      '"—"))';
    sheet.getRange(simRow, 5).setFormula(sf);
  }

  // Row 63: note
  sheet.getRange('A63:C63').merge()
    .setValue('Enter hypothetical grades above to see how your cumulative GPA would change.')
    .setFontColor(COLOR_GRAY).setFontStyle('italic');

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 220);

  sheet.setFrozenRows(3);

  // Populate summary bar with live values
  refreshGPASummaryBar();
}

// ---------------------------------------------------------------------------
// refreshGPASummaryBar
// ---------------------------------------------------------------------------

function refreshGPASummaryBar() {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var gpaSheet      = ss.getSheetByName(GPA_SHEET);
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!gpaSheet) return;

  var targetGPA = 0, priorCred = 0, priorGPA = 0, gradCred = 0;
  if (settingsSheet) {
    targetGPA = parseFloat(settingsSheet.getRange('B4').getValue()) || 0;
    priorCred = parseFloat(settingsSheet.getRange('B5').getValue()) || 0;
    priorGPA  = parseFloat(settingsSheet.getRange('B6').getValue()) || 0;
    gradCred  = parseFloat(settingsSheet.getRange('B7').getValue()) || 0;
  }

  var data        = gpaSheet.getRange(4, 1, 50, 7).getValues();
  var semCredits  = 0;
  var semWeighted = 0;

  for (var i = 0; i < data.length; i++) {
    var className = String(data[i][1]).trim();
    if (!className) continue;
    var credits  = parseFloat(data[i][2]) || 0;
    var gradeStr = String(data[i][3]).trim();
    var pts      = getGPAPoints(gradeStr);
    if (credits > 0 && pts !== null) {
      semCredits  += credits;
      semWeighted += credits * pts;
    }
  }

  var semGPA       = semCredits  > 0 ? (semWeighted / semCredits) : null;
  var totalCredits = priorCred   + semCredits;
  var totalWtd     = priorGPA * priorCred + semWeighted;
  var cumGPA       = totalCredits > 0 ? (totalWtd / totalCredits) : null;

  var semGPAText = semGPA !== null ? semGPA.toFixed(2) : '—';
  var cumGPAText = cumGPA !== null ? cumGPA.toFixed(2) : '—';
  var credText   = totalCredits > 0
    ? totalCredits + ' / ' + (gradCred > 0 ? gradCred : '—')
    : '— / —';
  var tgtText    = targetGPA > 0 ? targetGPA.toFixed(2) : '—';

  gpaSheet.getRange('A2:B2').setValue('Semester GPA: '   + semGPAText);
  gpaSheet.getRange('C2:D2').setValue('Cumulative GPA: ' + cumGPAText);
  gpaSheet.getRange('E2:F2').setValue('Credits: '        + credText);
  gpaSheet.getRange('G2:H2').setValue('Target GPA: '     + tgtText);
}

function getGPAPoints(letterGrade) {
  for (var i = 0; i < GPA_TABLE.length; i++) {
    if (GPA_TABLE[i][0] === letterGrade) return GPA_TABLE[i][1];
  }
  return null;
}

// ---------------------------------------------------------------------------
// setupGradeTrackerSheet
// ---------------------------------------------------------------------------

function setupGradeTrackerSheet() {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var sheet         = ss.getSheetByName(GRADE_SHEET);
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(GRADE_SHEET);
  }
  if (!settingsSheet) {
    ss.toast('Settings sheet not found. Please run Setup Settings Sheet first.', 'OptiSheets AI', 5);
    return;
  }

  sheet.setTabColor(COLOR_LIGHT_GREEN);
  sheet.clearContents();
  sheet.clearConditionalFormatRules();
  sheet.clearFormats();

  // ── Row 1: title banner ───────────────────────────────────────────────────
  sheet.getRange('A1:F1').merge()
    .setValue('OptiSheets — Grade Tracker')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // ── Row 2 ─────────────────────────────────────────────────────────────────
  sheet.getRange('A2').setValue('Select Class:').setFontWeight('bold').setFontColor(COLOR_DARK_GREEN);
  sheet.getRange('C2').setValue('Grading Scale:').setFontWeight('bold').setFontColor(COLOR_DARK_GREEN);
  sheet.getRange('D2').setFormula('=Settings!B8');

  // ── Row 3 — green background, white text; B3/D3/F3 filled after loop ──────
  sheet.getRange('A3').setValue('Current Grade:').setFontWeight('bold');
  sheet.getRange('C3').setValue('Letter Grade:').setFontWeight('bold');
  sheet.getRange('E3').setValue('Graded So Far:').setFontWeight('bold');
  sheet.getRange(3, 1, 1, 6).setBackground(COLOR_GREEN).setFontColor(COLOR_WHITE);
  sheet.setRowHeight(3, 28);

  // ── Row 4 ─────────────────────────────────────────────────────────────────
  sheet.getRange('A4').setValue('Grade Needed for Target GPA:').setFontWeight('bold').setFontColor(COLOR_DARK_GREEN);
  sheet.getRange('B4').setValue('Fill in scores to calculate').setFontStyle('italic').setFontColor(COLOR_GRAY);
  sheet.getRange(4, 1, 1, 6).setBackground(COLOR_VLIGHT_GREEN);
  sheet.setRowHeight(4, 26);

  // ── Row 5: column headers ─────────────────────────────────────────────────
  sheet.getRange(5, 1, 1, 6)
    .setValues([['Category', 'Item Name', 'Weight per Item (%)', 'Your Score (%)', 'Weighted Score', 'Status']])
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(5, 28);

  // ── Read category config from Settings ────────────────────────────────────
  var countVals  = settingsSheet.getRange(11, 5, 5, 1).getValues();
  var weightVals = settingsSheet.getRange(11, 6, 5, 1).getValues();
  var dropVals   = settingsSheet.getRange(11, 7, 5, 1).getValues();

  var currentRow   = 6;
  var firstDataRow = -1;
  var lastDataRow  = -1;
  var summaryRows  = [];

  // ── Dynamic category sections — one loop, i=0..4 ─────────────────────────
  for (var i = 0; i < 5; i++) {
    var catName, catCount, catWeightRef, catCountRef;

    if (i === 0) {
      catName      = 'Assignments';
      catCount     = parseInt(countVals[0][0]) || 0;
      catWeightRef = 'Settings!$F$11';
      catCountRef  = 'Settings!$E$11';
    } else if (i === 1) {
      catName      = 'Quizzes';
      catCount     = parseInt(countVals[1][0]) || 0;
      catWeightRef = 'Settings!$F$12';
      catCountRef  = 'Settings!$E$12';
    } else if (i === 2) {
      catName      = 'Exams';
      catCount     = parseInt(countVals[2][0]) || 0;
      catWeightRef = 'Settings!$F$13';
      catCountRef  = 'Settings!$E$13';
    } else if (i === 3) {
      catName      = 'Projects';
      catCount     = parseInt(countVals[3][0]) || 0;
      catWeightRef = 'Settings!$F$14';
      catCountRef  = 'Settings!$E$14';
    } else {
      catName      = 'Participation';
      catCount     = 1;
      catWeightRef = 'Settings!$F$15';
      catCountRef  = 'Settings!$E$15';
    }

    if (catCount <= 0) continue;

    // Sub-header row (merged)
    sheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue(catName)
      .setBackground(COLOR_GREEN)
      .setFontColor(COLOR_WHITE)
      .setFontWeight('bold');
    sheet.setRowHeight(currentRow, 24);
    currentRow++;

    var catStartRow = currentRow;
    if (firstDataRow === -1) firstDataRow = catStartRow;

    // Build value and formula arrays in memory
    var aVals     = [];
    var bVals     = [];
    var cFormulas = [];
    var eFormulas = [];
    var fFormulas = [];

    for (var j = 0; j < catCount; j++) {
      var dataRow  = currentRow + j;
      var itemName;
      if (i === 0)      { itemName = 'Assignment ' + (j + 1); }
      else if (i === 1) { itemName = 'Quiz '       + (j + 1); }
      else if (i === 2) { itemName = 'Exam '       + (j + 1); }
      else if (i === 3) { itemName = 'Project '    + (j + 1); }
      else              { itemName = 'Participation'; }

      aVals.push([catName]);
      bVals.push([itemName]);
      cFormulas.push(['=IFERROR(' + catWeightRef + '/' + catCountRef + ',0)']);
      eFormulas.push(['=IF(D' + dataRow + '="","",D' + dataRow + '*C' + dataRow + '/100)']);
      fFormulas.push([
        '=IF(D' + dataRow + '="","",IF(D' + dataRow + '>=90,"Excellent",' +
        'IF(D' + dataRow + '>=80,"Good",' +
        'IF(D' + dataRow + '>=70,"Satisfactory",' +
        'IF(D' + dataRow + '>=60,"Needs Work","At Risk")))))'
      ]);
    }

    // Write all arrays with single setValues/setFormulas calls per column
    sheet.getRange(catStartRow, 1, catCount, 1).setValues(aVals);
    sheet.getRange(catStartRow, 2, catCount, 1).setValues(bVals);
    sheet.getRange(catStartRow, 3, catCount, 1).setFormulas(cFormulas);
    sheet.getRange(catStartRow, 5, catCount, 1).setFormulas(eFormulas);
    sheet.getRange(catStartRow, 6, catCount, 1).setFormulas(fFormulas);

    // Alternating row backgrounds
    for (var k = 0; k < catCount; k++) {
      var rowBg = (k % 2 === 0) ? COLOR_WHITE : COLOR_VLIGHT_GREEN;
      sheet.getRange(catStartRow + k, 1, 1, 6).setBackground(rowBg);
      sheet.setRowHeight(catStartRow + k, 24);
    }

    lastDataRow = catStartRow + catCount - 1;
    currentRow  = catStartRow + catCount;

    // Summary row
    var sumRow = currentRow;
    summaryRows.push({ catName: catName, row: sumRow, startRow: catStartRow, count: catCount });

    sheet.getRange(sumRow, 1, 1, 6)
      .setValues([[catName + ' Summary', '', '', '', '', '']])
      .setBackground(COLOR_LIGHT_GREEN)
      .setFontColor(COLOR_DARK_GREEN)
      .setFontWeight('bold');
    sheet.getRange(sumRow, 4).setFormula(
      '=IFERROR(AVERAGEIF(D' + catStartRow + ':D' + lastDataRow + ',"<>"),"")');
    sheet.getRange(sumRow, 5).setFormula(
      '=IFERROR(SUM(E' + catStartRow + ':E' + lastDataRow + '),0)');
    sheet.setRowHeight(sumRow, 24);
    currentRow++;

    // Spacer row
    sheet.setRowHeight(currentRow, 8);
    currentRow++;
  }

  // ── Score validation — one single range call after the loop ───────────────
  if (firstDataRow !== -1 && lastDataRow !== -1) {
    sheet.getRange(firstDataRow, 4, lastDataRow - firstDataRow + 1, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, 150)
          .setAllowInvalid(false)
          .build()
      );
  }

  // ── Write live grade formulas into B3, D3, F3 using summaryRows ───────────
  if (summaryRows.length > 0) {
    // B3: current grade % = sum of all weighted scores (each category's E summary)
    var eParts = summaryRows.map(function(s) { return 'E' + s.row; });
    var b3Formula = '=IFERROR(SUM(' + eParts.join(',') + '),"—")';

    // F3: graded so far = count of entered scores across all categories
    var totalItems = 0;
    var cntParts   = summaryRows.map(function(s) {
      totalItems += s.count;
      return 'COUNTIF(D' + s.startRow + ':D' + (s.startRow + s.count - 1) + ',"<>")';
    });
    var f3Formula = '=IFERROR(SUM(' + cntParts.join(',') + ')&" of ' + totalItems + ' items","—")';

    // D3: letter grade — checks D2 (grading scale) for 10-pt vs 7-pt cutoffs
    var d3Formula =
      '=IF(B3="—","—",IF(D2="7-point scale",' +
        'IF(B3>=93,"A",IF(B3>=90,"A-",IF(B3>=87,"B+",IF(B3>=83,"B",IF(B3>=80,"B-",' +
        'IF(B3>=77,"C+",IF(B3>=73,"C",IF(B3>=70,"C-",IF(B3>=67,"D+",IF(B3>=63,"D","F")))))))))),' +
        'IF(B3>=90,"A",IF(B3>=87,"A-",IF(B3>=83,"B+",IF(B3>=80,"B",IF(B3>=77,"B-",' +
        'IF(B3>=73,"C+",IF(B3>=70,"C",IF(B3>=67,"C-",IF(B3>=63,"D+",IF(B3>=60,"D","F")))))))))))' +
      '))';

    sheet.getRange('B3').setFormula(b3Formula);
    sheet.getRange('D3').setFormula(d3Formula);
    sheet.getRange('F3').setFormula(f3Formula);
  }

  // ── Conditional formatting on score column D — one setConditionalFormatRules call ──
  if (firstDataRow !== -1 && lastDataRow !== -1) {
    var scoreCfRange = sheet.getRange(firstDataRow, 4, lastDataRow - firstDataRow + 1, 1);
    var scoreCfRules = [];
    scoreCfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(90).setBackground('#C8E6C9').setRanges([scoreCfRange]).build());
    scoreCfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(80, 89.99).setBackground('#FFF9C4').setRanges([scoreCfRange]).build());
    scoreCfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(70, 79.99).setBackground('#FFCCBC').setRanges([scoreCfRange]).build());
    scoreCfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(70).setBackground('#FFCDD2').setRanges([scoreCfRange]).build());
    sheet.setConditionalFormatRules(scoreCfRules);
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 110);

  sheet.setFrozenRows(5);
  ss.toast('Grade Tracker sheet is ready!', 'OptiSheets AI', 5);
}

// ---------------------------------------------------------------------------
// getAIRecommendations
// ---------------------------------------------------------------------------

function getAIRecommendations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Preparing your data…', 'OptiSheets AI', 5);

  var settings = readSettings(ss);
  if (!settings) return;

  // Read GPA tracker rows starting at row 4, skip blank class names
  var gpaSheet = ss.getSheetByName(GPA_SHEET);
  if (!gpaSheet) {
    showError('Missing sheet: "' + GPA_SHEET + '"', 'Please run Setup GPA Tracker Sheet from the OptiSheets AI menu.');
    return;
  }

  var gpaData = gpaSheet.getRange(4, 1, 50, 7).getValues();
  var classes = [];
  for (var i = 0; i < gpaData.length; i++) {
    var className = String(gpaData[i][1]).trim();
    if (!className) continue;
    classes.push({
      'Semester':        String(gpaData[i][0]).trim(),
      'Class Name':      className,
      'Credit Hours':    gpaData[i][2],
      'Letter Grade':    String(gpaData[i][3]).trim(),
      'Grade Points':    gpaData[i][4],
      'Weighted Points': gpaData[i][5],
      'Notes':           String(gpaData[i][6]).trim()
    });
  }

  if (classes.length === 0) {
    showError('No classes found', 'Your GPA Tracker sheet appears to be empty. Add some classes and try again.');
    return;
  }

  // Read grade tracker summary values
  var gradeSheet   = ss.getSheetByName(GRADE_SHEET);
  var currentGrade = '';
  var letterGrade  = '';
  var gradedSoFar  = '';
  if (gradeSheet) {
    currentGrade = String(gradeSheet.getRange('B3').getValue()).trim();
    letterGrade  = String(gradeSheet.getRange('D3').getValue()).trim();
    gradedSoFar  = String(gradeSheet.getRange('F3').getValue()).trim();
  }

  // Read settings values for AI context
  var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  var targetGPA    = settingsSheet ? String(settingsSheet.getRange('B4').getValue()).trim() : '';
  var priorCredits = settingsSheet ? String(settingsSheet.getRange('B5').getValue()).trim() : '';
  var priorGPA     = settingsSheet ? String(settingsSheet.getRange('B6').getValue()).trim() : '';
  var gradCredits  = settingsSheet ? String(settingsSheet.getRange('B7').getValue()).trim() : '';

  var payload = {
    private_key:   settings.privateKey,
    template_id:   TEMPLATE_ID,
    user_data: {
      classes:       classes,
      current_grade: currentGrade,
      letter_grade:  letterGrade,
      graded_so_far: gradedSoFar,
      target_gpa:    targetGPA,
      prior_credits: priorCredits,
      prior_gpa:     priorGPA,
      grad_credits:  gradCredits
    },
    system_prompt: buildSystemPrompt()
  };

  ss.toast('Calling OptiSheets AI (' + classes.length + ' class(es))…', 'OptiSheets AI', 30);

  var result = callBackend(settings.baseUrl, payload);
  if (!result) return;

  writeOutput(ss, result, classes.length, currentGrade, gradCredits, priorCredits);
  ss.toast('Done! ' + result.remaining_credits + ' credit(s) remaining.', 'OptiSheets AI', 8);
}

// ---------------------------------------------------------------------------
// buildSystemPrompt
// ---------------------------------------------------------------------------

function buildSystemPrompt() {
  return (
    'You are a supportive academic coach for college students. ' +
    'You will receive JSON data with a student\'s GPA tracker (classes, letter grades, credit hours) ' +
    'and grade tracker summary (current grade %, graded items so far, target GPA, prior credits, graduation credits). ' +
    'Analyze the data and produce a structured response with these six labeled sections:\n' +
    '1. GPA SNAPSHOT: Summarize the student\'s current semester GPA, cumulative GPA, and gap to their target GPA. Be encouraging but honest.\n' +
    '2. GRADE RESCUE ALERTS: Identify any classes where the grade is below a B- or trending at risk. Give specific, immediate action steps for each.\n' +
    '3. PRIORITY CLASS STRATEGY: Rank classes by where focused study effort will have the highest GPA impact. Explain why for each.\n' +
    '4. FINAL GRADE SIMULATIONS: Estimate what final scores are needed in weak classes to reach the target GPA, given remaining credits.\n' +
    '5. BURNOUT CHECK: Based on credit load and grade trends, assess overload risk and give one concrete wellness or time-management recommendation.\n' +
    '6. SEMESTER REFLECTION AND NEXT STEPS: Give 3 actionable steps the student should take this week to stay on track.\n' +
    'Be specific, encouraging, and direct. Use plain text only. 600 words max.'
  );
}

// ---------------------------------------------------------------------------
// readSettings
// ---------------------------------------------------------------------------

function readSettings(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    showError(
      'Missing sheet: "' + SETTINGS_SHEET + '"',
      'Please create a sheet named "' + SETTINGS_SHEET + '" with your license key in cell ' + PRIVATE_KEY_CELL + '.\n\nRun Setup All Sheets from the OptiSheets AI menu.'
    );
    return null;
  }

  var privateKey = String(sheet.getRange(PRIVATE_KEY_CELL).getValue()).trim();
  if (!privateKey) {
    showError(
      'License key not found',
      'Cell ' + PRIVATE_KEY_CELL + ' on the "' + SETTINGS_SHEET + '" sheet is empty.\n\nPaste your OptiSheets license key there and try again.'
    );
    return null;
  }

  var baseUrl = getScriptProperty('OPTISHEETS_BASE_URL');
  if (!baseUrl) {
    showError(
      'Backend URL not configured',
      'Go to Extensions → Apps Script → Project Settings → Script Properties\nand add:\n\n  Key:   OPTISHEETS_BASE_URL\n  Value: https://your-backend.vercel.app'
    );
    return null;
  }

  baseUrl = baseUrl.replace(/\/+$/, '');
  return { privateKey: privateKey, baseUrl: baseUrl };
}

// ---------------------------------------------------------------------------
// callBackend
// ---------------------------------------------------------------------------

function callBackend(baseUrl, payload) {
  var url = baseUrl + '/api/get-recommendations';
  var options = {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
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
      'The server returned a non-JSON response. This usually means the backend URL is wrong.\n\nRaw response:\n' + bodyText.slice(0, 300)
    );
    return null;
  }

  if (body.success) return body;

  var friendlyMessage = friendlyErrorMessage(statusCode, body.error || 'Unknown error');
  showError('OptiSheets AI error (HTTP ' + statusCode + ')', friendlyMessage);
  return null;
}

// ---------------------------------------------------------------------------
// writeOutput
// ---------------------------------------------------------------------------

function writeOutput(ss, result, classCount, currentGrade, gradCredits, priorCredits) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET);
  }

  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOR_LIGHT_GREEN);

  // Dark green title banner row 1
  sheet.getRange(1, 1, 1, 3).merge()
    .setValue('OptiSheets — AI Recommendations')
    .setBackground(COLOR_DARK_GREEN)
    .setFontColor(COLOR_WHITE)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM d, yyyy 'at' h:mm a");
  var cacheNote = result.cached ? ' (cached — no credit used)' : '';

  var credNote = '';
  if (gradCredits && priorCredits) {
    var remaining = parseFloat(gradCredits) - parseFloat(priorCredits);
    if (!isNaN(remaining)) credNote = '   |   Credits remaining toward graduation: ' + Math.max(0, remaining);
  }

  var headerData = [
    ['OptiSheets AI — GPA & Grade Tracker Recommendations'],
    ['Generated: ' + timestamp + cacheNote],
    ['Classes analysed: ' + classCount + '   |   Current grade: ' + (currentGrade || '—') + credNote + '   |   AI credits remaining: ' + result.remaining_credits],
    ['']
  ];

  sheet.getRange(2, 1, headerData.length, 1).setValues(headerData);
  sheet.getRange(2, 1).setFontSize(14).setFontWeight('bold');
  sheet.getRange(3, 1).setFontColor('#555555').setFontStyle('italic');
  sheet.getRange(4, 1).setFontColor('#555555');

  var outputRow  = 2 + headerData.length;
  var outputCell = sheet.getRange(outputRow, 1);
  outputCell
    .setValue(result.output)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(11)
    .setBackground(COLOR_VLIGHT_GREEN)
    .setBorder(null, true, null, null, null, null, COLOR_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);

  sheet.setColumnWidth(1, 720);
  sheet.setRowHeight(outputRow, 500);

  ss.setActiveSheet(sheet);
  sheet.setActiveRange(outputCell);
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

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
        'Your tracker has too much data for a single request.\n\n' +
        'Try removing old rows and re-run.'
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

function showError(title, message) {
  SpreadsheetApp.getUi().alert('⚠️ ' + title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function getScriptProperty(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? String(value).trim() : '';
}
