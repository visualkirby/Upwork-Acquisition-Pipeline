/**
 * ============================================================
 * 15. REPAIR / DIAGNOSTIC
 *
 * Formulas are applied automatically during FreelanceFlow Setup
 * (wizard_initialize -> applyJobScoringFormulas_/applyProposalGeneratorFormulas_
 * in 00_Setup_Wizard.gs, both powered by the Apps Script Library).
 *
 * REPAIR_FORMULAS() re-applies those same Library-built formulas
 * to an EXISTING sheet -- row 2 plus copied down to every data
 * row -- for when a formula fix ships after your sheet was
 * already set up. There's no way to push a fix to your sheet
 * automatically (Apps Script only runs under your own account),
 * so this is a "pull" you run yourself after updating your
 * Library version reference.
 *
 * SEND_DIAGNOSTIC_REPORT() reads your current Settings, column
 * headers, and the actual formula text in your scoring columns
 * so you can share it with support -- shown in a dialog to copy
 * into an email, or sent directly only if you click Yes.
 * ============================================================
 */
function REPAIR_FORMULAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var jsSheet = ss.getSheetByName('Job_Scoring');
  var pgSheet = ss.getSheetByName('Proposal_Generator');

  if (!jsSheet || !pgSheet) {
    ui.alert('Missing sheet. Confirm Job_Scoring and Proposal_Generator both exist.');
    return;
  }

  var repaired = [];

  if (jsSheet.getLastRow() >= 2) {
    var jsHeaders = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];
    applyJobScoringFormulas_(jsSheet, jsHeaders);
    var jsMap = getHeaderMap_(jsSheet);
    copyRowDown_(jsSheet, [
      getCol_(jsMap, ['Connects_Affordability']),
      getCol_(jsMap, ['Final_Decision'])
    ]);
    repaired.push('Job_Scoring (Connects_Affordability, Final_Decision)');
  }

  if (pgSheet.getLastRow() >= 2) {
    var pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
    applyProposalGeneratorFormulas_(pgSheet, pgHeaders);
    var pgMap = getHeaderMap_(pgSheet);
    copyRowDown_(pgSheet, [
      getCol_(pgMap, ['Tool_Detected']),
      getCol_(pgMap, ['Portfolio_Project'])
    ]);
    repaired.push('Proposal_Generator (Tool_Detected, Portfolio_Project)');
  }

  if (repaired.length === 0) {
    ui.alert('No data rows found to repair yet -- formulas will be set correctly as soon as row 2 is filled in.');
    return;
  }

  ui.alert('Formulas repaired:\n\n' + repaired.join('\n'));
}

function copyRowDown_(sheet, cols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  cols.forEach(function (col) {
    if (!col) return;
    sheet.getRange(2, col).copyTo(sheet.getRange(3, col, lastRow - 2, 1));
  });
}


var SUPPORT_EMAIL_ = 'support@benchlineanalytics.com';

function SEND_DIAGNOSTIC_REPORT() {
  var ui     = SpreadsheetApp.getUi();
  var report = buildDiagnosticReport_();

  var response = ui.alert(
    'Diagnostic Report',
    report + '\n\nClick Yes to email this report to Benchline Analytics support. Click No to copy it manually instead.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    MailApp.sendEmail({
      to: SUPPORT_EMAIL_,
      subject: 'FreelanceFlow Diagnostic Report',
      body: report
    });
    ui.alert('Report sent. Support will follow up by email.');
  } catch (err) {
    ui.alert('Could not send automatically (' + err.message + '). Please copy the report above into an email instead.');
  }
}

function buildDiagnosticReport_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var lines = [];

  lines.push('FreelanceFlow Diagnostic Report -- ' + new Date().toISOString());
  lines.push('');

  var settingsSheet = ss.getSheetByName('Settings');
  lines.push('SETTINGS:');
  if (settingsSheet && settingsSheet.getLastRow() > 1) {
    var data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
    data.forEach(function (row) {
      var key = String(row[0]).trim();
      if (!key) return;
      lines.push('  ' + key + ': ' + row[1]);
    });
  } else {
    lines.push('  (no Settings sheet or no data)');
  }
  lines.push('');

  var jsSheet = ss.getSheetByName('Job_Scoring');
  lines.push('JOB_SCORING:');
  if (jsSheet) {
    var jsHeaders = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];
    lines.push('  Headers: ' + jsHeaders.join(', '));
    if (jsSheet.getLastRow() >= 2) {
      var jsMap = getHeaderMap_(jsSheet);
      ['Connects_Affordability', 'Final_Decision'].forEach(function (name) {
        var col = getCol_(jsMap, [name]);
        if (col) lines.push('  ' + name + ' formula (row 2): ' + jsSheet.getRange(2, col).getFormula());
      });
    }
  } else {
    lines.push('  (sheet not found)');
  }
  lines.push('');

  var pgSheet = ss.getSheetByName('Proposal_Generator');
  lines.push('PROPOSAL_GENERATOR:');
  if (pgSheet) {
    var pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
    lines.push('  Headers: ' + pgHeaders.join(', '));
    if (pgSheet.getLastRow() >= 2) {
      var pgMap = getHeaderMap_(pgSheet);
      ['Tool_Detected', 'Portfolio_Project'].forEach(function (name) {
        var col = getCol_(pgMap, [name]);
        if (col) lines.push('  ' + name + ' formula (row 2): ' + pgSheet.getRange(2, col).getFormula());
      });
    }
  } else {
    lines.push('  (sheet not found)');
  }

  return lines.join('\n');
}
