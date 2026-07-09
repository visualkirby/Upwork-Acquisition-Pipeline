/**
 * ============================================================
 * 15. REPAIR / DIAGNOSTIC
 *
 * Formulas are applied automatically during FreelanceFlow Setup
 * (wizard_initialize -> applyJobDiscoveryFormulas_/applyJobScoringFormulas_/
 * applyProposalGeneratorFormulas_ in 00_Setup_Wizard.gs, all powered by the
 * Apps Script Library).
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

  // A template update can add a brand new pipeline sheet (Client_Chat_Log,
  // Contract_Tracker, Milestone_Tracker were added this way) -- an existing
  // customer's spreadsheet won't have it yet, and re-running the whole Setup
  // Wizard would be destructive. ensurePipelineSheets_ only creates sheets
  // that don't already exist, so calling it here is a safe way to backfill
  // any new sheet before the repair logic below tries to reference it.
  ensurePipelineSheets_(ss, []);

  var jdSheet = ss.getSheetByName('Job_Discovery');
  var jsSheet = ss.getSheetByName('Job_Scoring');
  var pgSheet = ss.getSheetByName('Proposal_Generator');
  var ptSheet = ss.getSheetByName('Proposal_Tracker');
  var ctSheet = ss.getSheetByName('Contract_Tracker');
  var msSheet = ss.getSheetByName('Milestone_Tracker');
  var hlSheet = ss.getSheetByName('Hourly_Log');

  if (!jdSheet || !jsSheet || !pgSheet) {
    ui.alert('Missing sheet. Confirm Job_Discovery, Job_Scoring, and Proposal_Generator all exist.');
    return;
  }

  var repaired = [];

  if (jdSheet.getLastRow() >= 2) {
    var jdHeaders = jdSheet.getRange(1, 1, 1, jdSheet.getLastColumn()).getValues()[0];
    applyJobDiscoveryFormulas_(jdSheet, jdHeaders);
    var jdMap = getHeaderMap_(jdSheet);
    copyRowDown_(jdSheet, [
      getCol_(jdMap, ['Discovery_ID']),
      getCol_(jdMap, ['Current_Age_Days']),
      getCol_(jdMap, ['Keyword_Fit_Score']),
      getCol_(jdMap, ['Tool_Detected']),
      getCol_(jdMap, ['Tool_Score']),
      getCol_(jdMap, ['Experience_Score']),
      getCol_(jdMap, ['Freshness_Score']),
      getCol_(jdMap, ['Competition_Score']),
      getCol_(jdMap, ['Verification_Score']),
      getCol_(jdMap, ['Client_History_Score']),
      getCol_(jdMap, ['Budget_Quick_Score']),
      getCol_(jdMap, ['Discovery_Priority_Score']),
      getCol_(jdMap, ['Discovery_Action']),
      getCol_(jdMap, ['Discovery_Status'])
    ]);
    applyJobDiscoveryConditionalFormatting_(jdSheet, jdHeaders);
    repaired.push('Job_Discovery (pre-screen chain + Discovery_Action)');
  }

  if (jsSheet.getLastRow() >= 2) {
    var jsHeaders = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];
    applyJobScoringPullFormula_(jsSheet, jsHeaders);
    applyJobScoringFormulas_(jsSheet, jsHeaders);
    var jsMap = getHeaderMap_(jsSheet);
    copyRowDown_(jsSheet, [
      getCol_(jsMap, ['Current_Age_Days']),
      getCol_(jsMap, ['Effort_Level']),
      getCol_(jsMap, ['Scope_Rating']),
      getCol_(jsMap, ['Portfolio_Match']),
      getCol_(jsMap, ['Estimated_Hours']),
      getCol_(jsMap, ['Estimated_Hourly_Rate']),
      getCol_(jsMap, ['Budget_Score']),
      getCol_(jsMap, ['Keyword_Score']),
      getCol_(jsMap, ['Tool_Score']),
      getCol_(jsMap, ['Experience_Score']),
      getCol_(jsMap, ['Freshness_Score']),
      getCol_(jsMap, ['Competition_Score']),
      getCol_(jsMap, ['Client_History_Score']),
      getCol_(jsMap, ['Scope_Score']),
      getCol_(jsMap, ['Portfolio_Score']),
      getCol_(jsMap, ['Connects_Affordability']),
      getCol_(jsMap, ['Total_Score']),
      getCol_(jsMap, ['Score_Per_Connect']),
      getCol_(jsMap, ['Final_Decision']),
      getCol_(jsMap, ['Proposal_Generator_Date'])
    ]);
    applyJobScoringValidation_(jsSheet, jsHeaders);
    applyJobScoringConditionalFormatting_(jsSheet, jsHeaders);
    repaired.push('Job_Scoring (Job_Discovery pull + full 9-score chain + Total_Score, Final_Decision)');
  }

  if (pgSheet.getLastRow() >= 2) {
    var pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
    applyProposalGeneratorPullFormula_(pgSheet, pgHeaders);
    applyProposalGeneratorFormulas_(pgSheet, pgHeaders);
    applyProposalGeneratorValidation_(pgSheet, pgHeaders);
    var pgMap = getHeaderMap_(pgSheet);
    copyRowDown_(pgSheet, [
      getCol_(pgMap, ['Tool_Detected']),
      getCol_(pgMap, ['Portfolio_Project'])
    ]);
    repaired.push('Proposal_Generator (Job_Scoring pull, Tool_Detected, Portfolio_Project)');
  }

  // Proposal_Tracker/Contract_Tracker/Milestone_Tracker only need row 1 --
  // validation applies to the whole prefilled range regardless of whether
  // any rows have been sent/contracted yet.
  if (ptSheet) {
    var ptHeaders = ptSheet.getRange(1, 1, 1, ptSheet.getLastColumn()).getValues()[0];
    applyProposalTrackerValidation_(ptSheet, ptHeaders);
    repaired.push('Proposal_Tracker (Viewed/Interview/Hired dropdowns, Revenue currency format)');
  }

  if (ctSheet) {
    var ctHeaders = ctSheet.getRange(1, 1, 1, ctSheet.getLastColumn()).getValues()[0];
    applyContractTrackerValidation_(ctSheet, ctHeaders);
    repaired.push('Contract_Tracker (Contract_Type/Status dropdowns, currency formats)');
  }

  if (msSheet) {
    var msHeaders = msSheet.getRange(1, 1, 1, msSheet.getLastColumn()).getValues()[0];
    applyMilestoneTrackerValidation_(msSheet, msHeaders);
    repaired.push('Milestone_Tracker (Status dropdown, Amount currency format)');
  }

  if (hlSheet) {
    ensureHourlyLogStatusColumn_(hlSheet);
    var hlHeaders = hlSheet.getRange(1, 1, 1, hlSheet.getLastColumn()).getValues()[0];
    applyHourlyLogValidation_(hlSheet, hlHeaders);
    repaired.push('Hourly_Log (Amount/Hours_Logged number formats, Status dropdown)');
  }

  // Additional_Questions flows backwards (Proposal_Generator -> Job_Scoring
  // -> Job_Discovery) -- read fresh headers unconditionally since this only
  // needs row 1 to exist, not any data rows, in either direction.
  var jdHeadersForAQ = jdSheet.getRange(1, 1, 1, jdSheet.getLastColumn()).getValues()[0];
  var jsHeadersForAQ = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];
  applyJobScoringAdditionalQuestionsLookup_(jsSheet, jsHeadersForAQ);
  applyJobDiscoveryAdditionalQuestionsLookup_(jdSheet, jdHeadersForAQ);
  repaired.push('Additional_Questions lookup (Proposal_Generator -> Job_Scoring -> Job_Discovery)');

  // Re-registering here clears out any stale trigger still pointing at the
  // old "onEdit" handler name. A bare onEdit function auto-runs as an
  // unauthorized simple trigger alongside the installable one, so every AI
  // call on edit (Boost_Connects, Bid_4th, Additional_Questions) could
  // randomly get overwritten by that simple trigger's permissions error.
  registerEditTrigger_();
  repaired.push('Edit trigger (re-registered to handleEdit, removes stale onEdit simple-trigger conflict)');

  if (repaired.length === 0) {
    ui.alert('No data rows found to repair yet -- formulas will be set correctly as soon as row 2 is filled in.');
    return;
  }

  ui.alert('Formulas repaired:\n\n' + repaired.join('\n'));
}

// Adds a Status column to an existing Hourly_Log sheet that predates it --
// appended after the last column so it doesn't shift any header a formula
// or trigger already references by position elsewhere (everything in this
// project resolves columns by header name, but no reason to disturb layout
// unnecessarily). Backfills existing rows from Contract_Tracker's current
// Status for that Discovery_ID (defaulting to Active) so old test/real rows
// don't sit blank against the new dropdown. No-ops if Status already exists.
function ensureHourlyLogStatusColumn_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Status') !== -1) return;

  var statusCol = lastCol + 1;
  sheet.getRange(1, statusCol).setValue('Status').setFontWeight('bold');

  if (sheet.getLastRow() < 2) return;

  var map   = getHeaderMap_(sheet);
  var idCol = getCol_(map, ['Discovery_ID']);
  if (!idCol) return;

  var contracts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Contract_Tracker');
  var ctStatusByDiscoveryId = {};
  if (contracts && contracts.getLastRow() > 1) {
    var ctMap       = getHeaderMap_(contracts);
    var ctIdCol     = getCol_(ctMap, ['Discovery_ID']);
    var ctStatusCol = getCol_(ctMap, ['Status']);
    if (ctIdCol && ctStatusCol) {
      var ctData = contracts.getRange(2, 1, contracts.getLastRow() - 1, contracts.getLastColumn()).getValues();
      ctData.forEach(function (row) {
        ctStatusByDiscoveryId[String(row[ctIdCol - 1])] = row[ctStatusCol - 1];
      });
    }
  }

  var lastRow  = sheet.getLastRow();
  var idValues = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  var statusValues = idValues.map(function (row) {
    return [ctStatusByDiscoveryId[String(row[0])] || 'Active'];
  });
  sheet.getRange(2, statusCol, statusValues.length, 1).setValues(statusValues);
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

  var jdSheet = ss.getSheetByName('Job_Discovery');
  lines.push('JOB_DISCOVERY:');
  if (jdSheet) {
    var jdHeaders = jdSheet.getRange(1, 1, 1, jdSheet.getLastColumn()).getValues()[0];
    lines.push('  Headers: ' + jdHeaders.join(', '));
    if (jdSheet.getLastRow() >= 2) {
      var jdMap = getHeaderMap_(jdSheet);
      ['Tool_Detected', 'Tool_Score', 'Discovery_Priority_Score', 'Discovery_Action'].forEach(function (name) {
        var col = getCol_(jdMap, [name]);
        if (col) lines.push('  ' + name + ' formula (row 2): ' + jdSheet.getRange(2, col).getFormula());
      });
    }
  } else {
    lines.push('  (sheet not found)');
  }
  lines.push('');

  var jsSheet = ss.getSheetByName('Job_Scoring');
  lines.push('JOB_SCORING:');
  if (jsSheet) {
    var jsHeaders = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];
    lines.push('  Headers: ' + jsHeaders.join(', '));
    if (jsSheet.getLastRow() >= 2) {
      var jsMap = getHeaderMap_(jsSheet);
      ['Job_Title', 'Tool_Score', 'Total_Score', 'Connects_Affordability', 'Final_Decision', 'Proposal_Generator_Date'].forEach(function (name) {
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
      ['Job_Title', 'Connects_Required', 'Tool_Detected', 'Portfolio_Project'].forEach(function (name) {
        var col = getCol_(pgMap, [name]);
        if (col) lines.push('  ' + name + ' formula (row 2): ' + pgSheet.getRange(2, col).getFormula());
      });
    }
  } else {
    lines.push('  (sheet not found)');
  }

  return lines.join('\n');
}
