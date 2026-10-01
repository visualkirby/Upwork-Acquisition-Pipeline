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
 *
 * RECONCILE_CONNECTS_HELPER() recomputes Connects_Helper's
 * Total_Connects_Used/Total_Proposal_Cost directly from
 * Proposal_Tracker, for when those two running totals drift out
 * of sync with the rows that actually fed them (see its own
 * comment below for why that happens).
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
  var settingsSheet = ss.getSheetByName('Settings');

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
    var pgMigrated = applyProposalGeneratorLookupFormulas_(pgSheet, pgHeaders);
    applyProposalGeneratorFormulas_(pgSheet, pgHeaders);
    applyProposalGeneratorValidation_(pgSheet, pgHeaders);
    var pgMap = getHeaderMap_(pgSheet);
    copyRowDown_(pgSheet, [
      getCol_(pgMap, ['Tool_Detected'])
    ]);
    // Migration clears Job_Scoring's Proposal_Generator_Date to break a
    // formula loop with the old spill, so the lookup goes back on here.
    applyJobScoringProposalDateLookup_(jsSheet,
      jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0]);
    if (pgMigrated > 0) {
      repaired.push('Proposal_Generator (moved ' + pgMigrated + ' row' + (pgMigrated === 1 ? '' : 's') +
        ' to fixed rows keyed by Discovery_ID; their Date is now the sent/skip date, or today if neither)');
    }
    repaired.push('Proposal_Generator (Discovery_ID lookups, Tool_Detected) -- run System Tools > Run Job Classification afterward to (re)fill Portfolio_Project, which is script-computed now instead of a formula');
  }

  // Appends a row for any APPLY job Proposal_Generator doesn't have yet and
  // tags rows whose job is no longer APPLY (28_Proposal_Sync.gs).
  var pgSync = syncProposalGenerator_(ss);
  if (pgSync.added.length > 0 || pgSync.flagged > 0) {
    repaired.push('Proposal_Generator sync (' + pgSync.added.length + ' APPLY job' + (pgSync.added.length === 1 ? '' : 's') +
      ' added, ' + pgSync.flagged + ' row' + (pgSync.flagged === 1 ? '' : 's') + ' tagged "Not APPLY now" in Notes)');
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

  // Backfills any Keyword_Search_List row (however it was added) that
  // predates the KEYWORD_SEARCH_LIST edit-trigger sync and never got a
  // matching Keyword_Strategy row -- see ensureKeywordStrategyRow_
  // (18_Keyword_Strategy.gs) and the KEYWORD_SEARCH_LIST block in
  // 14_Edit_Trigger.gs for the ongoing (post-repair) sync.
  var klSheet = ss.getSheetByName('Keyword_Search_List');
  if (klSheet && klSheet.getLastRow() >= 2) {
    var klMap      = getHeaderMap_(klSheet);
    var klQueryCol = getCol_(klMap, ['Search_Query']);
    if (klQueryCol) {
      var klValues = klSheet.getRange(2, klQueryCol, klSheet.getLastRow() - 1, 1).getValues();
      var klCreated = 0;
      klValues.forEach(function (row) {
        var q = row[0];
        if (q !== '' && q !== null && ensureKeywordStrategyRow_(ss, q)) {
          klCreated++;
        }
      });
      if (klCreated > 0) {
        repaired.push('Keyword_Strategy (backfilled ' + klCreated + ' missing tracking row' + (klCreated === 1 ? '' : 's') + ' from Keyword_Search_List -- Actual_Count computed from any matching jobs already in Job_Discovery)');
      }
    }
  }

  // Backfills the Scale column onto a Keyword_Strategy sheet that predates
  // it -- appended after the last column (Drop), same "add at the end, never
  // shift existing headers" convention as ensureHourlyLogStatusColumn_ below.
  // Nothing to backfill row-wise (Scale starts blank same as a fresh sheet),
  // just the header + its dropdown validation, so this is simpler than the
  // Hourly_Log case.
  var stratSheet = ss.getSheetByName('Keyword_Strategy');
  if (stratSheet) {
    var stratHeadersForScale = stratSheet.getRange(1, 1, 1, stratSheet.getLastColumn()).getValues()[0];
    if (stratHeadersForScale.indexOf('Scale') === -1) {
      var scaleHeaderCol = stratSheet.getLastColumn() + 1;
      stratSheet.getRange(1, scaleHeaderCol).setValue('Scale').setFontWeight('bold');
      repaired.push('Keyword_Strategy (added missing Scale column)');
    }
    var stratHeadersNow = stratSheet.getRange(1, 1, 1, stratSheet.getLastColumn()).getValues()[0];
    applyKeywordStrategyValidation_(stratSheet, stratHeadersNow);
  }

  // Re-syncs every EXISTING Keyword_Strategy row's Actual_Count against
  // Job_Discovery (recomputeKeywordStrategyActualCount_, 18_Keyword_Strategy.gs).
  // Covers rows that drifted under the old "+1 on first log" counter (a
  // resubmitted job, a misfired multi-row paste, or Keyword_Search corrected
  // after the fact could all leave Actual_Count wrong) -- this is the
  // one-time "pull" that re-syncs them after updating to the recompute-based
  // Library version.
  if (stratSheet && stratSheet.getLastRow() >= 2) {
    var stratMap2  = getHeaderMap_(stratSheet);
    var stratKwCol2 = getCol_(stratMap2, ['Keyword']);
    if (stratKwCol2) {
      var stratKwValues = stratSheet.getRange(2, stratKwCol2, stratSheet.getLastRow() - 1, 1).getValues();
      var stratResynced = 0;
      stratKwValues.forEach(function (row) {
        var kw = row[0];
        if (kw !== '' && kw !== null) {
          recomputeKeywordStrategyActualCount_(ss, kw);
          stratResynced++;
        }
      });
      if (stratResynced > 0) {
        repaired.push('Keyword_Strategy (re-synced Actual_Count for ' + stratResynced + ' keyword' + (stratResynced === 1 ? '' : 's') + ' against Job_Discovery)');
      }
    }
  }

  // Backfills a missing Proposal_Length setting (added after Proposal_Tone/
  // Journey_Stage already existed on some copies) and (re)applies the
  // Journey_Stage/Proposal_Tone/Proposal_Length dropdowns. Proposal_Length
  // belongs immediately after Proposal_Tone (matches initSettingsSheet_'s
  // row order for a fresh setup) -- self-healing on every run, whether the
  // row is missing entirely or already sitting somewhere else (e.g.
  // appended at the end by an earlier version of this repair step). Its
  // existing value, if any, is preserved when it gets repositioned.
  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    var settingsMap = getHeaderMap_(settingsSheet);
    var settingCol  = getCol_(settingsMap, ['Setting']);
    var valueCol    = getCol_(settingsMap, ['Value']);
    if (settingCol && valueCol) {
      var settingNames = settingsSheet
        .getRange(2, settingCol, settingsSheet.getLastRow() - 1, 1)
        .getValues()
        .map(function (row) { return String(row[0]).trim(); });

      var toneIdx   = settingNames.indexOf('Proposal_Tone');
      var lengthIdx = settingNames.indexOf('Proposal_Length');

      if (toneIdx !== -1 && lengthIdx !== toneIdx + 1) {
        var existingValue = 'Medium';
        if (lengthIdx !== -1) {
          existingValue = settingsSheet.getRange(lengthIdx + 2, valueCol).getValue() || 'Medium';
          settingsSheet.deleteRow(lengthIdx + 2);
          settingNames.splice(lengthIdx, 1);
          if (lengthIdx < toneIdx) toneIdx--;
        }

        var toneRow = toneIdx + 2;
        settingsSheet.insertRowAfter(toneRow);
        settingsSheet.getRange(toneRow + 1, settingCol).setValue('Proposal_Length');
        settingsSheet.getRange(toneRow + 1, valueCol).setValue(existingValue);
        settingNames.splice(toneIdx + 1, 0, 'Proposal_Length');
        repaired.push('Settings (Proposal_Length positioned right after Proposal_Tone, row ' + (toneRow + 1) + ')');
      }

      applySettingsValidation_(settingsSheet, settingNames);
      repaired.push('Settings (Journey_Stage/Proposal_Tone/Proposal_Length dropdowns)');
    }
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

  // ensurePipelineSheets_ above only creates sheets that don't exist yet --
  // a new one lands appended at the end regardless of where reorderPipelineTabs_'s
  // order list says it belongs, since that function is otherwise only called
  // once, during the Setup Wizard's initial run. Re-running it here fixes
  // tab position for any sheet a template update added after this copy was
  // already set up (Client_Chat_Log/Contract_Tracker/Milestone_Tracker
  // needed this before Keyword_Intelligence did) -- safe to call anytime,
  // it only repositions sheets that already exist and no-ops on ones that
  // don't (see its own `if (sheet)` guard).
  reorderPipelineTabs_(ss);

  if (repaired.length === 0) {
    ui.alert('No data rows found to repair yet -- formulas will be set correctly as soon as row 2 is filled in.');
    return;
  }

  ui.alert('Formulas repaired:\n\n' + repaired.join('\n'));
}

// Total_Connects_Used and Total_Proposal_Cost are meant to always equal a
// straight sum across every Proposal_Tracker row -- no month-scoping, no
// other inputs (see 26_Keyword_Intelligence.gs's own comment: "Proposal_Cost
// sums Proposal_Tracker's own Proposal_Cost column... same source the
// corresponding Connects_Helper metrics use"). That makes them safe to
// blindly recompute and overwrite here, unlike Current_Connect_Balance
// (also fed by manual Connect_Replenishment/Connect_Returned entries in
// Connects_Helper itself) or the MTD_*/Monthly_* metrics (period-scoped,
// reset by SNAPSHOT_MONTH_END/a new month) -- neither of those has a single
// source of truth to recompute from, so nothing here touches them.
//
// Why these two drift: handleProposalStatusChange_ (14_Edit_Trigger.gs)
// updates them via 5 separate read-modify-write calls to
// incrementConnectsHelperMetric_, run in sequence after the Proposal_Tracker
// row is appended. Google Sheets' own Ctrl+Z undo isn't part of Apps
// Script's execution model at all, so it can revert some of those writes
// without reverting others, or revert them without reverting the
// Proposal_Tracker row itself. Confirmed 2026-07-26: an undo/redo
// troubleshooting session left both metrics short of what Proposal_Tracker's
// own rows summed to, by exactly the proposals sent that same session.
//
// Returns null if Proposal_Tracker or its Connects_Used/Proposal_Cost
// columns aren't found -- the single place both RECONCILE_CONNECTS_HELPER's
// on-demand check and the silent auto-heal in BUILD_DASHBOARD/
// BUILD_KEYWORD_INTELLIGENCE get these two numbers from, so there's one
// definition of how they're derived.
function computeConnectsHelperTrueTotals_(ss) {
  var ptSheet = ss.getSheetByName('Proposal_Tracker');
  if (!ptSheet) return null;

  var ptMap       = getHeaderMap_(ptSheet);
  var connectsCol = getCol_(ptMap, ['Connects_Used']);
  var costCol     = getCol_(ptMap, ['Proposal_Cost']);
  if (!connectsCol || !costCol) return null;

  var totalConnectsUsed = 0;
  var totalProposalCost = 0;
  if (ptSheet.getLastRow() >= 2) {
    var ptData = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, ptSheet.getLastColumn()).getValues();
    ptData.forEach(function (row) {
      totalConnectsUsed += Number(row[connectsCol - 1]) || 0;
      totalProposalCost += parseDollarString_(row[costCol - 1]);
    });
  }

  return {
    Total_Connects_Used: totalConnectsUsed,
    Total_Proposal_Cost: Math.round(totalProposalCost * 100) / 100
  };
}

// Overwrites Connects_Helper's Total_Connects_Used/Total_Proposal_Cost with
// computeConnectsHelperTrueTotals_'s values wherever they've drifted -- no
// UI, so it's safe to call as a routine silent pre-step from anything that
// reads either metric (BUILD_DASHBOARD, BUILD_KEYWORD_INTELLIGENCE), not
// just the explicit on-demand menu item. Returns one {metric, from, to,
// diff, changed} entry per target metric found in Connects_Helper, whether
// or not it needed correcting, so a caller can report on it or just ignore
// the return value.
function reconcileConnectsHelperTotals_(ss) {
  var chSheet    = ss.getSheetByName('Connects_Helper');
  var trueTotals = computeConnectsHelperTrueTotals_(ss);
  if (!chSheet || !trueTotals || chSheet.getLastRow() < 2) return [];

  var metricData = chSheet.getRange(2, 1, chSheet.getLastRow() - 1, 2).getValues();
  var results = [];

  for (var i = 0; i < metricData.length; i++) {
    var metricName = String(metricData[i][0]).trim();
    if (!trueTotals.hasOwnProperty(metricName)) continue;

    var currentValue = Number(metricData[i][1]) || 0;
    var trueValue     = trueTotals[metricName];
    var diff          = Math.round((trueValue - currentValue) * 100) / 100;
    var changed       = Math.abs(diff) > 0.005;

    if (changed) chSheet.getRange(i + 2, 2).setValue(trueValue);

    results.push({ metric: metricName, from: currentValue, to: trueValue, diff: diff, changed: changed });
  }

  return results;
}

// Manual on-demand version (System Tools > Reconcile Connects_Helper
// Totals) -- same numbers and the same writes as the silent auto-heal
// above, but always shows a report (including metrics that were already
// correct) since this is something you run specifically to check.
function RECONCILE_CONNECTS_HELPER() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var ptSheet = ss.getSheetByName('Proposal_Tracker');
  var chSheet = ss.getSheetByName('Connects_Helper');
  if (!ptSheet || !chSheet) {
    ui.alert('Proposal_Tracker or Connects_Helper sheet not found.');
    return;
  }
  if (!computeConnectsHelperTrueTotals_(ss)) {
    ui.alert('Proposal_Tracker is missing its Connects_Used or Proposal_Cost column.');
    return;
  }
  if (chSheet.getLastRow() < 2) {
    ui.alert('Connects_Helper has no metric rows to reconcile.');
    return;
  }

  var results = reconcileConnectsHelperTotals_(ss);
  if (results.length === 0) {
    ui.alert('Total_Connects_Used / Total_Proposal_Cost rows not found in Connects_Helper.');
    return;
  }

  var report = results.map(function (r) {
    return r.changed
      ? r.metric + ': ' + r.from + ' -> ' + r.to + ' (was off by ' + r.diff + ')'
      : r.metric + ': ' + r.from + ' (already correct)';
  });

  ui.alert('Connects_Helper reconciled against Proposal_Tracker:\n\n' + report.join('\n'));
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
