/**
 * ============================================================
 * 3. HELPERS
 * ============================================================
 */

function normalizeHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    map[normalizeHeader_(h)] = i + 1;
  });
  return map;
}

function getCol_(headerMap, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var key = normalizeHeader_(possibleNames[i]);
    if (headerMap[key]) return headerMap[key];
  }
  return null;
}

function getCellValue_(sheet, row, headerMap, possibleNames) {
  var col = getCol_(headerMap, possibleNames);
  if (!col) return "";
  return sheet.getRange(row, col).getValue();
}

function setCellValue_(sheet, row, headerMap, possibleNames, value) {
  var col = getCol_(headerMap, possibleNames);
  if (!col) return;
  sheet.getRange(row, col).setValue(value);
}

function cleanJobText_(text) {
  if (text === null || text === undefined) return "";
  return String(text)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n+/g, " ")
    .replace(/\t+/g, " ")
    .replace(/ /g, " ")
    .replace(/[•▪■●◦]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function getLastRealRow_(sheet) {
  var maxRows  = sheet.getMaxRows();
  var lastReal = 1;
  for (var col = 1; col <= 3; col++) {
    var vals = sheet.getRange(2, col, maxRows - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]).trim() !== "") {
        var rowNum = i + 2;
        if (rowNum > lastReal) lastReal = rowNum;
        break;
      }
    }
  }
  return lastReal;
}

function findFirstEmptyRowByColumn_(sheet, col) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values  = sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) return i + 2;
  }
  return lastRow + 1;
}

function formatDuration_(ms) {
  var totalSeconds = Math.floor(ms / 1000);
  var hours        = Math.floor(totalSeconds / 3600);
  var minutes      = Math.floor((totalSeconds % 3600) / 60);
  var seconds      = totalSeconds % 60;
  if (hours > 0) {
    return hours + "h " + minutes + "m " + seconds + "s";
  }
  return minutes + "m " + seconds + "s";
}

function formatTime_(date) {
  var h    = date.getHours();
  var m    = date.getMinutes();
  var s    = date.getSeconds();
  var ampm = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return h + ":" + (m < 10 ? "0" + m : m) + ":" + (s < 10 ? "0" + s : s) + " " + ampm;
}

// Connects_Helper is metric-row structured (Metric name in column A, Value in
// column B) rather than header-column structured like the pipeline sheets, so
// updating one of its running totals from another sheet's edit trigger means
// scanning column A for the matching label rather than a header lookup.
//
// Locked because this is called from many concurrent-capable trigger paths
// (Milestone Released, Hourly_Log delta, Ended Early, Proposal_Tracker row
// creation) -- without a lock, two overlapping executions can both read the
// same "current" value before either writes, and one increment gets lost.
//
// SpreadsheetApp.flush() before releasing the lock is not optional here --
// Apps Script batches pending spreadsheet writes rather than committing them
// immediately, so releasing the lock right after setValue() can let the next
// execution acquire the lock and read this value before the write actually
// lands. Verified via testing on 2026-07-05: with the lock but no flush, two
// rapid Proposal_Tracker "Sent" edits each successfully appended their own
// row (a separate fix), but only one of their two increments here actually
// stuck -- the second execution's read of "current" was still the
// pre-increment value.
function incrementConnectsHelperMetric_(ss, metricName, amount) {
  if (!amount) return;
  var sheet = ss.getSheetByName('Connects_Helper');
  if (!sheet) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === metricName) {
        var current = Number(data[i][1]) || 0;
        sheet.getRange(i + 2, 2).setValue(current + amount);
        SpreadsheetApp.flush();
        return;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// Looks up an Hourly contract's rate from Contract_Tracker by Discovery_ID --
// used by the Hourly_Log handler to compute Amount from Hours_Logged without
// relying on a live formula (see 14_Edit_Trigger.gs's HOURLY_LOG block for why).
function getContractHourlyRate_(ss, discoveryId) {
  var sheet = ss.getSheetByName('Contract_Tracker');
  if (!sheet || sheet.getLastRow() < 2) return 0;

  var map    = getHeaderMap_(sheet);
  var idCol  = getCol_(map, ['Discovery_ID']);
  var rateCol = getCol_(map, ['Hourly_Rate']);
  if (!idCol || !rateCol) return 0;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idCol - 1]) === String(discoveryId)) {
      return Number(data[i][rateCol - 1]) || 0;
    }
  }
  return 0;
}

// Sums the revenue already recognized for a contract -- Released milestones
// for Fixed, all logged entries for Hourly. Used by 14_Edit_Trigger.gs's
// Contract_Tracker block both to roll up Total_Released on Completed and to
// find the delta still owed on Ended Early, so a contract that already had
// some milestones/hours paid out doesn't get that revenue counted twice.
function getContractRecognizedRevenue_(ss, discoveryId, contractType) {
  var total = 0;

  if (contractType === "Fixed") {
    var msSheet = ss.getSheetByName("Milestone_Tracker");
    if (msSheet && msSheet.getLastRow() > 1) {
      var msMap       = getHeaderMap_(msSheet);
      var msIdCol     = getCol_(msMap, ["Discovery_ID"]);
      var msStatusCol = getCol_(msMap, ["Status"]);
      var msAmountCol = getCol_(msMap, ["Amount"]);
      var msData = msSheet.getRange(2, 1, msSheet.getLastRow() - 1, msSheet.getLastColumn()).getValues();
      msData.forEach(function (r) {
        if (String(r[msIdCol - 1]) === String(discoveryId) && r[msStatusCol - 1] === "Released") {
          total += Number(r[msAmountCol - 1]) || 0;
        }
      });
    }
  } else if (contractType === "Hourly") {
    var hlSheet = ss.getSheetByName("Hourly_Log");
    if (hlSheet && hlSheet.getLastRow() > 1) {
      var hlMap       = getHeaderMap_(hlSheet);
      var hlIdCol     = getCol_(hlMap, ["Discovery_ID"]);
      var hlAmountCol = getCol_(hlMap, ["Amount"]);
      var hlData = hlSheet.getRange(2, 1, hlSheet.getLastRow() - 1, hlSheet.getLastColumn()).getValues();
      hlData.forEach(function (r) {
        if (String(r[hlIdCol - 1]) === String(discoveryId)) {
          total += Number(r[hlAmountCol - 1]) || 0;
        }
      });
    }
  }

  return total;
}

// True only if this Discovery_ID has at least one Milestone_Tracker row AND
// every one of them is Released -- a contract with zero milestones (data not
// set up yet) or any still-Pending/Funded/Delivered one is not done.
function allMilestonesReleased_(ss, discoveryId) {
  var sheet = ss.getSheetByName("Milestone_Tracker");
  if (!sheet || sheet.getLastRow() < 2) return false;

  var map       = getHeaderMap_(sheet);
  var idCol     = getCol_(map, ["Discovery_ID"]);
  var statusCol = getCol_(map, ["Status"]);
  if (!idCol || !statusCol) return false;

  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var found = false;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idCol - 1]) !== String(discoveryId)) continue;
    found = true;
    if (String(data[i][statusCol - 1]).trim() !== "Released") return false;
  }
  return found;
}

// Sets this contract's Contract_Tracker Status to Completed and rolls up
// Total_Released -- same computation the manual Status=Completed edit runs
// (see 14_Edit_Trigger.gs's CONTRACT_TRACKER block), factored out so
// Milestone_Tracker's auto-complete check can call it too. Guarded on
// "already Completed" so re-triggering (e.g. a milestone row getting
// re-saved) doesn't recompute/rewrite Status every time.
function autoCompleteContract_(ss, discoveryId) {
  var contracts = ss.getSheetByName("Contract_Tracker");
  if (!contracts || contracts.getLastRow() < 2) return;

  var ctMap        = getHeaderMap_(contracts);
  var ctIdCol      = getCol_(ctMap, ["Discovery_ID"]);
  var ctStatusCol  = getCol_(ctMap, ["Status"]);
  var ctTypeCol    = getCol_(ctMap, ["Contract_Type"]);
  var ctTotalRelCol = getCol_(ctMap, ["Total_Released"]);
  if (!ctIdCol || !ctStatusCol) return;

  var data = contracts.getRange(2, 1, contracts.getLastRow() - 1, contracts.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][ctIdCol - 1]) !== String(discoveryId)) continue;

    if (String(data[i][ctStatusCol - 1]).trim() === "Completed") return;

    var row = i + 2;
    if (ctTotalRelCol) {
      var ctType  = ctTypeCol ? data[i][ctTypeCol - 1] : "";
      var ctTotal = getContractRecognizedRevenue_(ss, discoveryId, ctType);
      contracts.getRange(row, ctTotalRelCol).setValue(ctTotal);
    }
    contracts.getRange(row, ctStatusCol).setValue("Completed");
    return;
  }
}

// First-session walkthrough plumbing. Checks+marks a one-time script-property
// flag (same FF_-prefixed convention as FF_SETUP_COMPLETE) and returns whether
// it had ALREADY been seen before this call -- so callers can gate on "was
// this the first time" with a single call, no separate check-then-set.
function showWalkthroughSeen_(key) {
  var prop = PropertiesService.getScriptProperties();
  var alreadySeen = prop.getProperty(key) === 'true';
  prop.setProperty(key, 'true');
  return alreadySeen;
}

// Alert-based walkthrough popup for handleEdit hook points -- fires once,
// the first time the given stage is ever reached, then never again.
function showWalkthroughOnce_(key, title, message) {
  if (showWalkthroughSeen_(key)) return;
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Splits FFLib.getQuickNotes' "[Effort], [Scope] & [Portfolio Match]" output
// (e.g. "Normal, Mostly Clear & Strong") into its three labeled parts for
// display -- same split points Lib_JobScoringFormulas.gs's Effort_Level/
// Scope_Rating/Portfolio_Match formulas use. Returns null for anything not
// in that shape (an error message, "Analyzing...", empty), so callers can
// just skip showing anything rather than display garbled text.
function parseAiFitNotes_(notes) {
  var text     = String(notes || '');
  var commaIdx = text.indexOf(',');
  var ampIdx   = text.indexOf('&');
  if (commaIdx === -1 || ampIdx === -1 || ampIdx < commaIdx) return null;

  return {
    effort:    text.substring(0, commaIdx).trim(),
    scope:     text.substring(commaIdx + 1, ampIdx).trim(),
    portfolio: text.substring(ampIdx + 1).trim().replace(/[.!]+$/, '')
  };
}

// ------------------------------------------------------------------------
// GUIDED FIRST-SESSION TOUR
// A fixed 10-step walkthrough chained across a freelancer's first job, start
// to finish (see trigger points in 18_Keyword_Strategy.gs,
// 12_Session_Management.gs, 21_Job_Discovery_Sidebar.gs, 14_Edit_Trigger.gs,
// 11_Job_Classifier.gs, 20_Contract_Setup.gs, and onSelectionChange below).
// Steps 8 (Chat Import) and 9 (Contract Setup) replaced those two sidebars'
// old standalone showWalkthroughOnce_-style info-box tips -- folded into this
// sequential chain instead of staying separate, one-off nudges. Each step is
// a standalone modal (OK continues, Cancel skips), title auto-prefixed with
// 👉 so it reads as part of the tour instead of blending into ordinary
// alerts/warnings, gated on its own one-time-seen flag, same convention as
// showWalkthroughOnce_ above -- but Cancel on ANY step sets FF_TOUR_SKIPPED,
// which silences every remaining step for good, not just that one.
// ------------------------------------------------------------------------
function showTourStep_(key, title, message) {
  var prop = PropertiesService.getScriptProperties();
  if (prop.getProperty('FF_TOUR_SKIPPED') === 'true') return;
  if (prop.getProperty(key) === 'true') return;
  prop.setProperty(key, 'true');

  var ui       = SpreadsheetApp.getUi();
  var response = ui.alert('👉 ' + title, message + '\n\n(Cancel skips the rest of this guided tour.)', ui.ButtonSet.OK_CANCEL);
  if (response === ui.Button.CANCEL) {
    prop.setProperty('FF_TOUR_SKIPPED', 'true');
  }
}

// Counts Job_Discovery rows tagged with the given Session_ID -- same method
// END_SESSION uses, so it stays accurate whether jobs were logged via the
// sidebar or pasted directly into cells. Shared by the yield/halfway
// one-shot checks below and by the sidebar's live countdown.
function getSessionJobCount_(ss, sessionId) {
  var discoverySheet = ss.getSheetByName('Job_Discovery');
  if (!discoverySheet || discoverySheet.getLastRow() < 2) return 0;

  var map          = getHeaderMap_(discoverySheet);
  var sessionIdCol = getCol_(map, ['Session_ID']);
  if (!sessionIdCol) return 0;

  var values = discoverySheet.getRange(2, sessionIdCol, discoverySheet.getLastRow() - 1, 1).getValues();
  var count  = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === sessionId) count++;
  }
  return count;
}

// True exactly once per session -- the call where the active session's
// logged-job count first reaches Session_Yield_Target.
function sessionYieldJustReached_(ss) {
  var prop = PropertiesService.getScriptProperties();
  if (prop.getProperty('SESSION_ACTIVE') !== 'true') return false;
  if (prop.getProperty('SESSION_YIELD_NOTIFIED') === 'true') return false;

  var sessionId   = prop.getProperty('SESSION_ID');
  var count       = getSessionJobCount_(ss, sessionId);
  var yieldTarget = parseInt(getSettings_()['Session_Yield_Target']) || 8;
  if (count < yieldTarget) return false;

  prop.setProperty('SESSION_YIELD_NOTIFIED', 'true');
  return true;
}

// True exactly once per session -- the call where the active session's
// logged-job count first reaches the halfway point to Session_Yield_Target
// (rounded up, so a target of 9 flags at 5).
function sessionHalfwayJustReached_(ss) {
  var prop = PropertiesService.getScriptProperties();
  if (prop.getProperty('SESSION_ACTIVE') !== 'true') return false;
  if (prop.getProperty('SESSION_HALFWAY_NOTIFIED') === 'true') return false;

  var sessionId   = prop.getProperty('SESSION_ID');
  var count       = getSessionJobCount_(ss, sessionId);
  var yieldTarget = parseInt(getSettings_()['Session_Yield_Target']) || 8;
  var halfway     = Math.ceil(yieldTarget / 2);
  if (count < halfway) return false;

  prop.setProperty('SESSION_HALFWAY_NOTIFIED', 'true');
  return true;
}

// Fires once per session, the moment the halfway point is reached --
// suggests switching keywords to keep results fresh for the back half of
// the session.
function handleSessionHalfwayReached_(ss) {
  if (!sessionHalfwayJustReached_(ss)) return;

  SpreadsheetApp.getUi().alert(
    'Halfway There',
    'You\'re halfway to this session\'s job target. Consider switching to a different keyword to keep results fresh for the rest of the session.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Dispatch point for the moment a session's yield target is reached --
// shared by both job-logging entry paths (sidebar and direct cell paste).
// Fires the first-session-only guided tour steps, then the every-session
// yield summary, both gated on the single sessionYieldJustReached_ check
// (it's stateful/one-shot per session, so it must only be called once here).
// Returns true the one time it actually fires, so callers (job_saveEntry)
// know the session target was just hit -- e.g. to close the Log New Job
// sidebar automatically.
function handleSessionYieldReached_(ss) {
  if (!sessionYieldJustReached_(ss)) return false;

  showTourStep_(
    'FF_TOUR_STEP4_YIELD_SEEN',
    'Session Target Reached',
    'You\'ve hit this session\'s job target. Head to Job_Scoring -- nothing needs to be entered there anymore, just check which jobs scored APPLY.'
  );
  showTourStep_(
    'FF_TOUR_STEP5_PROPOSAL_GEN_SEEN',
    'Next: Proposal_Generator',
    'Any job scored APPLY automatically moves to Proposal_Generator. Head there next to review the AI-drafted proposals.'
  );

  showSessionYieldSummary_(ss);
  return true;
}

// Every-session popup (not just the first) shown the moment the yield
// target is reached -- reports how many of THIS session's jobs moved to
// Job_Scoring (Discovery_Action = Move to Scoring) and how many of those
// were scored APPLY into Proposal_Generator. Job_Scoring carries no
// Session_ID of its own, so the APPLY count is cross-referenced by
// Discovery_ID against this session's Job_Discovery rows.
function showSessionYieldSummary_(ss) {
  var prop      = PropertiesService.getScriptProperties();
  var sessionId = prop.getProperty('SESSION_ID');

  var discoverySheet = ss.getSheetByName('Job_Discovery');
  if (!discoverySheet || discoverySheet.getLastRow() < 2) return;

  var discMap      = getHeaderMap_(discoverySheet);
  var discIdCol    = getCol_(discMap, ['Discovery_ID']);
  var sessionIdCol = getCol_(discMap, ['Session_ID']);
  var actionCol    = getCol_(discMap, ['Discovery_Action']);
  if (!discIdCol || !sessionIdCol || !actionCol) return;

  var discData = discoverySheet
    .getRange(2, 1, discoverySheet.getLastRow() - 1, discoverySheet.getLastColumn())
    .getValues();

  var sessionDiscoveryIds = {};
  var movedToScoring      = 0;

  for (var i = 0; i < discData.length; i++) {
    var rowSession = String(discData[i][sessionIdCol - 1]).trim().toUpperCase();
    if (rowSession !== sessionId) continue;

    sessionDiscoveryIds[String(discData[i][discIdCol - 1]).trim()] = true;
    if (String(discData[i][actionCol - 1]).trim() === 'Move to Scoring') movedToScoring++;
  }

  var applyCount   = 0;
  var scoringSheet = ss.getSheetByName('Job_Scoring');
  if (scoringSheet && scoringSheet.getLastRow() > 1) {
    var jsMap    = getHeaderMap_(scoringSheet);
    var jsIdCol  = getCol_(jsMap, ['Discovery_ID']);
    var jsDecCol = getCol_(jsMap, ['Final_Decision']);

    if (jsIdCol && jsDecCol) {
      var jsData = scoringSheet
        .getRange(2, 1, scoringSheet.getLastRow() - 1, scoringSheet.getLastColumn())
        .getValues();

      for (var j = 0; j < jsData.length; j++) {
        var jsDiscoveryId = String(jsData[j][jsIdCol - 1]).trim();
        if (sessionDiscoveryIds[jsDiscoveryId] && String(jsData[j][jsDecCol - 1]).trim() === 'APPLY') {
          applyCount++;
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    'Session Yield Summary',
    movedToScoring + ' job(s) from this session moved to Job_Scoring.\n' +
    applyCount + ' job(s) scored APPLY and moved to Proposal_Generator.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Simple trigger, no registration needed (auto-fires on any selection
// change, including switching the active sheet). Only ever used for the
// guided tour's last step -- fires once, the first time the active sheet
// becomes Proposal_Generator. Do NOT define onSelectionChange anywhere else
// in this project; Apps Script only recognizes one.
function onSelectionChange(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== 'Proposal_Generator') return;

  // The Setup Wizard creates/activates every pipeline sheet as it builds
  // them, including Proposal_Generator -- that alone counts as a "selection
  // change" and fired this step mid-wizard, stacked on top of the keyword
  // strategy summary dialog (confirmed during the 2026-07-05 walkthrough).
  // Only fire once setup is actually done, so this stays the tour's last
  // step instead of firing before the first one.
  if (PropertiesService.getScriptProperties().getProperty('FF_SETUP_COMPLETE') !== 'true') return;

  showTourStep_(
    'FF_TOUR_STEP6_RUN_CLASSIFICATION_SEEN',
    'Classify These Jobs',
    'Before writing proposals, run System Tools > Run Job Classification. It fills in Job_Type and picks a matching proposal template for each job.'
  );
}
