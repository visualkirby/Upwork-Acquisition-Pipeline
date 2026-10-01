/**
 * ============================================================
 * 12. SESSION MANAGEMENT
 * START_SESSION: prompts for Session ID and keywords, stores
 *   state in PropertiesService, confirms session target.
 * END_SESSION: computes all session stats, writes to Session_Log,
 *   updates Keyword_Search_List last-searched and yield.
 * ============================================================
 */
function START_SESSION() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var prop = PropertiesService.getScriptProperties();

  if (prop.getProperty('SESSION_ACTIVE') === 'true') {
    ui.alert(
      'A session is already active (Session ' + prop.getProperty('SESSION_ID') + ').\n' +
      'Use System Tools > End Session to close it before starting a new one.'
    );
    return;
  }

  // Every session starts by matching the sheet's Connects balance to
  // Upwork's (30_Connects_Sync.gs). Cancel stops the session start, same as
  // the steps below.
  var balance = promptConnectsBalanceSync_(ss, 'Start Session -- Step 1 of 3', '');
  if (balance.cancelled) return;

  var idResponse = ui.prompt(
    'Start Session -- Step 2 of 3',
    'Enter your Session ID (e.g. S001):',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;

  var sessionId = idResponse.getResponseText().trim().toUpperCase();
  if (!sessionId) {
    ui.alert('Session ID cannot be blank.');
    return;
  }

  var kwResponse = ui.prompt(
    'Start Session -- Step 3 of 3',
    'Which keywords are you searching this session?\n' +
    '(Enter comma-separated, e.g. Power BI Sales Dashboard, Excel Finance Report)',
    ui.ButtonSet.OK_CANCEL
  );
  if (kwResponse.getSelectedButton() !== ui.Button.OK) return;

  var keywords = kwResponse.getResponseText().trim();
  if (!keywords) {
    ui.alert('Please enter at least one keyword.');
    return;
  }

  var discoverySheet = ss.getSheetByName('Job_Discovery');
  var startRowCount  = discoverySheet ? Math.max(discoverySheet.getLastRow() - 1, 0) : 0;

  prop.setProperties({
    'SESSION_ACTIVE':           'true',
    'SESSION_ID':               sessionId,
    'SESSION_START_TIME':       new Date().toISOString(),
    'SESSION_KEYWORDS':         keywords,
    'SESSION_START_ROW_COUNT':  String(startRowCount),
    'SESSION_DUPE_COUNT':       '0'
  });

  var yieldTarget = parseInt(getSettings_()['Session_Yield_Target']) || 8;

  var connectsLine = balance.from !== undefined
    ? 'Connects: ' + balance.to + ' (synced from Upwork, was ' + balance.from + ')\n'
    : 'Connects: ' + (Number(getConnectsHelperValue_(ss, 'Current_Connect_Balance')) || 0) + '\n';

  ui.alert(
    'Session ' + sessionId + ' started.\n\n' +
    connectsLine +
    'Keywords: ' + keywords + '\n' +
    'Start time: ' + formatTime_(new Date()) + '\n\n' +
    'Session target: ' + yieldTarget + ' unique new jobs.\n' +
    'Go search Upwork -- every job you log will be tracked automatically.'
  );

  showTourStep_(
    'FF_TOUR_STEP3_LOG_JOB_SEEN',
    'Start Logging Jobs',
    'The Log New Job sidebar is opening now -- use it to log every job you find this session.'
  );

  openJobDiscoverySidebar_();
}


function END_SESSION() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var prop = PropertiesService.getScriptProperties();

  if (prop.getProperty('SESSION_ACTIVE') !== 'true') {
    ui.alert('No active session found.\nUse System Tools > Start Session to begin one.');
    return;
  }

  var sessionId    = prop.getProperty('SESSION_ID')           || '';
  var startTimeStr = prop.getProperty('SESSION_START_TIME')   || '';
  var keywords     = prop.getProperty('SESSION_KEYWORDS')     || '';
  var startCount   = parseInt(prop.getProperty('SESSION_START_ROW_COUNT') || '0', 10);
  var dupeCount    = parseInt(prop.getProperty('SESSION_DUPE_COUNT')      || '0', 10);

  var endTime   = new Date();
  var startTime = startTimeStr ? new Date(startTimeStr) : endTime;

  // A session left open past Session_Stale_Hours closes at its last real
  // activity instead of now, so Duration and the Sent/Skip window cover the
  // session itself rather than every day it sat open (S014 sat open 29 days).
  var staleNote = '';
  var stale     = getSessionStaleInfo_(ss);
  if (stale.stale) {
    var lastActivity = getSessionLastActivity_(ss, sessionId, startTime, stale.limitHours);
    endTime   = lastActivity || startTime;
    staleNote = 'Open ' + stale.ageLabel + ', so it was closed at its last activity (' +
      Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'MMM d, h:mm a') + '), not now.';
  }
  var durationMs = endTime - startTime;

  var discoverySheet = ss.getSheetByName('Job_Discovery');
  var jobsLogged     = 0;
  var movedToScoring = 0;
  var reviewLater    = 0;

  if (discoverySheet && discoverySheet.getLastRow() > 1) {
    var discMap      = getHeaderMap_(discoverySheet);
    var sessionIdCol = getCol_(discMap, ['Session_ID']);
    var actionCol    = getCol_(discMap, ['Discovery_Action']);
    var totalRows    = discoverySheet.getLastRow() - 1;

    if (sessionIdCol && actionCol && totalRows > 0) {
      var discData = discoverySheet
        .getRange(2, 1, totalRows, discoverySheet.getLastColumn())
        .getValues();

      for (var i = 0; i < discData.length; i++) {
        var rowSession = String(discData[i][sessionIdCol - 1]).trim().toUpperCase();
        var rowAction  = String(discData[i][actionCol   - 1]).trim();

        if (rowSession === sessionId) {
          jobsLogged++;
          if (rowAction === 'Move to Scoring') movedToScoring++;
          else if (rowAction === 'Review Later') reviewLater++;
        }
      }
    }
  }

  var YIELD_TARGET = parseInt(getSettings_()['Session_Yield_Target']) || 8;
  var sessionYield = jobsLogged;
  var saturating   = sessionYield < YIELD_TARGET;
  var satFlag      = saturating
    ? 'YES -- yield ' + sessionYield + '/' + YIELD_TARGET + ' (below target)'
    : 'No -- yield '  + sessionYield + '/' + YIELD_TARGET;

  // APPLY jobs still waiting on a decision: no Proposal_Generator row yet,
  // or a row with Proposal_Status blank or Ready.
  var scoringSheet    = ss.getSheetByName('Job_Scoring');
  var applyNoProposal = 0;

  if (scoringSheet && scoringSheet.getLastRow() > 1) {
    var jsMap         = getHeaderMap_(scoringSheet);
    var jsDecisionCol = getCol_(jsMap, ['Final_Decision']);
    var jsIdCol       = getCol_(jsMap, ['Discovery_ID']);
    var statusById    = getProposalStatusById_(ss);

    if (jsDecisionCol && jsIdCol) {
      var jsData = scoringSheet
        .getRange(2, 1, scoringSheet.getLastRow() - 1, scoringSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < jsData.length; i++) {
        var decision = String(jsData[i][jsDecisionCol - 1]).trim();
        var pgStatus0 = statusById[String(jsData[i][jsIdCol - 1]).trim()] || '';
        if (decision === 'APPLY' && pgStatus0 !== 'Sent' && pgStatus0 !== 'Skip') {
          applyNoProposal++;
        }
      }
    }
  }

  var proposalTrigger = applyNoProposal >= 5
    ? 'YES -- ' + applyNoProposal + ' APPLY jobs have no proposal sent. Send at least 2 before next session.'
    : 'No -- ' + applyNoProposal + ' APPLY jobs pending (' + (5 - applyNoProposal) + ' more needed to trigger).';

  var pgSheet          = ss.getSheetByName('Proposal_Generator');
  var proposalsSent    = 0;
  var proposalsSkipped = 0;
  var connectsSpent    = 0;

  if (pgSheet && pgSheet.getLastRow() > 1) {
    var pgMap         = getHeaderMap_(pgSheet);
    var pgStatusCol   = getCol_(pgMap, ['Proposal_Status']);
    var pgSentDateCol = getCol_(pgMap, ['Proposal_Sent_Date']);
    var pgSkipDateCol = getCol_(pgMap, ['Proposal_Skip_Date']);
    var pgConnectsCol = getCol_(pgMap, ['Total_Connects_Spent', 'Connects_Required']);

    if (pgStatusCol) {
      var pgData = pgSheet
        .getRange(2, 1, pgSheet.getLastRow() - 1, pgSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < pgData.length; i++) {
        var pgStatus   = String(pgData[i][pgStatusCol - 1]).trim();
        var pgSentDate = pgSentDateCol ? pgData[i][pgSentDateCol - 1] : null;
        var pgSkipDate = pgSkipDateCol ? pgData[i][pgSkipDateCol - 1] : null;

        if (pgStatus === 'Sent' && pgSentDate) {
          var sentDateObj = new Date(pgSentDate);
          if (sentDateObj >= startTime && sentDateObj <= endTime) {
            proposalsSent++;
            if (pgConnectsCol) {
              connectsSpent += Number(pgData[i][pgConnectsCol - 1]) || 0;
            }
          }
        } else if (pgStatus === 'Skip' && pgSkipDate) {
          var skipDateObj = new Date(pgSkipDate);
          if (skipDateObj >= startTime && skipDateObj <= endTime) {
            proposalsSkipped++;
          }
        }
      }
    }
  }

  var notesResponse = ui.prompt(
    'End Session -- Notes',
    'Any notes for this session? (optional -- press OK to skip)',
    ui.ButtonSet.OK_CANCEL
  );
  var sessionNotes = (notesResponse.getSelectedButton() === ui.Button.OK)
    ? notesResponse.getResponseText().trim()
    : '';

  var searchListSheet = ss.getSheetByName('Keyword_Search_List');
  if (searchListSheet && keywords) {
    var slMap           = getHeaderMap_(searchListSheet);
    var slQueryCol      = getCol_(slMap, ['Search_Query']);
    var slLastSearchCol = getCol_(slMap, ['Last_Searched']);
    var slYieldCol      = getCol_(slMap, ['Session_Yield']);
    var slLastRow       = searchListSheet.getLastRow();

    var keywordList = keywords.split(',').map(function(k) {
      return k.trim().toLowerCase();
    });

    if (slQueryCol && slLastRow > 1) {
      var slData = searchListSheet
        .getRange(2, 1, slLastRow - 1, searchListSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < slData.length; i++) {
        var rowQuery = String(slData[i][slQueryCol - 1]).trim().toLowerCase();
        if (keywordList.indexOf(rowQuery) !== -1) {
          var dataRow = i + 2;
          if (slLastSearchCol) {
            searchListSheet.getRange(dataRow, slLastSearchCol).setValue(endTime);
          }
          if (slYieldCol) {
            var perKeyword = Math.round(sessionYield / keywordList.length);
            searchListSheet.getRange(dataRow, slYieldCol).setValue(perKeyword);
          }
        }
      }
    }
  }

  var logSheet = ss.getSheetByName('Session_Log');
  if (logSheet) {
    var logMap     = getHeaderMap_(logSheet);
    var nextLogRow = findFirstEmptyRowByColumn_(logSheet, 1);
    if (logSheet.getLastRow() <= 1) nextLogRow = 2;

    setCellValue_(logSheet, nextLogRow, logMap, ['Session_ID'],            sessionId);
    setCellValue_(logSheet, nextLogRow, logMap, ['Date'],                  endTime);
    setCellValue_(logSheet, nextLogRow, logMap, ['Start_Time'],            formatTime_(startTime));
    setCellValue_(logSheet, nextLogRow, logMap, ['End_Time'],              formatTime_(endTime));
    setCellValue_(logSheet, nextLogRow, logMap, ['Duration'],              formatDuration_(durationMs));
    setCellValue_(logSheet, nextLogRow, logMap, ['Keywords_Searched'],     keywords);
    setCellValue_(logSheet, nextLogRow, logMap, ['Jobs_Logged'],           jobsLogged);
    setCellValue_(logSheet, nextLogRow, logMap, ['Jobs_Moved_To_Scoring'], movedToScoring);
    setCellValue_(logSheet, nextLogRow, logMap, ['Jobs_Review_Later'],     reviewLater);
    setCellValue_(logSheet, nextLogRow, logMap, ['Duplicates_Skipped'],    dupeCount);
    setCellValue_(logSheet, nextLogRow, logMap, ['Session_Yield'],         sessionYield);
    setCellValue_(logSheet, nextLogRow, logMap, ['Saturation_Flag'],       saturating ? 'Saturating' : 'Healthy');
    setCellValue_(logSheet, nextLogRow, logMap, ['Proposal_Trigger'],      applyNoProposal >= 5 ? 'TRIGGERED' : 'Clear');
    setCellValue_(logSheet, nextLogRow, logMap, ['Proposals_Sent'],        proposalsSent);
    setCellValue_(logSheet, nextLogRow, logMap, ['Proposals_Skipped'],     proposalsSkipped);
    setCellValue_(logSheet, nextLogRow, logMap, ['Connects_Spent'],        connectsSpent);
    setCellValue_(logSheet, nextLogRow, logMap, ['Notes'],
      [sessionNotes, staleNote].filter(function (n) { return n; }).join(' '));
  }

  // Session-scoped keys only -- deleteAllProperties() previously wiped every
  // script property on every End Session, including UPWORK_OPENAI_API_KEY,
  // FF_SETUP_COMPLETE (reopening the wizard next onOpen), and every
  // walkthrough/tour seen-flag.
  prop.deleteProperty('SESSION_ACTIVE');
  prop.deleteProperty('SESSION_ID');
  prop.deleteProperty('SESSION_START_TIME');
  prop.deleteProperty('SESSION_KEYWORDS');
  prop.deleteProperty('SESSION_START_ROW_COUNT');
  prop.deleteProperty('SESSION_DUPE_COUNT');
  prop.deleteProperty('SESSION_YIELD_NOTIFIED');
  prop.deleteProperty('SESSION_HALFWAY_NOTIFIED');

  var runLog =
    '════════════════════════════════\n' +
    '  SESSION RUN LOG -- ' + sessionId + '\n' +
    '════════════════════════════════\n\n' +
    'Date:           ' + endTime.toLocaleDateString()  + '\n' +
    'Start:          ' + formatTime_(startTime)         + '\n' +
    'End:            ' + formatTime_(endTime)           + '\n' +
    'Duration:       ' + formatDuration_(durationMs)    + '\n' +
    (staleNote ? staleNote + '\n' : '') + '\n' +
    'Keywords:\n  ' + keywords.split(',').join('\n  ')  + '\n\n' +
    '-- Discovery ──────────────────\n' +
    'Jobs logged:        ' + jobsLogged     + '\n' +
    'Moved to Scoring:   ' + movedToScoring + '\n' +
    'Review Later:       ' + reviewLater    + '\n' +
    'Duplicates skipped: ' + dupeCount      + '\n' +
    'Session yield:      ' + sessionYield   + ' / ' + YIELD_TARGET + ' target\n\n' +
    '-- Health ─────────────────────\n' +
    'Saturation flag:    ' + satFlag          + '\n' +
    'Proposal trigger:   ' + proposalTrigger  + '\n\n' +
    '-- Proposals ──────────────────\n' +
    'Sent this session:    ' + proposalsSent    + '\n' +
    'Skipped this session: ' + proposalsSkipped + '\n' +
    (connectsSpent > 0 ? 'Connects spent:     ' + connectsSpent + '\n' : '') +
    '\n' +
    (sessionNotes ? 'Notes: ' + sessionNotes + '\n\n' : '') +
    (logSheet ? 'Log written to Session_Log.' : 'Session_Log sheet not found -- log not saved.');

  ui.alert(runLog);
}

// { Discovery_ID: Proposal_Status } for every Proposal_Generator row.
function getProposalStatusById_(ss) {
  var pgSheet = ss.getSheetByName('Proposal_Generator');
  var out     = {};
  if (!pgSheet || pgSheet.getLastRow() < 2) return out;

  var pgMap     = getHeaderMap_(pgSheet);
  var idCol     = getCol_(pgMap, ['Discovery_ID']);
  var statusCol = getCol_(pgMap, ['Proposal_Status']);
  if (!idCol || !statusCol) return out;

  var data = pgSheet.getRange(2, 1, pgSheet.getLastRow() - 1, pgSheet.getLastColumn()).getValues();
  data.forEach(function (r) {
    var id = String(r[idCol - 1]).trim();
    if (id) out[id] = String(r[statusCol - 1]).trim();
  });
  return out;
}

// { stale, sessionId, hoursOpen, ageLabel } for the active session. stale
// is false with no active session. The threshold is Settings >
// Session_Stale_Hours, 12 when missing or blank.
function getSessionStaleInfo_(ss) {
  var prop = PropertiesService.getScriptProperties();
  var info = { stale: false, sessionId: '', hoursOpen: 0, ageLabel: '' };
  if (prop.getProperty('SESSION_ACTIVE') !== 'true') return info;

  var startStr = prop.getProperty('SESSION_START_TIME');
  if (!startStr) return info;

  var limit = parseFloat(getSettings_()['Session_Stale_Hours']);
  if (isNaN(limit) || limit <= 0) limit = 12;

  info.sessionId = prop.getProperty('SESSION_ID') || '';
  info.hoursOpen = (new Date() - new Date(startStr)) / 3600000;
  info.ageLabel  = info.hoursOpen >= 48
    ? Math.floor(info.hoursOpen / 24) + ' days'
    : Math.floor(info.hoursOpen) + ' hours';
  info.limitHours = limit;
  info.stale      = info.hoursOpen > limit;
  return info;
}

// When a session's real work ended. Anchored on its own logged jobs (the
// latest Date_Found tagged with its Session_ID, or its start if it has
// none), then extended by any proposal marked Sent/Skip within windowHours
// after that anchor -- bidding in the same sitting. Sent/Skip dates carry
// no Session_ID, so proposals sent days later while the session sat
// forgotten must not stretch it back out. null when the session has no
// jobs and no proposals in that window.
function getSessionLastActivity_(ss, sessionId, startTime, windowHours) {
  var asDate = function (v) {
    return (v instanceof Date && !isNaN(v.getTime())) ? v : null;
  };

  var lastJob = null;
  var jd = ss.getSheetByName('Job_Discovery');
  if (jd && jd.getLastRow() > 1) {
    var jdMap   = getHeaderMap_(jd);
    var sidCol  = getCol_(jdMap, ['Session_ID']);
    var dateCol = getCol_(jdMap, ['Date_Found']);
    if (sidCol && dateCol) {
      jd.getRange(2, 1, jd.getLastRow() - 1, jd.getLastColumn()).getValues().forEach(function (r) {
        if (String(r[sidCol - 1]).trim().toUpperCase() !== sessionId) return;
        var d = asDate(r[dateCol - 1]);
        if (d && d >= startTime && (!lastJob || d > lastJob)) lastJob = d;
      });
    }
  }

  var anchor   = lastJob || startTime;
  var windowMs = (windowHours || 12) * 3600000;
  var latest   = lastJob;

  var pg = ss.getSheetByName('Proposal_Generator');
  if (pg && pg.getLastRow() > 1) {
    var pgMap   = getHeaderMap_(pg);
    var sentCol = getCol_(pgMap, ['Proposal_Sent_Date']);
    var skipCol = getCol_(pgMap, ['Proposal_Skip_Date']);
    pg.getRange(2, 1, pg.getLastRow() - 1, pg.getLastColumn()).getValues().forEach(function (r) {
      [sentCol ? r[sentCol - 1] : null, skipCol ? r[skipCol - 1] : null].forEach(function (v) {
        var d = asDate(v);
        if (!d || d < anchor || d - anchor > windowMs) return;
        if (!latest || d > latest) latest = d;
      });
    });
  }
  return latest;
}

// Called before Log New Job opens. Returns false when the freelancer
// cancels. Yes ends the stale session (at its last activity) first.
function confirmStaleSessionBeforeLogging_(ss) {
  var info = getSessionStaleInfo_(ss);
  if (!info.stale) return true;

  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    'Session ' + info.sessionId + ' is still open',
    'Session ' + info.sessionId + ' has been open ' + info.ageLabel + '.\n\n' +
    'Yes = end it now (it closes at its last activity, then Log New Job opens)\n' +
    'No = keep logging jobs into it',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (answer === ui.Button.CANCEL || answer === ui.Button.CLOSE) return false;
  if (answer === ui.Button.YES) END_SESSION();
  return true;
}

// Non-blocking notice from onOpen. Toasts work in a simple trigger, unlike
// a modal alert; any failure is ignored so the menu always builds.
function showStaleSessionToast_() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var info = getSessionStaleInfo_(ss);
    if (!info.stale) return;
    ss.toast('Session ' + info.sessionId + ' has been open ' + info.ageLabel +
      '. System Tools > End Session closes it.', 'Session still open', 15);
  } catch (e) {}
}
