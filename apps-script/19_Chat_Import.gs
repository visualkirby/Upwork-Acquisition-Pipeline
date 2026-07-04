/**
 * ============================================================
 * 19. CHAT IMPORT SIDEBAR
 * Opens automatically when Proposal_Tracker's Interview column
 * flips to "Y" (see 14_Edit_Trigger.gs), or manually via the
 * "Import Client Chat" menu item. Lets the user paste a raw
 * copied Upwork chat transcript and have AI parse it into one
 * row per message in Client_Chat_Log.
 * ============================================================
 */
var PENDING_CHAT_IMPORT_PROP_ = 'PENDING_CHAT_IMPORT_DISCOVERY_ID';

function openChatImportSidebar_(discoveryId) {
  if (discoveryId) {
    PropertiesService.getScriptProperties().setProperty(PENDING_CHAT_IMPORT_PROP_, String(discoveryId));
  }
  var html = HtmlService.createHtmlOutputFromFile('ChatImportSidebar')
    .setTitle('Import Client Chat')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function IMPORT_CLIENT_CHAT() {
  openChatImportSidebar_(null);
}

function chat_getContext() {
  var prop = PropertiesService.getScriptProperties();
  var pendingId = prop.getProperty(PENDING_CHAT_IMPORT_PROP_) || '';
  prop.deleteProperty(PENDING_CHAT_IMPORT_PROP_);

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName('Proposal_Tracker');
  var jobs    = [];

  if (tracker && tracker.getLastRow() > 1) {
    var map          = getHeaderMap_(tracker);
    var idCol        = getCol_(map, ['Discovery_ID']);
    var titleCol     = getCol_(map, ['Job_Title']);
    var clientCol    = getCol_(map, ['Client_Name', 'Client Name']);
    var interviewCol = getCol_(map, ['Interview']);

    if (idCol && titleCol && clientCol && interviewCol) {
      var data = tracker.getRange(2, 1, tracker.getLastRow() - 1, tracker.getLastColumn()).getValues();
      data.forEach(function (row) {
        if (String(row[interviewCol - 1]).trim() === 'Y') {
          jobs.push({
            discoveryId: row[idCol - 1],
            jobTitle:    row[titleCol - 1],
            clientName:  row[clientCol - 1]
          });
        }
      });
    }
  }

  var showWalkthrough = !showWalkthroughSeen_('FF_WALKTHROUGH_CHAT_LOG_SEEN');

  return { pendingDiscoveryId: pendingId, jobs: jobs, showWalkthrough: showWalkthrough };
}

function chat_parseTranscript(discoveryId, rawText) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName('Proposal_Tracker');
  var log     = ss.getSheetByName('Client_Chat_Log');

  if (!tracker || !log) return { ok: false, message: 'Required sheets not found.' };

  var ptMap       = getHeaderMap_(tracker);
  var idCol       = getCol_(ptMap, ['Discovery_ID']);
  var titleCol    = getCol_(ptMap, ['Job_Title']);
  var clientCol   = getCol_(ptMap, ['Client_Name', 'Client Name']);

  var jobTitle   = '';
  var clientName = '';
  if (idCol && titleCol && clientCol && tracker.getLastRow() > 1) {
    var data = tracker.getRange(2, 1, tracker.getLastRow() - 1, tracker.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][idCol - 1]) === String(discoveryId)) {
        jobTitle   = data[i][titleCol - 1];
        clientName = data[i][clientCol - 1];
        break;
      }
    }
  }

  var apiKey;
  try {
    apiKey = getApiKey_();
  } catch (err) {
    return { ok: false, message: err.message };
  }

  var settings       = getSettings_();
  var freelancerName = settings['Freelancer_Name'] || 'the freelancer';

  var result = FFLib.parseChatTranscript(rawText, freelancerName, apiKey);
  if (!result.ok) return result;

  // Re-pasting replaces this job's messages -- filter its old rows out,
  // then rewrite the sheet as [every other job's rows unchanged] + [this
  // job's fresh messages]. Clearing old rows in place instead (leaving
  // gaps) would let findFirstEmptyRowByColumn_ scatter new rows into a
  // blank gap left by a DIFFERENT job's earlier re-paste, interleaving
  // jobs instead of keeping the log compact and append-ordered.
  var logMap    = getHeaderMap_(log);
  var logIdCol  = getCol_(logMap, ['Discovery_ID']);
  var lastCol   = log.getLastColumn();

  var keptRows = [];
  if (log.getLastRow() > 1 && logIdCol) {
    var existing = log.getRange(2, 1, log.getLastRow() - 1, lastCol).getValues();
    keptRows = existing.filter(function (r) {
      return String(r[logIdCol - 1]) !== String(discoveryId);
    });
  }

  // Resolve every column by header name (never a hardcoded position) so
  // this keeps working regardless of how Client_Chat_Log's columns are
  // ordered -- built into a fixed-width array per row since keptRows
  // (read straight off the live sheet) and these freshly-built rows get
  // concatenated into one setValues() call and must line up exactly.
  var colTitle  = getCol_(logMap, ['Job_Title']);
  var colClient = getCol_(logMap, ['Client_Name', 'Client Name']);
  var colMsgNum = getCol_(logMap, ['Message_Number']);
  var colSender = getCol_(logMap, ['Sender_Name']);
  var colDir    = getCol_(logMap, ['Direction']);
  var colTime   = getCol_(logMap, ['Message_Time']);
  var colText   = getCol_(logMap, ['Message_Text']);
  var colSynced = getCol_(logMap, ['Last_Synced']);

  var now          = new Date();
  var freelancerLc = String(freelancerName).trim().toLowerCase();
  var newRows = result.messages.map(function (m, i) {
    var senderLc  = String(m.sender || '').trim().toLowerCase();
    var direction = senderLc === freelancerLc ? 'Freelancer' : 'Client';
    var arr = new Array(lastCol).fill('');
    if (logIdCol) arr[logIdCol  - 1] = discoveryId;
    if (colTitle)  arr[colTitle  - 1] = jobTitle;
    if (colClient) arr[colClient - 1] = clientName;
    if (colMsgNum) arr[colMsgNum - 1] = i + 1;
    if (colSender) arr[colSender - 1] = m.sender || '';
    if (colDir)    arr[colDir    - 1] = direction;
    if (colTime)   arr[colTime   - 1] = m.time || '';
    if (colText)   arr[colText   - 1] = m.message || '';
    if (colSynced) arr[colSynced - 1] = now;
    return arr;
  });

  var allRows = keptRows.concat(newRows);

  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, lastCol).clearContent();
  }
  if (allRows.length > 0) {
    log.getRange(2, 1, allRows.length, lastCol).setValues(allRows);
  }

  return { ok: true, count: newRows.length };
}
