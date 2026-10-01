/**
 * ============================================================
 * 30. CONNECTS BALANCE SYNC
 * Current_Connect_Balance only moves when a proposal is marked Sent or a
 * Connect_Replenishment / Connect_Returned value is typed in. Upwork moves
 * the real balance on its own too (monthly free Connects, earned Connects,
 * refunds on jobs closed without a hire, boost refunds), so the sheet
 * drifts. On 2026-10-01 it reached -8 and flipped every job to "Cannot
 * Afford". There's no simple way for a Sheet to read the Upwork balance,
 * so the freelancer types it in:
 *   - every Start Session (first prompt)
 *   - System Tools > Sync Connects Balance, anytime
 *   - automatically when marking a proposal Sent pushes the balance below 0
 *
 * A sync overwrites Current_Connect_Balance, stamps
 * Connects_Balance_Last_Synced, and adds the gap to
 * Connects_Balance_Adjustments so drift stays visible. Purchases and the
 * Proposal_Tracker-derived totals are untouched.
 * ============================================================
 */
var CONNECTS_SYNC_ROWS_ = [
  ['Connects_Balance_Last_Synced', ''],
  ['Connects_Balance_Adjustments', 0]
];

// Appends any CONNECTS_SYNC_ROWS_ metric a Connects_Helper sheet is missing
// (copies created before these existed). Returns the names added.
function ensureConnectsSyncRows_(ss) {
  var sheet = ss.getSheetByName('Connects_Helper');
  if (!sheet || sheet.getLastRow() < 1) return [];

  var names = sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); })
    : [];
  var missing = CONNECTS_SYNC_ROWS_.filter(function (m) { return names.indexOf(m[0]) === -1; });
  if (missing.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 2).setValues(missing);
  }
  return missing.map(function (m) { return m[0]; });
}

function getConnectsHelperValue_(ss, metricName) {
  var sheet = ss.getSheetByName('Connects_Helper');
  if (!sheet || sheet.getLastRow() < 2) return null;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === metricName) return data[i][1];
  }
  return null;
}

// Asks for the balance Upwork shows and applies it. heading/reason frame
// the prompt for the caller. Returns { cancelled } when the freelancer
// pressed Cancel, { kept } for a blank answer, or { from, to } after a
// sync. A non-number re-asks once, then keeps the sheet value.
function promptConnectsBalanceSync_(ss, heading, reason) {
  var ui      = SpreadsheetApp.getUi();
  var current = Number(getConnectsHelperValue_(ss, 'Current_Connect_Balance')) || 0;
  var message =
    (reason ? reason + '\n\n' : '') +
    'What is your Connects balance on Upwork right now?\n' +
    '(Upwork: Find Work > Connects, or the Connects count on any job\'s Apply page)\n\n' +
    'The sheet shows ' + current + '. Leave blank to keep it.';

  for (var attempt = 0; attempt < 2; attempt++) {
    var response = ui.prompt(heading, message, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return { cancelled: true };

    var text = response.getResponseText().trim();
    if (text === '') return { kept: true };

    var value = Number(text.replace(/,/g, ''));
    if (!isNaN(value) && value >= 0 && Math.floor(value) === value) {
      applyConnectsBalanceSync_(ss, current, value);
      return { from: current, to: value };
    }
    message = '"' + text + '" isn\'t a whole number of Connects. Enter just the number, e.g. 84.\n\n' +
      'The sheet shows ' + current + '. Leave blank to keep it.';
  }
  return { kept: true };
}

function applyConnectsBalanceSync_(ss, from, to) {
  ensureConnectsSyncRows_(ss);
  var sheet = ss.getSheetByName('Connects_Helper');
  var lock  = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var name = String(data[i][0]).trim();
      if (name === 'Current_Connect_Balance')      sheet.getRange(i + 2, 2).setValue(to);
      if (name === 'Connects_Balance_Last_Synced') sheet.getRange(i + 2, 2).setValue(new Date());
      if (name === 'Connects_Balance_Adjustments') {
        sheet.getRange(i + 2, 2).setValue((Number(data[i][1]) || 0) + (to - from));
      }
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // A new balance changes Connects_Affordability, which can move jobs into
  // or out of APPLY. Script writes don't fire handleEdit, so sync here.
  syncProposalGenerator_(ss);
}

function SYNC_CONNECTS_BALANCE() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var result = promptConnectsBalanceSync_(ss, 'Sync Connects Balance', '');
  if (result.from !== undefined) {
    SpreadsheetApp.getUi().alert(
      'Connects balance updated: ' + result.from + ' -> ' + result.to +
      ' (' + (result.to - result.from >= 0 ? '+' : '') + (result.to - result.from) + ').'
    );
  }
}
