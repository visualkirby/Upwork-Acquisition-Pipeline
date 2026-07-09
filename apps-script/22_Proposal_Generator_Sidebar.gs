/**
 * ============================================================
 * 22. PROPOSAL GENERATOR SIDEBAR
 * Guided-entry alternative to editing Proposal_Generator cells by hand --
 * additive, not a replacement. Direct cell edits still work exactly as
 * before and still fire handleEdit's automation.
 *
 * Proposal_Generator rows aren't created here (they're auto-pulled from
 * Job_Scoring's APPLY rows via a FILTER formula) -- this sidebar is a row
 * picker + entry form for the sheet's manual fields only: Bid_1st-4th,
 * Boost_Connects, Additional_Questions, Proposal_Status, Notes. Since
 * script-driven setValue() writes never fire handleEdit, each save function
 * below calls the same automation functions handleEdit uses per-column
 * (computeBidRecommendation_, applyBoostConnects_, generateAdditionalAnswers_,
 * handleProposalStatusChange_ -- all in 14_Edit_Trigger.gs) so nothing about
 * the pipeline's behavior changes depending on which entry path was used.
 *
 * Split into 4 steps (Bids / Boost Connects / Additional Questions / Status &
 * Notes), each saving only its own fields -- a single combined form
 * previously resubmitted every field (including untouched ones re-prefilled
 * from the sheet) on every save, so e.g. adding Notes after bids were already
 * entered silently re-ran the bid recommendation AND proposal-regen AI calls
 * for the exact same values. Splitting by step keeps unrelated fields
 * structurally out of each save's payload -- Step 4 (Status & Notes) simply
 * has no way to carry Bid_4th along with it anymore -- so each AI-triggering
 * save below fires on "this step's field was provided" the same way
 * handleEdit fires on a direct cell edit, with no separate equality check
 * needed against the prior value.
 * ============================================================
 */
function openProposalGeneratorSidebar_() {
  var html = HtmlService.createHtmlOutputFromFile('ProposalGeneratorSidebar')
    .setTitle('Log Proposal Bid')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function LOG_PROPOSAL_BID() {
  openProposalGeneratorSidebar_();
}

// Dropdown source -- every Proposal_Generator row with a Discovery_ID and
// Job_Title, labeled with client name and current status for context.
function proposal_getRows() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var map       = getHeaderMap_(sheet);
  var idCol     = getCol_(map, ['Discovery_ID']);
  var titleCol  = getCol_(map, ['Job_Title']);
  var clientCol = getCol_(map, ['Client_Name', 'Client Name']);
  var statusCol = getCol_(map, ['Proposal_Status']);
  if (!idCol || !titleCol) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var discoveryId = data[i][idCol - 1];
    var jobTitle     = data[i][titleCol - 1];
    if (!discoveryId || !jobTitle) continue;

    var clientName = clientCol ? data[i][clientCol - 1] : '';
    var status     = statusCol ? data[i][statusCol - 1] : '';
    var label = jobTitle + (clientName ? ' -- ' + clientName : '') + (status ? ' [' + status + ']' : '');

    rows.push({ discoveryId: discoveryId, label: label });
  }

  return rows;
}

// Prefill data for a picked row, including the AI's current Bid_Recommendation
// and Additional_Answers for context (read-only display, not editable fields
// in this sidebar). Used to prefill all 4 steps up front, so stepping through
// the wizard never re-fetches from the sheet.
function proposal_getRowDetails(discoveryId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job. Refresh and try again.' };

  return {
    ok:                  true,
    bid1:                getCellValue_(sheet, row, map, ['Bid_1st']),
    bid2:                getCellValue_(sheet, row, map, ['Bid_2nd']),
    bid3:                getCellValue_(sheet, row, map, ['Bid_3rd']),
    bid4:                getCellValue_(sheet, row, map, ['Bid_4th']),
    boostConnects:       getCellValue_(sheet, row, map, ['Boost_Connects']),
    additionalQuestions: getCellValue_(sheet, row, map, ['Additional_Questions']),
    additionalAnswers:   getCellValue_(sheet, row, map, ['Additional_Answers']),
    proposalStatus:      getCellValue_(sheet, row, map, ['Proposal_Status']),
    notes:               getCellValue_(sheet, row, map, ['Notes']),
    bidRecommendation:   getCellValue_(sheet, row, map, ['Bid_Recommendation'])
  };
}

// Step 1: Bids. Writes only the fields actually provided (blank = leave
// existing value alone), and runs the bid-recommendation AI call whenever
// Bid_4th is provided -- same trigger condition handleEdit uses for a direct
// cell edit to Bid_4th (fires on the write, no equality check against the
// prior value). Safe from the original resend-waste bug because Step 1's
// payload can only ever contain bid fields -- unrelated edits made later
// (Boost Connects, Additional Questions, Status & Notes) are separate saves
// on separate steps and never touch Bid_4th at all.
function proposal_saveBids(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, data.discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job in Proposal_Generator. Refresh and try again.' };

  if (data.bid1 !== '') setCellValue_(sheet, row, map, ['Bid_1st'], Number(data.bid1));
  if (data.bid2 !== '') setCellValue_(sheet, row, map, ['Bid_2nd'], Number(data.bid2));
  if (data.bid3 !== '') setCellValue_(sheet, row, map, ['Bid_3rd'], Number(data.bid3));
  if (data.bid4 !== '') setCellValue_(sheet, row, map, ['Bid_4th'], Number(data.bid4));

  if (data.bid4 !== '') computeBidRecommendation_(ss, sheet, row, map);

  return { ok: true, bidRecommendation: getCellValue_(sheet, row, map, ['Bid_Recommendation']) };
}

// Step 2: Boost Connects (optional). Recalculates Total_Connects_Spent and
// runs the AI proposal-regen call whenever Boost_Connects is provided -- same
// trigger condition handleEdit uses for a direct cell edit (fires on the
// write, no equality check). Safe from the original resend-waste bug because
// Step 2's payload can only ever contain Boost_Connects -- edits made on
// other steps are separate saves and never touch this field.
function proposal_saveBoost(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, data.discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job in Proposal_Generator. Refresh and try again.' };

  if (data.boostConnects !== '') setCellValue_(sheet, row, map, ['Boost_Connects'], Number(data.boostConnects));

  if (data.boostConnects !== '') applyBoostConnects_(ss, sheet, row, map);

  return { ok: true };
}

// Step 3: Additional Questions (optional). Redrafts Additional_Answers
// whenever Additional_Questions is provided -- same trigger condition
// handleEdit uses. Safe from the original resend-waste bug for the same
// reason as Steps 1 and 2: this step's payload never carries any other
// step's fields.
function proposal_saveQuestions(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, data.discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job in Proposal_Generator. Refresh and try again.' };

  if (data.additionalQuestions !== '') setCellValue_(sheet, row, map, ['Additional_Questions'], data.additionalQuestions);

  if (data.additionalQuestions !== '') generateAdditionalAnswers_(ss, sheet, row, map);

  return { ok: true, additionalAnswers: getCellValue_(sheet, row, map, ['Additional_Answers']) };
}

// Step 4: Status & Notes. No AI calls at all -- handleProposalStatusChange_
// only stamps dates / syncs to Proposal_Tracker, and is already internally
// guarded (via Proposal_Sent_Date/Proposal_Skip_Date being blank) against
// running its side effects twice, so it's safe to call on every save here.
function proposal_saveStatusNotes(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, data.discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job in Proposal_Generator. Refresh and try again.' };

  if (data.notes !== '')   setCellValue_(sheet, row, map, ['Notes'], data.notes);
  if (data.proposalStatus) setCellValue_(sheet, row, map, ['Proposal_Status'], data.proposalStatus);
  if (data.proposalStatus) handleProposalStatusChange_(ss, sheet, row, map);

  return { ok: true };
}

function findProposalGeneratorRowByDiscoveryId_(sheet, map, discoveryId) {
  var idCol = getCol_(map, ['Discovery_ID']);
  if (!idCol || sheet.getLastRow() < 2) return null;

  var values = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(discoveryId).trim()) return i + 2;
  }
  return null;
}
