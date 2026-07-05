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
 * Boost_Connects, Proposal_Status, Notes. Since script-driven setValue()
 * writes never fire handleEdit, proposal_saveEntry calls the same
 * automation functions handleEdit uses per-column (computeBidRecommendation_,
 * applyBoostConnects_, handleProposalStatusChange_ -- all in
 * 14_Edit_Trigger.gs) so nothing about the pipeline's behavior changes
 * depending on which entry path was used.
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
// for context (read-only display, not an editable field in this sidebar).
function proposal_getRowDetails(discoveryId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job. Refresh and try again.' };

  return {
    ok:                true,
    bid1:              getCellValue_(sheet, row, map, ['Bid_1st']),
    bid2:              getCellValue_(sheet, row, map, ['Bid_2nd']),
    bid3:              getCellValue_(sheet, row, map, ['Bid_3rd']),
    bid4:              getCellValue_(sheet, row, map, ['Bid_4th']),
    boostConnects:     getCellValue_(sheet, row, map, ['Boost_Connects']),
    proposalStatus:    getCellValue_(sheet, row, map, ['Proposal_Status']),
    notes:             getCellValue_(sheet, row, map, ['Notes']),
    bidRecommendation: getCellValue_(sheet, row, map, ['Bid_Recommendation'])
  };
}

// Writes only the fields actually provided (blank = leave existing value
// alone, so a partial re-visit -- e.g. filling Bid_4th after already having
// entered Bid_1st-3rd earlier -- never clobbers what's already there), then
// runs the same per-field automation handleEdit would run on a direct cell edit.
function proposal_saveEntry(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Proposal_Generator');
  if (!sheet) return { ok: false, message: 'Proposal_Generator sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProposalGeneratorRowByDiscoveryId_(sheet, map, data.discoveryId);
  if (!row) return { ok: false, message: 'Could not find that job in Proposal_Generator. Refresh and try again.' };

  if (data.bid1 !== '')          setCellValue_(sheet, row, map, ['Bid_1st'], Number(data.bid1));
  if (data.bid2 !== '')          setCellValue_(sheet, row, map, ['Bid_2nd'], Number(data.bid2));
  if (data.bid3 !== '')          setCellValue_(sheet, row, map, ['Bid_3rd'], Number(data.bid3));
  if (data.bid4 !== '')          setCellValue_(sheet, row, map, ['Bid_4th'], Number(data.bid4));
  if (data.boostConnects !== '') setCellValue_(sheet, row, map, ['Boost_Connects'], Number(data.boostConnects));
  if (data.notes !== '')         setCellValue_(sheet, row, map, ['Notes'], data.notes);
  if (data.proposalStatus)       setCellValue_(sheet, row, map, ['Proposal_Status'], data.proposalStatus);

  if (data.bid4 !== '')          computeBidRecommendation_(ss, sheet, row, map);
  if (data.boostConnects !== '') applyBoostConnects_(ss, sheet, row, map);
  if (data.proposalStatus)       handleProposalStatusChange_(ss, sheet, row, map);

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
