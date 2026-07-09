/**
 * ============================================================
 * 23. CONTRACT PROGRESS SIDEBAR
 * Ongoing tracking for a contract already set up via Log New Contract --
 * separate from that sidebar, which is one-time setup (contract type, rate,
 * initial milestones). This one is for the recurring work afterward:
 * updating a Fixed contract's milestone statuses, or logging a new day's
 * hours against an Hourly contract.
 *
 * Milestone/Hourly writes here are script-driven, so they never fire
 * handleEdit on their own -- the save functions below call the same
 * automation functions handleEdit uses per-column (applyMilestoneStatusEffects_,
 * applyHourlyLogAmount_, applyHourlyLogStatusEffects_ -- all in
 * 14_Edit_Trigger.gs) so nothing about the pipeline's behavior changes
 * depending on which entry path was used.
 * ============================================================
 */
function openContractProgressSidebar_() {
  var html = HtmlService.createHtmlOutputFromFile('ContractProgressSidebar')
    .setTitle('Log Contract Progress')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function LOG_CONTRACT_PROGRESS() {
  openContractProgressSidebar_();
}

// Dropdown source -- every Contract_Tracker row, labeled with client name.
function progress_getJobs() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contract_Tracker');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var map        = getHeaderMap_(sheet);
  var idCol      = getCol_(map, ['Discovery_ID']);
  var titleCol   = getCol_(map, ['Job_Title']);
  var clientCol  = getCol_(map, ['Client_Name', 'Client Name']);
  var typeCol    = getCol_(map, ['Contract_Type']);
  var statusCol  = getCol_(map, ['Status']);
  if (!idCol || !titleCol) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var jobs = [];

  for (var i = 0; i < data.length; i++) {
    var discoveryId = data[i][idCol - 1];
    var jobTitle     = data[i][titleCol - 1];
    if (!discoveryId || !jobTitle) continue;

    var clientName = clientCol ? data[i][clientCol - 1] : '';
    var status     = statusCol ? data[i][statusCol - 1] : '';
    var label = jobTitle + (clientName ? ' -- ' + clientName : '') + (status ? ' [' + status + ']' : '');

    jobs.push({ discoveryId: discoveryId, label: label });
  }

  return jobs;
}

// Prefill data for a picked contract -- its type, its milestones (Fixed) or
// its 10 most recent Hourly_Log entries (Hourly), newest first.
function progress_getJobDetails(discoveryId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var contracts = ss.getSheetByName('Contract_Tracker');
  if (!contracts) return { ok: false, message: 'Contract_Tracker sheet not found.' };

  var ctMap       = getHeaderMap_(contracts);
  var ctIdCol     = getCol_(ctMap, ['Discovery_ID']);
  var ctTypeCol   = getCol_(ctMap, ['Contract_Type']);
  var ctStatusCol = getCol_(ctMap, ['Status']);
  if (!ctIdCol) return { ok: false, message: 'Discovery_ID column not found in Contract_Tracker.' };

  var ctData = contracts.getRange(2, 1, Math.max(contracts.getLastRow() - 1, 1), contracts.getLastColumn()).getValues();
  var contractType   = '';
  var contractStatus = '';
  for (var i = 0; i < ctData.length; i++) {
    if (String(ctData[i][ctIdCol - 1]) === String(discoveryId)) {
      contractType   = ctTypeCol   ? ctData[i][ctTypeCol   - 1] : '';
      contractStatus = ctStatusCol ? ctData[i][ctStatusCol - 1] : '';
      break;
    }
  }
  if (!contractType) return { ok: false, message: 'Could not find that contract. Refresh and try again.' };

  var result = { ok: true, contractType: contractType, status: contractStatus, milestones: [], hourlyEntries: [] };

  if (contractType === 'Fixed') {
    var msSheet = ss.getSheetByName('Milestone_Tracker');
    if (msSheet && msSheet.getLastRow() > 1) {
      var msMap    = getHeaderMap_(msSheet);
      var msIdCol  = getCol_(msMap, ['Discovery_ID']);
      var msNumCol = getCol_(msMap, ['Milestone_Number']);
      var msDescCol   = getCol_(msMap, ['Description']);
      var msAmountCol = getCol_(msMap, ['Amount']);
      var msStatusCol = getCol_(msMap, ['Status']);

      var msData = msSheet.getRange(2, 1, msSheet.getLastRow() - 1, msSheet.getLastColumn()).getValues();
      msData.forEach(function (row) {
        if (String(row[msIdCol - 1]) !== String(discoveryId)) return;
        result.milestones.push({
          milestoneNumber: msNumCol ? row[msNumCol - 1] : '',
          description:      msDescCol ? row[msDescCol - 1] : '',
          amount:            msAmountCol ? row[msAmountCol - 1] : 0,
          status:            msStatusCol ? row[msStatusCol - 1] : 'Pending'
        });
      });
      result.milestones.sort(function (a, b) { return Number(a.milestoneNumber) - Number(b.milestoneNumber); });
    }
  } else if (contractType === 'Hourly') {
    var hlSheet = ss.getSheetByName('Hourly_Log');
    if (hlSheet && hlSheet.getLastRow() > 1) {
      var hlMap    = getHeaderMap_(hlSheet);
      var hlIdCol  = getCol_(hlMap, ['Discovery_ID']);
      var hlDateCol   = getCol_(hlMap, ['Log_Date']);
      var hlHoursCol  = getCol_(hlMap, ['Hours_Logged']);
      var hlAmountCol = getCol_(hlMap, ['Amount']);
      var hlNotesCol  = getCol_(hlMap, ['Notes']);

      var hlData = hlSheet.getRange(2, 1, hlSheet.getLastRow() - 1, hlSheet.getLastColumn()).getValues();
      var entries = [];
      hlData.forEach(function (row) {
        if (String(row[hlIdCol - 1]) !== String(discoveryId)) return;
        var rawDate = hlDateCol ? row[hlDateCol - 1] : '';
        entries.push({
          // google.script.run does not reliably marshal Date objects back to
          // the client (silently drops the whole response instead of
          // throwing) -- send an ISO string instead, per Apps Script's own
          // guidance on returning date values across the bridge.
          logDate:     rawDate instanceof Date ? rawDate.toISOString() : (rawDate || ''),
          hoursLogged: hlHoursCol  ? row[hlHoursCol  - 1] : 0,
          amount:      hlAmountCol ? row[hlAmountCol - 1] : 0,
          notes:       hlNotesCol  ? row[hlNotesCol  - 1] : ''
        });
      });
      entries.sort(function (a, b) { return new Date(b.logDate) - new Date(a.logDate); });
      result.hourlyEntries = entries.slice(0, 10);
    }
  }

  return result;
}

// Updates an Hourly contract's Status by writing it onto that contract's
// most recent Hourly_Log row (matched by Log_Date), the same column a
// direct sheet edit would use -- Hourly_Log's Status column is the source
// of truth (see applyHourlyLogStatusEffects_ in 14_Edit_Trigger.gs), which
// then propagates to Contract_Tracker, rather than this sidebar writing to
// Contract_Tracker directly. Keeps both entry paths going through the same
// single mechanism.
function progress_saveContractStatus(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Hourly_Log');
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, message: 'Hourly_Log sheet not found.' };

  var map       = getHeaderMap_(sheet);
  var idCol     = getCol_(map, ['Discovery_ID']);
  var dateCol   = getCol_(map, ['Log_Date']);
  var statusCol = getCol_(map, ['Status']);
  if (!idCol || !statusCol) return { ok: false, message: 'Status column not found in Hourly_Log.' };

  var hlData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var targetRow  = -1;
  var latestDate = null;
  for (var i = 0; i < hlData.length; i++) {
    if (String(hlData[i][idCol - 1]) !== String(data.discoveryId)) continue;
    var rowDate = dateCol ? new Date(hlData[i][dateCol - 1]) : null;
    if (targetRow === -1 || (rowDate && (!latestDate || rowDate > latestDate))) {
      targetRow  = i + 2;
      latestDate = rowDate;
    }
  }
  if (targetRow === -1) return { ok: false, message: 'No Hourly_Log entries found for this contract yet.' };

  var oldStatus = sheet.getRange(targetRow, statusCol).getValue();
  if (String(oldStatus) === String(data.status)) return { ok: true };

  sheet.getRange(targetRow, statusCol).setValue(data.status);
  applyHourlyLogStatusEffects_(ss, sheet, targetRow, map, oldStatus, data.status);

  return { ok: true };
}

// Writes any milestone whose submitted status differs from the sheet's
// current one, then runs the same side effects a direct cell edit would.
function progress_saveMilestones(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Milestone_Tracker');
  if (!sheet || sheet.getLastRow() < 2) return { ok: false, message: 'Milestone_Tracker sheet not found.' };

  var map      = getHeaderMap_(sheet);
  var idCol    = getCol_(map, ['Discovery_ID']);
  var numCol   = getCol_(map, ['Milestone_Number']);
  var statusCol = getCol_(map, ['Status']);
  if (!idCol || !numCol || !statusCol) return { ok: false, message: 'Required columns not found in Milestone_Tracker.' };

  var lastRow = sheet.getLastRow();
  var idValues  = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  var numValues = sheet.getRange(2, numCol, lastRow - 1, 1).getValues();

  (data.milestones || []).forEach(function (m) {
    for (var i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0]) !== String(data.discoveryId)) continue;
      if (String(numValues[i][0]) !== String(m.milestoneNumber)) continue;

      var row    = i + 2;
      var oldVal = sheet.getRange(row, statusCol).getValue();
      if (String(oldVal) === String(m.status)) return;

      sheet.getRange(row, statusCol).setValue(m.status);
      applyMilestoneStatusEffects_(ss, sheet, row, map, oldVal, m.status);
      return;
    }
  });

  return { ok: true };
}

// Appends one new Hourly_Log row (never overwrites) so logging hours across
// multiple days just means calling this once per day.
function progress_addHourlyEntry(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Hourly_Log');
  if (!sheet) return { ok: false, message: 'Hourly_Log sheet not found.' };

  var contracts  = ss.getSheetByName('Contract_Tracker');
  var ctMap      = contracts ? getHeaderMap_(contracts) : {};
  var ctIdCol    = getCol_(ctMap, ['Discovery_ID']);
  var ctTitleCol = getCol_(ctMap, ['Job_Title']);
  var ctStatusCol = getCol_(ctMap, ['Status']);

  var jobTitle = '';
  var currentStatus = 'Active';
  if (contracts && ctIdCol && ctTitleCol && contracts.getLastRow() > 1) {
    var ctData = contracts.getRange(2, 1, contracts.getLastRow() - 1, contracts.getLastColumn()).getValues();
    for (var i = 0; i < ctData.length; i++) {
      if (String(ctData[i][ctIdCol - 1]) === String(data.discoveryId)) {
        jobTitle = ctData[i][ctTitleCol - 1];
        if (ctStatusCol && ctData[i][ctStatusCol - 1]) currentStatus = ctData[i][ctStatusCol - 1];
        break;
      }
    }
  }

  var map   = getHeaderMap_(sheet);
  var idCol = getCol_(map, ['Discovery_ID']);
  var row   = findFirstEmptyRowByColumn_(sheet, idCol || 1);

  setCellValue_(sheet, row, map, ['Discovery_ID'], data.discoveryId);
  setCellValue_(sheet, row, map, ['Job_Title'], jobTitle);
  setCellValue_(sheet, row, map, ['Log_Date'], data.logDate ? new Date(data.logDate) : new Date());
  setCellValue_(sheet, row, map, ['Hours_Logged'], Number(data.hoursLogged) || 0);
  setCellValue_(sheet, row, map, ['Status'], currentStatus);
  if (data.notes) setCellValue_(sheet, row, map, ['Notes'], data.notes);

  applyHourlyLogAmount_(ss, sheet, row, map);

  return progress_getJobDetails(data.discoveryId);
}
