/**
 * ============================================================
 * 20. CONTRACT SETUP SIDEBAR
 * Opens automatically when Proposal_Tracker's Hired column flips
 * to "Y" (see 14_Edit_Trigger.gs), or manually via the "Log New
 * Contract" menu item -- the latter also doubles as "add more
 * milestones to an existing contract," since the dropdown lists
 * every Hired = Y job, not just ones without a contract yet.
 * ============================================================
 */
var PENDING_CONTRACT_PROP_ = 'PENDING_CONTRACT_DISCOVERY_ID';

function openContractSetupSidebar_(discoveryId) {
  if (discoveryId) {
    PropertiesService.getScriptProperties().setProperty(PENDING_CONTRACT_PROP_, String(discoveryId));
  }
  var html = HtmlService.createHtmlOutputFromFile('ContractSetupSidebar')
    .setTitle('Log New Contract')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function LOG_NEW_CONTRACT() {
  openContractSetupSidebar_(null);
}

function contract_getContext() {
  var prop = PropertiesService.getScriptProperties();
  var pendingId = prop.getProperty(PENDING_CONTRACT_PROP_) || '';
  prop.deleteProperty(PENDING_CONTRACT_PROP_);

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tracker = ss.getSheetByName('Proposal_Tracker');
  var jobs    = [];

  if (tracker && tracker.getLastRow() > 1) {
    var map       = getHeaderMap_(tracker);
    var idCol     = getCol_(map, ['Discovery_ID']);
    var titleCol  = getCol_(map, ['Job_Title']);
    var clientCol = getCol_(map, ['Client_Name', 'Client Name']);
    var hiredCol  = getCol_(map, ['Hired']);

    if (idCol && titleCol && clientCol && hiredCol) {
      var data = tracker.getRange(2, 1, tracker.getLastRow() - 1, tracker.getLastColumn()).getValues();
      data.forEach(function (row) {
        if (String(row[hiredCol - 1]).trim() === 'Y') {
          jobs.push({
            discoveryId: row[idCol - 1],
            jobTitle:    row[titleCol - 1],
            clientName:  row[clientCol - 1]
          });
        }
      });
    }
  }

  return { pendingDiscoveryId: pendingId, contracts: jobs };
}

function contract_saveSetup(data) {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var tracker    = ss.getSheetByName('Proposal_Tracker');
  var contracts  = ss.getSheetByName('Contract_Tracker');
  var milestones = ss.getSheetByName('Milestone_Tracker');

  if (!tracker || !contracts || !milestones) return { ok: false, message: 'Required sheets not found.' };

  var ptMap     = getHeaderMap_(tracker);
  var idCol     = getCol_(ptMap, ['Discovery_ID']);
  var titleCol  = getCol_(ptMap, ['Job_Title']);
  var clientCol = getCol_(ptMap, ['Client_Name', 'Client Name']);

  var jobTitle   = '';
  var clientName = '';
  if (idCol && titleCol && clientCol && tracker.getLastRow() > 1) {
    var ptData = tracker.getRange(2, 1, tracker.getLastRow() - 1, tracker.getLastColumn()).getValues();
    for (var i = 0; i < ptData.length; i++) {
      if (String(ptData[i][idCol - 1]) === String(data.discoveryId)) {
        jobTitle   = ptData[i][titleCol - 1];
        clientName = ptData[i][clientCol - 1];
        break;
      }
    }
  }

  // Find-or-create the Contract_Tracker row -- handles both the auto-trigger
  // path (row doesn't exist yet) and re-opening the sidebar later to update
  // an existing contract's type/rate.
  var ctMap      = getHeaderMap_(contracts);
  var ctIdCol    = getCol_(ctMap, ['Discovery_ID']);
  var contractRow = null;
  if (ctIdCol && contracts.getLastRow() > 1) {
    var ctData = contracts.getRange(2, 1, contracts.getLastRow() - 1, contracts.getLastColumn()).getValues();
    for (var c = 0; c < ctData.length; c++) {
      if (String(ctData[c][ctIdCol - 1]) === String(data.discoveryId)) {
        contractRow = c + 2;
        break;
      }
    }
  }
  if (!contractRow) {
    contractRow = findFirstEmptyRowByColumn_(contracts, ctIdCol || 1);
    setCellValue_(contracts, contractRow, ctMap, ['Discovery_ID'], data.discoveryId);
    setCellValue_(contracts, contractRow, ctMap, ['Start_Date'], new Date());
    setCellValue_(contracts, contractRow, ctMap, ['Status'], 'Active');
  }

  setCellValue_(contracts, contractRow, ctMap, ['Job_Title'], jobTitle);
  setCellValue_(contracts, contractRow, ctMap, ['Client_Name', 'Client Name'], clientName);
  setCellValue_(contracts, contractRow, ctMap, ['Contract_Type'], data.contractType);
  setCellValue_(contracts, contractRow, ctMap, ['Contract_Value'], data.contractType === 'Fixed' ? data.contractValue : '');
  setCellValue_(contracts, contractRow, ctMap, ['Hourly_Rate'], data.contractType === 'Hourly' ? data.hourlyRate : '');

  if (data.contractType === 'Fixed' && data.milestones && data.milestones.length > 0) {
    var msMap    = getHeaderMap_(milestones);
    var msIdCol  = getCol_(msMap, ['Discovery_ID']);
    var msNumCol = getCol_(msMap, ['Milestone_Number']);

    var existingCount = 0;
    if (msIdCol && milestones.getLastRow() > 1) {
      var msData = milestones.getRange(2, 1, milestones.getLastRow() - 1, milestones.getLastColumn()).getValues();
      msData.forEach(function (row) {
        if (String(row[msIdCol - 1]) === String(data.discoveryId)) existingCount++;
      });
    }

    var msStartRow = findFirstEmptyRowByColumn_(milestones, msIdCol || 1);
    data.milestones.forEach(function (m, i) {
      var msRow = msStartRow + i;
      setCellValue_(milestones, msRow, msMap, ['Discovery_ID'], data.discoveryId);
      setCellValue_(milestones, msRow, msMap, ['Job_Title'], jobTitle);
      setCellValue_(milestones, msRow, msMap, ['Milestone_Number'], existingCount + i + 1);
      setCellValue_(milestones, msRow, msMap, ['Description'], m.description || '');
      setCellValue_(milestones, msRow, msMap, ['Amount'], m.amount || 0);
      setCellValue_(milestones, msRow, msMap, ['Status'], 'Pending');
    });
  }

  // Hourly contracts get one starter row in Hourly_Log, pre-linked with
  // Discovery_ID/Job_Title so the freelancer only has to type Hours_Logged --
  // same "land pre-filled, ready to work" treatment Fixed contracts get via
  // their milestone rows above. Guarded on "no row for this Discovery_ID
  // yet" so re-opening the sidebar to edit the hourly rate later doesn't
  // spawn a fresh blank row every time.
  if (data.contractType === 'Hourly') {
    var hourly = ss.getSheetByName('Hourly_Log');
    if (hourly) {
      var hlMap   = getHeaderMap_(hourly);
      var hlIdCol = getCol_(hlMap, ['Discovery_ID']);

      var hlAlreadyExists = false;
      if (hlIdCol && hourly.getLastRow() > 1) {
        var hlData = hourly.getRange(2, hlIdCol, hourly.getLastRow() - 1, 1).getValues();
        hlAlreadyExists = hlData.some(function (row) { return String(row[0]) === String(data.discoveryId); });
      }

      if (!hlAlreadyExists) {
        var hlRow = findFirstEmptyRowByColumn_(hourly, hlIdCol || 1);
        setCellValue_(hourly, hlRow, hlMap, ['Discovery_ID'], data.discoveryId);
        setCellValue_(hourly, hlRow, hlMap, ['Job_Title'], jobTitle);
        setCellValue_(hourly, hlRow, hlMap, ['Log_Date'], new Date());
        setCellValue_(hourly, hlRow, hlMap, ['Status'], 'Active');
      }
    }
  }

  showTourStep_(
    'FF_TOUR_STEP10_CONTRACT_PROGRESS_SEEN',
    'Track This Contract',
    'Your contract is set up. Come back to System Tools > Log Contract Progress whenever you update a milestone\'s status or log hours worked.'
  );

  return { ok: true };
}
