/**
 * ============================================================
 * 4. RESET
 * Two distinct reset depths:
 *   RESET_TO_AFTER_SETUP  -- clears all logged activity but keeps
 *     Settings/Proposal_Templates/Keyword_Strategy exactly as the
 *     Setup Wizard left them (re-prompts for Connect Balance the
 *     same way the wizard's own step does).
 *   RESET_TO_BEFORE_SETUP -- factory reset. Deletes every sheet the
 *     wizard creates and clears every script property (API key
 *     included), so onOpen reopens the Setup Wizard as if this were
 *     a spreadsheet that never had it run.
 * ============================================================
 */
function RESET_TO_AFTER_SETUP() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Reset to Right After Setup',
    'This clears every job, session, proposal, contract, and keyword-search-history ' +
    'you\'ve logged, but keeps your Settings, Proposal_Templates, Projects, and generated ' +
    'Keyword_Strategy exactly as the Setup Wizard left them.\n\n' +
    'You\'ll be asked to re-enter your Connect Balance. This cannot be undone. Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  clearByHeaders_(ss, 'Job_Discovery', [
    'Job_Title', 'Description', 'Additional_Questions', 'Client_Name', 'Client Name',
    'Keyword_Search', 'Experience_Level', 'Hours_Since_Posted', 'Minutes_Since_Posted',
    'Days_Since_Posted', 'Proposal_Count', 'Payment_Verified', 'Client_Hires',
    'Budget_Type', 'Budget', 'Hourly_Rate', 'Job_Link', 'Connects_Required',
    'AI_Fit_Notes', 'Date_Found', 'Session_ID'
  ]);

  // Job_Scoring's raw-data + score columns are all formulas driven off
  // Job_Discovery (applyJobScoringPullFormula_/applyJobScoringFormulas_ in
  // 00_Setup_Wizard.gs) -- clearing Job_Discovery above already empties
  // them, nothing to clear here directly.

  clearByHeaders_(ss, 'Proposal_Generator', [
    'Job_Type', 'Recommended_Template', 'Hook_Version', 'CTA_Version',
    'Bid_1st', 'Bid_2nd', 'Bid_3rd', 'Bid_4th', 'Boost_Connects', 'Total_Connects_Spent',
    'Bid_Recommendation', 'Additional_Questions', 'AI_Generated_Proposal', 'Additional_Answers',
    'Proposal_Status', 'Proposal_Sent_Date', 'Proposal_Skip_Date', 'Notes'
  ]);
  // Discovery_ID/Date/Job_Title/Client_Name/Description/Job_Link/Keyword_Search/
  // Connects_Required/Proposal_Count/Budget are a FILTER pull from Job_Scoring,
  // and Tool_Detected/Portfolio_Project are formulas off that pull -- both
  // self-empty once Job_Discovery above has nothing left for Job_Scoring to match.

  // These sheets are entirely script-written rows (no formula/identity
  // column worth preserving), so a full clear is the right shape.
  clearAllRows_(ss, 'Proposal_Tracker');
  clearAllRows_(ss, 'Client_Chat_Log');
  clearAllRows_(ss, 'Contract_Tracker');
  clearAllRows_(ss, 'Milestone_Tracker');
  clearAllRows_(ss, 'Hourly_Log');
  clearAllRows_(ss, 'Session_Log');
  clearAllRows_(ss, 'Monthly_Performance');

  // Keyword_Search_List/Keyword_Strategy keep the identity columns the
  // wizard's keyword strategy generated (Tool/Business_Area/Intent/
  // Search_Query, Keyword/Recommended_Action/Target_Count/Notes) -- only the
  // session-tracking columns reset, matching a fresh wizard run exactly.
  clearByHeaders_(ss, 'Keyword_Search_List', ['Last_Searched', 'Session_Yield']);
  setColumnValue_(ss, 'Keyword_Strategy', 'Actual_Count', 0);

  var balanceResponse = ui.prompt(
    'Connect Balance',
    'Enter your current Upwork Connect Balance:',
    ui.ButtonSet.OK_CANCEL
  );
  if (balanceResponse.getSelectedButton() === ui.Button.OK) {
    initConnectsHelper_(ss, Number(balanceResponse.getResponseText()) || 0);
  }

  resetScriptPropertiesKeepSetup_();

  ui.alert('Reset complete. The system now matches the state right after Setup Wizard finished.');
}

function RESET_TO_BEFORE_SETUP() {
  var ui = SpreadsheetApp.getUi();

  var warning = ui.alert(
    'Factory Reset -- Before Setup Wizard',
    'This deletes EVERY sheet FreelanceFlow created (Settings, Connects_Helper, ' +
    'Proposal_Templates, Projects, Job_Discovery, Job_Scoring, Proposal_Generator, Proposal_Tracker, ' +
    'Client_Chat_Log, Contract_Tracker, Milestone_Tracker, Hourly_Log, Session_Log, ' +
    'Keyword_Search_List, Keyword_Strategy, Monthly_Performance) and clears your API key. ' +
    'The Setup Wizard reopens immediately after.\n\n' +
    'This cannot be undone. Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (warning !== ui.Button.OK) return;

  var confirmResponse = ui.prompt(
    'Type RESET to confirm',
    'This permanently deletes all FreelanceFlow data and settings. Type RESET (all caps) to confirm.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirmResponse.getSelectedButton() !== ui.Button.OK) return;
  if (confirmResponse.getResponseText().trim() !== 'RESET') {
    ui.alert('Confirmation text did not match. Reset cancelled.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = [
    'Settings', 'Connects_Helper', 'Proposal_Templates', 'Projects',
    'Job_Discovery', 'Job_Scoring', 'Proposal_Generator', 'Proposal_Tracker',
    'Client_Chat_Log', 'Contract_Tracker', 'Milestone_Tracker', 'Hourly_Log',
    'Session_Log', 'Keyword_Search_List', 'Keyword_Strategy', 'Monthly_Performance'
  ];

  sheetNames.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    // Sheets won't allow deleting the last remaining sheet in a spreadsheet --
    // guarantee a placeholder survives if every remaining sheet is one of ours.
    if (ss.getSheets().length <= 1) ss.insertSheet('Sheet1');
    ss.deleteSheet(sh);
  });

  // Removes the installable handleEdit trigger -- registerEditTrigger_() in
  // 00_Setup_Wizard.gs recreates it the moment the wizard finishes again.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'handleEdit') ScriptApp.deleteTrigger(t);
  });

  PropertiesService.getScriptProperties().deleteAllProperties();

  ui.alert('Factory reset complete. The Setup Wizard will open now.');
  openSetupWizard_();
}

// ---- Shared helpers --------------------------------------------------------

function clearByHeaders_(ss, sheetName, headerNames) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  var map = getHeaderMap_(sh);
  headerNames.forEach(function (name) {
    var col = getCol_(map, [name]);
    if (col) sh.getRange(2, col, lastRow - 1, 1).clearContent();
  });
}

function clearAllRows_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;
  sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
}

function setColumnValue_(ss, sheetName, headerName, value) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  var map = getHeaderMap_(sh);
  var col = getCol_(map, [headerName]);
  if (!col) return;

  var values = [];
  for (var i = 0; i < lastRow - 1; i++) values.push([value]);
  sh.getRange(2, col, lastRow - 1, 1).setValues(values);
}

// Deletes every script property except the two that define "setup already
// ran" -- API key and FF_SETUP_COMPLETE. Iterates rather than hardcoding
// every SESSION_*/FF_TOUR_*/FF_WALKTHROUGH_* key name, so it stays correct
// as new one-time flags get added elsewhere in the project.
function resetScriptPropertiesKeepSetup_() {
  var prop = PropertiesService.getScriptProperties();
  var keep = ['UPWORK_OPENAI_API_KEY', 'FF_SETUP_COMPLETE'];
  var all  = prop.getProperties();

  Object.keys(all).forEach(function (key) {
    if (keep.indexOf(key) === -1) prop.deleteProperty(key);
  });
}
