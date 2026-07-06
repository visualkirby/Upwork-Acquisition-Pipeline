/**
 * ============================================================
 * 5. AI CONTEXT
 * Reads Settings sheet dynamically to build freelancer context
 * for all AI prompt functions. Call getSettings_() at the top
 * of any AI function to get the latest values without restarting.
 * ============================================================
 */
function getSettings_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName('Settings');
  var settings = {};

  if (!sheet || sheet.getLastRow() < 2) return settings;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = String(data[i][1]).trim();
    if (key) settings[key] = val;
  }

  settings['Portfolio_All'] = buildPortfolioAllFromProjects_(ss);

  return settings;
}

// Portfolio_All feeds the AI proposal prompts (see Lib_ProposalGenerator.gs's
// getPortfolioContext) -- pulling Description in here, not just Project_Name,
// is what lets the AI reference what a project actually did instead of just
// its title.
function buildPortfolioAllFromProjects_(ss) {
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() < 2) return '';

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf('Project_Name');
  var descCol = headers.indexOf('Description');
  if (nameCol < 0) return '';

  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var parts = [];

  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][nameCol]).trim();
    if (!name) continue;
    var desc = descCol >= 0 ? String(data[i][descCol]).trim() : '';
    parts.push(desc ? name + ': ' + desc : name);
  }

  return parts.join('; ');
}


// Journey-stage narrative logic moved to the Apps Script Library
// (FFLib.buildJourneyStage(settings)). getSettings_() above still
// runs client-side (sheet I/O) and its result is passed in.
