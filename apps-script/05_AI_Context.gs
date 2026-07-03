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

  var portfolioParts = [];
  for (var p = 1; p <= 6; p++) {
    var pVal = settings['Portfolio_' + p];
    if (pVal && pVal !== '') portfolioParts.push(pVal);
  }
  settings['Portfolio_All'] = portfolioParts.join('; ');

  return settings;
}


// Journey-stage narrative logic moved to the Apps Script Library
// (FFLib.buildJourneyStage(settings)). getSettings_() above still
// runs client-side (sheet I/O) and its result is passed in.
