/**
 * ============================================================
 * 2. API KEY SETUP & DIAGNOSTICS
 *
 * Run SETUP_API_KEY once from the Apps Script editor to store
 * your OpenAI API key securely in PropertiesService.
 *
 * Steps:
 *   1. Paste your key between the quotes in SETUP_API_KEY
 *   2. Click Run -> SETUP_API_KEY
 *   3. When the alert says "stored", delete your key from
 *      the function and save -- it is now stored securely
 * ============================================================
 */
function SETUP_API_KEY() {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt("Setup API Key", "Paste your OpenAI API key below:", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var key = response.getResponseText().trim();

  if (!key) {
    ui.alert("No key entered. Setup cancelled.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
  ui.alert("API key stored. You do not need to run this again unless you rotate your key.");
}


function CHECK_API_KEY() {
  var ui     = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    ui.alert("No API key found. Run System Tools > Setup API Key.");
    return;
  }

  ui.alert("API key is set.\nLength: " + apiKey.length + " chars\nPrefix: " + apiKey.substring(0, 8) + "...");
}


function TEST_JOB_TYPE_API() {
  var ui     = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    ui.alert("No API key found. Run Setup API Key first.");
    return;
  }

  ui.alert("API key found. Length: " + apiKey.length + "\nFirst 8 chars: " + apiKey.substring(0, 8));

  var testDesc  = "We are looking for a skilled developer to build a new Power BI dashboard tracking sales KPIs.";
  var testTitle = "Power BI Dashboard Developer";
  var result    = getJobType_(testDesc, testTitle);

  ui.alert("Test classification result: \"" + result + "\"\n\n" +
    (result ? "API call working correctly." : "API call returned empty -- check key or quota."));
}


function DIAGNOSE_PROPOSAL_GENERATOR() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet) { ui.alert("Proposal_Generator not found."); return; }

  var map        = getHeaderMap_(sheet);
  var descCol    = getCol_(map, ["Description"]);
  var jobTypeCol = getCol_(map, ["Job_Type"]);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var lastRow    = sheet.getLastRow();

  var actualRows  = 0;
  var lastDescRow = 1;
  if (descCol && lastRow > 1) {
    var descVals = sheet.getRange(2, descCol, lastRow - 1, 1).getValues();
    for (var t = 0; t < descVals.length; t++) {
      if (String(descVals[t][0]).trim() !== "") {
        actualRows++;
        lastDescRow = t + 2;
      }
    }
  }

  var msg = "Proposal_Generator Diagnostic\n";
  msg += "getLastRow(): " + lastRow + "\n";
  msg += "Actual rows with Description: " + actualRows + "\n";
  msg += "Last description row: " + lastDescRow + "\n";
  msg += "Description col: " + descCol + "\n";
  msg += "Job_Type col: " + jobTypeCol + "\n";
  msg += "Job_Title col: " + titleCol + "\n\n";

  if (lastRow < 2) { ui.alert(msg + "No data rows found."); return; }

  var data = sheet.getRange(2, 1, Math.min(lastRow - 1, 5), sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var title = titleCol   ? String(data[i][titleCol   - 1]).substring(0, 30) : "N/A";
    var desc  = descCol    ? String(data[i][descCol    - 1]).trim()            : "N/A";
    var jtype = jobTypeCol ? String(data[i][jobTypeCol - 1]).trim()            : "N/A";
    msg += "Row " + (i + 2) + ": title=" + title + "\n";
    msg += "  Job_Type=\"" + jtype + "\" desc_length=" + desc.length + "\n";
    msg += "  desc_preview=\"" + desc.substring(0, 60) + "\"\n\n";
  }

  ui.alert(msg);
}
