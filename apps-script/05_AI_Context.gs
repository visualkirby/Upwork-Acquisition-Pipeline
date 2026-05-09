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
  var sheet   = ss.getSheetByName("Settings");
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
    var pVal = settings["Portfolio_" + p];
    if (pVal && pVal !== "") portfolioParts.push(pVal);
  }
  settings["Portfolio_All"] = portfolioParts.join("; ");

  return settings;
}


function getJourneyStage_() {
  var s = getSettings_();

  var contracts = s["Contracts_Completed"] || "0";
  var reviews   = s["Reviews_Count"]       || "0";
  var score     = s["Job_Success_Score"]   || "0";
  var tools     = s["Primary_Tools"]       || "Power BI, Tableau, Looker Studio, Excel, SQL, Google Sheets";
  var stage     = s["Journey_Stage"]       || "";

  if (stage && stage.length > 20) return stage;

  var context =
    "The freelancer is Sawandi, a data analytics freelancer on Upwork. " +
    "He currently has " + contracts + " completed contract(s) and " + reviews + " review(s). " +
    "His Job Success Score is " + (score !== "0" ? score + "%." : "not yet established.") + " " +
    "His core tools are: " + tools + ". ";

  if (contracts === "0") {
    context +=
      "Because he has no reviews yet, winning depends heavily on proposal quality and job fit, " +
      "not just bid position. He should be conservative with boost spending until he has at least 1-2 reviews. " +
      "Lower-competition jobs where his proposal can stand out on merit are higher priority than aggressive outbidding.";
  } else if (parseInt(contracts) < 5) {
    context +=
      "He has some early contracts and is building his reputation. " +
      "Moderate boost spending is acceptable on strong-fit jobs. " +
      "Focus on maintaining high review scores and job completion rate.";
  } else {
    context +=
      "He has an established track record. " +
      "Strategic boosting is appropriate on high-value jobs. " +
      "Prioritize quality clients and long-term relationships over volume.";
  }

  return context;
}


// Legacy variable kept for backward compatibility -- now reads dynamically
var JOURNEY_STAGE_ = "";
