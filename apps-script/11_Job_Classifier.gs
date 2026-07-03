/**
 * ============================================================
 * 11. JOB CLASSIFIER & BATCH PROPOSALS
 *
 * Classification logic (FFLib.getJobType) and proposal generation
 * (FFLib.generateAIProposal) live in the Apps Script Library.
 * RUN_JOB_CLASSIFICATION: fills Job_Type column in Proposal_Generator
 * RUN_AI_PROPOSALS: batch-generates AI proposals for all unfilled rows
 * ============================================================
 */
function RUN_JOB_CLASSIFICATION() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No rows found in Proposal_Generator.");
    return;
  }

  var map        = getHeaderMap_(sheet);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var descCol    = getCol_(map, ["Description"]);
  var jobTypeCol = getCol_(map, ["Job_Type"]);

  if (!descCol || !jobTypeCol || !titleCol) {
    ui.alert("Required columns not found. Confirm Job_Title, Description, and Job_Type columns exist.");
    return;
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty("UPWORK_OPENAI_API_KEY");

  var validTypes = ["Dashboard Build", "Dashboard Fix", "Data to Dashboard", "Reporting"];
  var lastRow    = sheet.getLastRow();
  var jtValues   = sheet.getRange(2, jobTypeCol, lastRow - 1, 1).getValues();
  var filled     = 0;
  var skipped    = 0;

  for (var i = 0; i < jtValues.length; i++) {
    var current = String(jtValues[i][0]).trim();
    var done    = false;
    for (var v = 0; v < validTypes.length; v++) {
      if (current === validTypes[v]) { done = true; break; }
    }
    if (done) { skipped++; continue; }

    var r     = i + 2;
    var desc  = String(sheet.getRange(r, descCol).getValue()).trim();
    var title = String(sheet.getRange(r, titleCol).getValue()).trim();

    if (!desc && !title) { continue; }

    var result = FFLib.getJobType(desc, title, apiKey);
    sheet.getRange(r, jobTypeCol).setValue(result);
    filled++;

    if (filled % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert("Done.\n\n✓ " + filled + " rows classified.\n-> " + skipped + " already had a Job_Type.");
}


function RUN_AI_PROPOSALS() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No rows found in Proposal_Generator.");
    return;
  }

  var map        = getHeaderMap_(sheet);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var descCol    = getCol_(map, ["Description"]);
  var toolCol    = getCol_(map, ["Tool_Detected"]);
  var jobTypeCol = getCol_(map, ["Job_Type"]);
  var tmplCol    = getCol_(map, ["Recommended_Template"]);
  var hookCol    = getCol_(map, ["Hook_Version"]);
  var ctaCol     = getCol_(map, ["CTA_Version"]);
  var aiPropCol  = getCol_(map, ["AI_Generated_Proposal"]);

  if (!descCol || !aiPropCol) {
    ui.alert("Required columns not found. Make sure Description and AI_Generated_Proposal columns exist.");
    return;
  }

  var apiKey;
  try {
    apiKey = getApiKey_();
  } catch (err) {
    ui.alert(err.message);
    return;
  }
  var settings       = getSettings_();
  var journeyContext = FFLib.buildJourneyStage(settings);
  var freelancerName = settings['Freelancer_Name'] || 'the freelancer';
  var proposalTone   = settings['Proposal_Tone']   || 'Direct';
  var portfolioAll   = settings['Portfolio_All']   || '';

  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var count   = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var desc     = descCol    ? String(data[i][descCol    - 1]).trim() : "";
    var aiProp   = aiPropCol  ? String(data[i][aiPropCol  - 1]).trim() : "";
    var jobTitle = titleCol   ? String(data[i][titleCol   - 1]).trim() : "";
    var tool     = toolCol    ? String(data[i][toolCol    - 1]).trim() : "";
    var jobType  = jobTypeCol ? String(data[i][jobTypeCol - 1]).trim() : "";
    var tmplId   = tmplCol    ? String(data[i][tmplCol    - 1]).trim().substring(0, 2) : "T1";
    var hookVer  = hookCol    ? String(data[i][hookCol    - 1]).trim() : "A";
    var ctaVer   = ctaCol     ? String(data[i][ctaCol     - 1]).trim() : "A";

    if (!desc || (aiProp && aiProp !== "" && aiProp !== "Drafting proposal...")) {
      skipped++;
      continue;
    }

    var dataRow = i + 2;
    sheet.getRange(dataRow, aiPropCol).setValue("Drafting proposal...");

    var template = lookupProposalTemplate_(tmplId || "T1", hookVer || "A", ctaVer || "A");
    var result = FFLib.generateAIProposal(jobTitle, desc, tool, jobType, template,
                                          apiKey, journeyContext, portfolioAll, proposalTone, freelancerName);
    sheet.getRange(dataRow, aiPropCol).setValue(result);
    count++;

    if (count % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert(
    "Done.\n\n" +
    "✓ " + count + " AI proposals generated.\n" +
    "-> " + skipped + " rows skipped (no description or already had a proposal)."
  );
}
