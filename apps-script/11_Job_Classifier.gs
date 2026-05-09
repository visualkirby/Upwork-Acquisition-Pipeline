/**
 * ============================================================
 * 11. JOB CLASSIFIER & BATCH PROPOSALS
 *
 * getJobType_: classifies a job into one of four categories
 * RUN_JOB_CLASSIFICATION: fills Job_Type column in Proposal_Generator
 * RUN_AI_PROPOSALS: batch-generates AI proposals for all unfilled rows
 * ============================================================
 */
function getJobType_(description, jobTitle) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  function regexFallback_() {
    var t = (jobTitle + " " + description).toLowerCase();
    if (/fix|improve|update|modify|redesign|optimize|existing/.test(t)) return "Dashboard Fix";
    if (/clean|spreadsheet|raw data|data cleaning|csv|excel file|export/.test(t)) return "Data to Dashboard";
    if (/report|analysis|analytics(?! dashboard)|insight/.test(t) && !/dashboard/.test(t)) return "Reporting";
    return "Dashboard Build";
  }

  if (!apiKey) return regexFallback_();

  var prompt =
    "Classify this Upwork job into exactly one of these four categories: " +
    "Dashboard Build, Dashboard Fix, Data to Dashboard, Reporting. " +
    "Dashboard Build = new dashboard needed from scratch. " +
    "Dashboard Fix = existing dashboard needs fixing or improving. " +
    "Data to Dashboard = raw data needs cleaning then turned into a dashboard. " +
    "Reporting = data analysis or reporting without a dashboard. " +
    "Reply with only the category name and nothing else. " +
    "Job: " + jobTitle + ". " + description.substring(0, 800);

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        max_tokens: 10,
        temperature: 0
      }),
      muteHttpExceptions: true
    });

    var parsed = JSON.parse(response.getContentText());
    if (!parsed.choices || !parsed.choices[0]) return regexFallback_();

    var text = parsed.choices[0].message.content.trim();

    if (text.indexOf("Data to Dashboard") !== -1) return "Data to Dashboard";
    if (text.indexOf("Dashboard Fix")     !== -1) return "Dashboard Fix";
    if (text.indexOf("Dashboard Build")   !== -1) return "Dashboard Build";
    if (text.indexOf("Reporting")         !== -1) return "Reporting";
    return regexFallback_();

  } catch (err) {
    return regexFallback_();
  }
}


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

    var result = getJobType_(desc, title);
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

    var result = generateAIProposal_(jobTitle, desc, tool, jobType,
                                      tmplId || "T1", hookVer || "A", ctaVer || "A");
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
