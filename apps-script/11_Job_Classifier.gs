/**
 * ============================================================
 * 11. JOB CLASSIFIER & BATCH PROPOSALS
 *
 * Classification logic (FFLib.getJobType) and proposal generation
 * (FFLib.generateAIProposal) live in the Apps Script Library.
 * RUN_JOB_CLASSIFICATION: fills Job_Type, then picks a matching
 *   Recommended_Template/Hook_Version/CTA_Version from Proposal_Templates
 *   (FFLib.pickWeightedTemplate) for any row still missing one --
 *   including rows that already had a Job_Type from an earlier run.
 *   Also fills Portfolio_Project (FFLib.pickPortfolioProject) for any row
 *   still missing one -- same "batch-fill what the FILTER pull can't"
 *   reasoning as Job_Type: Job_Title/Description arrive via a live FILTER
 *   pull, which never fires an edit event, so nothing recomputes a formula
 *   automatically the way Job_Discovery's Tool_Detected does. This is the
 *   catch-up pass for both.
 * RUN_AI_PROPOSALS: batch-generates AI proposals for all unfilled rows
 * ============================================================
 */
function RUN_JOB_CLASSIFICATION() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || getLastRealRow_(sheet) < 2) {
    ui.alert("No rows found in Proposal_Generator.");
    return;
  }

  var map          = getHeaderMap_(sheet);
  var titleCol     = getCol_(map, ["Job_Title"]);
  var descCol      = getCol_(map, ["Description"]);
  var jobTypeCol   = getCol_(map, ["Job_Type"]);
  var tmplCol      = getCol_(map, ["Recommended_Template"]);
  var hookCol      = getCol_(map, ["Hook_Version"]);
  var ctaCol       = getCol_(map, ["CTA_Version"]);
  var portfolioCol = getCol_(map, ["Portfolio_Project"]);

  if (!descCol || !jobTypeCol || !titleCol) {
    ui.alert("Required columns not found. Confirm Job_Title, Description, and Job_Type columns exist.");
    return;
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty("UPWORK_OPENAI_API_KEY");

  var templateRows  = getProposalTemplateRows_();
  var categoryList   = getJobTypeCategories_(templateRows);
  var categoryNames  = categoryList.map(function (c) { return c.name; });
  var trackerStats   = getTemplateTrackerStats_();
  var portfolioMap   = portfolioCol ? getPortfolioMapFromProjects_() : null;

  var lastRow      = getLastRealRow_(sheet);
  var jtValues     = sheet.getRange(2, jobTypeCol, lastRow - 1, 1).getValues();
  var tmplValues   = tmplCol      ? sheet.getRange(2, tmplCol, lastRow - 1, 1).getValues()      : null;
  var portfValues  = portfolioCol ? sheet.getRange(2, portfolioCol, lastRow - 1, 1).getValues() : null;

  var filled          = 0;
  var skipped         = 0;
  var templatesSet    = 0;
  var noTemplateMatch = 0;
  var portfolioSet    = 0;

  for (var i = 0; i < jtValues.length; i++) {
    var r       = i + 2;
    var current = String(jtValues[i][0]).trim();
    var isValid = categoryNames.indexOf(current) >= 0;
    var jobType = current;

    var desc  = String(sheet.getRange(r, descCol).getValue()).trim();
    var title = String(sheet.getRange(r, titleCol).getValue()).trim();

    if (!isValid) {
      if (desc || title) {
        jobType = FFLib.getJobType(desc, title, apiKey, categoryList);
        sheet.getRange(r, jobTypeCol).setValue(jobType);
        filled++;
        if (filled % 5 === 0) Utilities.sleep(1000);
      }
    } else {
      skipped++;
    }

    if (tmplCol && hookCol && ctaCol && (desc || title)) {
      var currentTmpl = String(tmplValues[i][0]).trim();
      if (!currentTmpl) {
        var picked = FFLib.pickWeightedTemplate(jobType, templateRows, trackerStats);
        if (picked) {
          sheet.getRange(r, tmplCol).setValue(picked.templateId);
          sheet.getRange(r, hookCol).setValue(picked.hookVersion);
          sheet.getRange(r, ctaCol).setValue(picked.ctaVersion);
          templatesSet++;
        } else {
          noTemplateMatch++;
        }
      }
    }

    if (portfolioCol && (desc || title)) {
      var currentPortfolio = String(portfValues[i][0]).trim();
      if (!currentPortfolio) {
        sheet.getRange(r, portfolioCol).setValue(FFLib.pickPortfolioProject(title, desc, portfolioMap));
        portfolioSet++;
      }
    }
  }

  var msg = "Done.\n\n" +
    "✓ " + filled + " rows classified.\n" +
    "-> " + skipped + " already had a Job_Type.\n" +
    "✓ " + templatesSet + " templates assigned.\n" +
    "✓ " + portfolioSet + " Portfolio_Project values filled.";
  if (noTemplateMatch > 0) {
    msg += "\n⚠ " + noTemplateMatch + " rows had no matching template and no Is_Default row is set in Proposal_Templates.";
  }
  ui.alert(msg);

  showTourStep_(
    'FF_TOUR_STEP7_LOG_BID_SEEN',
    'Log Your Bids',
    'Each classified job needs its competing bid info. Use System Tools > Log Proposal Bid to enter the top 4 bids and get an AI bid recommendation.'
  );
}

// Reads Proposal_Templates into plain rows for FFLib.pickWeightedTemplate.
function getProposalTemplateRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Proposal_Templates');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var map        = getHeaderMap_(sheet);
  var idCol      = getCol_(map, ['Template_ID']);
  var typeCol    = getCol_(map, ['Job_Type']);
  var hookCol    = getCol_(map, ['Hook_Version']);
  var ctaCol     = getCol_(map, ['CTA_Version']);
  var notesCol   = getCol_(map, ['Notes']);
  var defaultCol = getCol_(map, ['Is_Default']);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return data.map(function (row) {
    return {
      templateId:  idCol      ? String(row[idCol      - 1]).trim() : '',
      jobType:     typeCol    ? String(row[typeCol    - 1]).trim() : '',
      hookVersion: hookCol    ? String(row[hookCol    - 1]).trim() : '',
      ctaVersion:  ctaCol     ? String(row[ctaCol     - 1]).trim() : '',
      notes:       notesCol   ? String(row[notesCol   - 1]).trim() : '',
      isDefault:   defaultCol ? String(row[defaultCol - 1]).trim().toLowerCase() === 'yes' : false
    };
  }).filter(function (r) { return r.templateId; });
}

// Distinct {name, notes} pairs from Proposal_Templates' Job_Type column, in
// first-seen order -- the classifier's category list. Notes doubles as
// classifier guidance (FFLib.getJobType) and human documentation in the sheet.
function getJobTypeCategories_(templateRows) {
  var seen = {};
  var list = [];
  templateRows.forEach(function (r) {
    if (!r.jobType || seen[r.jobType]) return;
    seen[r.jobType] = true;
    list.push({ name: r.jobType, notes: r.notes || '' });
  });
  return list;
}

// Aggregates Proposal_Tracker into { "TemplateID|Hook|CTA": {sent, viewed} }
// for FFLib.pickWeightedTemplate's view-rate weighting. Always read fresh --
// Proposal_Tracker changes continuously as proposals get viewed/replied to.
function getTemplateTrackerStats_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Proposal_Tracker');
  if (!sheet || sheet.getLastRow() < 2) return {};

  var map     = getHeaderMap_(sheet);
  var tmplCol = getCol_(map, ['Template_Used']);
  var hookCol = getCol_(map, ['Hook_Version']);
  var ctaCol  = getCol_(map, ['CTA_Version']);
  var viewCol = getCol_(map, ['Viewed']);

  if (!tmplCol || !hookCol || !ctaCol) return {};

  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var stats = {};

  data.forEach(function (row) {
    var tmpl = String(row[tmplCol - 1]).trim();
    var hook = String(row[hookCol - 1]).trim();
    var cta  = String(row[ctaCol  - 1]).trim();
    if (!tmpl) return;

    var key = tmpl + '|' + hook + '|' + cta;
    if (!stats[key]) stats[key] = { sent: 0, viewed: 0 };
    stats[key].sent++;
    if (viewCol && String(row[viewCol - 1]).trim() === 'Y') stats[key].viewed++;
  });

  return stats;
}


function RUN_AI_PROPOSALS() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || getLastRealRow_(sheet) < 2) {
    ui.alert("No rows found in Proposal_Generator.");
    return;
  }

  var map          = getHeaderMap_(sheet);
  var titleCol     = getCol_(map, ["Job_Title"]);
  var descCol      = getCol_(map, ["Description"]);
  var toolCol      = getCol_(map, ["Tool_Detected"]);
  var jobTypeCol   = getCol_(map, ["Job_Type"]);
  var tmplCol      = getCol_(map, ["Recommended_Template"]);
  var hookCol      = getCol_(map, ["Hook_Version"]);
  var ctaCol       = getCol_(map, ["CTA_Version"]);
  var aiPropCol    = getCol_(map, ["AI_Generated_Proposal"]);
  var questionsCol = getCol_(map, ["Additional_Questions"]);
  var answersCol   = getCol_(map, ["Additional_Answers"]);

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
  var settings         = getSettings_();
  var journeyContext   = FFLib.buildJourneyStage(settings);
  var freelancerName   = settings['Freelancer_Name'] || 'the freelancer';
  var proposalTone     = settings['Proposal_Tone']   || 'Direct';
  var proposalLength   = settings['Proposal_Length'] || 'Medium';
  var portfolioAll     = settings['Portfolio_All']   || '';
  var portfolioContext = FFLib.getPortfolioContext(settings);

  var lastRow      = getLastRealRow_(sheet);
  var data         = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var count        = 0;
  var answersCount = 0;
  var skipped      = 0;

  for (var i = 0; i < data.length; i++) {
    var desc      = descCol      ? String(data[i][descCol      - 1]).trim() : "";
    var aiProp    = aiPropCol    ? String(data[i][aiPropCol    - 1]).trim() : "";
    var jobTitle  = titleCol     ? String(data[i][titleCol     - 1]).trim() : "";
    var tool      = toolCol      ? String(data[i][toolCol      - 1]).trim() : "";
    var jobType   = jobTypeCol   ? String(data[i][jobTypeCol   - 1]).trim() : "";
    var tmplId    = tmplCol      ? String(data[i][tmplCol      - 1]).trim().substring(0, 2) : "T1";
    var hookVer   = hookCol      ? String(data[i][hookCol      - 1]).trim() : "A";
    var ctaVer    = ctaCol       ? String(data[i][ctaCol       - 1]).trim() : "A";
    var questions = questionsCol ? String(data[i][questionsCol - 1]).trim() : "";
    var answers   = answersCol   ? String(data[i][answersCol   - 1]).trim() : "";

    var dataRow       = i + 2;
    var needsProposal = desc && !(aiProp && aiProp !== "" && aiProp !== "Drafting proposal...");
    var needsAnswers  = questions && answersCol && !(answers && answers !== "" && answers !== "Drafting answers...");

    if (!needsProposal && !needsAnswers) {
      skipped++;
      continue;
    }

    if (needsProposal) {
      sheet.getRange(dataRow, aiPropCol).setValue("Drafting proposal...");
      var template = lookupProposalTemplate_(tmplId || "T1", hookVer || "A", ctaVer || "A");
      var result = FFLib.generateAIProposal(jobTitle, desc, tool, jobType, template,
                                            apiKey, journeyContext, portfolioAll, proposalTone, freelancerName, proposalLength);
      sheet.getRange(dataRow, aiPropCol).setValue(result);
      count++;
    }

    if (needsAnswers) {
      sheet.getRange(dataRow, answersCol).setValue("Drafting answers...");
      var answerResult = FFLib.generateAdditionalAnswers(questions, jobTitle, desc, portfolioContext, freelancerName, apiKey);
      sheet.getRange(dataRow, answersCol).setValue(answerResult);
      answersCount++;
    }

    if ((count + answersCount) % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert(
    "Done.\n\n" +
    "✓ " + count + " AI proposals generated.\n" +
    "✓ " + answersCount + " additional-question answers drafted.\n" +
    "-> " + skipped + " rows skipped (nothing to do)."
  );
}
