/**
 * ============================================================
 * 14. MAIN EDIT TRIGGER
 * Handles all sheet-specific onEdit automation:
 *   Job_Discovery  -- auto-timestamp, session stamp, Quick_Notes, dupe check
 *   Job_Scoring    -- date stamp on title entry, APPLY auto-proposal
 *   Connects_Helper -- replenishment accumulation + date stamp
 *   Proposal_Generator -- bid recommendation, proposal regen, Sent -> Tracker/Followup
 * ============================================================
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var ss        = e.source;
  var sheet     = e.range.getSheet();
  var sheetName = sheet.getName();
  var row       = e.range.getRow();
  var col       = e.range.getColumn();

  if (row <= 1) return;

  var map = getHeaderMap_(sheet);

  // ----------------------------------------------------------
  // JOB_DISCOVERY
  // ----------------------------------------------------------
  if (sheetName === "Job_Discovery") {
    var descColJD    = getCol_(map, ["Description"]);
    var dateFoundCol = getCol_(map, ["Date_Found"]);
    var linkColJD    = getCol_(map, ["Job_Link"]);
    var sessionIdCol = getCol_(map, ["Session_ID"]);

    if (descColJD && col === descColJD) {
      var descriptionCell = sheet.getRange(row, descColJD);
      var rawText         = descriptionCell.getValue();

      if (rawText !== "" && rawText !== null && rawText !== undefined) {
        var cleanedText = cleanJobText_(rawText);
        if (cleanedText !== rawText) {
          descriptionCell.setValue(cleanedText);
        }

        if (dateFoundCol) {
          var timestampCell = sheet.getRange(row, dateFoundCol);
          if (timestampCell.getValue() === "") {
            timestampCell.setValue(new Date());
          }
        }

        if (sessionIdCol) {
          var prop      = PropertiesService.getScriptProperties();
          var active    = prop.getProperty("SESSION_ACTIVE");
          var sessionId = prop.getProperty("SESSION_ID");
          if (active === "true" && sessionId) {
            var sessionCell = sheet.getRange(row, sessionIdCol);
            if (sessionCell.getValue() === "") {
              sessionCell.setValue(sessionId);
            }
          }
        }

        var quickNotesCol = getCol_(map, ["Quick_Notes"]);
        if (quickNotesCol) {
          sheet.getRange(row, quickNotesCol).setValue("Analyzing...");
          var quickResult = getQuickNotes_(cleanedText || rawText);
          sheet.getRange(row, quickNotesCol).setValue(quickResult);
        }
      }
    }

    if (linkColJD && col === linkColJD) {
      var pastedLink = sheet.getRange(row, linkColJD).getValue();
      var prop2      = PropertiesService.getScriptProperties();
      var active2    = prop2.getProperty("SESSION_ACTIVE");

      if (active2 === "true" && pastedLink) {
        var lastRow2   = sheet.getLastRow();
        var linkValues = sheet.getRange(2, linkColJD, lastRow2 - 1, 1).getValues();
        var linkStr    = String(pastedLink).trim();
        var matchCount = 0;

        for (var i = 0; i < linkValues.length; i++) {
          if (String(linkValues[i][0]).trim() === linkStr) matchCount++;
        }

        if (matchCount > 1) {
          var dupeCount = parseInt(prop2.getProperty("SESSION_DUPE_COUNT") || "0", 10);
          prop2.setProperty("SESSION_DUPE_COUNT", String(dupeCount + 1));

          SpreadsheetApp.getUi().alert(
            "⚠ Duplicate Detected -- Session " + prop2.getProperty("SESSION_ID") + "\n\n" +
            "This job link already exists in Job_Discovery.\n" +
            "You can delete this row and skip to the next job.\n\n" +
            "Duplicate count this session: " + (dupeCount + 1)
          );
        }
      }

      colorDuplicateJobLinks();
    }

    return;
  }

  // ----------------------------------------------------------
  // JOB_SCORING
  // ----------------------------------------------------------
  if (sheetName === "Job_Scoring") {
    var jobTitleColJS      = getCol_(map, ["Job_Title"]);
    var dateScoredCol      = getCol_(map, ["Date_Scored"]);
    var finalDecisionCol   = getCol_(map, ["Final_Decision"]);
    var proposalGenDateCol = getCol_(map, ["Proposal_Generator_Date"]);

    if (jobTitleColJS && col === jobTitleColJS && dateScoredCol) {
      var dateScoredCell = sheet.getRange(row, dateScoredCol);
      if (dateScoredCell.getValue() === "") {
        dateScoredCell.setValue(new Date());
      }
    }

    if (finalDecisionCol && proposalGenDateCol) {
      var finalDecision       = sheet.getRange(row, finalDecisionCol).getValue();
      var proposalGenDateCell = sheet.getRange(row, proposalGenDateCol);

      if (finalDecision === "APPLY" && proposalGenDateCell.getValue() === "") {
        proposalGenDateCell.setValue(new Date());

        var jsMap2         = getHeaderMap_(sheet);
        var jsTitleCol2    = getCol_(jsMap2, ["Job_Title"]);
        var jsDescCol2     = getCol_(jsMap2, ["Description"]);
        var jsToolCol2     = getCol_(jsMap2, ["Tool_Detected"]);
        var jsKeywordCol2  = getCol_(jsMap2, ["Keyword_Search"]);
        var jsProposalCol2 = getCol_(jsMap2, ["Proposal_Count"]);
        var jsBudgetCol2   = getCol_(jsMap2, ["Budget"]);
        var jsQuickCol2    = getCol_(jsMap2, ["Quick_Notes"]);

        var aiJobTitle      = jsTitleCol2    ? sheet.getRange(row, jsTitleCol2).getValue()    : "";
        var aiDescription   = jsDescCol2     ? sheet.getRange(row, jsDescCol2).getValue()     : "";
        var aiTool          = jsToolCol2     ? sheet.getRange(row, jsToolCol2).getValue()     : "";
        var aiKeyword       = jsKeywordCol2  ? sheet.getRange(row, jsKeywordCol2).getValue()  : "";
        var aiProposalCount = jsProposalCol2 ? sheet.getRange(row, jsProposalCol2).getValue() : "";
        var aiBudget        = jsBudgetCol2   ? sheet.getRange(row, jsBudgetCol2).getValue()   : "";
        var aiQuickNotes    = jsQuickCol2    ? sheet.getRange(row, jsQuickCol2).getValue()    : "";

        var aiJobType = "Dashboard Build";
        if (aiQuickNotes) {
          var qn = String(aiQuickNotes).toLowerCase();
          if (qn.indexOf("fix") !== -1 || qn.indexOf("update") !== -1 || qn.indexOf("improve") !== -1) {
            aiJobType = "Dashboard Fix";
          } else if (qn.indexOf("data") !== -1 && qn.indexOf("dashboard") === -1) {
            aiJobType = "Data to Dashboard";
          }
        }

        if (aiJobTitle && aiDescription) {
          var pgSheet = ss.getSheetByName("Proposal_Generator");
          if (pgSheet && pgSheet.getLastRow() > 1) {
            var pgMap2      = getHeaderMap_(pgSheet);
            var pgTitleCol2 = getCol_(pgMap2, ["Job_Title"]);
            var pgAiCol2    = getCol_(pgMap2, ["AI_Generated_Proposal"]);

            if (pgTitleCol2 && pgAiCol2) {
              var pgData = pgSheet
                .getRange(2, 1, pgSheet.getLastRow() - 1, pgSheet.getLastColumn())
                .getValues();

              for (var p = 0; p < pgData.length; p++) {
                if (String(pgData[p][pgTitleCol2 - 1]).trim() === String(aiJobTitle).trim()) {
                  var pgRow = p + 2;
                  pgSheet.getRange(pgRow, pgAiCol2).setValue("Generating proposal...");
                  var aiProposalText = generateAiProposal_(
                    aiJobTitle, aiDescription, aiTool,
                    aiJobType, aiProposalCount, aiBudget, aiKeyword
                  );
                  pgSheet.getRange(pgRow, pgAiCol2).setValue(aiProposalText);
                  break;
                }
              }
            }
          }
        }
      }
    }

    return;
  }

  // ----------------------------------------------------------
  // CONNECTS_HELPER
  // ----------------------------------------------------------
  if (sheetName === "Connects_Helper") {
    var metricCol = getCol_(map, ["Metric"]);
    var valueCol  = getCol_(map, ["Value"]);

    if (metricCol && valueCol && col === valueCol) {
      var metricLabel = sheet.getRange(row, metricCol).getValue();

      if (String(metricLabel).trim() === "Connect_Replenishment") {
        var newVal = sheet.getRange(row, valueCol).getValue();
        if (newVal !== "" && newVal !== null) {
          var lastRow    = sheet.getLastRow();
          var allMetrics = sheet.getRange(2, metricCol, lastRow - 1, 1).getValues();

          for (var m = 0; m < allMetrics.length; m++) {
            var mLabel = String(allMetrics[m][0]).trim();

            if (mLabel === "Connect_Replenishment_Date") {
              sheet.getRange(m + 2, valueCol).setValue(new Date());
            }

            if (mLabel === "Total_Connects_Purchased") {
              var currentTotal    = sheet.getRange(m + 2, valueCol).getValue();
              var currentTotalNum = Number(currentTotal) || 0;
              sheet.getRange(m + 2, valueCol).setValue(currentTotalNum + Number(newVal));
            }
          }
        }
      }
    }
    return;
  }

  // ----------------------------------------------------------
  // PROPOSAL_GENERATOR
  // ----------------------------------------------------------
  if (sheetName === "Proposal_Generator") {

    // Bid recommendation fires when Bid_3rd is entered
    var bid3Col = getCol_(map, ["Bid_3rd"]);
    if (bid3Col && col === bid3Col) {
      var bid3Val = sheet.getRange(row, bid3Col).getValue();
      if (bid3Val !== "" && bid3Val !== null) {
        var pgMap2       = map;
        var bid1Col      = getCol_(pgMap2, ["Bid_1st"]);
        var bid2Col      = getCol_(pgMap2, ["Bid_2nd"]);
        var baseConCol   = getCol_(pgMap2, ["Connects_Required"]);
        var propCountCol = getCol_(pgMap2, ["Proposal_Count"]);
        var titleCol2    = getCol_(pgMap2, ["Job_Title"]);
        var recCol       = getCol_(pgMap2, ["Bid_Recommendation"]);

        var bid1Val      = bid1Col      ? sheet.getRange(row, bid1Col).getValue()      : "";
        var bid2Val      = bid2Col      ? sheet.getRange(row, bid2Col).getValue()      : "";
        var baseConVal   = baseConCol   ? sheet.getRange(row, baseConCol).getValue()   : "";
        var propCountVal = propCountCol ? sheet.getRange(row, propCountCol).getValue() : "";
        var titleVal     = titleCol2    ? sheet.getRange(row, titleCol2).getValue()    : "";

        var totalScoreVal  = "";
        var scoringSheet2  = ss.getSheetByName("Job_Scoring");
        if (scoringSheet2 && titleVal) {
          var jsMap2      = getHeaderMap_(scoringSheet2);
          var jsTitleCol2 = getCol_(jsMap2, ["Job_Title"]);
          var jsScoreCol2 = getCol_(jsMap2, ["Total_Score"]);
          if (jsTitleCol2 && jsScoreCol2 && scoringSheet2.getLastRow() > 1) {
            var jsRows = scoringSheet2
              .getRange(2, 1, scoringSheet2.getLastRow() - 1, scoringSheet2.getLastColumn())
              .getValues();
            for (var s = 0; s < jsRows.length; s++) {
              if (String(jsRows[s][jsTitleCol2 - 1]).trim() === String(titleVal).trim()) {
                totalScoreVal = jsRows[s][jsScoreCol2 - 1];
                break;
              }
            }
          }
        }

        if (recCol) {
          sheet.getRange(row, recCol).setValue("Analyzing...");
          var recommendation = getBidRecommendation_(
            titleVal, baseConVal, propCountVal,
            totalScoreVal, bid1Val, bid2Val, bid3Val
          );
          sheet.getRange(row, recCol).setValue(recommendation);
        }
      }
    }

    // Proposal regen fires when Boost_Connects is entered
    var boostColPG = getCol_(map, ["Boost_Connects"]);
    if (boostColPG && col === boostColPG) {
      var boostVal = sheet.getRange(row, boostColPG).getValue();
      if (boostVal !== "" && boostVal !== null) {
        var pgMap3       = map;
        var titleColPG   = getCol_(pgMap3, ["Job_Title"]);
        var descColPG    = getCol_(pgMap3, ["Description"]);
        var toolColPG    = getCol_(pgMap3, ["Tool_Detected"]);
        var jobTypeColPG = getCol_(pgMap3, ["Job_Type"]);
        var tmplColPG    = getCol_(pgMap3, ["Recommended_Template"]);
        var hookColPG    = getCol_(pgMap3, ["Hook_Version"]);
        var ctaColPG     = getCol_(pgMap3, ["CTA_Version"]);
        var aiPropColPG  = getCol_(pgMap3, ["AI_Generated_Proposal"]);

        var pgTitle    = titleColPG   ? sheet.getRange(row, titleColPG).getValue()   : "";
        var pgDesc     = descColPG    ? sheet.getRange(row, descColPG).getValue()    : "";
        var pgTool     = toolColPG    ? sheet.getRange(row, toolColPG).getValue()    : "";
        var pgJobType  = jobTypeColPG ? sheet.getRange(row, jobTypeColPG).getValue() : "";
        var pgTemplate = tmplColPG    ? sheet.getRange(row, tmplColPG).getValue()    : "";
        var pgHook     = hookColPG    ? sheet.getRange(row, hookColPG).getValue()    : "";
        var pgCta      = ctaColPG     ? sheet.getRange(row, ctaColPG).getValue()     : "";

        if (pgDesc && aiPropColPG) {
          var tmplId  = String(pgTemplate).trim().substring(0, 2) || "T1";
          var hookVer = pgHook || "A";
          var ctaVer  = pgCta  || "A";

          sheet.getRange(row, aiPropColPG).setValue("Drafting proposal...");
          var aiProposal = generateAIProposal_(
            pgTitle, pgDesc, pgTool, pgJobType, tmplId, hookVer, ctaVer
          );
          sheet.getRange(row, aiPropColPG).setValue(aiProposal);
        }
      }
    }

    // Proposal_Status = "Sent" -> write to Proposal_Tracker + Followup_Tracker
    var proposalStatusCol   = getCol_(map, ["Proposal_Status"]);
    var proposalSentDateCol = getCol_(map, ["Proposal_Sent_Date"]);

    if (!proposalStatusCol || !proposalSentDateCol) return;
    if (col !== proposalStatusCol) return;

    var proposalStatus = sheet.getRange(row, proposalStatusCol).getValue();
    if (proposalStatus !== "Sent") return;

    var proposalSentDateCell = sheet.getRange(row, proposalSentDateCol);
    if (proposalSentDateCell.getValue() !== "") return;

    var tracker  = ss.getSheetByName("Proposal_Tracker");
    var followup = ss.getSheetByName("Followup_Tracker");
    var scoring  = ss.getSheetByName("Job_Scoring");

    if (!tracker || !followup || !scoring) return;

    var pgMap = map;
    var ptMap = getHeaderMap_(tracker);
    var ftMap = getHeaderMap_(followup);
    var jsMap = getHeaderMap_(scoring);

    var dateInGenerator    = getCellValue_(sheet, row, pgMap, ["Date"]);
    var jobTitle           = getCellValue_(sheet, row, pgMap, ["Job_Title"]);
    var clientName         = getCellValue_(sheet, row, pgMap, ["Client_Name", "Client Name"]);
    var toolRequested      = getCellValue_(sheet, row, pgMap, ["Tool_Detected", "Tool_Requested"]);
    var templateUsed       = getCellValue_(sheet, row, pgMap, ["Recommended_Template", "Template_Used"]);
    var hookVersion        = getCellValue_(sheet, row, pgMap, ["Hook_Version"]);
    var ctaVersion         = getCellValue_(sheet, row, pgMap, ["CTA_Version"]);
    var notes              = getCellValue_(sheet, row, pgMap, ["Notes"]);
    var jobLink            = getCellValue_(sheet, row, pgMap, ["Job_Link"]);
    var boostConnects      = getCellValue_(sheet, row, pgMap, ["Boost_Connects"]);
    var totalConnectsSpent = getCellValue_(sheet, row, pgMap, ["Total_Connects_Spent"]);

    var scoringLastRow = scoring.getLastRow();
    var scoringLastCol = scoring.getLastColumn();
    var scoringData    = [];
    if (scoringLastRow > 1 && scoringLastCol > 0) {
      scoringData = scoring.getRange(2, 1, scoringLastRow - 1, scoringLastCol).getValues();
    }

    var keywordSearch    = "";
    var daysSincePosted  = "";
    var hoursSincePosted = "";
    var proposalCount    = "";
    var totalScore       = "";
    var ageDays          = "";
    var currentAgeDays   = "";
    var connectsUsed     = totalConnectsSpent !== "" ? totalConnectsSpent
                           : (boostConnects !== "" ? (Number(boostConnects) + 0) : "");

    var jsJobTitleCol      = getCol_(jsMap, ["Job_Title"]);
    var jsClientCol        = getCol_(jsMap, ["Client_Name", "Client Name"]);
    var jsKeywordCol       = getCol_(jsMap, ["Keyword_Search"]);
    var jsDaysCol          = getCol_(jsMap, ["Days_Since_Posted"]);
    var jsHoursCol         = getCol_(jsMap, ["Hours_Since_Posted"]);
    var jsProposalCountCol = getCol_(jsMap, ["Proposal_Count"]);
    var jsConnectsCol      = getCol_(jsMap, ["Connects_Required"]);
    var jsJobLinkCol       = getCol_(jsMap, ["Job_Link"]);
    var jsTotalScoreCol    = getCol_(jsMap, ["Total_Score"]);
    var jsDateScoredCol    = getCol_(jsMap, ["Date_Scored"]);

    for (var i = 0; i < scoringData.length; i++) {
      var rowTitle  = jsJobTitleCol ? scoringData[i][jsJobTitleCol - 1] : "";
      var rowClient = jsClientCol   ? scoringData[i][jsClientCol   - 1] : "";

      var titleMatch  = rowTitle === jobTitle;
      var clientMatch = !clientName || !rowClient || rowClient === clientName;

      if (titleMatch && clientMatch) {
        keywordSearch    = jsKeywordCol       ? scoringData[i][jsKeywordCol       - 1] : "";
        daysSincePosted  = jsDaysCol          ? scoringData[i][jsDaysCol          - 1] : "";
        hoursSincePosted = jsHoursCol         ? scoringData[i][jsHoursCol         - 1] : "";
        proposalCount    = jsProposalCountCol ? scoringData[i][jsProposalCountCol - 1] : "";
        if (connectsUsed === "") {
          connectsUsed = jsConnectsCol ? scoringData[i][jsConnectsCol - 1] : "";
        }
        if (!jobLink) {
          jobLink = jsJobLinkCol ? scoringData[i][jsJobLinkCol - 1] : "";
        }
        totalScore = jsTotalScoreCol ? scoringData[i][jsTotalScoreCol - 1] : "";
        if (jsDateScoredCol && scoringData[i][jsDateScoredCol - 1]) {
          var scoredDate = new Date(scoringData[i][jsDateScoredCol - 1]);
          var today      = new Date();
          ageDays = Math.floor((today - scoredDate) / (1000 * 60 * 60 * 24));
        }
        if (jsDateScoredCol && scoringData[i][jsDateScoredCol - 1]) {
          var anchorDate  = new Date(scoringData[i][jsDateScoredCol - 1]);
          var now         = new Date();
          var elapsedDays = (now - anchorDate) / (1000 * 60 * 60 * 24);
          var hVal        = parseFloat(hoursSincePosted) || 0;
          var dVal        = parseFloat(daysSincePosted)  || 0;
          if (hVal > 0) {
            currentAgeDays = Math.round(((hVal / 24) + elapsedDays) * 10) / 10;
          } else if (dVal > 0) {
            currentAgeDays = Math.round((dVal + elapsedDays) * 10) / 10;
          }
        }
        break;
      }
    }

    var sentDate    = new Date();
    var appliedDate = dateInGenerator || sentDate;

    function existsInProposalTracker_() {
      var tJobCol      = getCol_(ptMap, ["Job_Title"]);
      var tClientCol   = getCol_(ptMap, ["Client_Name", "Client Name"]);
      var tTemplateCol = getCol_(ptMap, ["Template_Used", "Recommended_Template"]);
      var tHookCol     = getCol_(ptMap, ["Hook_Version"]);
      var tCtaCol      = getCol_(ptMap, ["CTA_Version"]);

      if (!tJobCol || !tClientCol || !tTemplateCol || !tHookCol || !tCtaCol) return false;

      var lastRealRow = tracker.getLastRow();
      if (lastRealRow < 2) return false;

      var data = tracker.getRange(2, 1, lastRealRow - 1, tracker.getLastColumn()).getValues();
      for (var i = 0; i < data.length; i++) {
        if (
          data[i][tJobCol      - 1] === jobTitle     &&
          data[i][tClientCol   - 1] === clientName   &&
          data[i][tTemplateCol - 1] === templateUsed &&
          data[i][tHookCol     - 1] === hookVersion  &&
          data[i][tCtaCol      - 1] === ctaVersion
        ) {
          return true;
        }
      }
      return false;
    }

    function existsInFollowupTracker_() {
      var fJobCol      = getCol_(ftMap, ["Job_Title"]);
      var fClientCol   = getCol_(ftMap, ["Client_Name", "Client Name"]);
      var fTemplateCol = getCol_(ftMap, ["Template_Used"]);

      if (!fJobCol || !fClientCol || !fTemplateCol) return false;

      var lastRealRow = followup.getLastRow();
      if (lastRealRow < 2) return false;

      var data = followup.getRange(2, 1, lastRealRow - 1, followup.getLastColumn()).getValues();
      for (var i = 0; i < data.length; i++) {
        if (
          data[i][fJobCol      - 1] === jobTitle     &&
          data[i][fClientCol   - 1] === clientName   &&
          data[i][fTemplateCol - 1] === templateUsed
        ) {
          return true;
        }
      }
      return false;
    }

    if (!existsInProposalTracker_()) {
      var ptJobTitleCol  = getCol_(ptMap, ["Job_Title"]);
      var nextTrackerRow = findFirstEmptyRowByColumn_(tracker, ptJobTitleCol);

      setCellValue_(tracker, nextTrackerRow, ptMap, ["Date_Applied"],                    appliedDate);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Job_Title"],                       jobTitle);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Client_Name", "Client Name"],      clientName);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Keyword_Search"],                  keywordSearch);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Tool_Requested", "Tool_Detected"], toolRequested);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Days_Since_Posted"],               daysSincePosted);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Proposal_Count"],                  proposalCount);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Total_Score"],                     totalScore);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Template_Used"],                   templateUsed);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Hook_Version"],                    hookVersion);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["CTA_Version"],                     ctaVersion);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Client_Replied"],                  "N");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Interview"],                       "N");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Hired"],                           "N");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Revenue"],                         "");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Notes"],                           notes);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Age_Days"],                        ageDays);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Current_Age_Days"],                currentAgeDays);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Connects_Used"],                   connectsUsed);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Boost_Connects"],                  boostConnects !== "" ? boostConnects : 0);
      var totalForCost = connectsUsed !== "" ? Number(connectsUsed) : 0;
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Proposal_Cost"],                   totalForCost > 0 ? "$" + (totalForCost * 0.15).toFixed(2) : "");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Job_Link"],                        jobLink);
    }

    if (!existsInFollowupTracker_()) {
      var fJobTitleCol    = getCol_(ftMap, ["Job_Title"]);
      var nextFollowupRow = findFirstEmptyRowByColumn_(followup, fJobTitleCol);

      setCellValue_(followup, nextFollowupRow, ftMap, ["Date_Applied"],               appliedDate);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Job_Title"],                  jobTitle);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Client_Name", "Client Name"], clientName);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Template_Used"],              templateUsed);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup1_Sent"],             "");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup1_Template"],         "F1");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup2_Sent"],             "");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup2_Template"],         "F2");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup3_Sent"],             "");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup3_Template"],         "F3");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Client_Replied"],             "N");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Interview"],                  "N");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Hired"],                      "N");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Notes"],                      notes);
    }

    proposalSentDateCell.setValue(sentDate);
  }
}
