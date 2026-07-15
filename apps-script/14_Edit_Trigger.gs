/**
 * ============================================================
 * 14. MAIN EDIT TRIGGER
 * Handles all sheet-specific edit automation:
 *   Job_Discovery  -- auto-timestamp, session stamp, AI_Fit_Notes, dupe check,
 *                     Keyword_Strategy Actual_Count increment
 *   Job_Scoring    -- date stamp on title entry, APPLY auto-proposal
 *   Connects_Helper -- replenishment/return accumulation + date stamp, both of
 *                      which also add back into Current_Connect_Balance
 *   Proposal_Tracker -- Viewed/Interview/Hired feed Connects_Helper's MTD_Replies/
 *                       MTD_Interviews/MTD_Hires; Interview=Y opens Chat Import,
 *                       Hired=Y opens Contract Setup. Revenue is manual/informational only.
 *   Milestone_Tracker -- Status date-stamps (Funded/Delivered/Released); Released feeds revenue
 *   Contract_Tracker -- Status=Completed rolls up Total_Released; Ended Early prompts for reconciliation
 *   Hourly_Log -- Hours_Logged computes Amount from Contract_Tracker's rate, feeds revenue by delta
 *   Proposal_Generator -- bid recommendation, proposal regen, Sent -> Proposal_Tracker,
 *                         Current_Connect_Balance/Total_Connects_Used decrement
 *
 * Named handleEdit (not onEdit) so Apps Script never auto-registers it as
 * a simple trigger. Simple triggers run in a restricted authorization mode
 * that can't call UrlFetchApp, so every AI call here would randomly fail
 * (or race against the installable trigger and overwrite its result with
 * an error) if this were ever named onEdit. registerEditTrigger_() in
 * 00_Setup_Wizard.gs installs this as an installable trigger instead, which
 * always runs with full authorization. Not trailing-underscored either --
 * ScriptApp.newTrigger() must reference it by name, and private (_-suffixed)
 * functions are excluded from the trigger function picker.
 * ============================================================
 */
function handleEdit(e) {
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
    var descColJD       = getCol_(map, ["Description"]);
    var dateFoundCol    = getCol_(map, ["Date_Found"]);
    var linkColJD       = getCol_(map, ["Job_Link"]);
    var sessionIdCol    = getCol_(map, ["Session_ID"]);
    var keywordColJD    = getCol_(map, ["Keyword_Search"]);

    if (descColJD && col === descColJD) {
      var descriptionCell = sheet.getRange(row, descColJD);
      var rawText         = descriptionCell.getValue();

      if (rawText !== "" && rawText !== null && rawText !== undefined) {
        var cleanedText = cleanJobText_(rawText);
        if (cleanedText !== rawText) {
          descriptionCell.setValue(cleanedText);
        }

        // First-time-logged guard -- Description can be re-edited later (e.g.
        // cleaning it up again), which would re-run this whole block. Only
        // count the keyword once, the same moment Date_Found first gets stamped.
        //
        // Uses e.oldValue (the snapshot for THIS specific edit to Description),
        // not a live read of Date_Found -- two Description edits landing on the
        // same row in quick succession queue two async executions, and by the
        // time either runs, a live read of Date_Found could see it still blank
        // for both (the first execution's write hasn't committed yet), double-
        // counting the keyword. e.oldValue is fixed at dispatch time, so both
        // executions can't agree it was blank.
        var isFirstLog = (e.oldValue === undefined || e.oldValue === null || e.oldValue === "");

        if (dateFoundCol) {
          var timestampCell = sheet.getRange(row, dateFoundCol);
          if (timestampCell.getValue() === "") {
            timestampCell.setValue(new Date());
          }
        }

        if (isFirstLog && keywordColJD) {
          incrementKeywordStrategyActualCount_(ss, sheet.getRange(row, keywordColJD).getValue());
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

        var aiFitNotesCol = getCol_(map, ["AI_Fit_Notes"]);
        if (aiFitNotesCol) {
          sheet.getRange(row, aiFitNotesCol).setValue("Analyzing...");
          var qnApiKey  = PropertiesService.getScriptProperties().getProperty("UPWORK_OPENAI_API_KEY");
          var qnSettings = getSettings_();
          var quickResult = FFLib.getQuickNotes(cleanedText || rawText, qnApiKey, qnSettings);
          sheet.getRange(row, aiFitNotesCol).setValue(quickResult);
        }

        showWalkthroughOnce_(
          "FF_WALKTHROUGH_JOB_DISCOVERY_SEEN",
          "First job logged!",
          "AI_Fit_Notes just analyzed how well this job matches your profile.\n\n" +
          "Next: head to Job_Scoring to score this job and decide whether to apply."
        );

        handleSessionYieldReached_(ss);
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
    var descColJS          = getCol_(map, ["Description"]);
    var aiFitNotesColJS    = getCol_(map, ["AI_Fit_Notes"]);

    if (jobTitleColJS && col === jobTitleColJS && dateScoredCol) {
      var dateScoredCell = sheet.getRange(row, dateScoredCol);
      if (dateScoredCell.getValue() === "") {
        dateScoredCell.setValue(new Date());
      }
    }

    if (descColJS && col === descColJS && aiFitNotesColJS) {
      var jsDescText = sheet.getRange(row, descColJS).getValue();
      if (jsDescText !== "" && jsDescText !== null && jsDescText !== undefined) {
        sheet.getRange(row, aiFitNotesColJS).setValue("Analyzing...");
        var jsQnApiKey   = PropertiesService.getScriptProperties().getProperty("UPWORK_OPENAI_API_KEY");
        var jsQnSettings = getSettings_();
        var jsQuickResult = FFLib.getQuickNotes(jsDescText, jsQnApiKey, jsQnSettings);
        sheet.getRange(row, aiFitNotesColJS).setValue(jsQuickResult);
      }
    }

    if (finalDecisionCol && proposalGenDateCol) {
      var finalDecision       = sheet.getRange(row, finalDecisionCol).getValue();
      var proposalGenDateCell = sheet.getRange(row, proposalGenDateCol);

      if (finalDecision === "APPLY" && proposalGenDateCell.getValue() === "") {
        proposalGenDateCell.setValue(new Date());

        showWalkthroughOnce_(
          "FF_WALKTHROUGH_JOB_SCORING_SEEN",
          "First job scored APPLY!",
          "This job is strong enough to apply to, so it's being sent to Proposal_Generator.\n\n" +
          "Next: head to Proposal_Generator to review and send the AI-drafted proposal."
        );

        var jsMap2         = getHeaderMap_(sheet);
        var jsTitleCol2    = getCol_(jsMap2, ["Job_Title"]);
        var jsDescCol2     = getCol_(jsMap2, ["Description"]);
        var jsToolCol2     = getCol_(jsMap2, ["Tool_Detected"]);
        var jsKeywordCol2  = getCol_(jsMap2, ["Keyword_Search"]);
        var jsProposalCol2 = getCol_(jsMap2, ["Proposal_Count"]);
        var jsBudgetCol2   = getCol_(jsMap2, ["Budget"]);

        var aiJobTitle      = jsTitleCol2    ? sheet.getRange(row, jsTitleCol2).getValue()    : "";
        var aiDescription   = jsDescCol2     ? sheet.getRange(row, jsDescCol2).getValue()     : "";
        var aiTool          = jsToolCol2     ? sheet.getRange(row, jsToolCol2).getValue()     : "";
        var aiKeyword       = jsKeywordCol2  ? sheet.getRange(row, jsKeywordCol2).getValue()  : "";
        var aiProposalCount = jsProposalCol2 ? sheet.getRange(row, jsProposalCol2).getValue() : "";
        var aiBudget        = jsBudgetCol2   ? sheet.getRange(row, jsBudgetCol2).getValue()   : "";

        var jtApiKey     = PropertiesService.getScriptProperties().getProperty("UPWORK_OPENAI_API_KEY");
        var jtCategories = getJobTypeCategories_(getProposalTemplateRows_());
        var aiJobType    = FFLib.getJobType(aiDescription, aiJobTitle, jtApiKey, jtCategories);

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
                  var aiProposalText;
                  try {
                    var autoApiKey          = getApiKey_();
                    var autoSettings        = getSettings_();
                    var autoPortfolioContext = FFLib.getPortfolioContext(autoSettings);
                    var autoFreelancerName   = autoSettings['Freelancer_Name'] || 'the freelancer';
                    aiProposalText = FFLib.generateAiProposal(
                      aiJobTitle, aiDescription, aiTool,
                      aiJobType, aiProposalCount, aiBudget, aiKeyword,
                      autoApiKey, autoPortfolioContext, autoFreelancerName
                    );
                  } catch (err) {
                    aiProposalText = err.message;
                  }
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

      if (String(metricLabel).trim() === "Connect_Returned") {
        var retVal = sheet.getRange(row, valueCol).getValue();
        var retOld = e.oldValue;
        if (retVal !== "" && retVal !== null && (retOld === undefined || retOld === null || retOld === "")) {
          var retLastRow    = sheet.getLastRow();
          var retAllMetrics = sheet.getRange(2, metricCol, retLastRow - 1, 1).getValues();

          for (var r = 0; r < retAllMetrics.length; r++) {
            if (String(retAllMetrics[r][0]).trim() === "Total_Connects_Returned") {
              var currentReturned = sheet.getRange(r + 2, valueCol).getValue();
              sheet.getRange(r + 2, valueCol).setValue((Number(currentReturned) || 0) + Number(retVal));
            }
            if (String(retAllMetrics[r][0]).trim() === "Connect_Returned_Date") {
              sheet.getRange(r + 2, valueCol).setValue(new Date());
            }
          }

          incrementConnectsHelperMetric_(ss, "Current_Connect_Balance", Number(retVal));
        }
      }

      if (String(metricLabel).trim() === "Connect_Replenishment") {
        var newVal = sheet.getRange(row, valueCol).getValue();
        var oldVal = e.oldValue;
        // Only accumulate when entering a fresh value into a blank cell.
        // Editing an existing value would double-count the replenishment.
        if (newVal !== "" && newVal !== null && (oldVal === undefined || oldVal === null || oldVal === "")) {
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

          incrementConnectsHelperMetric_(ss, "Current_Connect_Balance", Number(newVal));
        }
      }
    }
    return;
  }

  // ----------------------------------------------------------
  // PROPOSAL_TRACKER
  // Viewed/Interview/Hired flipping to "Y" for the first time feeds
  // Connects_Helper's MTD_Replies/MTD_Interviews/MTD_Hires -- guarded on
  // oldValue so re-saving an already-"Y" cell doesn't double count, same
  // guard style as Connect_Replenishment above. Revenue uses a delta
  // instead, since it's a manually-entered number that may get corrected
  // after the fact rather than a one-time flag.
  //
  // Uses e.value/e.oldValue (the snapshot pair for THIS specific edit),
  // not a live sheet.getRange().getValue() re-read -- two edits landing
  // on the same cell in quick succession (e.g. clear then retype) queue
  // two async executions, and by the time either runs, a live read would
  // see the SAME final value for both, pairing it against two different
  // oldValues and double-counting the delta.
  // ----------------------------------------------------------
  if (sheetName === "Proposal_Tracker") {
    var ptViewedCol    = getCol_(map, ["Viewed"]);
    var ptInterviewCol = getCol_(map, ["Interview"]);
    var ptHiredCol     = getCol_(map, ["Hired"]);
    var ptIdCol        = getCol_(map, ["Discovery_ID"]);
    var ptOldVal       = e.oldValue;
    var ptNewVal       = e.value;

    if (ptViewedCol && col === ptViewedCol && ptNewVal === "Y" && ptOldVal !== "Y") {
      incrementConnectsHelperMetric_(ss, "MTD_Replies", 1);
    }
    if (ptInterviewCol && col === ptInterviewCol && ptNewVal === "Y" && ptOldVal !== "Y") {
      incrementConnectsHelperMetric_(ss, "MTD_Interviews", 1);
      // Client replied -- Upwork opens an ongoing chat thread at this point.
      // Auto-open the Chat Import sidebar so the freelancer can paste it in
      // right away instead of hunting for the menu item later. Tour step
      // fires first (same order Step 3 uses for Log New Job) so the alert
      // resolves before the sidebar takes focus.
      if (ptIdCol) {
        showTourStep_(
          "FF_TOUR_STEP8_CHAT_IMPORT_SEEN",
          "Import the Chat",
          "The client replied, so the Import Client Chat sidebar is opening now -- paste the conversation and it'll parse it into Client_Chat_Log automatically."
        );
        openChatImportSidebar_(sheet.getRange(row, ptIdCol).getValue());
      }
    }
    if (ptHiredCol && col === ptHiredCol && ptNewVal === "Y" && ptOldVal !== "Y") {
      incrementConnectsHelperMetric_(ss, "MTD_Hires", 1);
      // Hired -- auto-open the Contract Setup sidebar so the freelancer can
      // log the contract type and milestones right away. contract_saveSetup
      // creates the Contract_Tracker row itself on submit, so nothing needs
      // to be pre-created here.
      if (ptIdCol) {
        showTourStep_(
          "FF_TOUR_STEP9_CONTRACT_SETUP_SEEN",
          "Log the Contract",
          "You're hired! The Log New Contract sidebar is opening now -- for fixed-price work, add each milestone's description and amount. For hourly work, just set a rate; a starter row lands in Hourly_Log ready for you to log hours."
        );
        openContractSetupSidebar_(sheet.getRange(row, ptIdCol).getValue());
      }
    }
    // Revenue is a plain manual/informational field here -- it no longer
    // drives Connects_Helper. Milestone_Tracker's Released status and
    // Hourly_Log entries are the authoritative revenue sources now (see the
    // MILESTONE_TRACKER/CONTRACT_TRACKER/HOURLY_LOG blocks below), since
    // they reflect the real escrow/payment lifecycle instead of a freeform
    // number a freelancer might type in before funds have actually cleared.
    return;
  }

  // ----------------------------------------------------------
  // MILESTONE_TRACKER
  // Status transitions stamp their date column once, and only "Released"
  // (funds actually paid out, not merely "Delivered") feeds Connects_Helper's
  // revenue metrics -- delivering work and being paid are different moments
  // in Upwork's real escrow flow, and counting revenue at Delivered would
  // overstate income before it's certain.
  // ----------------------------------------------------------
  if (sheetName === "Milestone_Tracker") {
    var msStatusCol = getCol_(map, ["Status"]);
    if (msStatusCol && col === msStatusCol) {
      applyMilestoneStatusEffects_(ss, sheet, row, map, e.oldValue, e.value);
    }
    return;
  }

  // ----------------------------------------------------------
  // CONTRACT_TRACKER
  // Status -> Completed computes a Total_Released rollup for a clean final
  // total on the row (sum of this contract's Released milestones, or its
  // Hourly_Log entries) -- no separate Connects_Helper push here, since
  // those amounts already fed revenue individually as each one cleared.
  // Status -> Ended Early runs a short manual reconciliation prompt, since
  // a contract ending early has no clean milestone/hourly trail to sum.
  // ----------------------------------------------------------
  if (sheetName === "Contract_Tracker") {
    var ctStatusCol = getCol_(map, ["Status"]);
    if (ctStatusCol && col === ctStatusCol) {
      var ctOldVal      = e.oldValue;
      var ctNewVal      = e.value;
      var ctIdCol       = getCol_(map, ["Discovery_ID"]);
      var ctTypeCol     = getCol_(map, ["Contract_Type"]);
      var ctTotalRelCol = getCol_(map, ["Total_Released"]);
      var ctDiscoveryId = ctIdCol ? sheet.getRange(row, ctIdCol).getValue() : "";
      var ctType        = ctTypeCol ? sheet.getRange(row, ctTypeCol).getValue() : "";

      if (ctNewVal === "Completed" && ctOldVal !== "Completed" && ctTotalRelCol) {
        var ctTotal = getContractRecognizedRevenue_(ss, ctDiscoveryId, ctType);
        sheet.getRange(row, ctTotalRelCol).setValue(ctTotal);
      }

      if (ctNewVal === "Ended Early" && ctOldVal !== "Ended Early") {
        var ctUi = SpreadsheetApp.getUi();
        var ctAmountResponse = ctUi.prompt(
          "Contract Ended Early",
          "Enter the actual amount funded/received for this contract (0 if none):",
          ctUi.ButtonSet.OK_CANCEL
        );
        if (ctAmountResponse.getSelectedButton() === ctUi.Button.OK) {
          var ctEndedAmount = Number(ctAmountResponse.getResponseText()) || 0;
          var ctReleasedResponse = ctUi.prompt(
            "Contract Ended Early",
            "Were those funds released to you? (Y/N)",
            ctUi.ButtonSet.OK_CANCEL
          );
          var ctReleased = ctReleasedResponse.getSelectedButton() === ctUi.Button.OK &&
            String(ctReleasedResponse.getResponseText()).trim().toUpperCase() === "Y";

          endContractEarly_(ss, ctDiscoveryId, ctType, ctEndedAmount, ctReleased);
        }
      }
    }
    return;
  }

  // ----------------------------------------------------------
  // HOURLY_LOG
  // Amount is script-computed (not a live formula) from Hours_Logged x this
  // contract's Hourly_Rate, looked up from Contract_Tracker by Discovery_ID.
  // A formula wouldn't work here -- e.value/e.oldValue only capture the cell
  // actually edited (Hours_Logged), not a dependent formula cell, so there'd
  // be no way to compute the delta needed to avoid double-counting a
  // correction. Delta-based, same safety pattern the old
  // Proposal_Tracker.Revenue trigger used.
  //
  // Reads Hours_Logged/Amount fresh off the sheet per row in e.range rather
  // than trusting e.value/col -- a fast multi-cell commit (row-fill, paste,
  // or several Tab-committed cells landing as one edit) reports e.range
  // spanning multiple columns/rows with e.value undefined and col/row set to
  // the range's top-left cell, not the cell that actually changed.
  //
  // Recomputes for EVERY row in e.range regardless of which column was
  // touched -- gating on "only if Hours_Logged is within the edited column
  // range" (an earlier version of this fix) still lost rows. Confirmed via
  // added logging on 2026-07-05: fast Tab-across-row entry can commit a
  // row's cells as several separate single-cell edits, but Google Sheets
  // silently never dispatched onEdit at all for that row's Discovery_ID or
  // Hours_Logged cells specifically -- only the Job_Title/Log_Date columns'
  // edits fired. No in-code range check can compensate for a trigger that
  // never invokes. Instead, any edit anywhere on this sheet now recomputes
  // Amount for every touched row from whatever Hours_Logged currently holds
  // -- since at least one of a row's several cell commits reliably fires
  // (confirmed in the same test), that's enough to self-heal the row even
  // when the Hours_Logged cell's own edit event never arrives.
  // ----------------------------------------------------------
  if (sheetName === "Hourly_Log") {
    var hlEditStartRow = e.range.getRow();
    var hlEditNumRows  = e.range.getNumRows();

    for (var hlRow = hlEditStartRow; hlRow < hlEditStartRow + hlEditNumRows; hlRow++) {
      if (hlRow <= 1) continue;
      applyHourlyLogAmount_(ss, sheet, hlRow, map);
    }

    // Status is a single dropdown cell (not a fill/paste range like
    // Hours_Logged commonly is), so this checks the specific edited column
    // rather than looping every row in e.range the way the Amount recompute
    // above does.
    var hlStatusCol = getCol_(map, ["Status"]);
    if (hlStatusCol && col === hlStatusCol) {
      applyHourlyLogStatusEffects_(ss, sheet, row, map, e.oldValue, e.value);
    }
    return;
  }

  // ----------------------------------------------------------
  // PROPOSAL_GENERATOR
  // Per-behavior logic lives in computeBidRecommendation_/applyBoostConnects_/
  // generateAdditionalAnswers_/handleProposalStatusChange_ below (not inlined
  // here) -- both this trigger AND the Proposal_Generator sidebar's
  // proposal_saveBids/proposal_saveBoost/proposal_saveQuestions/
  // proposal_saveStatusNotes (22_Proposal_Generator_Sidebar.gs) need to fire
  // the same automation, since script-driven writes from that sidebar never
  // trigger handleEdit on their own.
  // ----------------------------------------------------------
  if (sheetName === "Proposal_Generator") {

    // Bid recommendation fires when Bid_4th is entered -- Upwork shows the
    // top 4 competing bids, so that's the last one visible before deciding.
    var bid4Col = getCol_(map, ["Bid_4th"]);
    if (bid4Col && col === bid4Col) computeBidRecommendation_(ss, sheet, row, map);

    // Proposal regen fires when Boost_Connects is entered
    var boostColPG = getCol_(map, ["Boost_Connects"]);
    if (boostColPG && col === boostColPG) applyBoostConnects_(ss, sheet, row, map);

    // Additional_Answers regenerates whenever Additional_Questions changes
    var questionsColPG = getCol_(map, ["Additional_Questions"]);
    if (questionsColPG && col === questionsColPG) generateAdditionalAnswers_(ss, sheet, row, map);

    // Proposal_Status changing (Skip stamps a date; Sent syncs to Proposal_Tracker)
    var proposalStatusCol = getCol_(map, ["Proposal_Status"]);
    if (proposalStatusCol && col === proposalStatusCol) {
      handleProposalStatusChange_(ss, sheet, row, map);
    }
  }
}

// Fires the Upwork top-4-bids AI recommendation -- self-contained (reads
// Bid_4th itself and no-ops if blank) so it can be called either from
// handleEdit's per-column check above or directly from the Proposal_Generator
// sidebar after it writes all four bid fields at once.
function computeBidRecommendation_(ss, sheet, row, map) {
  var bid4Col = getCol_(map, ["Bid_4th"]);
  if (!bid4Col) return;
  var bid4Val = sheet.getRange(row, bid4Col).getValue();
  if (bid4Val === "" || bid4Val === null) return;

  var bid1Col      = getCol_(map, ["Bid_1st"]);
  var bid2Col      = getCol_(map, ["Bid_2nd"]);
  var bid3Col      = getCol_(map, ["Bid_3rd"]);
  var baseConCol   = getCol_(map, ["Connects_Required"]);
  var propCountCol = getCol_(map, ["Proposal_Count"]);
  var titleCol2    = getCol_(map, ["Job_Title"]);
  var recCol       = getCol_(map, ["Bid_Recommendation"]);

  var bid1Val      = bid1Col      ? sheet.getRange(row, bid1Col).getValue()      : "";
  var bid2Val      = bid2Col      ? sheet.getRange(row, bid2Col).getValue()      : "";
  var bid3Val      = bid3Col      ? sheet.getRange(row, bid3Col).getValue()      : "";
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

  if (!recCol) return;
  sheet.getRange(row, recCol).setValue("Analyzing...");
  var recommendation;
  try {
    var bidApiKey          = getApiKey_();
    var bidSettings        = getSettings_();
    var bidJourneyContext  = FFLib.buildJourneyStage(bidSettings);
    var bidNoBoostMaxProp  = parseInt(bidSettings['Apply_Max_Proposals']) || 35;
    var bidNoBoostMinScore = parseFloat(bidSettings['Apply_Min_Score'])   || 0.60;
    recommendation = FFLib.getBidRecommendation(
      titleVal, baseConVal, propCountVal,
      totalScoreVal, bid1Val, bid2Val, bid3Val, bid4Val,
      bidApiKey, bidJourneyContext, bidNoBoostMaxProp, bidNoBoostMinScore
    );
  } catch (err) {
    recommendation = err.message;
  }
  sheet.getRange(row, recCol).setValue(recommendation);
}

// Recomputes Total_Connects_Spent from Boost_Connects + Connects_Required,
// then regenerates the AI proposal with the boosted context. Self-contained
// (reads Boost_Connects itself) so it can be called from handleEdit or
// directly from the Proposal_Generator sidebar.
function applyBoostConnects_(ss, sheet, row, map) {
  var boostColPG = getCol_(map, ["Boost_Connects"]);
  if (!boostColPG) return;
  var boostVal = sheet.getRange(row, boostColPG).getValue();

  // Total_Connects_Spent = base connects + boost, recalculated every time
  // Boost_Connects changes, independent of whether the proposal regen
  // conditions below are met.
  var totalSpentColPG = getCol_(map, ["Total_Connects_Spent"]);
  var connReqColPG    = getCol_(map, ["Connects_Required"]);
  if (totalSpentColPG && connReqColPG) {
    var connReqValPG = sheet.getRange(row, connReqColPG).getValue();
    sheet.getRange(row, totalSpentColPG).setValue((Number(connReqValPG) || 0) + (Number(boostVal) || 0));
  }

  if (boostVal === "" || boostVal === null) return;

  var titleColPG   = getCol_(map, ["Job_Title"]);
  var descColPG    = getCol_(map, ["Description"]);
  var toolColPG    = getCol_(map, ["Tool_Detected"]);
  var jobTypeColPG = getCol_(map, ["Job_Type"]);
  var tmplColPG    = getCol_(map, ["Recommended_Template"]);
  var hookColPG    = getCol_(map, ["Hook_Version"]);
  var ctaColPG     = getCol_(map, ["CTA_Version"]);
  var aiPropColPG  = getCol_(map, ["AI_Generated_Proposal"]);

  var pgTitle    = titleColPG   ? sheet.getRange(row, titleColPG).getValue()   : "";
  var pgDesc     = descColPG    ? sheet.getRange(row, descColPG).getValue()    : "";
  var pgTool     = toolColPG    ? sheet.getRange(row, toolColPG).getValue()    : "";
  var pgJobType  = jobTypeColPG ? sheet.getRange(row, jobTypeColPG).getValue() : "";
  var pgTemplate = tmplColPG    ? sheet.getRange(row, tmplColPG).getValue()    : "";
  var pgHook     = hookColPG    ? sheet.getRange(row, hookColPG).getValue()    : "";
  var pgCta      = ctaColPG     ? sheet.getRange(row, ctaColPG).getValue()     : "";

  if (!pgDesc || !aiPropColPG) return;

  var tmplId  = String(pgTemplate).trim().substring(0, 2) || "T1";
  var hookVer = pgHook || "A";
  var ctaVer  = pgCta  || "A";

  sheet.getRange(row, aiPropColPG).setValue("Drafting proposal...");
  var aiProposal;
  try {
    var boostApiKey         = getApiKey_();
    var boostSettings       = getSettings_();
    var boostJourneyContext = FFLib.buildJourneyStage(boostSettings);
    var boostTemplate       = lookupProposalTemplate_(tmplId, hookVer, ctaVer);
    var boostFreelancerName = boostSettings['Freelancer_Name'] || 'the freelancer';
    var boostProposalTone   = boostSettings['Proposal_Tone']   || 'Direct';
    var boostPortfolioAll   = boostSettings['Portfolio_All']   || '';
    aiProposal = FFLib.generateAIProposal(
      pgTitle, pgDesc, pgTool, pgJobType, boostTemplate,
      boostApiKey, boostJourneyContext, boostPortfolioAll, boostProposalTone, boostFreelancerName
    );
  } catch (err) {
    aiProposal = err.message;
  }
  sheet.getRange(row, aiPropColPG).setValue(aiProposal);
}

// Drafts answers to Additional_Questions -- self-contained (reads
// Additional_Questions itself and no-ops if blank) so it can be called either
// from handleEdit's per-column check above or directly from the
// Proposal_Generator sidebar. Always regenerates when called (not just once)
// since the freelancer may add or edit a line after already getting an
// answer back -- callers are responsible for only calling this when the
// question text actually changed.
function generateAdditionalAnswers_(ss, sheet, row, map) {
  var questionsColPG = getCol_(map, ["Additional_Questions"]);
  var answersColPG   = getCol_(map, ["Additional_Answers"]);
  if (!questionsColPG || !answersColPG) return;

  var questionsVal = sheet.getRange(row, questionsColPG).getValue();
  if (questionsVal === "" || questionsVal === null) return;

  var titleColAQ = getCol_(map, ["Job_Title"]);
  var descColAQ  = getCol_(map, ["Description"]);
  var titleValAQ = titleColAQ ? sheet.getRange(row, titleColAQ).getValue() : "";
  var descValAQ  = descColAQ  ? sheet.getRange(row, descColAQ).getValue()  : "";

  sheet.getRange(row, answersColPG).setValue("Drafting answers...");
  var answersResult;
  try {
    var aqApiKey         = getApiKey_();
    var aqSettings       = getSettings_();
    var aqPortfolioCtx   = FFLib.getPortfolioContext(aqSettings);
    var aqFreelancerName = aqSettings['Freelancer_Name'] || 'the freelancer';
    answersResult = FFLib.generateAdditionalAnswers(
      questionsVal, titleValAQ, descValAQ, aqPortfolioCtx, aqFreelancerName, aqApiKey
    );
  } catch (err) {
    answersResult = err.message;
  }
  sheet.getRange(row, answersColPG).setValue(answersResult);
}

// Skip stamps Proposal_Skip_Date. Sent syncs the row into Proposal_Tracker
// (creating it once, guarded on Proposal_Sent_Date already being set) and
// feeds Connects_Helper's MTD metrics. Self-contained (reads Proposal_Status
// itself) so it can be called from handleEdit or directly from the
// Proposal_Generator sidebar.
function handleProposalStatusChange_(ss, sheet, row, map) {
  var proposalStatusCol   = getCol_(map, ["Proposal_Status"]);
  var proposalSentDateCol = getCol_(map, ["Proposal_Sent_Date"]);
  var proposalSkipDateCol = getCol_(map, ["Proposal_Skip_Date"]);
  if (!proposalStatusCol) return;

  var proposalStatus = sheet.getRange(row, proposalStatusCol).getValue();

  if (proposalStatus === "Skip" && proposalSkipDateCol) {
    var skipDateCell = sheet.getRange(row, proposalSkipDateCol);
    if (skipDateCell.getValue() === "") {
      skipDateCell.setValue(new Date());
    }
    return;
  }

  if (proposalStatus !== "Sent") return;
  if (!proposalSentDateCol) return;

  var proposalSentDateCell = sheet.getRange(row, proposalSentDateCol);
  if (proposalSentDateCell.getValue() !== "") return;

  var tracker = ss.getSheetByName("Proposal_Tracker");
  var scoring = ss.getSheetByName("Job_Scoring");

  if (!tracker || !scoring) return;

  var pgMap = map;
  var ptMap = getHeaderMap_(tracker);
  var jsMap = getHeaderMap_(scoring);

  var discoveryId        = getCellValue_(sheet, row, pgMap, ["Discovery_ID"]);
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

  // Locked because "check it doesn't exist yet, then write it" is a
  // read-then-act sequence against the shared Proposal_Tracker sheet.
  //
  // A prior version of this fix found the "first empty row" via a manual
  // full-column scan (findFirstEmptyRowByColumn_) and wrote each field with
  // a separate setCellValue_ call. Verified via added logging on 2026-07-05
  // that this still lost rows even with the lock in place: two rows marked
  // Sent within the same moment both independently computed the SAME "first
  // empty row" before either had written (confirmed via debug logs showing
  // identical target rows from two concurrent executions ~1s apart), so the
  // second write clobbered the first. appendRow() removes that race at its
  // root -- Sheets determines the true current last row on its own server
  // side at the moment of the call, instead of the script computing it in
  // advance from an earlier snapshot read, so two near-simultaneous appends
  // can't collide on the same target row the way two manual scans could.
  var ptLock = LockService.getScriptLock();
  ptLock.waitLock(30000);
  try {
    if (!existsInProposalTracker_()) {
      var ptRowValues = new Array(tracker.getLastColumn()).fill("");
      var ptSetCol = function (headerNames, value) {
        var col = getCol_(ptMap, headerNames);
        if (col) ptRowValues[col - 1] = value;
      };

      ptSetCol(["Discovery_ID"],                    discoveryId);
      ptSetCol(["Date_Applied"],                    appliedDate);
      ptSetCol(["Job_Title"],                       jobTitle);
      ptSetCol(["Client_Name", "Client Name"],      clientName);
      ptSetCol(["Keyword_Search"],                  keywordSearch);
      ptSetCol(["Tool_Requested", "Tool_Detected"], toolRequested);
      ptSetCol(["Days_Since_Posted"],               daysSincePosted);
      ptSetCol(["Proposal_Count"],                  proposalCount);
      ptSetCol(["Total_Score"],                     totalScore);
      ptSetCol(["Template_Used"],                   templateUsed);
      ptSetCol(["Hook_Version"],                    hookVersion);
      ptSetCol(["CTA_Version"],                     ctaVersion);
      ptSetCol(["Viewed"],                          "N");
      ptSetCol(["Interview"],                       "N");
      ptSetCol(["Hired"],                           "N");
      ptSetCol(["Revenue"],                         "");
      ptSetCol(["Notes"],                           notes);
      ptSetCol(["Age_Days"],                        ageDays);
      ptSetCol(["Current_Age_Days"],                currentAgeDays);
      ptSetCol(["Connects_Used"],                   connectsUsed);
      ptSetCol(["Boost_Connects"],                  boostConnects !== "" ? boostConnects : 0);
      var totalForCost = connectsUsed !== "" ? Number(connectsUsed) : 0;
      ptSetCol(["Proposal_Cost"],                   totalForCost > 0 ? "$" + (totalForCost * 0.15).toFixed(2) : "");
      ptSetCol(["Job_Link"],                        jobLink);

      tracker.appendRow(ptRowValues);
      // Without this, the next concurrent execution's existsInProposalTracker_()
      // check can run before this appended row is actually committed and read
      // it as still not-existing -- same flush-before-release reasoning as
      // incrementConnectsHelperMetric_ below.
      SpreadsheetApp.flush();

      // Connects_Helper's MTD/Total metrics only ever move here, at the
      // moment a fresh Proposal_Tracker row is created -- this whole "Sent"
      // path is already guarded to run once per row, so no double count.
      incrementConnectsHelperMetric_(ss, "MTD_Proposals_Sent", 1);
      incrementConnectsHelperMetric_(ss, "MTD_Connects_Used",  totalForCost);
      incrementConnectsHelperMetric_(ss, "Total_Connects_Used", totalForCost);
      incrementConnectsHelperMetric_(ss, "Current_Connect_Balance", -totalForCost);
      incrementConnectsHelperMetric_(ss, "Total_Proposal_Cost", totalForCost > 0 ? totalForCost * 0.15 : 0);
    }
  } finally {
    ptLock.releaseLock();
  }

  proposalSentDateCell.setValue(sentDate);

  showWalkthroughOnce_(
    "FF_WALKTHROUGH_PROPOSAL_SENT_SEEN",
    "First proposal sent!",
    "This job now shows up in Proposal_Tracker so you can track what happens next.\n\n" +
    "When the client views your proposal, mark Viewed. When they reply, mark Interview -- " +
    "that automatically opens a sidebar to import the chat. When you're hired, mark Hired -- " +
    "that automatically opens a sidebar to log the contract."
  );
}

// Runs Milestone_Tracker's Status-transition side effects (date stamps,
// revenue, Contract_Tracker auto-complete) -- assumes the Status cell has
// already been written with newVal by the caller. Self-contained so it can
// be called from handleEdit's per-column check above or directly from the
// Contract Progress sidebar.
function applyMilestoneStatusEffects_(ss, sheet, row, map, oldVal, newVal) {
  if (newVal === "Funded" && oldVal !== "Funded") {
    var msFundedCol = getCol_(map, ["Funded_Date"]);
    if (msFundedCol && sheet.getRange(row, msFundedCol).getValue() === "") {
      sheet.getRange(row, msFundedCol).setValue(new Date());
    }
  }
  if (newVal === "Delivered" && oldVal !== "Delivered") {
    var msDeliveredCol = getCol_(map, ["Delivered_Date"]);
    if (msDeliveredCol && sheet.getRange(row, msDeliveredCol).getValue() === "") {
      sheet.getRange(row, msDeliveredCol).setValue(new Date());
    }
  }
  if (newVal === "Released" && oldVal !== "Released") {
    var msReleasedCol = getCol_(map, ["Released_Date"]);
    if (msReleasedCol && sheet.getRange(row, msReleasedCol).getValue() === "") {
      sheet.getRange(row, msReleasedCol).setValue(new Date());
    }
    var msAmountCol = getCol_(map, ["Amount"]);
    if (msAmountCol) {
      var msAmountVal = Number(sheet.getRange(row, msAmountCol).getValue()) || 0;
      incrementConnectsHelperMetric_(ss, "MTD_Revenue", msAmountVal);
      incrementConnectsHelperMetric_(ss, "Monthly_Revenue", msAmountVal);
    }

    // Once every milestone belonging to this contract is Released, the
    // contract itself is done -- auto-complete it instead of leaving
    // Contract_Tracker's Status as a separate manual step.
    var msIdColForComplete = getCol_(map, ["Discovery_ID"]);
    var msDiscoveryId = msIdColForComplete ? sheet.getRange(row, msIdColForComplete).getValue() : "";
    if (msDiscoveryId && allMilestonesReleased_(ss, msDiscoveryId)) {
      autoCompleteContract_(ss, msDiscoveryId);
    }
  }
}

// Computes Amount = Hours_Logged x this contract's Hourly_Rate for one
// Hourly_Log row and feeds the delta into Connects_Helper's revenue metrics.
// Self-contained (reads Hours_Logged/Discovery_ID itself) so it can be
// called from handleEdit's per-row loop above or directly from the Contract
// Progress sidebar after appending a new row.
function applyHourlyLogAmount_(ss, sheet, row, map) {
  var hlHoursCol  = getCol_(map, ["Hours_Logged"]);
  var hlIdCol     = getCol_(map, ["Discovery_ID"]);
  var hlAmountCol = getCol_(map, ["Amount"]);
  if (!hlHoursCol || !hlIdCol || !hlAmountCol) return;

  var hlHoursVal    = Number(sheet.getRange(row, hlHoursCol).getValue()) || 0;
  var hlDiscoveryId = sheet.getRange(row, hlIdCol).getValue();
  if (!hlDiscoveryId) return;

  var hlRate      = getContractHourlyRate_(ss, hlDiscoveryId);
  var hlNewAmount = hlHoursVal * hlRate;
  var hlOldAmount = Number(sheet.getRange(row, hlAmountCol).getValue()) || 0;
  var hlDelta     = hlNewAmount - hlOldAmount;

  if (hlNewAmount !== hlOldAmount) {
    sheet.getRange(row, hlAmountCol).setValue(hlNewAmount);
  }
  if (hlDelta !== 0) {
    incrementConnectsHelperMetric_(ss, "MTD_Revenue", hlDelta);
    incrementConnectsHelperMetric_(ss, "Monthly_Revenue", hlDelta);
  }
}

// Propagates an Hourly_Log row's Status to that contract's Contract_Tracker
// Status -- the Hourly equivalent of Milestone_Tracker driving auto-complete,
// since Hourly contracts have no milestones to signal "done" off of. Self-
// contained (reads Discovery_ID itself) so it can be called from handleEdit's
// per-column check above or directly from the Contract Progress sidebar.
function applyHourlyLogStatusEffects_(ss, sheet, row, map, oldVal, newVal) {
  if (String(oldVal) === String(newVal)) return;

  var idCol = getCol_(map, ["Discovery_ID"]);
  var discoveryId = idCol ? sheet.getRange(row, idCol).getValue() : "";
  if (!discoveryId) return;

  if (newVal === "Completed") {
    autoCompleteContract_(ss, discoveryId);
    return;
  }

  var contracts = ss.getSheetByName("Contract_Tracker");
  if (!contracts || contracts.getLastRow() < 2) return;

  var ctMap       = getHeaderMap_(contracts);
  var ctIdCol     = getCol_(ctMap, ["Discovery_ID"]);
  var ctStatusCol = getCol_(ctMap, ["Status"]);
  if (!ctIdCol || !ctStatusCol) return;

  var ctIdValues = contracts.getRange(2, ctIdCol, contracts.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ctIdValues.length; i++) {
    if (String(ctIdValues[i][0]) === String(discoveryId)) {
      contracts.getRange(i + 2, ctStatusCol).setValue(newVal);
      return;
    }
  }
}
