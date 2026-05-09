/**
 * ============================================================
 * 12. SESSION MANAGEMENT
 * START_SESSION: prompts for Session ID and keywords, stores
 *   state in PropertiesService, confirms session target.
 * END_SESSION: computes all session stats, writes to Session_Log,
 *   updates Keyword_Search_List last-searched and yield.
 * ============================================================
 */
function START_SESSION() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var prop = PropertiesService.getScriptProperties();

  if (prop.getProperty("SESSION_ACTIVE") === "true") {
    ui.alert(
      "A session is already active (Session " + prop.getProperty("SESSION_ID") + ").\n" +
      "Use System Tools > End Session to close it before starting a new one."
    );
    return;
  }

  var idResponse = ui.prompt(
    "Start Session -- Step 1 of 2",
    "Enter your Session ID (e.g. S001):",
    ui.ButtonSet.OK_CANCEL
  );
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;

  var sessionId = idResponse.getResponseText().trim().toUpperCase();
  if (!sessionId) {
    ui.alert("Session ID cannot be blank.");
    return;
  }

  var kwResponse = ui.prompt(
    "Start Session -- Step 2 of 2",
    "Which keywords are you searching this session?\n" +
    "(Enter comma-separated, e.g. Excel Finance Dashboard Creation, Power BI Sales Dashboard)",
    ui.ButtonSet.OK_CANCEL
  );
  if (kwResponse.getSelectedButton() !== ui.Button.OK) return;

  var keywords = kwResponse.getResponseText().trim();
  if (!keywords) {
    ui.alert("Please enter at least one keyword.");
    return;
  }

  var discoverySheet = ss.getSheetByName("Job_Discovery");
  var startRowCount  = discoverySheet ? Math.max(discoverySheet.getLastRow() - 1, 0) : 0;

  prop.setProperties({
    "SESSION_ACTIVE":          "true",
    "SESSION_ID":              sessionId,
    "SESSION_START_TIME":      new Date().toISOString(),
    "SESSION_KEYWORDS":        keywords,
    "SESSION_START_ROW_COUNT": String(startRowCount),
    "SESSION_DUPE_COUNT":      "0"
  });

  ui.alert(
    "✓ Session " + sessionId + " started.\n\n" +
    "Keywords: " + keywords + "\n" +
    "Start time: " + formatTime_(new Date()) + "\n\n" +
    "Session target: 8 unique new jobs.\n" +
    "Go search Upwork -- every job you log will be tracked automatically."
  );
}


function END_SESSION() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var prop = PropertiesService.getScriptProperties();

  if (prop.getProperty("SESSION_ACTIVE") !== "true") {
    ui.alert("No active session found.\nUse System Tools > Start Session to begin one.");
    return;
  }

  var sessionId    = prop.getProperty("SESSION_ID")           || "";
  var startTimeStr = prop.getProperty("SESSION_START_TIME")   || "";
  var keywords     = prop.getProperty("SESSION_KEYWORDS")     || "";
  var startCount   = parseInt(prop.getProperty("SESSION_START_ROW_COUNT") || "0", 10);
  var dupeCount    = parseInt(prop.getProperty("SESSION_DUPE_COUNT")      || "0", 10);

  var endTime    = new Date();
  var startTime  = startTimeStr ? new Date(startTimeStr) : endTime;
  var durationMs = endTime - startTime;

  var discoverySheet = ss.getSheetByName("Job_Discovery");
  var jobsLogged     = 0;
  var movedToScoring = 0;
  var reviewLater    = 0;

  if (discoverySheet && discoverySheet.getLastRow() > 1) {
    var discMap      = getHeaderMap_(discoverySheet);
    var sessionIdCol = getCol_(discMap, ["Session_ID"]);
    var actionCol    = getCol_(discMap, ["Discovery_Action"]);
    var totalRows    = discoverySheet.getLastRow() - 1;

    if (sessionIdCol && actionCol && totalRows > 0) {
      var discData = discoverySheet
        .getRange(2, 1, totalRows, discoverySheet.getLastColumn())
        .getValues();

      for (var i = 0; i < discData.length; i++) {
        var rowSession = String(discData[i][sessionIdCol - 1]).trim().toUpperCase();
        var rowAction  = String(discData[i][actionCol   - 1]).trim();

        if (rowSession === sessionId) {
          jobsLogged++;
          if (rowAction === "Move to Scoring") movedToScoring++;
          else if (rowAction === "Review Later") reviewLater++;
        }
      }
    }
  }

  var sessionYield = jobsLogged;
  var YIELD_TARGET = 8;
  var saturating   = sessionYield < YIELD_TARGET;
  var satFlag      = saturating
    ? "YES -- yield " + sessionYield + "/" + YIELD_TARGET + " (below target)"
    : "No -- yield " + sessionYield + "/" + YIELD_TARGET;

  var scoringSheet    = ss.getSheetByName("Job_Scoring");
  var applyNoProposal = 0;

  if (scoringSheet && scoringSheet.getLastRow() > 1) {
    var jsMap         = getHeaderMap_(scoringSheet);
    var jsDecisionCol = getCol_(jsMap, ["Final_Decision"]);
    var jsPropDateCol = getCol_(jsMap, ["Proposal_Generator_Date"]);

    if (jsDecisionCol && jsPropDateCol) {
      var jsData = scoringSheet
        .getRange(2, 1, scoringSheet.getLastRow() - 1, scoringSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < jsData.length; i++) {
        var decision = String(jsData[i][jsDecisionCol - 1]).trim();
        var propDate = jsData[i][jsPropDateCol - 1];
        if (decision === "APPLY" && (propDate === "" || propDate === null)) {
          applyNoProposal++;
        }
      }
    }
  }

  var proposalTrigger = applyNoProposal >= 5
    ? "⚠ YES -- " + applyNoProposal + " APPLY jobs have no proposal sent. Send at least 2 before next session."
    : "No -- " + applyNoProposal + " APPLY jobs pending (" + (5 - applyNoProposal) + " more needed to trigger).";

  var pgSheet          = ss.getSheetByName("Proposal_Generator");
  var proposalsSent    = 0;
  var proposalsSkipped = 0;
  var connectsSpent    = 0;

  if (pgSheet && pgSheet.getLastRow() > 1) {
    var pgMap         = getHeaderMap_(pgSheet);
    var pgStatusCol   = getCol_(pgMap, ["Proposal_Status"]);
    var pgSentDateCol = getCol_(pgMap, ["Proposal_Sent_Date"]);
    var pgConnectsCol = getCol_(pgMap, ["Total_Connects_Spent", "Connects_Required"]);

    if (pgStatusCol && pgSentDateCol) {
      var pgData = pgSheet
        .getRange(2, 1, pgSheet.getLastRow() - 1, pgSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < pgData.length; i++) {
        var pgStatus   = String(pgData[i][pgStatusCol   - 1]).trim();
        var pgSentDate = pgData[i][pgSentDateCol - 1];

        if (pgSentDate) {
          var sentDateObj = new Date(pgSentDate);
          if (sentDateObj >= startTime && sentDateObj <= endTime) {
            if (pgStatus === "Sent") {
              proposalsSent++;
              if (pgConnectsCol) {
                connectsSpent += Number(pgData[i][pgConnectsCol - 1]) || 0;
              }
            } else if (pgStatus === "Skip") {
              proposalsSkipped++;
            }
          }
        }
      }
    }
  }

  var notesResponse = ui.prompt(
    "End Session -- Notes",
    "Any notes for this session? (optional -- press OK to skip)",
    ui.ButtonSet.OK_CANCEL
  );
  var sessionNotes = (notesResponse.getSelectedButton() === ui.Button.OK)
    ? notesResponse.getResponseText().trim()
    : "";

  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  if (searchListSheet && keywords) {
    var slMap           = getHeaderMap_(searchListSheet);
    var slQueryCol      = getCol_(slMap, ["Search_Query"]);
    var slLastSearchCol = getCol_(slMap, ["Last_Searched"]);
    var slYieldCol      = getCol_(slMap, ["Session_Yield"]);
    var slLastRow       = searchListSheet.getLastRow();

    var keywordList = keywords.split(",").map(function (k) {
      return k.trim().toLowerCase();
    });

    if (slQueryCol && slLastRow > 1) {
      var slData = searchListSheet
        .getRange(2, 1, slLastRow - 1, searchListSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < slData.length; i++) {
        var rowQuery = String(slData[i][slQueryCol - 1]).trim().toLowerCase();
        if (keywordList.indexOf(rowQuery) !== -1) {
          var dataRow = i + 2;
          if (slLastSearchCol) {
            searchListSheet.getRange(dataRow, slLastSearchCol).setValue(endTime);
          }
          if (slYieldCol) {
            var perKeyword = Math.round(sessionYield / keywordList.length);
            searchListSheet.getRange(dataRow, slYieldCol).setValue(perKeyword);
          }
        }
      }
    }
  }

  var logSheet = ss.getSheetByName("Session_Log");
  if (logSheet) {
    var logMap     = getHeaderMap_(logSheet);
    var nextLogRow = findFirstEmptyRowByColumn_(logSheet, 1);
    if (logSheet.getLastRow() <= 1) nextLogRow = 2;

    setCellValue_(logSheet, nextLogRow, logMap, ["Session_ID"],            sessionId);
    setCellValue_(logSheet, nextLogRow, logMap, ["Date"],                  endTime);
    setCellValue_(logSheet, nextLogRow, logMap, ["Start_Time"],            formatTime_(startTime));
    setCellValue_(logSheet, nextLogRow, logMap, ["End_Time"],              formatTime_(endTime));
    setCellValue_(logSheet, nextLogRow, logMap, ["Duration"],              formatDuration_(durationMs));
    setCellValue_(logSheet, nextLogRow, logMap, ["Keywords_Searched"],     keywords);
    setCellValue_(logSheet, nextLogRow, logMap, ["Jobs_Logged"],           jobsLogged);
    setCellValue_(logSheet, nextLogRow, logMap, ["Jobs_Moved_To_Scoring"], movedToScoring);
    setCellValue_(logSheet, nextLogRow, logMap, ["Jobs_Review_Later"],     reviewLater);
    setCellValue_(logSheet, nextLogRow, logMap, ["Duplicates_Skipped"],    dupeCount);
    setCellValue_(logSheet, nextLogRow, logMap, ["Session_Yield"],         sessionYield);
    setCellValue_(logSheet, nextLogRow, logMap, ["Saturation_Flag"],       saturating ? "Saturating" : "Healthy");
    setCellValue_(logSheet, nextLogRow, logMap, ["Proposal_Trigger"],      applyNoProposal >= 5 ? "TRIGGERED" : "Clear");
    setCellValue_(logSheet, nextLogRow, logMap, ["Proposals_Sent"],        proposalsSent);
    setCellValue_(logSheet, nextLogRow, logMap, ["Proposals_Skipped"],     proposalsSkipped);
    setCellValue_(logSheet, nextLogRow, logMap, ["Connects_Spent"],        connectsSpent);
    setCellValue_(logSheet, nextLogRow, logMap, ["Notes"],                 sessionNotes);
  }

  prop.deleteAllProperties();

  var runLog =
    "════════════════════════════════\n" +
    "  SESSION RUN LOG -- " + sessionId + "\n" +
    "════════════════════════════════\n\n" +
    "Date:           " + endTime.toLocaleDateString() + "\n" +
    "Start:          " + formatTime_(startTime) + "\n" +
    "End:            " + formatTime_(endTime) + "\n" +
    "Duration:       " + formatDuration_(durationMs) + "\n\n" +
    "Keywords:\n  " + keywords.split(",").join("\n  ") + "\n\n" +
    "-- Discovery ──────────────────\n" +
    "Jobs logged:        " + jobsLogged + "\n" +
    "Moved to Scoring:   " + movedToScoring + "\n" +
    "Review Later:       " + reviewLater + "\n" +
    "Duplicates skipped: " + dupeCount + "\n" +
    "Session yield:      " + sessionYield + " / " + YIELD_TARGET + " target\n\n" +
    "-- Health ─────────────────────\n" +
    "Saturation flag:    " + satFlag + "\n" +
    "Proposal trigger:   " + proposalTrigger + "\n\n" +
    "-- Proposals ──────────────────\n" +
    "Sent this session:    " + proposalsSent + "\n" +
    "Skipped this session: " + proposalsSkipped + "\n" +
    (connectsSpent > 0 ? "Connects spent:     " + connectsSpent + "\n" : "") +
    "\n" +
    (sessionNotes ? "Notes: " + sessionNotes + "\n\n" : "") +
    (logSheet ? "✓ Log written to Session_Log." : "⚠ Session_Log sheet not found -- log not saved.");

  ui.alert(runLog);
}
