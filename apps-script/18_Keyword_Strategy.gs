/**
 * ============================================================
 * 18. KEYWORD STRATEGY BUILDER
 * Generates an AI Primary/Secondary/Negative keyword strategy
 * from the freelancer's niche + portfolio (already on file in
 * the Projects sheet) and writes it into Keyword_Search_List (so it's
 * ready to search Upwork with) and Keyword_Strategy (so it's
 * tracked against a target job count per keyword).
 *
 * The AI call + Recommended_Action formula builder both live in
 * the Apps Script Library (FFLib.generateKeywordStrategy,
 * FFLib.buildKeywordStrategyFormula). This file is the UI + sheet
 * I/O wrapper.
 * ============================================================
 */
function GENERATE_KEYWORD_STRATEGY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var DEFAULT_TARGET_COUNT = 15;

  var apiKey;
  try {
    apiKey = getApiKey_();
  } catch (err) {
    ui.alert(err.message);
    return;
  }

  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var strategySheet   = ss.getSheetByName("Keyword_Strategy");

  if (!searchListSheet || !strategySheet) {
    ui.alert("Missing sheet. Confirm Keyword_Search_List and Keyword_Strategy both exist.");
    return;
  }

  var settings = getSettings_();
  var niche    = settings["Freelancer_Background"] || "";

  var portfolioMap   = getPortfolioMapFromProjects_();
  var portfolioParts = Object.keys(portfolioMap).sort().map(function (n) {
    var proj = portfolioMap[n];
    if (!proj.name) return "";
    var kw = proj.keywords && proj.keywords.length ? proj.keywords.join(", ") : "";
    return proj.name + (kw ? " (keywords: " + kw + ")" : "");
  }).filter(function (s) { return s !== ""; });

  var portfolioSummary = portfolioParts.join("; ");

  var result = FFLib.generateKeywordStrategy(niche, portfolioSummary, apiKey);
  if (!result.ok) {
    ui.alert("Error: " + result.message);
    return;
  }

  if (result.primary.length === 0 && result.secondary.length === 0 && result.negative.length === 0) {
    ui.alert("AI returned no keyword suggestions. Try again, or add keywords manually to Keyword_Search_List.");
    return;
  }

  var searchListStart = getLastRealRow_(searchListSheet) + 1;
  var searchRows = result.primary.concat(result.secondary).map(function (kw) {
    return [kw, "", "", kw];
  });
  if (searchRows.length > 0) {
    searchListSheet.getRange(searchListStart, 1, searchRows.length, 4).setValues(searchRows);
  }

  var stratMap        = getHeaderMap_(strategySheet);
  var stratKeywordCol = getCol_(stratMap, ["Keyword"]);
  var stratActualCol  = getCol_(stratMap, ["Actual_Count"]);
  var stratTargetCol  = getCol_(stratMap, ["Target_Count"]);
  var stratNotesCol   = getCol_(stratMap, ["Notes"]);
  var stratHeaders    = strategySheet.getRange(1, 1, 1, strategySheet.getLastColumn()).getValues()[0];

  var tiers = [
    { keywords: result.primary,   target: DEFAULT_TARGET_COUNT, note: "Primary tier" },
    { keywords: result.secondary, target: DEFAULT_TARGET_COUNT, note: "Secondary tier" },
    { keywords: result.negative,  target: 0,                    note: "Negative -- avoid" }
  ];

  var writeRow = getLastRealRow_(strategySheet) + 1;

  tiers.forEach(function (tier) {
    tier.keywords.forEach(function (kw) {
      if (stratKeywordCol) strategySheet.getRange(writeRow, stratKeywordCol).setValue(kw);
      if (stratActualCol)  strategySheet.getRange(writeRow, stratActualCol).setValue(0);
      if (stratTargetCol)  strategySheet.getRange(writeRow, stratTargetCol).setValue(tier.target);
      if (stratNotesCol)   strategySheet.getRange(writeRow, stratNotesCol).setValue(tier.note);

      var formula = FFLib.buildKeywordStrategyFormula(stratHeaders, writeRow);
      if (formula.recommendedActionFormula) {
        strategySheet.getRange(writeRow, formula.recommendedActionCol).setFormula(formula.recommendedActionFormula);
      }

      writeRow++;
    });
  });

  var summary = "Keyword strategy generated.\n\n" +
    "Primary: " + result.primary.length + "\n" +
    "Secondary: " + result.secondary.length + "\n" +
    "Negative (avoid): " + result.negative.length + "\n\n" +
    "Primary + Secondary added to Keyword_Search_List -- start searching Upwork with those.\n" +
    "All three tracked in Keyword_Strategy against a target of " + DEFAULT_TARGET_COUNT + " jobs each.";

  ui.alert(summary);

  showTourStep_(
    'FF_TOUR_STEP1_SESSIONS_SEEN',
    'Keyword Strategy Ready',
    'Keyword_Strategy and Keyword_Search_List are ready to use.\n\n' +
    'Use System Tools > Start Session before you start searching Upwork, and End Session when you\'re done -- that logs the session\'s results.'
  );
  showTourStep_(
    'FF_TOUR_STEP2_MINING_SEEN',
    'Keyword Mining Strategy',
    'Keyword_Strategy tracks which keywords are working:\n\n' +
    '- Generate Keyword Strategy creates new keyword ideas (what you just ran)\n' +
    '- Mine Keywords scans your logged jobs for new keyword candidates\n' +
    '- Any keyword marked "Avoid" in Recommended_Action is underperforming -- set its Drop column to "Drop" and run Drop Keywords to stop searching it.'
  );
  showTourStep_(
    'FF_TOUR_STEP1B_MANAGE_PROJECTS_SEEN',
    'Manage Your Portfolio Projects',
    'The keyword strategy above pulled from your Projects sheet -- Upwork allows up to 10 portfolio projects. Use System Tools > Manage Projects anytime to add or edit one; Proposal_Generator\'s AI matching picks it up automatically, no repair needed.'
  );
  showTourStep_(
    'FF_TOUR_STEP1C_DASHBOARD_SEEN',
    'Build Your Dashboard',
    'Once you\'ve logged a few jobs and sessions, run System Tools > Build/Refresh Dashboard for a one-page snapshot: KPI tiles, pipeline funnel, and revenue trend with charts. It only updates when you run it, not live.'
  );
}

// Removes Keyword_Search_List rows whose Search_Query matches a keyword
// marked "Drop" in Keyword_Strategy's Drop column (case/whitespace-insensitive
// match, since that's how the keyword text was originally written into both
// sheets). Shared by DROP_KEYWORDS() below and MINE_KEYWORDS() (08_Keyword_
// Mining.gs), which purges before mining so a dropped keyword doesn't get
// re-suggested. Keyword_Strategy itself is never touched -- the Drop marker
// and Recommended_Action history stay exactly as they are.
function purgeDroppedKeywords_(ss) {
  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var strategySheet   = ss.getSheetByName("Keyword_Strategy");
  if (!searchListSheet || !strategySheet) return 0;

  var stratLastRow = strategySheet.getLastRow();
  if (stratLastRow <= 1) return 0;

  var stratMap = getHeaderMap_(strategySheet);
  var kwCol    = getCol_(stratMap, ["Keyword"]);
  var dropCol  = getCol_(stratMap, ["Drop"]);
  if (!kwCol || !dropCol) return 0;

  var stratData = strategySheet
    .getRange(2, 1, stratLastRow - 1, strategySheet.getLastColumn())
    .getValues();

  var dropSet = {};
  for (var i = 0; i < stratData.length; i++) {
    var kw   = String(stratData[i][kwCol - 1]).trim().toLowerCase();
    var drop = String(stratData[i][dropCol - 1]).trim();
    if (drop === "Drop" && kw !== "") dropSet[kw] = true;
  }

  if (Object.keys(dropSet).length === 0) return 0;

  var slMap      = getHeaderMap_(searchListSheet);
  var slQueryCol = getCol_(slMap, ["Search_Query"]);
  if (!slQueryCol) return 0;

  var purgedCount = 0;
  var slLastRow   = searchListSheet.getLastRow();
  for (var r = slLastRow; r >= 2; r--) {
    var cellQuery = String(searchListSheet.getRange(r, slQueryCol).getValue()).trim().toLowerCase();
    if (dropSet[cellQuery]) {
      searchListSheet.deleteRow(r);
      purgedCount++;
    }
  }
  return purgedCount;
}

function DROP_KEYWORDS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var strategySheet   = ss.getSheetByName("Keyword_Strategy");
  if (!searchListSheet || !strategySheet) {
    ui.alert("Missing sheet. Confirm Keyword_Search_List and Keyword_Strategy both exist.");
    return;
  }

  var purgedCount = purgeDroppedKeywords_(ss);

  if (purgedCount === 0) {
    ui.alert(
      "No keywords marked \"Drop\" in Keyword_Strategy were found in Keyword_Search_List.\n\n" +
      "Set a keyword's Drop column to \"Drop\" in Keyword_Strategy, then run this again."
    );
    return;
  }

  ui.alert("Done.\n\n✓ " + purgedCount + " row(s) removed from Keyword_Search_List.");
}
