/**
 * ============================================================
 * 18. KEYWORD STRATEGY BUILDER
 * Generates an AI Primary/Secondary/Negative keyword strategy
 * from the freelancer's niche + portfolio (already on file in
 * Settings) and writes it into Keyword_Search_List (so it's
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

  var portfolioMap   = getPortfolioMapFromSettings_();
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
}
