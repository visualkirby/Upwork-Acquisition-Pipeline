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
// Shared with ensureKeywordStrategyRow_ below, so a keyword synced in later
// (typed directly into Keyword_Search_List, or mined) tracks against the
// same default as the AI-generated tiers.
var DEFAULT_KEYWORD_TARGET_COUNT_ = 15;

function GENERATE_KEYWORD_STRATEGY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var DEFAULT_TARGET_COUNT = DEFAULT_KEYWORD_TARGET_COUNT_;

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

      // Covers a keyword that already has matching Job_Discovery rows before
      // this AI-generated strategy was built (e.g. re-running the generator).
      if (stratActualCol) recomputeKeywordStrategyActualCount_(ss, kw);

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
// sheets), AND deletes that keyword's own row from Keyword_Strategy -- Drop
// is a full removal from both sheets, not just a Keyword_Search_List purge.
// Its Actual_Count/Notes/Recommended_Action history is discarded, but
// nothing is actually lost: if the same keyword is ever re-added later
// (Generate Keyword Strategy, Mine Keywords, Scale Keywords), its row gets
// rebuilt fresh with Actual_Count recomputed straight from Job_Discovery --
// the real source of truth -- same as any other newly-added keyword.
//
// Shared by DROP_KEYWORDS() below, MINE_KEYWORDS() (08_Keyword_Mining.gs),
// and SCALE_KEYWORDS() above, all of which purge before doing their own work
// so a dropped keyword doesn't get re-suggested or re-scaled.
//
// Returns { searchListRemoved, strategyRemoved } -- separate counts since
// callers report on both sheets.
function purgeDroppedKeywords_(ss) {
  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var strategySheet   = ss.getSheetByName("Keyword_Strategy");
  var empty = { searchListRemoved: 0, strategyRemoved: 0 };
  if (!searchListSheet || !strategySheet) return empty;

  var stratLastRow = strategySheet.getLastRow();
  if (stratLastRow <= 1) return empty;

  var stratMap = getHeaderMap_(strategySheet);
  var kwCol    = getCol_(stratMap, ["Keyword"]);
  var dropCol  = getCol_(stratMap, ["Drop"]);
  if (!kwCol || !dropCol) return empty;

  var stratData = strategySheet
    .getRange(2, 1, stratLastRow - 1, strategySheet.getLastColumn())
    .getValues();

  var dropSet  = {};
  var dropRows = []; // Keyword_Strategy sheet row numbers to delete
  for (var i = 0; i < stratData.length; i++) {
    var kw   = String(stratData[i][kwCol - 1]).trim().toLowerCase();
    var drop = String(stratData[i][dropCol - 1]).trim();
    if (drop === "Drop" && kw !== "") {
      dropSet[kw] = true;
      dropRows.push(i + 2);
    }
  }

  if (Object.keys(dropSet).length === 0) return empty;

  var searchListRemoved = 0;
  var slMap      = getHeaderMap_(searchListSheet);
  var slQueryCol = getCol_(slMap, ["Search_Query"]);
  if (slQueryCol) {
    var slLastRow = searchListSheet.getLastRow();
    for (var r = slLastRow; r >= 2; r--) {
      var cellQuery = String(searchListSheet.getRange(r, slQueryCol).getValue()).trim().toLowerCase();
      if (dropSet[cellQuery]) {
        searchListSheet.deleteRow(r);
        searchListRemoved++;
      }
    }
  }

  // Delete bottom-up so earlier row numbers in dropRows stay valid.
  dropRows.sort(function (a, b) { return b - a; });
  dropRows.forEach(function (row) { strategySheet.deleteRow(row); });

  return { searchListRemoved: searchListRemoved, strategyRemoved: dropRows.length };
}

// Recomputes Keyword_Strategy's Actual_Count for the matching keyword by
// counting how many Job_Discovery rows currently carry that Keyword_Search
// value -- same case/whitespace-insensitive match, and the same counting
// logic Keyword_Intelligence's Total_Jobs uses (26_Keyword_Intelligence.gs's
// computeKeywordIntelligenceRow_), so the two always agree.
//
// Replaces an earlier running "+1 on first log" counter, which couldn't
// recover from a resubmitted/duplicated row, a paste that misfired across
// multiple rows, or Keyword_Search being corrected after Description was
// already logged -- all three happened in the same real session and left
// Actual_Count permanently wrong for two keywords. Recomputing straight from
// Job_Discovery (the source of truth) instead of incrementing is self-healing
// against all of those, and against Job_Discovery rows being deleted, since
// there's nothing to "undo".
//
// Called whenever a Job_Discovery row's Keyword_Search takes on a value --
// from the direct-paste edit trigger (14_Edit_Trigger.gs's JOB_DISCOVERY
// block, both on first Description log and on any later Keyword_Search edit)
// and from the Log New Job sidebar (job_saveEntry in
// 21_Job_Discovery_Sidebar.gs). Silent no-op if the keyword doesn't match any
// row in Keyword_Strategy (e.g. a one-off manual search outside the generated
// strategy) -- Recommended_Action only tracks keywords that are actually in
// the strategy.
function recomputeKeywordStrategyActualCount_(ss, keyword) {
  if (!keyword) return;
  var strategySheet = ss.getSheetByName("Keyword_Strategy");
  var jdSheet        = ss.getSheetByName("Job_Discovery");
  if (!strategySheet || !jdSheet) return;

  var stratMap  = getHeaderMap_(strategySheet);
  var kwCol     = getCol_(stratMap, ["Keyword"]);
  var actualCol = getCol_(stratMap, ["Actual_Count"]);
  if (!kwCol || !actualCol) return;

  var stratLastRow = strategySheet.getLastRow();
  if (stratLastRow < 2) return;

  var keywordStr = String(keyword).trim().toLowerCase();
  var kwValues    = strategySheet.getRange(2, kwCol, stratLastRow - 1, 1).getValues();

  var stratRow = -1;
  for (var i = 0; i < kwValues.length; i++) {
    if (String(kwValues[i][0]).trim().toLowerCase() === keywordStr) {
      stratRow = i + 2;
      break;
    }
  }
  if (stratRow === -1) return;

  var jdMap   = getHeaderMap_(jdSheet);
  var jdKwCol = getCol_(jdMap, ["Keyword_Search"]);
  if (!jdKwCol) return;

  var jdLastRow = jdSheet.getLastRow();
  var count = 0;
  if (jdLastRow >= 2) {
    var jdValues = jdSheet.getRange(2, jdKwCol, jdLastRow - 1, 1).getValues();
    for (var j = 0; j < jdValues.length; j++) {
      if (String(jdValues[j][0]).trim().toLowerCase() === keywordStr) count++;
    }
  }

  strategySheet.getRange(stratRow, actualCol).setValue(count);
}

// Ensures a Keyword_Strategy row exists for the given keyword text, creating
// one (Actual_Count recomputed from Job_Discovery, Target_Count
// DEFAULT_KEYWORD_TARGET_COUNT_, the same Recommended_Action formula
// GENERATE_KEYWORD_STRATEGY writes) if no match is found. Idempotent -- safe
// to call on every edit or a full backfill pass without duplicating rows,
// since it checks for an existing match first. Case/whitespace-insensitive,
// same matching style as recomputeKeywordStrategyActualCount_ above.
//
// Called from 14_Edit_Trigger.gs's KEYWORD_SEARCH_LIST block (fires once
// Search_Query is filled in, whichever way it got there -- typed directly,
// mined, or otherwise -- Keyword_Search_List has no single "how it got
// added" gate, so this doesn't assume one either) and from REPAIR_FORMULAS()
// (15_Formula_Fixes.gs) to backfill rows that predate this sync. Returns
// true if a new row was created, false if a match already existed or the
// keyword/sheet was invalid -- callers use this to report an accurate
// backfill count.
function ensureKeywordStrategyRow_(ss, keyword) {
  if (!keyword) return false;
  var strategySheet = ss.getSheetByName("Keyword_Strategy");
  if (!strategySheet) return false;

  var stratMap = getHeaderMap_(strategySheet);
  var kwCol    = getCol_(stratMap, ["Keyword"]);
  if (!kwCol) return false;

  var keywordStr = String(keyword).trim().toLowerCase();
  if (!keywordStr) return false;

  var lastRow = strategySheet.getLastRow();
  if (lastRow >= 2) {
    var kwValues = strategySheet.getRange(2, kwCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < kwValues.length; i++) {
      if (String(kwValues[i][0]).trim().toLowerCase() === keywordStr) return false;
    }
  }

  var actualCol    = getCol_(stratMap, ["Actual_Count"]);
  var targetCol    = getCol_(stratMap, ["Target_Count"]);
  var notesCol     = getCol_(stratMap, ["Notes"]);
  var stratHeaders = strategySheet.getRange(1, 1, 1, strategySheet.getLastColumn()).getValues()[0];

  var writeRow = getLastRealRow_(strategySheet) + 1;
  strategySheet.getRange(writeRow, kwCol).setValue(keyword);
  if (actualCol) strategySheet.getRange(writeRow, actualCol).setValue(0);
  if (targetCol) strategySheet.getRange(writeRow, targetCol).setValue(DEFAULT_KEYWORD_TARGET_COUNT_);
  if (notesCol)  strategySheet.getRange(writeRow, notesCol).setValue("Synced from Keyword_Search_List");

  var formula = FFLib.buildKeywordStrategyFormula(stratHeaders, writeRow);
  if (formula.recommendedActionFormula) {
    strategySheet.getRange(writeRow, formula.recommendedActionCol).setFormula(formula.recommendedActionFormula);
  }

  // Covers a keyword that already has matching Job_Discovery rows before its
  // Keyword_Strategy row existed (e.g. mined from an already-logged job).
  if (actualCol) recomputeKeywordStrategyActualCount_(ss, keyword);

  return true;
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

  var purged = purgeDroppedKeywords_(ss);

  if (purged.strategyRemoved === 0) {
    ui.alert(
      "No keywords marked \"Drop\" were found in Keyword_Strategy.\n\n" +
      "Set a keyword's Drop column to \"Drop\", then run this again."
    );
    return;
  }

  ui.alert(
    "Done.\n\n" +
    "✓ " + purged.strategyRemoved + " keyword(s) removed from Keyword_Strategy.\n" +
    "✓ " + purged.searchListRemoved + " row(s) removed from Keyword_Search_List."
  );
}

// Scale is Drop's opposite number: instead of removing a keyword, this
// generates related search-phrase variations off every keyword marked
// "Scale" in Keyword_Strategy and adds them to both Keyword_Search_List and
// Keyword_Strategy. The AI call itself (FFLib.generateKeywordVariations)
// lives in the Library, same split as GENERATE_KEYWORD_STRATEGY above.
//
// New rows are written directly into both sheets here (not left to the
// KEYWORD_SEARCH_LIST edit-trigger block in 14_Edit_Trigger.gs), same
// direct-write pattern GENERATE_KEYWORD_STRATEGY uses -- ensureKeywordStrategyRow_
// would still be safe to rely on since it's idempotent, but writing directly
// lets this set a Scale-specific Notes value instead of the trigger's generic
// "Synced from Keyword_Search_List".
//
// The Scale flag itself is never cleared, matching Drop's convention -- it's
// a manual toggle Sawandi controls, not a one-shot switch this function
// resets. Running it again on the same keyword calls the AI again and skips
// any duplicate variation text, so it stays safe to re-run.
function SCALE_KEYWORDS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

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

  var stratMap    = getHeaderMap_(strategySheet);
  var stratKwCol  = getCol_(stratMap, ["Keyword"]);
  var scaleCol    = getCol_(stratMap, ["Scale"]);
  var actualCol   = getCol_(stratMap, ["Actual_Count"]);
  var targetCol   = getCol_(stratMap, ["Target_Count"]);
  var notesCol    = getCol_(stratMap, ["Notes"]);

  if (!stratKwCol || !scaleCol) {
    ui.alert("Keyword_Strategy is missing its Scale column. Run System Tools > Repair Formulas first.");
    return;
  }

  var stratLastRow = strategySheet.getLastRow();
  if (stratLastRow < 2) {
    ui.alert("Keyword_Strategy has no keywords yet.");
    return;
  }

  var stratData = strategySheet
    .getRange(2, 1, stratLastRow - 1, strategySheet.getLastColumn())
    .getValues();

  var toScale = [];
  stratData.forEach(function (row) {
    var kw    = String(row[stratKwCol - 1]).trim();
    var scale = String(row[scaleCol - 1]).trim();
    if (kw !== "" && scale === "Scale") toScale.push(kw);
  });

  if (toScale.length === 0) {
    ui.alert(
      "No keywords marked \"Scale\" were found in Keyword_Strategy.\n\n" +
      "Set a keyword's Scale column to \"Scale\", then run this again."
    );
    return;
  }

  // Same purge-before-generating pattern MINE_KEYWORDS uses, so a keyword
  // that's also marked Drop doesn't get its variations suggested.
  purgeDroppedKeywords_(ss);

  var settings = getSettings_();
  var niche    = settings["Freelancer_Background"] || "";

  var slMap      = getHeaderMap_(searchListSheet);
  var slQueryCol = getCol_(slMap, ["Search_Query"]);

  var existingKeys = {};
  if (slQueryCol && searchListSheet.getLastRow() >= 2) {
    searchListSheet
      .getRange(2, slQueryCol, searchListSheet.getLastRow() - 1, 1)
      .getValues()
      .forEach(function (row) {
        var key = String(row[0]).trim().toLowerCase();
        if (key !== "") existingKeys[key] = true;
      });
  }

  var stratHeaders   = strategySheet.getRange(1, 1, 1, strategySheet.getLastColumn()).getValues()[0];
  var addedTotal     = 0;
  var skippedTotal   = 0;
  var failedKeywords = [];

  toScale.forEach(function (keyword) {
    var result = FFLib.generateKeywordVariations(keyword, niche, apiKey);
    if (!result.ok) {
      failedKeywords.push(keyword + " (" + result.message + ")");
      return;
    }

    var newVariations = result.variations.filter(function (v) {
      var key = String(v).trim().toLowerCase();
      if (key === "" || existingKeys[key]) return false;
      existingKeys[key] = true;
      return true;
    });

    skippedTotal += (result.variations.length - newVariations.length);
    if (newVariations.length === 0) return;

    var searchStart = getLastRealRow_(searchListSheet) + 1;
    var searchRows = newVariations.map(function (v) { return [v, "", "", v]; });
    searchListSheet.getRange(searchStart, 1, searchRows.length, 4).setValues(searchRows);

    var writeRow = getLastRealRow_(strategySheet) + 1;
    newVariations.forEach(function (v) {
      strategySheet.getRange(writeRow, stratKwCol).setValue(v);
      if (actualCol) strategySheet.getRange(writeRow, actualCol).setValue(0);
      if (targetCol) strategySheet.getRange(writeRow, targetCol).setValue(DEFAULT_KEYWORD_TARGET_COUNT_);
      if (notesCol)  strategySheet.getRange(writeRow, notesCol).setValue("Scaled variation of \"" + keyword + "\"");

      var formula = FFLib.buildKeywordStrategyFormula(stratHeaders, writeRow);
      if (formula.recommendedActionFormula) {
        strategySheet.getRange(writeRow, formula.recommendedActionCol).setFormula(formula.recommendedActionFormula);
      }

      writeRow++;
    });

    addedTotal += newVariations.length;
  });

  var summary = "Scale Keywords done.\n\n" +
    "✓ " + addedTotal + " new variation(s) added to Keyword_Search_List and Keyword_Strategy.\n" +
    "Scaled from " + toScale.length + " keyword(s) marked \"Scale\".";

  if (skippedTotal > 0) {
    summary += "\n" + skippedTotal + " suggested variation(s) skipped -- already in Keyword_Search_List.";
  }
  if (failedKeywords.length > 0) {
    summary += "\n\n⚠ Failed for: " + failedKeywords.join(", ");
  }

  ui.alert(summary);
}
