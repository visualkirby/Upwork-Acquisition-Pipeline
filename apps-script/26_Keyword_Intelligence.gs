/**
 * ============================================================
 * 26. KEYWORD INTELLIGENCE
 * BUILD_KEYWORD_INTELLIGENCE: refresh-on-demand per-keyword ROI/priority
 * analytics -- ported from the older "Upwork Client Acquisition System"
 * gsheet's Keyword_Intelligence tab (a separate, pre-FreelanceFlow system),
 * where the same metrics were native Sheets formulas hardcoded to that
 * system's own column letters. Those letters don't line up with
 * FreelanceFlow's actual layout, and hardcoding them here would violate
 * this project's header-based-lookup convention (see getCol_/getHeaderMap_
 * usage throughout), so this is a script-computed rebuild instead, same
 * refresh-on-demand model as BUILD_DASHBOARD (25_Dashboard.gs) -- it only
 * updates when you run it, not live, and clears/rewrites the sheet each
 * time.
 *
 * Revenue/ROI columns (Revenue, Proposal_Cost, Net_Value, EV_per_Proposal,
 * EV_per_Connect, ROI, Net_ROI) do NOT mirror the source system's formulas,
 * which summed Proposal_Tracker!Revenue directly -- that field is explicitly
 * non-authoritative in FreelanceFlow (see 14_Edit_Trigger.gs's
 * PROPOSAL_TRACKER block comment: "Revenue is a plain manual/informational
 * field here... Milestone_Tracker's Released status and Hourly_Log entries
 * are the authoritative revenue sources now"). Revenue here is computed the
 * same way Connects_Helper's own MTD_Revenue/Monthly_Revenue are: summed
 * from Milestone_Tracker rows with Status="Released" plus Hourly_Log rows'
 * Amount, joined back to a keyword via Job_Discovery's Discovery_ID (which
 * carries both Keyword_Search and Discovery_ID on the same row).
 * Proposal_Cost sums Proposal_Tracker's own Proposal_Cost column (stored as
 * a "$X.XX" string, parsed here), same source the corresponding
 * Connects_Helper metrics use.
 * ============================================================
 */
function BUILD_KEYWORD_INTELLIGENCE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var kiSheet = ss.getSheetByName("Keyword_Intelligence");
  if (!kiSheet) {
    ui.alert("Keyword_Intelligence sheet not found. Run System Tools > Repair Formulas first to create it.");
    return;
  }

  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var jdSheet          = ss.getSheetByName("Job_Discovery");
  var jsSheet          = ss.getSheetByName("Job_Scoring");
  var ptSheet          = ss.getSheetByName("Proposal_Tracker");
  var msSheet          = ss.getSheetByName("Milestone_Tracker");
  var hlSheet          = ss.getSheetByName("Hourly_Log");
  var chSheet          = ss.getSheetByName("Connects_Helper");

  if (!searchListSheet || !jdSheet || !jsSheet) {
    ui.alert("Missing sheet. Confirm Keyword_Search_List, Job_Discovery, and Job_Scoring all exist.");
    return;
  }

  // Self-heals Total_Connects_Used/Total_Proposal_Cost against Proposal_Tracker
  // before the cross-check below reads Total_Proposal_Cost -- see
  // 15_Formula_Fixes.gs's reconcileConnectsHelperTotals_ for why those two can
  // drift out of sync (an undo mid-way through handleProposalStatusChange_'s
  // several Connects_Helper writes). Without this, the cross-check below could
  // misreport counter drift as an orphaned keyword.
  reconcileConnectsHelperTotals_(ss);

  var keywords = getUniqueKeywords_(searchListSheet);
  if (keywords.length === 0) {
    ui.alert("Keyword_Search_List has no Search_Query values yet -- nothing to analyze.");
    return;
  }

  var jd = readSheetAsObjects_(jdSheet);
  var js = readSheetAsObjects_(jsSheet);
  var pt = ptSheet ? readSheetAsObjects_(ptSheet) : [];
  var ms = msSheet ? readSheetAsObjects_(msSheet) : [];
  var hl = hlSheet ? readSheetAsObjects_(hlSheet) : [];

  // Discovery_ID -> total Released Milestone revenue for that contract.
  var releasedByDiscoveryId = {};
  ms.forEach(function (row) {
    if (String(row.Status).trim() !== "Released") return;
    var id = String(row.Discovery_ID).trim();
    if (!id) return;
    releasedByDiscoveryId[id] = (releasedByDiscoveryId[id] || 0) + (Number(row.Amount) || 0);
  });

  // Discovery_ID -> total Hourly_Log revenue for that contract.
  var hourlyByDiscoveryId = {};
  hl.forEach(function (row) {
    var id = String(row.Discovery_ID).trim();
    if (!id) return;
    hourlyByDiscoveryId[id] = (hourlyByDiscoveryId[id] || 0) + (Number(row.Amount) || 0);
  });

  var results = keywords.map(function (keyword) {
    return computeKeywordIntelligenceRow_(keyword, jd, js, pt, releasedByDiscoveryId, hourlyByDiscoveryId);
  });

  writeKeywordIntelligenceSheet_(kiSheet, results);

  var summary = "Keyword Intelligence refreshed for " + results.length + " keyword(s).";

  // Cross-check against Connects_Helper's Total_Proposal_Cost -- both are
  // all-time cumulative totals (unlike MTD_Revenue/Monthly_Revenue, which
  // are period-scoped and reset, so they're not a valid comparison here),
  // so they should match closely. A real mismatch usually means a keyword
  // in Proposal_Tracker no longer matches anything in Keyword_Search_List
  // (renamed/retyped), so its cost isn't attributed to any row above.
  if (chSheet) {
    var chMetrics = readMetricValueSheet_(chSheet);
    var chTotalCost = Number(chMetrics["Total_Proposal_Cost"]) || 0;
    var kiTotalCost = results.reduce(function (sum, r) { return sum + r.Proposal_Cost; }, 0);
    var diff = Math.abs(chTotalCost - kiTotalCost);
    if (diff > 0.01) {
      summary += "\n\nNote: Keyword_Intelligence's total Proposal_Cost ($" + kiTotalCost.toFixed(2) +
        ") doesn't match Connects_Helper's Total_Proposal_Cost ($" + chTotalCost.toFixed(2) + "). " +
        "This usually means a Proposal_Tracker row's Keyword_Search text doesn't match any current " +
        "Keyword_Search_List entry, so its cost isn't attributed to a keyword above.";
    }
  }

  ui.alert(summary);
}

// Search_Query values from Keyword_Search_List, non-blank, de-duplicated
// case/whitespace-insensitively (first-seen casing kept).
function getUniqueKeywords_(searchListSheet) {
  var map      = getHeaderMap_(searchListSheet);
  var queryCol = getCol_(map, ["Search_Query"]);
  if (!queryCol || searchListSheet.getLastRow() < 2) return [];

  var values = searchListSheet.getRange(2, queryCol, searchListSheet.getLastRow() - 1, 1).getValues();
  var seen = {};
  var out  = [];
  values.forEach(function (row) {
    var raw = row[0];
    if (raw === "" || raw === null) return;
    var key = String(raw).trim().toLowerCase();
    if (key === "" || seen[key]) return;
    seen[key] = true;
    out.push(String(raw).trim());
  });
  return out;
}

// Reads every data row of a sheet into an array of {Header: value} objects,
// keyed by that sheet's own header row -- so callers can read row.Keyword_Search,
// row.Discovery_Action, etc. by name instead of by position.
function readSheetAsObjects_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values  = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      if (h !== "") obj[h] = row[i];
    });
    return obj;
  });
}

// Connects_Helper-style Metric/Value sheet -> plain {Metric: Value} object.
function readMetricValueSheet_(sheet) {
  var map       = getHeaderMap_(sheet);
  var metricCol = getCol_(map, ["Metric"]);
  var valueCol  = getCol_(map, ["Value"]);
  var out = {};
  if (!metricCol || !valueCol || sheet.getLastRow() < 2) return out;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  values.forEach(function (row) {
    var key = String(row[metricCol - 1]).trim();
    if (key !== "") out[key] = row[valueCol - 1];
  });
  return out;
}

// Parses a Proposal_Cost cell, which is stored as a formatted "$X.XX"
// string (see handleProposalStatusChange_, 14_Edit_Trigger.gs) or blank,
// into a plain number.
function parseDollarString_(value) {
  if (value === "" || value === null || value === undefined) return 0;
  var num = Number(String(value).replace(/[^0-9.\-]/g, ""));
  return isNaN(num) ? 0 : num;
}

function classifyToolFocus_(keyword) {
  var k = String(keyword).toLowerCase();
  if (/tableau/.test(k)) return "Tableau";
  if (/power\s*bi/.test(k)) return "Power BI";
  if (/looker/.test(k)) return "Looker Studio";
  if (/excel|spreadsheet/.test(k)) return "Excel";
  if (/sql/.test(k)) return "SQL";
  return "Other";
}

function computeKeywordIntelligenceRow_(keyword, jd, js, pt, releasedByDiscoveryId, hourlyByDiscoveryId) {
  var keyLower = String(keyword).trim().toLowerCase();
  var matchesKeyword = function (val) { return String(val).trim().toLowerCase() === keyLower; };

  var jdRows = jd.filter(function (row) { return matchesKeyword(row.Keyword_Search); });
  var jsRows = js.filter(function (row) { return matchesKeyword(row.Keyword_Search); });
  var ptRows = pt.filter(function (row) { return matchesKeyword(row.Keyword_Search); });

  var totalJobs = jdRows.length;
  var moveToScoring = jdRows.filter(function (r) { return String(r.Discovery_Action).trim() === "Move to Scoring"; }).length;
  var reviewLater    = jdRows.filter(function (r) { return String(r.Discovery_Action).trim() === "Review Later"; }).length;
  var skip            = jdRows.filter(function (r) { return String(r.Discovery_Action).trim() === "Skip"; }).length;

  var scoringApply = jsRows.filter(function (r) { return String(r.Final_Decision).trim() === "APPLY"; }).length;
  var scoringHold   = jsRows.filter(function (r) { return String(r.Final_Decision).trim() === "HOLD"; }).length;
  var scoringSkip   = jsRows.filter(function (r) { return String(r.Final_Decision).trim() === "SKIP"; }).length;

  var applyRate = totalJobs > 0 ? scoringApply / totalJobs : "";

  var proposalCounts = jdRows.map(function (r) { return Number(r.Proposal_Count) || 0; }).filter(function (n) { return !isNaN(n); });
  var avgProposals = average_(proposalCounts);
  var competitionScore = avgProposals === "" ? "" : tierCompetitionScore_(avgProposals);

  var avgBudgetScore = average_(jsRows.map(function (r) { return r.Budget_Score; }).filter(hasNumericValue_));

  var avgTotalScore = average_(jsRows.map(function (r) { return r.Total_Score; }).filter(hasNumericValue_));

  var connectsValues = jdRows.map(function (r) { return r.Connects_Required; }).filter(hasNumericValue_);
  var avgConnects = average_(connectsValues);

  var connectEfficiency = (avgTotalScore === "" || avgConnects === "" || avgConnects === 0) ? "" : avgTotalScore / avgConnects;

  var priorityScore = "";
  if (connectEfficiency !== "" && applyRate !== "" && avgTotalScore !== "" && avgBudgetScore !== "" && competitionScore !== "") {
    priorityScore = Math.round((
      applyRate * 0.3 +
      avgTotalScore * 0.25 +
      avgBudgetScore * 0.15 +
      connectEfficiency * 0.15 +
      (totalJobs > 0 ? (moveToScoring / totalJobs) : 0) * 0.05 +
      (totalJobs > 0 ? (scoringHold / totalJobs) : 0) * 0.05 +
      competitionScore * 0.05
    ) * 100) / 100;
  }

  var status = priorityScore === "" ? "" : (priorityScore >= 0.6 ? "Scale" : (priorityScore >= 0.4 ? "Test More" : "Drop"));

  var revenue = jdRows.reduce(function (sum, row) {
    var id = String(row.Discovery_ID).trim();
    if (!id) return sum;
    return sum + (releasedByDiscoveryId[id] || 0) + (hourlyByDiscoveryId[id] || 0);
  }, 0);

  var proposalCost = ptRows.reduce(function (sum, row) { return sum + parseDollarString_(row.Proposal_Cost); }, 0);

  var netValue = revenue - proposalCost;
  var evPerProposal = scoringApply > 0 ? revenue / scoringApply : "";
  var evPerConnect = (avgConnects !== "" && avgConnects > 0 && scoringApply > 0) ? revenue / (avgConnects * scoringApply) : "";
  var roi = proposalCost > 0 ? revenue / proposalCost : "";
  var netROI = proposalCost > 0 ? netValue / proposalCost : "";

  return {
    Keyword: keyword,
    Tool_Focus: classifyToolFocus_(keyword),
    Total_Jobs: totalJobs,
    Discovery_Move_to_Scoring: moveToScoring,
    Discovery_Review_Later: reviewLater,
    Discovery_Skip: skip,
    Scoring_APPLY: scoringApply,
    Scoring_HOLD: scoringHold,
    Scoring_SKIP: scoringSkip,
    APPLY_Rate: applyRate,
    Avg_Proposals: avgProposals,
    Competition_Score: competitionScore,
    Avg_Budget_Score: avgBudgetScore,
    Avg_Total_Score: avgTotalScore,
    Avg_Connects: avgConnects,
    Connect_Efficiency: connectEfficiency,
    Priority_Score: priorityScore,
    Status: status,
    Revenue: revenue,
    Proposal_Cost: proposalCost,
    Net_Value: netValue,
    EV_per_Proposal: evPerProposal,
    EV_per_Connect: evPerConnect,
    ROI: roi,
    Net_ROI: netROI
  };
}

function hasNumericValue_(v) {
  return v !== "" && v !== null && v !== undefined && !isNaN(Number(v));
}

function average_(numericValues) {
  if (!numericValues || numericValues.length === 0) return "";
  var sum = numericValues.reduce(function (a, b) { return a + Number(b); }, 0);
  return sum / numericValues.length;
}

function tierCompetitionScore_(avgProposals) {
  if (avgProposals <= 5) return 1;
  if (avgProposals <= 10) return 0.9;
  if (avgProposals <= 20) return 0.7;
  if (avgProposals <= 35) return 0.5;
  return 0.3;
}

var KEYWORD_INTELLIGENCE_HEADERS_ = [
  "Keyword", "Tool_Focus", "Total_Jobs", "Discovery_Move_to_Scoring", "Discovery_Review_Later",
  "Discovery_Skip", "Scoring_APPLY", "Scoring_HOLD", "Scoring_SKIP", "APPLY_Rate", "Avg_Proposals",
  "Competition_Score", "Avg_Budget_Score", "Avg_Total_Score", "Avg_Connects", "Connect_Efficiency",
  "Priority_Score", "Status", "Revenue", "Proposal_Cost", "Net_Value", "EV_per_Proposal",
  "EV_per_Connect", "ROI", "Net_ROI"
];

function writeKeywordIntelligenceSheet_(kiSheet, results) {
  kiSheet.clear();
  kiSheet.getRange(1, 1, 1, KEYWORD_INTELLIGENCE_HEADERS_.length)
    .setValues([KEYWORD_INTELLIGENCE_HEADERS_])
    .setFontWeight("bold");

  if (results.length === 0) return;

  var rows = results.map(function (r) {
    return KEYWORD_INTELLIGENCE_HEADERS_.map(function (h) { return r[h]; });
  });
  kiSheet.getRange(2, 1, rows.length, KEYWORD_INTELLIGENCE_HEADERS_.length).setValues(rows);
  kiSheet.autoResizeColumns(1, KEYWORD_INTELLIGENCE_HEADERS_.length);
}
