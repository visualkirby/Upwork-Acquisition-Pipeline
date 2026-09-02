/**
 * ============================================================
 * 25. DASHBOARD
 * BUILD_DASHBOARD: refresh-on-demand snapshot. Layout/labels come from
 *   FFLib.buildDashboardLayoutSpec (Library) -- this file just reads sheet
 *   data and writes/charts it per that spec. Clears and rewrites the
 *   Dashboard sheet, and removes existing charts before inserting new ones
 *   so repeated refreshes don't duplicate them.
 *
 * Layout: label/value tile sections (KPIs, Rates, Connects/Cost, Expected
 * Value, Weekly Metrics, Top Performers) stack vertically in columns 1-2,
 * independent of everything below -- charts never touch columns 1-2, so
 * that stack can grow as tall as it needs to without affecting the grid.
 *
 * All 5 charts (Funnel, Revenue Trend, Keyword Priority Score, Priority-vs-
 * Competition scatter, Tool Market Share) sit in a compact 3-row x 2-column
 * grid instead of each claiming its own far-right column band -- an earlier
 * version spread them from column D out past column AM, which meant
 * scrolling through mostly-empty columns to see the whole dashboard.
 * DASHBOARD_GRID_ROWS_/DASHBOARD_GRID_COLS_ below are that grid's anchor
 * points; each write*Section_ call in BUILD_DASHBOARD picks one cell.
 * Row/column spacing is sized for the explicit chart width/height set
 * below, but Apps Script can't preview exact pixel layout -- after
 * refreshing, visually check for any chart overlap and adjust the grid
 * constants if needed.
 * ============================================================
 */
// Explicit size on every chart below -- Sheets' default chart size isn't
// documented/guaranteed, and an earlier version of this file spaced chart
// columns by guessing at the default (~600px), which came out far too wide
// in practice once actually viewed in the sheet. Setting an explicit width/
// height makes the ~4.5-column, ~12-row footprint per chart predictable, so
// DASHBOARD_GRID_ROWS_/DASHBOARD_GRID_COLS_ can be sized tightly on purpose
// instead of guessed defensively.
var DASHBOARD_CHART_WIDTH_ = 450;
var DASHBOARD_CHART_HEIGHT_ = 260;

// The 3-row x 2-column chart grid's anchor points (see header comment).
// Row spacing (15) clears a ~260px-tall chart (~12 rows at Sheets' ~21px
// default row height) with a little buffer; column spacing (9) clears a
// ~450px-wide chart (~4.5 cols at ~100px default column width) plus its
// backing data table.
var DASHBOARD_GRID_ROWS_ = [4, 19, 34];
var DASHBOARD_GRID_COLS_ = [4, 13];

function BUILD_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) {
    ui.alert("Dashboard sheet not found. Run FreelanceFlow Setup first.");
    return;
  }

  // Self-heals Total_Connects_Used/Total_Proposal_Cost against
  // Proposal_Tracker before reading them below -- see
  // 15_Formula_Fixes.gs's reconcileConnectsHelperTotals_ for why those two
  // can drift (an undo mid-way through handleProposalStatusChange_'s several
  // Connects_Helper writes). Silent since this runs on every routine
  // Dashboard refresh, not just when checking for drift on purpose.
  reconcileConnectsHelperTotals_(ss);

  var spec   = FFLib.buildDashboardLayoutSpec();
  var colors = FFLib.getDashboardColors();

  clearDashboardSheet_(dashboard);

  dashboard.getRange(1, 1).setValue(spec.title).setFontSize(16).setFontWeight("bold");
  dashboard.getRange(2, 1).setValue(
    spec.generatedLabel + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
  ).setFontStyle("italic");

  var chMetrics     = readConnectsHelperMetrics_(ss);
  var rates         = FFLib.computeDashboardRates(chMetrics);
  var mergedMetrics = mergeMetrics_(chMetrics, rates);

  var gridRows = DASHBOARD_GRID_ROWS_;
  var gridCols = DASHBOARD_GRID_COLS_;
  var row = gridRows[0];

  row = writeKpiSection_(ss, dashboard, row, spec.kpiSection);
  row = writeMetricValueSection_(dashboard, row, spec.ratesSection, mergedMetrics);
  row = writeMetricValueSection_(dashboard, row, spec.connectsCostSection, mergedMetrics);
  row = writeMetricValueSection_(dashboard, row, spec.expectedValueSection, mergedMetrics);
  row = writeWeeklyMetricsSection_(ss, dashboard, row, spec.weeklyMetricsSection);
  row = writeTopPerformersSection_(ss, dashboard, row, spec.topPerformersSection);

  // Grid: row 1 = Funnel | Revenue Trend, row 2 = Keyword Priority Score |
  // Priority-vs-Competition, row 3 = Tool Market Share (alone).
  row = writeFunnelSection_(ss, dashboard, row, spec.funnelSection, gridRows[0], gridCols[0]);
  writeRevenueTrendSection_(ss, dashboard, row, spec.revenueTrendSection, gridRows[0], gridCols[1]);

  writeKeywordPriorityBarSection_(ss, dashboard, gridRows[1], gridCols[0], spec.keywordIntelligenceSection, colors);
  writeCompetitionScatterSection_(ss, dashboard, gridRows[1], gridCols[1], spec.keywordIntelligenceSection, colors);
  writeToolMarketShareSection_(ss, dashboard, gridRows[2], gridCols[0], spec.toolMarketShareSection, colors);

  dashboard.autoResizeColumns(1, 2);
  ui.alert("Dashboard refreshed.");
}

function formatDashboardValue_(value, format) {
  var num = Number(value) || 0;
  if (format === "currency") return "$" + num.toFixed(2);
  if (format === "percent") return Math.round(num * 100) + "%";
  return num;
}

// Connects_Helper's Metric/Value rows as a plain {Metric: Value} object --
// shared by writeKpiSection_ and BUILD_DASHBOARD's rates/EV computation so
// both read the sheet the same way instead of duplicating the read.
function readConnectsHelperMetrics_(ss) {
  var ch = ss.getSheetByName("Connects_Helper");
  var metrics = {};
  if (ch && ch.getLastRow() > 1) {
    var data = ch.getRange(2, 1, ch.getLastRow() - 1, 2).getValues();
    data.forEach(function (r) {
      var key = String(r[0]).trim();
      if (key !== "") metrics[key] = r[1];
    });
  }
  return metrics;
}

function mergeMetrics_(a, b) {
  var out = {};
  Object.keys(a || {}).forEach(function (k) { out[k] = a[k]; });
  Object.keys(b || {}).forEach(function (k) { out[k] = b[k]; });
  return out;
}

function writeKpiSection_(ss, dashboard, startRow, kpiSpec) {
  var metrics = readConnectsHelperMetrics_(ss);

  dashboard.getRange(startRow, 1).setValue(kpiSpec.title).setFontWeight("bold");
  var row = startRow + 1;
  kpiSpec.metrics.forEach(function (m) {
    dashboard.getRange(row, 1).setValue(m.label);
    dashboard.getRange(row, 2).setValue(formatDashboardValue_(metrics[m.key], m.format));
    row++;
  });
  return row + 1;
}

// Generic label/value tile writer for any section whose values already sit
// in a flat metrics object (Rates, Connects/Cost, Expected Value all share
// this shape) -- keeps their spec-driven label/format handling identical to
// writeKpiSection_ without re-reading Connects_Helper each time.
function writeMetricValueSection_(dashboard, startRow, sectionSpec, metricsObj) {
  dashboard.getRange(startRow, 1).setValue(sectionSpec.title).setFontWeight("bold");
  var row = startRow + 1;
  sectionSpec.metrics.forEach(function (m) {
    dashboard.getRange(row, 1).setValue(m.label);
    dashboard.getRange(row, 2).setValue(formatDashboardValue_(metricsObj[m.key], m.format));
    row++;
  });
  return row + 1;
}

// Last-7-days counts -- nothing stores a rolling weekly window anywhere
// (Connects_Helper's MTD_*/Monthly_* counters reset on a calendar boundary,
// not a rolling one), so this recomputes fresh from Job_Discovery/
// Proposal_Tracker dates every refresh.
function writeWeeklyMetricsSection_(ss, dashboard, startRow, sectionSpec) {
  dashboard.getRange(startRow, 1).setValue(sectionSpec.title).setFontWeight("bold");
  var row = startRow + 1;

  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  var jdSheet = ss.getSheetByName("Job_Discovery");
  var jobsFoundWeek = 0;
  if (jdSheet && jdSheet.getLastRow() > 1) {
    var jdMap = getHeaderMap_(jdSheet);
    var dateFoundCol = getCol_(jdMap, ["Date_Found"]);
    if (dateFoundCol) {
      var jdDates = jdSheet.getRange(2, dateFoundCol, jdSheet.getLastRow() - 1, 1).getValues();
      jdDates.forEach(function (r) {
        if (r[0] instanceof Date && r[0] >= sevenDaysAgo) jobsFoundWeek++;
      });
    }
  }

  var ptSheet = ss.getSheetByName("Proposal_Tracker");
  var proposalsSentWeek = 0, repliesWeek = 0, interviewsWeek = 0, hiresWeek = 0;
  if (ptSheet && ptSheet.getLastRow() > 1) {
    var ptMap          = getHeaderMap_(ptSheet);
    var dateAppliedCol = getCol_(ptMap, ["Date_Applied"]);
    var viewedCol      = getCol_(ptMap, ["Viewed"]);
    var interviewCol   = getCol_(ptMap, ["Interview"]);
    var hiredCol       = getCol_(ptMap, ["Hired"]);
    var ptData = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, ptSheet.getLastColumn()).getValues();
    ptData.forEach(function (r) {
      var d = dateAppliedCol ? r[dateAppliedCol - 1] : null;
      if (d instanceof Date && d >= sevenDaysAgo) {
        proposalsSentWeek++;
        if (viewedCol && String(r[viewedCol - 1]).trim() === "Y") repliesWeek++;
        if (interviewCol && String(r[interviewCol - 1]).trim() === "Y") interviewsWeek++;
        if (hiredCol && String(r[hiredCol - 1]).trim() === "Y") hiresWeek++;
      }
    });
  }

  var values = {
    jobsFoundWeek: jobsFoundWeek,
    proposalsSentWeek: proposalsSentWeek,
    repliesWeek: repliesWeek,
    interviewsWeek: interviewsWeek,
    hiresWeek: hiresWeek
  };
  sectionSpec.metrics.forEach(function (m) {
    dashboard.getRange(row, 1).setValue(m.label);
    dashboard.getRange(row, 2).setValue(values[m.key] || 0);
    row++;
  });
  return row + 1;
}

// Best keyword by 4 Keyword_Intelligence metrics, best Template/Hook/CTA
// combo and top requested tool from Proposal_Tracker. "Viewed" is this
// codebase's own definition of a reply (see 14_Edit_Trigger.gs's
// PROPOSAL_TRACKER block -- MTD_Replies increments when Viewed flips to Y),
// so reply rate here is Viewed-count / sent-count, same meaning throughout.
function writeTopPerformersSection_(ss, dashboard, startRow, sectionSpec) {
  dashboard.getRange(startRow, 1).setValue(sectionSpec.title).setFontWeight("bold");
  var row = startRow + 1;

  var bestByMetric = {};
  var kiSheet = ss.getSheetByName("Keyword_Intelligence");
  if (kiSheet && kiSheet.getLastRow() > 1) {
    var kiHeaders = kiSheet.getRange(1, 1, 1, kiSheet.getLastColumn()).getValues()[0];
    var kiData    = kiSheet.getRange(2, 1, kiSheet.getLastRow() - 1, kiSheet.getLastColumn()).getValues();
    var kwCol     = kiHeaders.indexOf("Keyword");

    ["Priority_Score", "ROI", "EV_per_Proposal", "EV_per_Connect"].forEach(function (metricName) {
      var col = kiHeaders.indexOf(metricName);
      if (col === -1 || kwCol === -1) return;
      var best = null;
      kiData.forEach(function (r) {
        var val = Number(r[col]);
        if (r[col] !== "" && !isNaN(val) && (best === null || val > best.value)) {
          best = { keyword: r[kwCol], value: val };
        }
      });
      bestByMetric[metricName] = best;
    });
  }

  var bestTemplate = null;
  var topTool      = null;
  var ptSheet = ss.getSheetByName("Proposal_Tracker");
  if (ptSheet && ptSheet.getLastRow() > 1) {
    var ptMap     = getHeaderMap_(ptSheet);
    var tmplCol   = getCol_(ptMap, ["Template_Used"]);
    var hookCol   = getCol_(ptMap, ["Hook_Version"]);
    var ctaCol    = getCol_(ptMap, ["CTA_Version"]);
    var viewedCol = getCol_(ptMap, ["Viewed"]);
    var toolCol   = getCol_(ptMap, ["Tool_Requested", "Tool_Detected"]);
    var ptData    = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, ptSheet.getLastColumn()).getValues();

    var comboStats  = {};
    var toolCounts  = {};
    ptData.forEach(function (r) {
      if (tmplCol && hookCol && ctaCol) {
        var key = String(r[tmplCol - 1]) + "|" + String(r[hookCol - 1]) + "|" + String(r[ctaCol - 1]);
        if (!comboStats[key]) {
          comboStats[key] = { sent: 0, viewed: 0, template: r[tmplCol - 1], hook: r[hookCol - 1], cta: r[ctaCol - 1] };
        }
        comboStats[key].sent++;
        if (viewedCol && String(r[viewedCol - 1]).trim() === "Y") comboStats[key].viewed++;
      }
      if (toolCol) {
        var tool = String(r[toolCol - 1]).trim();
        if (tool) toolCounts[tool] = (toolCounts[tool] || 0) + 1;
      }
    });

    Object.keys(comboStats).forEach(function (key) {
      var c = comboStats[key];
      var rate = c.sent > 0 ? c.viewed / c.sent : 0;
      if (c.sent > 0 && (!bestTemplate || rate > bestTemplate.rate)) {
        bestTemplate = { template: c.template, hook: c.hook, cta: c.cta, rate: rate };
      }
    });

    Object.keys(toolCounts).forEach(function (tool) {
      if (!topTool || toolCounts[tool] > topTool.count) {
        topTool = { tool: tool, count: toolCounts[tool] };
      }
    });
  }

  var lines = [
    ["Best Keyword by Priority Score",   bestByMetric.Priority_Score   ? bestByMetric.Priority_Score.keyword   : "--"],
    ["Best Keyword by ROI",              bestByMetric.ROI              ? bestByMetric.ROI.keyword              : "--"],
    ["Best Keyword by EV per Proposal",  bestByMetric.EV_per_Proposal  ? bestByMetric.EV_per_Proposal.keyword  : "--"],
    ["Best Keyword by EV per Connect",   bestByMetric.EV_per_Connect   ? bestByMetric.EV_per_Connect.keyword   : "--"],
    ["Best Template/Hook/CTA Combo",     bestTemplate ? (bestTemplate.template + " / " + bestTemplate.hook + " / " + bestTemplate.cta) : "--"],
    ["Top Tool Requested",               topTool ? topTool.tool : "--"]
  ];
  lines.forEach(function (pair) {
    dashboard.getRange(row, 1).setValue(pair[0]);
    dashboard.getRange(row, 2).setValue(pair[1]);
    row++;
  });
  return row + 1;
}

function writeFunnelSection_(ss, dashboard, startRow, funnelSpec, chartAnchorRow, chartCol) {
  dashboard.getRange(startRow, 1).setValue(funnelSpec.title).setFontWeight("bold");
  var headerRow = startRow + 1;
  dashboard.getRange(headerRow, 1, 1, 2).setValues([["Stage", "Count"]]).setFontWeight("bold");

  var funnel = readFunnelStagesFromSheets_(ss);
  var firstDataRow = headerRow + 1;
  var row = firstDataRow;
  funnelSpec.stages.forEach(function (stage) {
    var count = funnel ? (funnel[stage.key] ? funnel[stage.key][stage.countKey] : 0) : 0;
    dashboard.getRange(row, 1).setValue(stage.label);
    dashboard.getRange(row, 2).setValue(count || 0);
    row++;
  });

  if (funnel) {
    var chartRange = dashboard.getRange(headerRow, 1, row - headerRow, 2);
    var chart = dashboard.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(chartRange)
      .setOption("title", funnelSpec.title)
      .setOption("width", DASHBOARD_CHART_WIDTH_)
      .setOption("height", DASHBOARD_CHART_HEIGHT_)
      .setPosition(chartAnchorRow, chartCol, 0, 0)
      .build();
    dashboard.insertChart(chart);
  }

  return row + 1;
}

function writeRevenueTrendSection_(ss, dashboard, startRow, trendSpec, chartAnchorRow, chartCol) {
  dashboard.getRange(startRow, 1).setValue(trendSpec.title).setFontWeight("bold");
  var headerRow = startRow + 1;
  dashboard.getRange(headerRow, 1, 1, 2).setValues([["Month", "Revenue"]]).setFontWeight("bold");

  var mp = ss.getSheetByName("Monthly_Performance");
  var row = headerRow + 1;
  if (mp && mp.getLastRow() > 1) {
    var mpMap = getHeaderMap_(mp);
    var monthCol = getCol_(mpMap, ["Month"]);
    var yearCol = getCol_(mpMap, ["Year"]);
    var revenueCol = getCol_(mpMap, ["Revenue"]);
    if (monthCol && yearCol && revenueCol) {
      var mpData = mp.getRange(2, 1, mp.getLastRow() - 1, mp.getLastColumn()).getValues();
      mpData.forEach(function (r) {
        var label = String(r[monthCol - 1]) + " " + String(r[yearCol - 1]);
        dashboard.getRange(row, 1).setValue(label);
        dashboard.getRange(row, 2).setValue(Number(r[revenueCol - 1]) || 0);
        row++;
      });
    }
  }

  if (row > headerRow + 1) {
    var chartRange = dashboard.getRange(headerRow, 1, row - headerRow, 2);
    var chart = dashboard.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
      .setOption("title", trendSpec.title)
      .setOption("width", DASHBOARD_CHART_WIDTH_)
      .setOption("height", DASHBOARD_CHART_HEIGHT_)
      .setPosition(chartAnchorRow, chartCol, 0, 0)
      .build();
    dashboard.insertChart(chart);
  }

  return row + 1;
}

// Shared Keyword_Intelligence read for the bar + scatter sections below --
// top N keywords by Priority_Score, same ranked set feeds both charts.
function getRankedKeywordIntelligenceRows_(ss, maxKeywords) {
  var kiSheet = ss.getSheetByName("Keyword_Intelligence");
  if (!kiSheet || kiSheet.getLastRow() < 2) return null;

  var headers = kiSheet.getRange(1, 1, 1, kiSheet.getLastColumn()).getValues()[0];
  var data    = kiSheet.getRange(2, 1, kiSheet.getLastRow() - 1, kiSheet.getLastColumn()).getValues();
  var kwCol         = headers.indexOf("Keyword");
  var priorityCol   = headers.indexOf("Priority_Score");
  var competitionCol = headers.indexOf("Competition_Score");
  var statusCol     = headers.indexOf("Status");
  if (kwCol === -1 || priorityCol === -1) return null;

  var rows = data
    .filter(function (r) { return r[priorityCol] !== "" && !isNaN(Number(r[priorityCol])); })
    .sort(function (a, b) { return Number(b[priorityCol]) - Number(a[priorityCol]); })
    .slice(0, maxKeywords);

  if (rows.length === 0) return null;

  return { rows: rows, kwCol: kwCol, priorityCol: priorityCol, competitionCol: competitionCol, statusCol: statusCol };
}

// Keyword Priority Score ranking -- single hue, this is a magnitude/ranking
// question, not an identity question, so one color is correct per the
// dataviz method.
function writeKeywordPriorityBarSection_(ss, dashboard, row, col, sectionSpec, colors) {
  var ki = getRankedKeywordIntelligenceRows_(ss, sectionSpec.maxKeywords);
  if (!ki) return;

  dashboard.getRange(row, col, 1, 2).setValues([["Keyword", "Priority_Score"]]).setFontWeight("bold");
  var barValues = ki.rows.map(function (r) { return [r[ki.kwCol], Number(r[ki.priorityCol])]; });
  dashboard.getRange(row + 1, col, barValues.length, 2).setValues(barValues);

  var barChart = dashboard.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dashboard.getRange(row, col, barValues.length + 1, 2))
    .setOption("title", sectionSpec.title)
    .setOption("colors", [colors.singleSeries])
    .setOption("legend", { position: "none" })
    .setOption("width", DASHBOARD_CHART_WIDTH_)
    .setOption("height", DASHBOARD_CHART_HEIGHT_)
    .setPosition(row, col + 3, 0, 0)
    .build();
  dashboard.insertChart(barChart);
}

// Priority-vs-Competition scatter colored by Status. Sheets scatter charts
// color by series, not by per-point value, so the backing table is
// structured as Competition_Score | Scale | Test More | Drop, with only the
// matching status column populated per row -- the standard "one series per
// category" technique for categorical point coloring in Sheets/Google Charts.
function writeCompetitionScatterSection_(ss, dashboard, row, col, sectionSpec, colors) {
  var ki = getRankedKeywordIntelligenceRows_(ss, sectionSpec.maxKeywords);
  if (!ki || ki.competitionCol === -1 || ki.statusCol === -1) return;

  var statusOrder = ["Scale", "Test More", "Drop"];
  var scatterHeader = ["Competition_Score"].concat(statusOrder);
  dashboard.getRange(row, col, 1, scatterHeader.length)
    .setValues([scatterHeader]).setFontWeight("bold");

  var scatterRows = ki.rows.map(function (r) {
    var line = [Number(r[ki.competitionCol]) || 0, "", "", ""];
    var idx = statusOrder.indexOf(String(r[ki.statusCol]).trim());
    if (idx !== -1) line[idx + 1] = Number(r[ki.priorityCol]);
    return line;
  });
  dashboard.getRange(row + 1, col, scatterRows.length, scatterHeader.length).setValues(scatterRows);

  var scatterChart = dashboard.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(dashboard.getRange(row, col, scatterRows.length + 1, scatterHeader.length))
    .setOption("title", sectionSpec.scatterTitle)
    .setOption("colors", statusOrder.map(function (s) { return colors.status[s]; }))
    .setOption("hAxis", { title: "Competition Score" })
    .setOption("vAxis", { title: "Priority Score" })
    .setOption("width", DASHBOARD_CHART_WIDTH_)
    .setOption("height", DASHBOARD_CHART_HEIGHT_)
    .setPosition(row, col + 5, 0, 0)
    .build();
  dashboard.insertChart(scatterChart);
}

// Top N tools by Job_Discovery volume, remainder folded into "Other" --
// single hue, same magnitude-not-identity reasoning as the keyword bar.
function writeToolMarketShareSection_(ss, dashboard, row, col, sectionSpec, colors) {
  var jdSheet = ss.getSheetByName("Job_Discovery");
  if (!jdSheet || jdSheet.getLastRow() < 2) return;

  var jdMap = getHeaderMap_(jdSheet);
  var toolCol = getCol_(jdMap, ["Tool_Detected"]);
  if (!toolCol) return;

  var toolValues = jdSheet.getRange(2, toolCol, jdSheet.getLastRow() - 1, 1).getValues();
  var counts = {};
  toolValues.forEach(function (r) {
    var tool = String(r[0]).trim();
    if (tool) counts[tool] = (counts[tool] || 0) + 1;
  });

  var ranked = Object.keys(counts)
    .map(function (tool) { return { tool: tool, count: counts[tool] }; })
    .sort(function (a, b) { return b.count - a.count; });

  if (ranked.length === 0) return;

  var top = ranked.slice(0, sectionSpec.maxTools);
  var otherCount = ranked.slice(sectionSpec.maxTools).reduce(function (sum, r) { return sum + r.count; }, 0);
  if (otherCount > 0) {
    // Tool_Detected already has its own legitimate "Other" category (the
    // classifier's catch-all) -- if that real "Other" is already in the top
    // N, merge the overflow into it instead of adding a second, separately-
    // labeled "Other" bar.
    var existingOther = top.filter(function (r) { return r.tool === "Other"; })[0];
    if (existingOther) {
      existingOther.count += otherCount;
    } else {
      top.push({ tool: "Other", count: otherCount });
    }
  }

  dashboard.getRange(row, col, 1, 2).setValues([["Tool", "Job_Count"]]).setFontWeight("bold");
  var values = top.map(function (r) { return [r.tool, r.count]; });
  dashboard.getRange(row + 1, col, values.length, 2).setValues(values);

  var chart = dashboard.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dashboard.getRange(row, col, values.length + 1, 2))
    .setOption("title", sectionSpec.title)
    .setOption("colors", [colors.singleSeries])
    .setOption("legend", { position: "none" })
    .setOption("width", DASHBOARD_CHART_WIDTH_)
    .setOption("height", DASHBOARD_CHART_HEIGHT_)
    .setPosition(row, col + 3, 0, 0)
    .build();
  dashboard.insertChart(chart);
}
