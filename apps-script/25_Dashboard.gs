/**
 * ============================================================
 * 25. DASHBOARD
 * BUILD_DASHBOARD: refresh-on-demand snapshot -- KPI tiles (from
 *   Connects_Helper), pipeline funnel (via readFunnelStagesFromSheets_,
 *   03_Helpers.gs, same computation ANALYZE_JOB_WORKFLOW uses), and a
 *   revenue trend (from Monthly_Performance). Layout/labels come from
 *   FFLib.buildDashboardLayoutSpec (Library) -- this file just reads
 *   sheet data and writes/charts it per that spec. Clears and rewrites
 *   the Dashboard sheet, and removes existing charts before inserting
 *   new ones so repeated refreshes don't duplicate them.
 * ============================================================
 */
function BUILD_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) {
    ui.alert("Dashboard sheet not found. Run FreelanceFlow Setup first.");
    return;
  }

  var spec = FFLib.buildDashboardLayoutSpec();

  clearDashboardSheet_(dashboard);

  dashboard.getRange(1, 1).setValue(spec.title).setFontSize(16).setFontWeight("bold");
  dashboard.getRange(2, 1).setValue(
    spec.generatedLabel + ": " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
  ).setFontStyle("italic");

  var chartAnchorRow = 4;
  var row = 4;

  row = writeKpiSection_(ss, dashboard, row, spec.kpiSection);
  row = writeFunnelSection_(ss, dashboard, row, spec.funnelSection, chartAnchorRow, 4);
  writeRevenueTrendSection_(ss, dashboard, row, spec.revenueTrendSection, chartAnchorRow, 10);

  dashboard.autoResizeColumns(1, 2);
  ui.alert("Dashboard refreshed.");
}

function formatDashboardValue_(value, format) {
  var num = Number(value) || 0;
  if (format === "currency") return "$" + num.toFixed(2);
  if (format === "percent") return Math.round(num * 100) + "%";
  return num;
}

function writeKpiSection_(ss, dashboard, startRow, kpiSpec) {
  var ch = ss.getSheetByName("Connects_Helper");
  var metrics = {};
  if (ch && ch.getLastRow() > 1) {
    var data = ch.getRange(2, 1, ch.getLastRow() - 1, 2).getValues();
    data.forEach(function (r) {
      var key = String(r[0]).trim();
      if (key !== "") metrics[key] = r[1];
    });
  }

  dashboard.getRange(startRow, 1).setValue(kpiSpec.title).setFontWeight("bold");
  var row = startRow + 1;
  kpiSpec.metrics.forEach(function (m) {
    dashboard.getRange(row, 1).setValue(m.label);
    dashboard.getRange(row, 2).setValue(formatDashboardValue_(metrics[m.key], m.format));
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
      .setPosition(chartAnchorRow, chartCol, 0, 0)
      .build();
    dashboard.insertChart(chart);
  }

  return row + 1;
}
