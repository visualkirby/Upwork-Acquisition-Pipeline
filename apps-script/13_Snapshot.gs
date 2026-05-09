/**
 * ============================================================
 * 13. SNAPSHOT MONTH END
 * Reads MTD metrics from Connects_Helper, writes a row to
 * Monthly_Performance, then resets Monthly_Revenue to 0.
 * ============================================================
 */
function SNAPSHOT_MONTH_END() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ch = ss.getSheetByName("Connects_Helper");
  var mp = ss.getSheetByName("Monthly_Performance");

  if (!ch || !mp) {
    ui.alert("Connects_Helper or Monthly_Performance sheet not found.");
    return;
  }

  var lastRow    = ch.getLastRow();
  var metricData = ch.getRange(2, 1, lastRow - 1, 2).getValues();
  var metrics    = {};
  for (var i = 0; i < metricData.length; i++) {
    var key = String(metricData[i][0]).trim();
    var val = metricData[i][1];
    if (key !== "") metrics[key] = val;
  }

  var today            = new Date();
  var monthName        = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM");
  var year             = today.getFullYear();
  var totalSessions    = metrics["MTD_Sessions"]                || 0;
  var jobsLogged       = metrics["MTD_Jobs_Logged"]             || 0;
  var propsSent        = metrics["MTD_Proposals_Sent"]          || 0;
  var connectsUsed     = metrics["MTD_Connects_Used"]           || 0;
  var proposalCost     = metrics["Total_Proposal_Cost"]         || 0;
  var replies          = metrics["MTD_Replies"]                 || 0;
  var interviews       = metrics["MTD_Interviews"]              || 0;
  var hires            = metrics["MTD_Hires"]                   || 0;
  var revenue          = metrics["MTD_Revenue"]                 || 0;
  var monthlyCost      = metrics["Monthly_Cost"]                || 0;
  var monthlyROI       = metrics["Monthly_ROI"]                 || 0;
  var monthlyROIDollar = metrics["Monthly_ROI_Dollar"]          || 0;
  var evPerProposal    = metrics["Expected_Value_per_Proposal"] || 0;
  var revenuePerConn   = metrics["Revenue_per_Connect"]         || 0;
  var netValuePerConn  = metrics["Net_Value_per_Connect"]       || 0;
  var cpr              = metrics["Cost_per_Reply"]              || 0;
  var cpi              = metrics["Cost_per_Interview"]          || 0;
  var cph              = metrics["Cost_per_Hire"]               || 0;

  var replyRate     = propsSent > 0 ? Math.round((replies    / propsSent) * 1000) / 10 : 0;
  var interviewRate = propsSent > 0 ? Math.round((interviews / propsSent) * 1000) / 10 : 0;
  var hireRate      = propsSent > 0 ? Math.round((hires      / propsSent) * 1000) / 10 : 0;

  var summary =
    "MONTH END SNAPSHOT -- " + monthName + " " + year + "\n" +
    "════════════════════════════════\n" +
    "Sessions:                  " + totalSessions  + "\n" +
    "Jobs Logged:               " + jobsLogged     + "\n" +
    "Proposals Sent:            " + propsSent      + "\n" +
    "Connects Used:             " + connectsUsed   + "\n" +
    "Proposal Cost:             $" + Number(proposalCost).toFixed(2)     + "\n\n" +
    "Replies:                   " + replies        + "\n" +
    "Interviews:                " + interviews     + "\n" +
    "Hires:                     " + hires          + "\n" +
    "Reply Rate:                " + replyRate      + "%\n" +
    "Interview Rate:            " + interviewRate  + "%\n" +
    "Hire Rate:                 " + hireRate       + "%\n\n" +
    "Revenue:                   $" + Number(revenue).toFixed(2)          + "\n" +
    "Cost:                      $" + Number(monthlyCost).toFixed(2)      + "\n" +
    "ROI:                       "  + Math.round(Number(monthlyROI) * 100) + "%\n" +
    "ROI Dollar:                $" + Number(monthlyROIDollar).toFixed(2) + "\n" +
    "EV per Proposal:           $" + Number(evPerProposal).toFixed(2)    + "\n" +
    "Revenue per Connect:       "  + Number(revenuePerConn).toFixed(4)   + "\n" +
    "Net Value per Connect:     "  + Number(netValuePerConn).toFixed(4)  + "\n\n" +
    "Snapshot this month and reset Monthly_Revenue to 0?";

  var response = ui.alert(
    "Month End Snapshot",
    summary,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    ui.alert("Snapshot cancelled. No changes made.");
    return;
  }

  var mpMap   = getHeaderMap_(mp);
  var nextRow = mp.getLastRow() + 1;
  if (mp.getLastRow() <= 1) nextRow = 2;

  setCellValue_(mp, nextRow, mpMap, ["Month"],                        monthName);
  setCellValue_(mp, nextRow, mpMap, ["Year"],                         year);
  setCellValue_(mp, nextRow, mpMap, ["Total_Sessions"],               totalSessions);
  setCellValue_(mp, nextRow, mpMap, ["Jobs_Logged"],                  jobsLogged);
  setCellValue_(mp, nextRow, mpMap, ["Proposals_Sent"],               propsSent);
  setCellValue_(mp, nextRow, mpMap, ["Connects_Used"],                connectsUsed);
  setCellValue_(mp, nextRow, mpMap, ["Proposal_Cost"],                proposalCost);
  setCellValue_(mp, nextRow, mpMap, ["Replies"],                      replies);
  setCellValue_(mp, nextRow, mpMap, ["Interviews"],                   interviews);
  setCellValue_(mp, nextRow, mpMap, ["Hires"],                        hires);
  setCellValue_(mp, nextRow, mpMap, ["Reply_Rate_Pct"],               replyRate);
  setCellValue_(mp, nextRow, mpMap, ["Interview_Rate_Pct"],           interviewRate);
  setCellValue_(mp, nextRow, mpMap, ["Hire_Rate_Pct"],                hireRate);
  setCellValue_(mp, nextRow, mpMap, ["Revenue"],                      revenue);
  setCellValue_(mp, nextRow, mpMap, ["Cost"],                         monthlyCost);
  setCellValue_(mp, nextRow, mpMap, ["ROI"],                          monthlyROI);
  setCellValue_(mp, nextRow, mpMap, ["Cost_per_Reply"],               cpr);
  setCellValue_(mp, nextRow, mpMap, ["Cost_per_Interview"],           cpi);
  setCellValue_(mp, nextRow, mpMap, ["Cost_per_Hire"],                cph);
  setCellValue_(mp, nextRow, mpMap, ["Monthly_ROI_Dollar"],           monthlyROIDollar);
  setCellValue_(mp, nextRow, mpMap, ["Expected_Value_per_Proposal"],  evPerProposal);
  setCellValue_(mp, nextRow, mpMap, ["Revenue_per_Connect"],          revenuePerConn);
  setCellValue_(mp, nextRow, mpMap, ["Net_Value_per_Connect"],        netValuePerConn);

  for (var i = 0; i < metricData.length; i++) {
    if (String(metricData[i][0]).trim() === "Monthly_Revenue") {
      ch.getRange(i + 2, 2).setValue(0);
      break;
    }
  }

  ui.alert(
    "✓ Snapshot complete.\n\n" +
    monthName + " " + year + " saved to Monthly_Performance.\n" +
    "Monthly_Revenue reset to 0."
  );
}
