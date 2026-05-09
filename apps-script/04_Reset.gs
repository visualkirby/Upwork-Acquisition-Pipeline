/**
 * ============================================================
 * 4. SMART RESET
 * Clears only true input columns. Preserves formulas / auto columns.
 * ============================================================
 */
function RESET_SYSTEM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function clearByHeaders(sheetName, headerNames) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow <= 1) return;
    var map = getHeaderMap_(sh);
    headerNames.forEach(function (name) {
      var col = getCol_(map, [name]);
      if (col) {
        sh.getRange(2, col, lastRow - 1, 1).clearContent();
      }
    });
  }

  clearByHeaders("Job_Discovery", [
    "Job_Title", "Description", "Client_Name", "Client Name",
    "Keyword_Search", "Experience_Level", "Hours_Since_Posted",
    "Days_Since_Posted", "Proposal_Count", "Payment_Verified",
    "Client_Hires", "Budget_Type", "Budget", "Hourly_Rate",
    "Job_Link", "Connects_Required", "Quick_Notes"
  ]);

  clearByHeaders("Job_Scoring", [
    "Effort_Level", "Scope_Rating", "Portfolio_Match"
  ]);

  clearByHeaders("Proposal_Generator", [
    "Notes", "Proposal_Status"
  ]);

  clearByHeaders("Proposal_Tracker", [
    "Client_Replied", "Interview", "Hired", "Revenue", "Notes"
  ]);

  clearByHeaders("Followup_Tracker", [
    "Followup1_Sent", "Followup2_Sent", "Followup3_Sent",
    "Client_Replied", "Interview", "Hired", "Notes"
  ]);
}
