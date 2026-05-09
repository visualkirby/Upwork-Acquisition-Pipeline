/**
 * ============================================================
 * 15. FORMULA FIXES
 * One-time runner that corrects four formula bugs identified
 * during the May 2026 system audit. Safe to re-run -- it
 * overwrites the target cells with the corrected formula and
 * copies down to all data rows.
 *
 * F4a -- Affordability_Check ($B$14 -> $B$16 in rows 3+)
 * F4b -- Final_Decision gate (AF=0 -> AF="Cannot Afford")
 * F7  -- Tool_Detected (adds Google Sheets, SQL, Python)
 * F9  -- Portfolio_Project (real names, expanded keywords)
 * ============================================================
 */
function APPLY_FORMULA_FIXES() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var jsSheet = ss.getSheetByName("Job_Scoring");
  var pgSheet = ss.getSheetByName("Proposal_Generator");

  if (!jsSheet || !pgSheet) {
    ui.alert("Job_Scoring or Proposal_Generator sheet not found.");
    return;
  }

  var jsMap  = getHeaderMap_(jsSheet);
  var pgMap  = getHeaderMap_(pgSheet);
  var jsLast = jsSheet.getLastRow();
  var pgLast = pgSheet.getLastRow();

  var applied = [];
  var skipped = [];

  // ----------------------------------------------------------
  // F4a: Affordability_Check
  // All rows should reference $B$16 (Current_Connect_Balance).
  // Rows 3+ were using $B$14 (Total_Connects_Used) by mistake.
  // Fix: set the correct formula on row 2 and copy down.
  // ----------------------------------------------------------
  var affordCol = getCol_(jsMap, ["Connects_Affordability", "Affordability_Check", "Affordability", "Can_Afford"]);
  if (affordCol && jsLast >= 2) {
    var affordSrc = jsSheet.getRange(2, affordCol);
    affordSrc.setFormula('=IF(P2="","",IF(P2<=Connects_Helper!$B$16,"Can Afford","Cannot Afford"))');
    if (jsLast > 2) {
      affordSrc.copyTo(
        jsSheet.getRange(3, affordCol, jsLast - 2, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
      );
    }
    applied.push("F4a: Affordability_Check uses Current_Connect_Balance ($B$16) in all rows");
  } else {
    skipped.push("F4a: column not found (tried: Connects_Affordability, Affordability_Check, Affordability, Can_Afford)");
  }

  // ----------------------------------------------------------
  // F4b: Final_Decision affordability gate
  // IF(AF2=0,...) compares text to 0 -- never TRUE.
  // Fix: IF(AF2="Cannot Afford",...).
  // Column letters (AG=Total_Score, AF=Affordability_Check,
  // W=Tool_Score, X=Experience_Score, P=Connects_Required)
  // come from the formula export and match the live sheet.
  // ----------------------------------------------------------
  var finalDecCol = getCol_(jsMap, ["Final_Decision"]);
  if (finalDecCol && jsLast >= 2) {
    var finalSrc = jsSheet.getRange(2, finalDecCol);
    finalSrc.setFormula(
      '=IF(AG2="","",IF(AF2="Cannot Afford","SKIP",' +
      'IFS(' +
        'AND(AG2>=0.6,W2>=0.8,X2>=0.7,P2<=14),"APPLY",' +
        'AND(AG2>=0.5,P2<=16),"HOLD",' +
        'TRUE,"SKIP"' +
      ')))'
    );
    if (jsLast > 2) {
      finalSrc.copyTo(
        jsSheet.getRange(3, finalDecCol, jsLast - 2, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
      );
    }
    applied.push("F4b: Final_Decision gate fixed -- Cannot Afford jobs now correctly get SKIP");
  } else {
    skipped.push("F4b: Final_Decision column not found in Job_Scoring");
  }

  // ----------------------------------------------------------
  // F7: Tool_Detected in Proposal_Generator
  // Original formula missed Google Sheets, SQL, and Python.
  // Added after Excel; priority order preserved.
  // ----------------------------------------------------------
  var toolCol = getCol_(pgMap, ["Tool_Detected"]);
  if (toolCol && pgLast >= 2) {
    var toolSrc = pgSheet.getRange(2, toolCol);
    toolSrc.setFormula(
      '=IF(B2="","",IFS(' +
        'ISNUMBER(SEARCH("power bi",C2)),"Power BI",' +
        'ISNUMBER(SEARCH("tableau",C2)),"Tableau",' +
        'ISNUMBER(SEARCH("looker",C2)),"Looker Studio",' +
        'ISNUMBER(SEARCH("google sheets",C2)),"Google Sheets",' +
        'ISNUMBER(SEARCH("excel",C2)),"Excel",' +
        'ISNUMBER(SEARCH("sql",C2)),"SQL",' +
        'ISNUMBER(SEARCH("python",C2)),"Python",' +
        'TRUE,"Unknown"' +
      '))'
    );
    if (pgLast > 2) {
      toolSrc.copyTo(
        pgSheet.getRange(3, toolCol, pgLast - 2, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
      );
    }
    applied.push("F7: Tool_Detected now detects Google Sheets, SQL, Python");
  } else {
    skipped.push("F7: Tool_Detected column not found in Proposal_Generator");
  }

  // ----------------------------------------------------------
  // F9: Portfolio_Project in Proposal_Generator
  // Original formula referenced "Customer Service Analytics
  // Dashboard" and "Email Marketing Analytics Dashboard" --
  // neither exists in the portfolio. Also used short project
  // names that don't match what the AI prompt requires.
  // Fix: use full project names, remove fake projects, expand
  // keyword matching for all three real projects.
  // ----------------------------------------------------------
  var portfolioCol = getCol_(pgMap, ["Portfolio_Project"]);
  if (portfolioCol && pgLast >= 2) {
    var portSrc = pgSheet.getRange(2, portfolioCol);
    portSrc.setFormula(
      '=IF(B2="","",IFS(' +
        'ISNUMBER(SEARCH("inventory",C2)),"Inventory Optimization & Revenue Strategy Dashboard",' +
        'ISNUMBER(SEARCH("stock",C2)),"Inventory Optimization & Revenue Strategy Dashboard",' +
        'ISNUMBER(SEARCH("warehouse",C2)),"Inventory Optimization & Revenue Strategy Dashboard",' +
        'ISNUMBER(SEARCH("reorder",C2)),"Inventory Optimization & Revenue Strategy Dashboard",' +
        'ISNUMBER(SEARCH("supply chain",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("logistics",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("shipping",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("carrier",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("3pl",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("freight",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("delivery",C2)),"3PL Logistics Cost & Performance Analytics Dashboard",' +
        'ISNUMBER(SEARCH("operations",C2)),"Operations Performance & KPI Monitoring Dashboard",' +
        'ISNUMBER(SEARCH("kpi",C2)),"Operations Performance & KPI Monitoring Dashboard",' +
        'ISNUMBER(SEARCH("throughput",C2)),"Operations Performance & KPI Monitoring Dashboard",' +
        'ISNUMBER(SEARCH("performance",C2)),"Operations Performance & KPI Monitoring Dashboard",' +
        'TRUE,"Operations Performance & KPI Monitoring Dashboard"' +
      '))'
    );
    if (pgLast > 2) {
      portSrc.copyTo(
        pgSheet.getRange(3, portfolioCol, pgLast - 2, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
      );
    }
    applied.push("F9: Portfolio_Project uses real project names with expanded keyword matching");
  } else {
    skipped.push("F9: Portfolio_Project column not found in Proposal_Generator");
  }

  var msg = "";
  if (applied.length > 0) {
    msg += "Applied (" + applied.length + "):\n" +
           applied.map(function (s) { return "checkmark " + s; }).join("\n");
  }
  if (skipped.length > 0) {
    msg += (msg ? "\n\n" : "") +
           "Skipped -- column not found (" + skipped.length + "):\n" +
           skipped.map(function (s) { return "warning " + s; }).join("\n");
  }
  ui.alert("Formula Fixes", msg || "Nothing to apply.", ui.ButtonSet.OK);
}
