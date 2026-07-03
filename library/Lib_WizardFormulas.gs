/**
 * ============================================================
 * FreelanceFlow Library -- Wizard Formula Builders
 * Pure formula-string builders -- headers/portfolioMap in,
 * formula-string out, no SpreadsheetApp access. The thin client
 * (00_Setup_Wizard.gs's apply*_ functions at setup time, and
 * 15_Formula_Fixes.gs's REPAIR_FORMULAS() for an existing sheet)
 * calls these and writes the returned formula strings via
 * setFormula(). This is the actual "how do we build a working
 * pipeline automatically" mechanism -- the highest-value piece
 * of IP to keep out of customers' hands.
 *
 * buildJobScoringFormulas and buildProposalGeneratorFormulas are
 * the two entry points the client calls, so both are public (no
 * trailing underscore). buildPortfolioFormula_ and colLetter_ are
 * only ever called internally by those two, so they keep the
 * underscore and stay private to the Library.
 * ============================================================
 */
function buildJobScoringFormulas(headers) {
  var connReqIdx  = headers.indexOf('Connects_Required');
  var connAffIdx  = headers.indexOf('Connects_Affordability');
  var toolScoIdx  = headers.indexOf('Tool_Score');
  var expScoIdx   = headers.indexOf('Experience_Score');
  var totScoIdx   = headers.indexOf('Total_Score');
  var finalDecIdx = headers.indexOf('Final_Decision');

  var result = {};

  if (connAffIdx >= 0 && connReqIdx >= 0) {
    var crL = colLetter_(connReqIdx + 1);
    result.connectsAffordabilityCol = connAffIdx + 1;
    result.connectsAffordabilityFormula =
      '=IF(' + crL + '2="","",IF(' + crL + '2<=Connects_Helper!$B$2,"Can Afford","Cannot Afford"))';
  }

  if (finalDecIdx >= 0 && totScoIdx >= 0 && connAffIdx >= 0 && toolScoIdx >= 0 && expScoIdx >= 0 && connReqIdx >= 0) {
    var totL  = colLetter_(totScoIdx + 1);
    var affL  = colLetter_(connAffIdx + 1);
    var toolL = colLetter_(toolScoIdx + 1);
    var expL  = colLetter_(expScoIdx + 1);
    var crL2  = colLetter_(connReqIdx + 1);
    result.finalDecisionCol = finalDecIdx + 1;
    result.finalDecisionFormula =
      '=IF(' + totL + '2="","",IF(' + affL + '2="Cannot Afford","SKIP",' +
      'IFS(' +
        'AND(' +
          totL  + '2>=VLOOKUP("Apply_Min_Score",Settings!$A:$B,2,0),' +
          toolL + '2>=VLOOKUP("Apply_Min_Tool_Score",Settings!$A:$B,2,0),' +
          expL  + '2>=VLOOKUP("Apply_Min_Exp_Score",Settings!$A:$B,2,0),' +
          crL2  + '2<=VLOOKUP("Apply_Max_Connects",Settings!$A:$B,2,0)' +
        '),"APPLY",' +
        'AND(' +
          totL + '2>=VLOOKUP("Hold_Min_Score",Settings!$A:$B,2,0),' +
          crL2 + '2<=VLOOKUP("Hold_Max_Connects",Settings!$A:$B,2,0)' +
        '),"HOLD",' +
        'TRUE,"SKIP"' +
      ')))';
  }

  return result;
}


function buildProposalGeneratorFormulas(headers, portfolioMap) {
  var jobTitleIdx  = headers.indexOf('Job_Title');
  var descIdx      = headers.indexOf('Description');
  var toolDetIdx   = headers.indexOf('Tool_Detected');
  var portfolioIdx = headers.indexOf('Portfolio_Project');

  var result = {};
  if (jobTitleIdx < 0 || descIdx < 0) return result;

  var jtL = colLetter_(jobTitleIdx + 1);
  var dcL = colLetter_(descIdx + 1);

  if (toolDetIdx >= 0) {
    result.toolDetectedCol = toolDetIdx + 1;
    result.toolDetectedFormula =
      '=IF(' + jtL + '2="","",IFS(' +
        'ISNUMBER(SEARCH("power bi",'    + dcL + '2)),"Power BI",' +
        'ISNUMBER(SEARCH("tableau",'     + dcL + '2)),"Tableau",' +
        'ISNUMBER(SEARCH("looker",'      + dcL + '2)),"Looker Studio",' +
        'ISNUMBER(SEARCH("google sheets",' + dcL + '2)),"Google Sheets",' +
        'ISNUMBER(SEARCH("excel",'       + dcL + '2)),"Excel",' +
        'ISNUMBER(SEARCH("sql",'         + dcL + '2)),"SQL",' +
        'ISNUMBER(SEARCH("python",'      + dcL + '2)),"Python",' +
        'ISNUMBER(SEARCH("bigquery",'    + dcL + '2)),"BigQuery",' +
        'ISNUMBER(SEARCH("dbt",'         + dcL + '2)),"dbt",' +
        'ISNUMBER(SEARCH("snowflake",'   + dcL + '2)),"Snowflake",' +
        'TRUE,"Unknown"' +
      '))';
  }

  if (portfolioIdx >= 0) {
    result.portfolioProjectCol     = portfolioIdx + 1;
    result.portfolioProjectFormula = buildPortfolioFormula_(jtL, dcL, portfolioMap);
  }

  return result;
}


function buildPortfolioFormula_(jobTitleLetter, descLetter, portfolioMap) {
  var jtL = jobTitleLetter || 'B';
  var dcL = descLetter     || 'D';
  var map = portfolioMap   || {};

  var clauses     = [];
  var defaultName = '';
  var pNums       = Object.keys(map).sort();

  if (pNums.length === 0) {
    return '=IF(' + jtL + '2="","","Add portfolio projects via FreelanceFlow Setup")';
  }

  for (var p = 0; p < pNums.length; p++) {
    var proj = map[pNums[p]];
    if (!proj.name) continue;
    if (!defaultName) defaultName = proj.name;
    var safeName = proj.name.replace(/"/g, '""');
    for (var k = 0; k < proj.keywords.length; k++) {
      var kw = proj.keywords[k].replace(/"/g, '""');
      if (kw) clauses.push('ISNUMBER(SEARCH("' + kw + '",' + dcL + '2)),"' + safeName + '"');
    }
  }

  if (clauses.length === 0) {
    return '=IF(' + jtL + '2="","","Add portfolio keyword mappings in Settings sheet")';
  }

  var safeDefault = (defaultName || 'Portfolio Project').replace(/"/g, '""');
  return '=IF(' + jtL + '2="","",IFS(' + clauses.join(',') + ',TRUE,"' + safeDefault + '"))';
}


function colLetter_(n) {
  var s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}
