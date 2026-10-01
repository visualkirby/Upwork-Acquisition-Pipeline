/**
 * ============================================================
 * FreelanceFlow Library -- Job Scoring Formulas
 * Pure formula-string builders -- headers/primaryTools in,
 * formula strings out, no SpreadsheetApp access. Ported from the
 * proven UpWork_Acquisition_System sheet's Job_Scoring chain
 * (Budget/Keyword/Tool/Experience/Freshness/Competition/
 * Client_History/Scope/Portfolio scores -> weighted Total_Score
 * -> Final_Decision), with Tool_Score made Settings-driven, dropdown
 * labels matched to Upwork's real UI, and Effort_Level/Scope_Rating/
 * Portfolio_Match auto-parsed from the AI_Fit_Notes classifier
 * output instead of manually re-typed.
 *
 * Tool_Score is the one formula that must be built fresh per user
 * (its IFS clauses/VLOOKUP depend on that user's own Primary_Tools)
 * -- same precedent as buildJobDiscoveryFormulas' Tool_Detected.
 * ============================================================
 */
// Upwork's Proposal_Count ranges and the lowest count each one can mean.
// Shared by Final_Decision's Apply_Max_Proposals check (formula) and the
// client's Proposal_Generator tagging (proposalCountFloor), so both read a
// range the same way. A job is over the cap when its floor is at or above
// Apply_Max_Proposals: with a cap of 15, "15 to 20" is over and "10 to 15"
// is not.
var PROPOSAL_COUNT_FLOORS_ = [
  ['Fewer than 5', 0], ['5 to 10', 5], ['10 to 15', 10],
  ['15 to 20', 15], ['20 to 50', 20], ['50+', 50]
];

// Floor for one Proposal_Count value: a range label, a plain number, or -1
// when it's blank or unrecognized (never over any cap).
function proposalCountFloor(value) {
  var v = String(value === null || value === undefined ? '' : value).trim();
  for (var i = 0; i < PROPOSAL_COUNT_FLOORS_.length; i++) {
    if (PROPOSAL_COUNT_FLOORS_[i][0] === v) return PROPOSAL_COUNT_FLOORS_[i][1];
  }
  var n = Number(v);
  return v !== '' && !isNaN(n) ? n : -1;
}

// A Settings value, or fallback when the row is missing or blank. A bare
// VLOOKUP on a blank cell reads as 0, which would make every cap fire.
function settingOr_(name, fallback) {
  var look = 'VLOOKUP("' + name + '",Settings!$A:$B,2,0)';
  return 'IF(IFERROR(' + look + ',"")="",' + fallback + ',' + look + ')';
}

// pgHeaders (optional) are Proposal_Generator's headers. With them,
// Final_Decision exempts a job that already has a Proposal_Generator row
// from the age and proposal caps, so jobs you already applied to stay
// APPLY as they age instead of rewriting your history to SKIP. Without them
// (the Setup Wizard builds Job_Scoring before Proposal_Generator exists),
// there's no exemption until ensurePipelineSheets_ re-applies it.
function buildJobScoringFormulas(headers, pgHeaders) {
  function idx(name) { return headers.indexOf(name); }
  function L(name) {
    var i = idx(name);
    return i >= 0 ? colLetter_(i + 1) : null;
  }

  var result = {};

  var dateScoredL     = L('Date_Scored');
  var descL           = L('Description');
  var kwL             = L('Keyword_Search');
  var toolDetL        = L('Tool_Detected');
  var expL            = L('Experience_Level');
  var hoursL          = L('Hours_Since_Posted');
  var daysL           = L('Days_Since_Posted');
  var propCountL      = L('Proposal_Count');
  var payVerL         = L('Payment_Verified');
  var clientHiresL    = L('Client_Hires');
  var budgetTypeL     = L('Budget_Type');
  var budgetL         = L('Budget');
  var hourlyRateL     = L('Hourly_Rate');
  var connectsReqL    = L('Connects_Required');
  var effortL         = L('Effort_Level');
  var estHoursL       = L('Estimated_Hours');
  var estHourlyRateL  = L('Estimated_Hourly_Rate');
  var budgetScoreL    = L('Budget_Score');
  var keywordScoreL   = L('Keyword_Score');
  var toolScoreL      = L('Tool_Score');
  var expScoreL       = L('Experience_Score');
  var freshL          = L('Freshness_Score');
  var compL           = L('Competition_Score');
  var clHistL         = L('Client_History_Score');
  var scopeRatingL    = L('Scope_Rating');
  var scopeScoreL     = L('Scope_Score');
  var portfolioMatchL = L('Portfolio_Match');
  var portfolioScoreL = L('Portfolio_Score');
  var connAffordL     = L('Connects_Affordability');
  var totalScoreL     = L('Total_Score');
  var notesL          = L('AI_Fit_Notes');

  // Current_Age_Days -- anchored on Date_Scored (Job_Scoring's own
  // re-entry point), not Date_Found.
  var ageIdx = idx('Current_Age_Days');
  if (ageIdx >= 0 && dateScoredL && hoursL && daysL) {
    result.currentAgeDaysCol = ageIdx + 1;
    result.currentAgeDaysFormula =
      '=IF(' + dateScoredL + '2="","",IF(' + hoursL + '2<>"",(' + hoursL + '2/24)+(TODAY()-' + dateScoredL + '2),' +
      'IF(' + daysL + '2<>"",' + daysL + '2+(TODAY()-' + dateScoredL + '2),"")))';
  }

  // Effort_Level / Scope_Rating / Portfolio_Match -- auto-parsed from
  // "Effort, Scope & Portfolio Match" written by FFLib.getQuickNotes
  // into AI_Fit_Notes. No longer manual dropdowns.
  var effortIdx = idx('Effort_Level');
  if (effortIdx >= 0 && notesL) {
    result.effortLevelCol = effortIdx + 1;
    result.effortLevelFormula =
      '=IF(' + notesL + '2="","",TRIM(LEFT(' + notesL + '2,FIND(",",' + notesL + '2&",")-1)))';
  }

  var scopeRatingIdx = idx('Scope_Rating');
  if (scopeRatingIdx >= 0 && notesL) {
    result.scopeRatingCol = scopeRatingIdx + 1;
    result.scopeRatingFormula =
      '=IF(' + notesL + '2="","",TRIM(MID(' + notesL + '2,FIND(",",' + notesL + '2&",")+1,' +
      'FIND("&",' + notesL + '2&"&")-FIND(",",' + notesL + '2&",")-1)))';
  }

  var portfolioMatchIdx = idx('Portfolio_Match');
  if (portfolioMatchIdx >= 0 && notesL) {
    result.portfolioMatchCol = portfolioMatchIdx + 1;
    result.portfolioMatchFormula =
      '=IF(' + notesL + '2="","",TRIM(REGEXREPLACE(MID(' + notesL + '2,FIND("&",' + notesL + '2&"&")+1,999),"[.!]+$","")))';
  }

  // Estimated_Hours -- derived from Effort_Level
  var estHoursIdx = idx('Estimated_Hours');
  if (estHoursIdx >= 0 && effortL) {
    result.estimatedHoursCol = estHoursIdx + 1;
    result.estimatedHoursFormula =
      '=IFS(' + effortL + '2="Simple",4,' + effortL + '2="Normal",8,' +
      effortL + '2="Complex",15,' + effortL + '2="Large",25,TRUE,"")';
  }

  // Estimated_Hourly_Rate -- Budget/Estimated_Hours for Fixed jobs, else Hourly_Rate directly
  // (blank for Hourly_Unknown -- Budget_Score below gives that its own neutral score)
  var estHourlyRateIdx = idx('Estimated_Hourly_Rate');
  if (estHourlyRateIdx >= 0 && budgetTypeL && budgetL && estHoursL && hourlyRateL) {
    result.estimatedHourlyRateCol = estHourlyRateIdx + 1;
    result.estimatedHourlyRateFormula =
      '=IF(' + budgetTypeL + '2="","",IF(' + budgetTypeL + '2="Fixed",' +
      'IF(OR(' + budgetL + '2="",' + estHoursL + '2=""),"",' + budgetL + '2/' + estHoursL + '2),' +
      'IF(' + budgetTypeL + '2="Hourly_Unknown","",' + hourlyRateL + '2)))';
  }

  // Budget_Score -- scores the effective/estimated hourly rate. Hourly_Unknown gets its own
  // neutral score (no rate to judge), same precedent as Job_Discovery's Budget_Quick_Score.
  var budgetScoreIdx = idx('Budget_Score');
  if (budgetScoreIdx >= 0 && estHourlyRateL && budgetTypeL) {
    result.budgetScoreCol = budgetScoreIdx + 1;
    result.budgetScoreFormula =
      '=IF(' + budgetTypeL + '2="","",IF(' + budgetTypeL + '2="Hourly_Unknown",0.5,IF(' +
      estHourlyRateL + '2="","",IFS(' +
      estHourlyRateL + '2>=50,1,' + estHourlyRateL + '2>=35,0.85,' + estHourlyRateL + '2>=25,0.7,' +
      estHourlyRateL + '2>=18,0.55,' + estHourlyRateL + '2>=12,0.4,TRUE,0.25))))';
  }

  // Keyword_Score
  var keywordScoreIdx = idx('Keyword_Score');
  if (keywordScoreIdx >= 0 && kwL && descL) {
    result.keywordScoreCol = keywordScoreIdx + 1;
    result.keywordScoreFormula =
      '=IF(' + kwL + '2="","",IFS(REGEXMATCH(LOWER(' + descL + '2),LOWER(' + kwL + '2)),0.8,' +
      'REGEXMATCH(LOWER(' + descL + '2),"dashboard|report|kpi|analytics|visualization"),0.55,TRUE,0.25))';
  }

  // Tool_Score -- Settings-driven: is the detected tool one of the user's own?
  var toolScoreIdx = idx('Tool_Score');
  if (toolScoreIdx >= 0 && toolDetL) {
    result.toolScoreCol = toolScoreIdx + 1;
    result.toolScoreFormula =
      '=IF(' + toolDetL + '2="","",IF(ISNUMBER(SEARCH(LOWER(' + toolDetL + '2),' +
      'LOWER(VLOOKUP("Primary_Tools",Settings!$A:$B,2,0)))),1,IF(' + toolDetL + '2="Other",0.4,0.6)))';
  }

  // Experience_Score -- relative to the freelancer's own profile level, see
  // buildExperienceScoreFormula_ in Lib_WizardFormulas.gs.
  var expScoreIdx = idx('Experience_Score');
  if (expScoreIdx >= 0 && expL) {
    result.experienceScoreCol = expScoreIdx + 1;
    result.experienceScoreFormula = buildExperienceScoreFormula_(expL);
  }

  // Freshness_Score -- no Minutes_Since_Posted column in Job_Scoring
  var freshIdx = idx('Freshness_Score');
  if (freshIdx >= 0 && hoursL && daysL) {
    result.freshnessScoreCol = freshIdx + 1;
    result.freshnessScoreFormula =
      '=IF(AND(' + hoursL + '2="",' + daysL + '2=""),"",IFS(' +
      'AND(' + hoursL + '2<>"",' + hoursL + '2<=24),0.9,' +
      'AND(' + hoursL + '2<>"",' + hoursL + '2<=72),1,' +
      daysL + '2<=3,1,' + daysL + '2<=7,0.8,' + daysL + '2<=14,0.6,TRUE,0.4))';
  }

  // Competition_Score -- Proposal_Count is a dropdown of Upwork's own ranges
  var compIdx = idx('Competition_Score');
  if (compIdx >= 0 && propCountL) {
    result.competitionScoreCol = compIdx + 1;
    result.competitionScoreFormula =
      '=IF(' + propCountL + '2="","",IFS(' +
      propCountL + '2="Fewer than 5",1,' +
      propCountL + '2="5 to 10",0.8,' +
      propCountL + '2="10 to 15",0.6,' +
      propCountL + '2="15 to 20",0.4,' +
      propCountL + '2="20 to 50",0.2,' +
      propCountL + '2="50+",0.1,' +
      'TRUE,0.5))';
  }

  // Client_History_Score
  var clHistIdx = idx('Client_History_Score');
  if (clHistIdx >= 0 && clientHiresL) {
    result.clientHistoryScoreCol = clHistIdx + 1;
    result.clientHistoryScoreFormula =
      '=IF(' + clientHiresL + '2="","",IFS(' + clientHiresL + '2>=20,1,' + clientHiresL + '2>=5,0.7,' +
      clientHiresL + '2>=1,0.5,TRUE,0.15))';
  }

  // Scope_Score -- derived from Scope_Rating
  var scopeScoreIdx = idx('Scope_Score');
  if (scopeScoreIdx >= 0 && scopeRatingL) {
    result.scopeScoreCol = scopeScoreIdx + 1;
    result.scopeScoreFormula =
      '=IF(' + scopeRatingL + '2="","",IFS(' +
      scopeRatingL + '2="Clear",1,' + scopeRatingL + '2="Mostly Clear",0.8,' +
      scopeRatingL + '2="Vague",0.5,' + scopeRatingL + '2="Very Vague",0.2,TRUE,0.2))';
  }

  // Portfolio_Score -- derived from Portfolio_Match
  var portfolioScoreIdx = idx('Portfolio_Score');
  if (portfolioScoreIdx >= 0 && portfolioMatchL) {
    result.portfolioScoreCol = portfolioScoreIdx + 1;
    result.portfolioScoreFormula =
      '=IF(' + portfolioMatchL + '2="","",IFS(' +
      portfolioMatchL + '2="Exact",1,' + portfolioMatchL + '2="Strong",0.85,' +
      portfolioMatchL + '2="Partial",0.6,' + portfolioMatchL + '2="Weak",0.35,' +
      portfolioMatchL + '2="None",0.1,TRUE,0.1))';
  }

  // Connects_Affordability -- VLOOKUP against Connects_Helper's Current_Connect_Balance,
  // never a hardcoded cell reference
  var connAffordIdx = idx('Connects_Affordability');
  if (connAffordIdx >= 0 && connectsReqL) {
    result.connectsAffordabilityCol = connAffordIdx + 1;
    result.connectsAffordabilityFormula =
      '=IF(' + connectsReqL + '2="","",IF(' + connectsReqL + '2<=' +
      'VLOOKUP("Current_Connect_Balance",Connects_Helper!$A:$B,2,0),"Can Afford","Cannot Afford"))';
  }

  // Total_Score -- weighted composite, weights sum to 1.0
  var totalScoreIdx = idx('Total_Score');
  if (totalScoreIdx >= 0 && portfolioScoreL && budgetScoreL && keywordScoreL && toolScoreL &&
      expScoreL && freshL && compL && clHistL && scopeScoreL) {
    result.totalScoreCol = totalScoreIdx + 1;
    result.totalScoreFormula =
      '=IF(' + portfolioScoreL + '2="","",ROUND(' +
      budgetScoreL + '2*0.14+' + keywordScoreL + '2*0.06+' + toolScoreL + '2*0.2+' +
      expScoreL + '2*0.18+' + freshL + '2*0.12+' + compL + '2*0.1+' +
      clHistL + '2*0.06+' + scopeScoreL + '2*0.07+' + portfolioScoreL + '2*0.07,2))';
  }

  // Score_Per_Connect
  var scorePerConnectIdx = idx('Score_Per_Connect');
  if (scorePerConnectIdx >= 0 && totalScoreL && connectsReqL) {
    result.scorePerConnectCol = scorePerConnectIdx + 1;
    result.scorePerConnectFormula =
      '=IF(OR(' + totalScoreL + '2="",' + connectsReqL + '2=""),"",' + totalScoreL + '2/' + connectsReqL + '2)';
  }

  // Final_Decision -- Settings-driven thresholds (Apply_Min_Score, Apply_Min_Tool_Score,
  // Apply_Min_Exp_Score, Apply_Max_Connects, Hold_Min_Score, Hold_Max_Connects), never hardcoded.
  // A job past Apply_Max_Age_Days or Apply_Max_Proposals is SKIP unless it
  // already has a Proposal_Generator row (see pgHeaders above). A blank
  // Current_Age_Days or Proposal_Count never trips a cap.
  var finalDecisionIdx = idx('Final_Decision');
  if (finalDecisionIdx >= 0 && totalScoreL && connAffordL && toolScoreL && expScoreL && connectsReqL) {
    var idL    = L('Discovery_ID');
    var ageL   = L('Current_Age_Days');
    var pgIdIx = pgHeaders ? pgHeaders.indexOf('Discovery_ID') : -1;
    var pgIdL  = pgIdIx >= 0 ? colLetter_(pgIdIx + 1) : null;

    var inPg = (idL && pgIdL)
      ? 'ISNUMBER(MATCH(' + idL + '2,Proposal_Generator!$' + pgIdL + ':$' + pgIdL + ',0))'
      : 'FALSE';
    var tooOld = ageL
      ? 'AND(ISNUMBER(' + ageL + '2),' + ageL + '2>' + settingOr_('Apply_Max_Age_Days', 9999) + ')'
      : 'FALSE';
    var floorExpr = propCountL
      ? 'IFS(' + PROPOSAL_COUNT_FLOORS_.map(function (f) {
          return propCountL + '2="' + f[0] + '",' + f[1];
        }).join(',') + ',ISNUMBER(' + propCountL + '2),' + propCountL + '2,TRUE,-1)'
      : '-1';
    var tooMany = '(' + floorExpr + ')>=' + settingOr_('Apply_Max_Proposals', 9999);

    result.finalDecisionCol = finalDecisionIdx + 1;
    result.finalDecisionFormula =
      '=IF(' + totalScoreL + '2="","",' +
      'IF(AND(NOT(' + inPg + '),OR(' + tooOld + ',' + tooMany + ')),"SKIP",' +
      'IF(' + connAffordL + '2="Cannot Afford","SKIP",IFS(' +
      'AND(' + totalScoreL + '2>=VLOOKUP("Apply_Min_Score",Settings!$A:$B,2,0),' +
      toolScoreL + '2>=VLOOKUP("Apply_Min_Tool_Score",Settings!$A:$B,2,0),' +
      expScoreL + '2>=VLOOKUP("Apply_Min_Exp_Score",Settings!$A:$B,2,0),' +
      connectsReqL + '2<=VLOOKUP("Apply_Max_Connects",Settings!$A:$B,2,0)),"APPLY",' +
      'AND(' + totalScoreL + '2>=VLOOKUP("Hold_Min_Score",Settings!$A:$B,2,0),' +
      connectsReqL + '2<=VLOOKUP("Hold_Max_Connects",Settings!$A:$B,2,0)),"HOLD",' +
      'TRUE,"SKIP"))))';
  }

  // Proposal_Generator_Date is not built here. It's a Discovery_ID lookup of
  // Proposal_Generator's static Date (buildJobScoringProposalDateLookup in
  // Lib_WizardFormulas.gs), applied once all three pipeline sheets exist.

  return result;
}

/**
 * Job_Scoring's raw-data columns (Job_Title, Description, Budget, etc.) are
 * meant to auto-populate from Job_Discovery wherever Discovery_Action="Move
 * to Scoring" -- ported from the proven UpWork_Acquisition_System sheet's
 * Job_Scoring!A2 FILTER. Rebuilt here as a dynamic, header-driven version so
 * it can't silently break the way the hand-typed original did when a new
 * Job_Discovery column shifted every reference after it.
 * Additional_Questions is excluded -- it flows backwards from
 * Proposal_Generator instead (see buildDiscoveryIdLookup_ in
 * Lib_WizardFormulas.gs), since it's only knowable once the freelancer is on
 * Upwork's submission page, one stage past Job_Scoring.
 */
function buildJobScoringPullFormulas(jobDiscoveryHeaders, jobScoringHeaders) {
  var fieldPairs = [
    ['Discovery_ID', 'Discovery_ID'],
    ['Date_Scored', 'Date_Found'],
    ['Job_Title', 'Job_Title'],
    ['Description', 'Description'],
    ['Client_Name', 'Client_Name'],
    ['Keyword_Search', 'Keyword_Search'],
    ['Tool_Detected', 'Tool_Detected'],
    ['Experience_Level', 'Experience_Level'],
    ['Hours_Since_Posted', 'Hours_Since_Posted'],
    ['Days_Since_Posted', 'Days_Since_Posted'],
    ['Proposal_Count', 'Proposal_Count'],
    ['Payment_Verified', 'Payment_Verified'],
    ['Client_Hires', 'Client_Hires'],
    ['Budget_Type', 'Budget_Type'],
    ['Budget', 'Budget'],
    ['Hourly_Rate', 'Hourly_Rate'],
    ['Connects_Required', 'Connects_Required'],
    ['Job_Link', 'Job_Link'],
    ['AI_Fit_Notes', 'AI_Fit_Notes']
  ];

  return buildFilterPull_(
    fieldPairs, 'Job_Discovery', jobDiscoveryHeaders, jobScoringHeaders,
    'Discovery_Action', 'Move to Scoring'
  );
}
