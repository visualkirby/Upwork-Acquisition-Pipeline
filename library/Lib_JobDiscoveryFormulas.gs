/**
 * ============================================================
 * FreelanceFlow Library -- Job Discovery Pre-Screen Formulas
 * Pure formula-string builders -- headers/primaryTools in,
 * formula strings out, no SpreadsheetApp access. Ported from the
 * proven UpWork_Acquisition_System sheet's Discovery-stage
 * pre-screen chain (Keyword/Tool/Experience/Freshness/Competition/
 * Verification/Client_History/Budget_Quick scores -> weighted
 * Discovery_Priority_Score -> Discovery_Action), with Tool_Detected
 * and Tool_Score made Settings-driven instead of hardcoded to a
 * fixed BI-tool list.
 *
 * Tool_Detected is the one formula that must be built fresh per
 * user (its IFS clauses depend on however many tools *this* user
 * listed in Primary_Tools) -- same precedent as pickPortfolioProject
 * (Lib_WizardFormulas.gs) reading that user's own Projects sheet at
 * call time. REPAIR_FORMULAS() re-calls this if the user updates
 * Primary_Tools later.
 * ============================================================
 */
function buildJobDiscoveryFormulas(headers, primaryToolsCsv, aliasMap) {
  function idx(name) { return headers.indexOf(name); }
  function L(name) {
    var i = idx(name);
    return i >= 0 ? colLetter_(i + 1) : null;
  }

  var result = {};

  var titleL      = L('Job_Title');
  var descL       = L('Description');
  var kwL         = L('Keyword_Search');
  var toolDetIdx  = idx('Tool_Detected');
  var toolDetL    = L('Tool_Detected');
  var expL        = L('Experience_Level');
  var minsL       = L('Minutes_Since_Posted');
  var hoursL      = L('Hours_Since_Posted');
  var daysL       = L('Days_Since_Posted');
  var dateFoundL  = L('Date_Found');
  var propCountL  = L('Proposal_Count');
  var payVerL     = L('Payment_Verified');
  var clientHireL = L('Client_Hires');
  var budgetTypeL = L('Budget_Type');
  var budgetL     = L('Budget');
  var hourlyL     = L('Hourly_Rate');
  var jobLinkL    = L('Job_Link');
  var kwFitL      = L('Keyword_Fit_Score');
  var toolScoreL  = L('Tool_Score');
  var expScoreL   = L('Experience_Score');
  var freshL      = L('Freshness_Score');
  var compL       = L('Competition_Score');
  var verL        = L('Verification_Score');
  var clHistL     = L('Client_History_Score');
  var budgetQkL   = L('Budget_Quick_Score');
  var priorityL   = L('Discovery_Priority_Score');
  var actionL     = L('Discovery_Action');

  // Discovery_ID -- stable per-row identifier so a Proposal_Generator row can
  // be traced back to its exact Job_Discovery source row even when Job_Title
  // repeats (two different jobs can share a title). ROW()-1 needs no counter
  // to maintain and no onEdit bookkeeping; it only shows once Description is
  // filled, matching the pattern other Job_Discovery formulas use to detect
  // "this row has real data."
  var discoveryIdIdx = idx('Discovery_ID');
  if (discoveryIdIdx >= 0 && descL) {
    result.discoveryIdCol = discoveryIdIdx + 1;
    result.discoveryIdFormula = '=IF(' + descL + '2="","",ROW()-1)';
  }

  // Current_Age_Days -- live aging, no Settings dependency. The job's age
  // when logged is Days + Hours + Minutes added together (the sidebar takes
  // all three, e.g. 3 days 14 hours), plus the time since Date_Found. NOW(),
  // not TODAY(): Date_Found carries a time, and TODAY() is midnight, which
  // made a job found that morning show a negative age.
  // Minutes_Since_Posted is optional -- older sheets built before it existed
  // just leave it out of the sum.
  var ageIdx = idx('Current_Age_Days');
  var postedHours = postedHoursExpr_(minsL, hoursL, daysL);
  if (ageIdx >= 0 && dateFoundL && postedHours) {
    result.currentAgeDaysCol = ageIdx + 1;
    result.currentAgeDaysFormula =
      '=IF(OR(' + dateFoundL + '2="",' + postedBlankExpr_(minsL, hoursL, daysL) + '),"",' +
      '(' + postedHours + ')/24+(NOW()-' + dateFoundL + '2))';
  }

  // Keyword_Fit_Score
  var kwFitIdx = idx('Keyword_Fit_Score');
  if (kwFitIdx >= 0 && kwL && descL) {
    result.keywordFitScoreCol = kwFitIdx + 1;
    result.keywordFitScoreFormula =
      '=IF(' + kwL + '2="","",IFS(REGEXMATCH(LOWER(' + descL + '2),LOWER(' + kwL + '2)),1,' +
      'REGEXMATCH(LOWER(' + descL + '2),"dashboard|report|kpi|analytics|visualization"),0.65,TRUE,0.3))';
  }

  // Tool_Detected -- dynamic, one clause per tool the user listed in Settings.
  // See buildToolDetectedIfsArgs_ (Lib_WizardFormulas.gs, shared with
  // Proposal_Generator) for the title-priority, word-boundary matching logic.
  if (toolDetIdx >= 0 && descL) {
    var toolArgs = buildToolDetectedIfsArgs_(titleL, descL, primaryToolsCsv, aliasMap);
    if (toolArgs) {
      result.toolDetectedCol = toolDetIdx + 1;
      result.toolDetectedFormula =
        '=IF(' + descL + '2="","",IFS(' + toolArgs.join(',') + ',TRUE,"Other"))';
    }
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

  // Freshness_Score -- scored on the job's total age when logged (Days +
  // Hours + Minutes), so a filled Minutes box no longer outranks the Days
  // box. Under an hour scores maximum freshness. Optional column, same as above.
  var freshIdx = idx('Freshness_Score');
  if (freshIdx >= 0 && hoursL && daysL) {
    result.freshnessScoreCol = freshIdx + 1;
    result.freshnessScoreFormula = buildFreshnessFormula_(minsL, hoursL, daysL);
  }

  // Competition_Score -- Proposal_Count is a dropdown of Upwork's own
  // ranges, not a raw number.
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

  // Verification_Score
  var verIdx = idx('Verification_Score');
  if (verIdx >= 0 && payVerL) {
    result.verificationScoreCol = verIdx + 1;
    result.verificationScoreFormula =
      '=IF(' + payVerL + '2="","",IF(OR(' + payVerL + '2=TRUE,' + payVerL + '2="Yes"),1,0.3))';
  }

  // Client_History_Score
  var clHistIdx = idx('Client_History_Score');
  if (clHistIdx >= 0 && clientHireL) {
    result.clientHistoryScoreCol = clHistIdx + 1;
    result.clientHistoryScoreFormula =
      '=IF(' + clientHireL + '2="","",IFS(' + clientHireL + '2>=20,1,' + clientHireL + '2>=5,0.7,' +
      clientHireL + '2>=1,0.5,TRUE,0.15))';
  }

  // Budget_Quick_Score
  var budgetQkIdx = idx('Budget_Quick_Score');
  if (budgetQkIdx >= 0 && budgetTypeL && budgetL && hourlyL) {
    result.budgetQuickScoreCol = budgetQkIdx + 1;
    result.budgetQuickScoreFormula =
      '=IF(' + budgetTypeL + '2="","",IFS(' +
      budgetTypeL + '2="Fixed",IFS(' + budgetL + '2>=500,1,' + budgetL + '2>=300,0.8,' + budgetL + '2>=150,0.6,' + budgetL + '2>=75,0.4,TRUE,0.2),' +
      budgetTypeL + '2="Hourly_Range",IFS(' + hourlyL + '2>=30,1,' + hourlyL + '2>=20,0.8,' + hourlyL + '2>=15,0.6,' + hourlyL + '2>=10,0.4,TRUE,0.2),' +
      budgetTypeL + '2="Hourly_Unknown",0.5))';
  }

  // Discovery_Priority_Score
  var priorityIdx = idx('Discovery_Priority_Score');
  if (priorityIdx >= 0 && jobLinkL && kwFitL && toolScoreL && expScoreL && freshL && compL && verL && clHistL && budgetQkL) {
    result.discoveryPriorityScoreCol = priorityIdx + 1;
    result.discoveryPriorityScoreFormula =
      '=IF(' + jobLinkL + '2="","",ROUND(' +
      kwFitL + '2*0.08+' + toolScoreL + '2*0.22+' + expScoreL + '2*0.18+' + freshL + '2*0.14+' +
      compL + '2*0.12+' + verL + '2*0.08+' + clHistL + '2*0.08+' + budgetQkL + '2*0.1,2))';
  }

  // Discovery_Action
  var actionIdx = idx('Discovery_Action');
  if (actionIdx >= 0 && priorityL && toolScoreL && expScoreL) {
    result.discoveryActionCol = actionIdx + 1;
    result.discoveryActionFormula =
      '=IF(' + priorityL + '2="","",IFS(' +
      'AND(' + priorityL + '2>=0.58,' + toolScoreL + '2>=0.8,' + expScoreL + '2>=0.7),"Move to Scoring",' +
      priorityL + '2>=0.45,"Review Later",TRUE,"Skip"))';
  }

  // Discovery_Status -- friendly display alias for Discovery_Action
  var statusIdx = idx('Discovery_Status');
  if (statusIdx >= 0 && actionL) {
    result.discoveryStatusCol = statusIdx + 1;
    result.discoveryStatusFormula =
      '=IF(' + actionL + '2="","",IFS(' +
      actionL + '2="Move to Scoring","Ready",' +
      actionL + '2="Review Later","Review",' +
      actionL + '2="Skip","Skip",TRUE,"New"))';
  }

  return result;
}

// A job's age when logged, in hours: Days + Hours + Minutes added together.
// N() turns a blank box into 0. Any column letter passed as null (an older
// sheet without that column) is left out. Shared with Job_Scoring's
// formulas (Lib_JobScoringFormulas.gs), which have no Minutes column.
function postedHoursExpr_(minsL, hoursL, daysL) {
  var parts = [];
  if (daysL)  parts.push('N(' + daysL + '2)*24');
  if (hoursL) parts.push('N(' + hoursL + '2)');
  if (minsL)  parts.push('N(' + minsL + '2)/60');
  return parts.join('+');
}

// TRUE when none of the age boxes were filled in.
function postedBlankExpr_(minsL, hoursL, daysL) {
  var checks = [];
  [daysL, hoursL, minsL].forEach(function (c) { if (c) checks.push(c + '2=""'); });
  return 'AND(' + checks.join(',') + ')';
}

// Freshness_Score on the job's total age when logged. Same bands as before:
// under an hour 1, under a day 0.9, up to 3 days 1, up to a week 0.8, up to
// two weeks 0.6, older 0.4.
function buildFreshnessFormula_(minsL, hoursL, daysL) {
  var h = '(' + postedHoursExpr_(minsL, hoursL, daysL) + ')';
  return '=IF(' + postedBlankExpr_(minsL, hoursL, daysL) + ',"",IFS(' +
    h + '<1,1,' + h + '<24,0.9,' + h + '<=72,1,' +
    h + '<=168,0.8,' + h + '<=336,0.6,TRUE,0.4))';
}
