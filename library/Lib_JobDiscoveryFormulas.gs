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
function buildJobDiscoveryFormulas(headers, primaryToolsCsv) {
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

  // Current_Age_Days -- live aging, no Settings dependency. Minutes takes
  // priority (job posted under an hour ago), then Hours, then Days.
  // Minutes_Since_Posted is optional -- older sheets built before it existed
  // fall back to the Hours/Days-only version instead of breaking.
  var ageIdx = idx('Current_Age_Days');
  if (ageIdx >= 0 && dateFoundL) {
    result.currentAgeDaysCol = ageIdx + 1;
    var hoursOrDays =
      'IF(' + hoursL + '2<>"",(' + hoursL + '2/24)+(TODAY()-' + dateFoundL + '2),' +
      'IF(' + daysL + '2<>"",' + daysL + '2+(TODAY()-' + dateFoundL + '2),""))';
    result.currentAgeDaysFormula = minsL
      ? '=IF(' + dateFoundL + '2="","",IF(' + minsL + '2<>"",(' + minsL + '2/1440)+(TODAY()-' + dateFoundL + '2),' + hoursOrDays + '))'
      : '=IF(' + dateFoundL + '2="","",' + hoursOrDays + ')';
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
    var toolArgs = buildToolDetectedIfsArgs_(titleL, descL, primaryToolsCsv);
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

  // Freshness_Score -- Minutes_Since_Posted (job posted under an hour ago)
  // always scores maximum freshness. Optional column, same as above.
  var freshIdx = idx('Freshness_Score');
  if (freshIdx >= 0 && hoursL && daysL) {
    result.freshnessScoreCol = freshIdx + 1;
    var blankCheck = minsL
      ? 'AND(' + minsL + '2="",' + hoursL + '2="",' + daysL + '2="")'
      : 'AND(' + hoursL + '2="",' + daysL + '2="")';
    var minsClause = minsL ? (minsL + '2<>"",1,') : '';
    result.freshnessScoreFormula =
      '=IF(' + blankCheck + ',"",IFS(' + minsClause +
      'AND(' + hoursL + '2<>"",' + hoursL + '2<=24),0.9,' +
      'AND(' + hoursL + '2<>"",' + hoursL + '2<=72),1,' +
      daysL + '2<=3,1,' + daysL + '2<=7,0.8,' + daysL + '2<=14,0.6,TRUE,0.4))';
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
