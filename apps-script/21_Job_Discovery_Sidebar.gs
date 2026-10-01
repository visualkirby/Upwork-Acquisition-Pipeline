/**
 * ============================================================
 * 21. JOB DISCOVERY SIDEBAR
 * Guided-entry alternative to pasting a job directly into cells --
 * additive, not a replacement. Direct cell paste into Job_Discovery
 * still works exactly as before; this is a second entry path for
 * the highest manual-entry-burden step in the pipeline.
 *
 * Since script-driven setValues()/setValue() writes never fire
 * handleEdit (only real user edits do), job_saveEntry inlines the
 * same side effects the Description/Job_Link paste triggers today
 * (see 14_Edit_Trigger.gs's JOB_DISCOVERY block) rather than relying
 * on the trigger to catch a sidebar-driven write.
 * ============================================================
 */
function openJobDiscoverySidebar_() {
  var html = HtmlService.createHtmlOutputFromFile('LogJobSidebar')
    .setTitle('Log New Job')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function LOG_NEW_JOB() {
  if (!confirmStaleSessionBeforeLogging_(SpreadsheetApp.getActiveSpreadsheet())) return;
  openJobDiscoverySidebar_();
}

function jobDiscovery_getContext() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ctx = buildSessionCountdown_(ss);
  ctx.showWalkthrough = !showWalkthroughSeen_('FF_WALKTHROUGH_JOB_DISCOVERY_SIDEBAR_SEEN');
  return ctx;
}

// Live countdown payload for the sidebar -- { sessionActive, yieldTarget,
// loggedCount, remaining }. remaining/loggedCount are 0 when no session is
// active (the sidebar still works stand-alone, same additive precedent as
// direct cell paste).
function buildSessionCountdown_(ss) {
  var prop         = PropertiesService.getScriptProperties();
  var sessionActive = prop.getProperty('SESSION_ACTIVE') === 'true';
  var yieldTarget    = parseInt(getSettings_()['Session_Yield_Target']) || 8;
  var loggedCount    = sessionActive ? getSessionJobCount_(ss, prop.getProperty('SESSION_ID')) : 0;

  return {
    sessionActive: sessionActive,
    yieldTarget:   yieldTarget,
    loggedCount:   loggedCount,
    remaining:     Math.max(yieldTarget - loggedCount, 0)
  };
}

function job_saveEntry(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Job_Discovery');
  if (!sheet) return { ok: false, message: 'Job_Discovery sheet not found.' };

  var map        = getHeaderMap_(sheet);
  var titleCol   = getCol_(map, ['Job_Title']);
  var row        = findFirstEmptyRowByColumn_(sheet, titleCol);

  var cleanedDescription = cleanJobText_(data.description);

  setCellValue_(sheet, row, map, ['Job_Title'], data.jobTitle);
  setCellValue_(sheet, row, map, ['Description'], cleanedDescription);
  setCellValue_(sheet, row, map, ['Client_Name', 'Client Name'], data.clientName);
  setCellValue_(sheet, row, map, ['Job_Link'], data.jobLink);
  setCellValue_(sheet, row, map, ['Keyword_Search'], data.keywordSearch);
  recomputeKeywordStrategyActualCount_(ss, data.keywordSearch);
  setCellValue_(sheet, row, map, ['Experience_Level'], data.experienceLevel);
  setCellValue_(sheet, row, map, ['Proposal_Count'], data.proposalCount);
  setCellValue_(sheet, row, map, ['Budget_Type'], data.budgetType);
  setCellValue_(sheet, row, map, ['Payment_Verified'], data.paymentVerified);
  setCellValue_(sheet, row, map, ['Budget'], data.budget);
  // Upwork shows hourly jobs as a range ("$25 - $50/hr"). The range is kept
  // in Hourly_Rate_Min/Max; Hourly_Rate holds its midpoint, which is what
  // Budget_Quick_Score and Job_Scoring's Estimated_Hourly_Rate score. A
  // single rate goes in either box.
  if (ensureHourlyRangeColumns_(sheet)) map = getHeaderMap_(sheet);
  var rateMin = Number(data.hourlyRateMin) || 0;
  var rateMax = Number(data.hourlyRateMax) || 0;
  var rate    = (rateMin && rateMax) ? (rateMin + rateMax) / 2 : (rateMin || rateMax);
  setCellValue_(sheet, row, map, ['Hourly_Rate'], rate);
  setCellValue_(sheet, row, map, ['Hourly_Rate_Min'], rateMin || '');
  setCellValue_(sheet, row, map, ['Hourly_Rate_Max'], rateMax || '');
  setCellValue_(sheet, row, map, ['Client_Hires'], data.clientHires);
  setCellValue_(sheet, row, map, ['Connects_Required'], data.connectsRequired);
  setCellValue_(sheet, row, map, ['Days_Since_Posted'], data.daysSincePosted);
  setCellValue_(sheet, row, map, ['Hours_Since_Posted'], data.hoursSincePosted);
  setCellValue_(sheet, row, map, ['Minutes_Since_Posted'], data.minutesSincePosted);
  setCellValue_(sheet, row, map, ['Date_Found'], new Date());

  var prop   = PropertiesService.getScriptProperties();
  var active = prop.getProperty('SESSION_ACTIVE');
  var sessionId = prop.getProperty('SESSION_ID');
  if (active === 'true' && sessionId) {
    setCellValue_(sheet, row, map, ['Session_ID'], sessionId);
  }

  var fitNotes         = null;
  var duplicateWarning = '';

  var aiFitNotesCol = getCol_(map, ['AI_Fit_Notes']);
  if (aiFitNotesCol) {
    sheet.getRange(row, aiFitNotesCol).setValue('Analyzing...');
    var qnApiKey   = PropertiesService.getScriptProperties().getProperty('UPWORK_OPENAI_API_KEY');
    var qnSettings = getSettings_();
    var quickResult;
    try {
      quickResult = FFLib.getQuickNotes(cleanedDescription, qnApiKey, qnSettings);
    } catch (err) {
      quickResult = err.message;
    }
    sheet.getRange(row, aiFitNotesCol).setValue(quickResult);

    // Fit notes, the duplicate warning, and the session milestones all go
    // back to the sidebar instead of popping up: a popup only renders while
    // the Sheets tab is in front, so with the freelancer on Upwork the save
    // sat waiting on a popup they couldn't see.
    fitNotes = parseAiFitNotes_(quickResult);
  }

  if (active === 'true' && data.jobLink) {
    var linkCol = getCol_(map, ['Job_Link']);
    if (linkCol) {
      var lastRow2   = sheet.getLastRow();
      var linkValues = sheet.getRange(2, linkCol, lastRow2 - 1, 1).getValues();
      var linkStr    = String(data.jobLink).trim();
      var matchCount = 0;
      for (var i = 0; i < linkValues.length; i++) {
        if (String(linkValues[i][0]).trim() === linkStr) matchCount++;
      }
      if (matchCount > 1) {
        var dupeCount = parseInt(prop.getProperty('SESSION_DUPE_COUNT') || '0', 10);
        prop.setProperty('SESSION_DUPE_COUNT', String(dupeCount + 1));
        duplicateWarning = 'Duplicate: this job link is already in Job_Discovery. ' +
          'You can delete the new row and skip to the next job. Duplicates this session: ' + (dupeCount + 1) + '.';
      }
    }
    colorDuplicateJobLinks();
  }

  // Script writes never fire handleEdit, so the Proposal_Generator sync its
  // Job_Discovery block runs has to happen here too.
  syncProposalGenerator_(ss);

  var halfwayTip     = handleSessionHalfwayReached_(ss, true);
  var sessionSummary = handleSessionYieldReached_(ss, true);

  var result = buildSessionCountdown_(ss);
  result.ok               = true;
  result.sessionComplete  = !!sessionSummary;
  result.sessionSummary   = sessionSummary;
  result.halfwayTip       = halfwayTip;
  result.duplicateWarning = duplicateWarning;
  result.fitNotes         = fitNotes;
  return result;
}

// Appends Hourly_Rate_Min / Hourly_Rate_Max to a Job_Discovery created
// before they existed (added at the end, so no existing column moves).
// Returns true when it added either one.
function ensureHourlyRangeColumns_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  var missing = ['Hourly_Rate_Min', 'Hourly_Rate_Max'].filter(function (h) {
    return headers.indexOf(h) === -1;
  });
  if (missing.length === 0) return false;
  sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length)
    .setValues([missing]).setFontWeight('bold');
  return true;
}
