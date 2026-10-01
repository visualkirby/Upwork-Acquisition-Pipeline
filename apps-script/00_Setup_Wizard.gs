/**
 * ============================================================
 * 00. SETUP WIZARD
 * Defines onOpen for the entire spreadsheet.
 * Do NOT define onOpen in any other module.
 * ============================================================
 */

// How many rows to prefill scoring/detection formulas down when a
// pipeline sheet is created, so they're already live before the user
// pastes in their first job -- rather than only existing in row 2.
var FORMULA_PREFILL_ROWS = 500;

function onOpen() {
  buildSystemMenu_();
}

function openSetupWizard_() {
  var html = HtmlService.createHtmlOutputFromFile('SetupWizard')
    .setTitle('FreelanceFlow Setup')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function OPEN_SETUP_WIZARD() {
  openSetupWizard_();
}


// ---- Step 2: API key validation ------------------------------------------------

function wizard_validateApiKey(key) {
  if (!key || key.trim().length < 20) {
    return { ok: false, message: 'Key looks too short. Paste the full key.' };
  }
  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + key.trim() },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 200) return { ok: true,  message: 'Key validated.' };
    if (code === 401) return { ok: false, message: 'Invalid key. Check that you copied the full key from OpenAI.' };
    if (code === 429) return { ok: true,  message: 'Key valid but quota exceeded. Add billing at platform.openai.com then continue.' };
    return { ok: false, message: 'Unexpected response (' + code + '). Try again.' };
  } catch (e) {
    return { ok: false, message: 'Connection error: ' + e.message };
  }
}


// ---- Step 3: AI niche analysis (optional) --------------------------------------

function wizard_analyzeNiche(description, apiKey) {
  var key = apiKey || PropertiesService.getScriptProperties().getProperty('UPWORK_OPENAI_API_KEY');
  if (!key) return { ok: false, message: 'API key not found. Complete Step 2 first.' };
  // Sidebar HTML can only call top-level bound-script functions via
  // google.script.run, not Library functions directly -- this stays
  // as the entry point and delegates the actual prompt/parsing.
  return FFLib.analyzeNiche(description, key);
}

// Fires automatically right after wizard_analyzeNiche succeeds -- generates
// niche-specific Job_Type categories + template content instead of shipping
// fixed BI/dashboard sample templates. Result flows into data.generatedTemplates
// at Step 7, written by initProposalTemplates_.
function wizard_generateTemplates(description, tools, background, apiKey) {
  var key = apiKey || PropertiesService.getScriptProperties().getProperty('UPWORK_OPENAI_API_KEY');
  if (!key) return { ok: false, message: 'API key not found. Complete Step 2 first.' };
  return FFLib.generateNicheTemplates(description, tools, background, key);
}


// ---- Main initialization (Step 7 button) ---------------------------------------

function wizard_initialize(data) {
  var prop = PropertiesService.getScriptProperties();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();

  prop.setProperty('UPWORK_OPENAI_API_KEY', data.apiKey.trim());

  var thresholds = getScoringThresholds_(data.scoringProfile);

  initSettingsSheet_(ss, data, thresholds);
  initConnectsHelper_(ss, Number(data.connectBalance) || 0);
  initProposalTemplates_(ss, data.portfolio, data.generatedTemplates);
  initProjectsSheet_(ss, data.portfolio);
  ensurePipelineSheets_(ss, data.starterKeywords || []);
  registerEditTrigger_();
  reorderPipelineTabs_(ss);

  // Auto-run the AI keyword strategy off the niche/portfolio just written to
  // Settings, so Keyword_Strategy and Keyword_Search_List are already
  // populated when the wizard closes -- no separate manual menu click needed.
  // Wrapped so a failure here (bad key, quota, network) never blocks setup
  // from completing; GENERATE_KEYWORD_STRATEGY() surfaces its own ui.alert
  // on failure, same message you'd see running it from the menu later.
  try {
    GENERATE_KEYWORD_STRATEGY();
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Setup complete, but keyword strategy generation failed: ' + err.message + '\n\n' +
      'Run System Tools > Generate Keyword Strategy to try again.'
    );
  }

  prop.setProperty('FF_SETUP_COMPLETE', 'true');
  return { ok: true };
}


// ---- Sheet builders ------------------------------------------------------------

function getScoringThresholds_(profile) {
  if (profile === 'Conservative') {
    return {
      applyMin: 0.65, applyMinTool: 0.8, applyMinExp: 0.7,
      applyMaxConn: 12, applyMaxProp: 25, applyMaxAge: 14,
      holdMin: 0.55, holdMaxConn: 14, yieldTarget: 6
    };
  }
  if (profile === 'Aggressive') {
    return {
      applyMin: 0.55, applyMinTool: 0.7, applyMinExp: 0.6,
      applyMaxConn: 18, applyMaxProp: 50, applyMaxAge: 30,
      holdMin: 0.45, holdMaxConn: 20, yieldTarget: 10
    };
  }
  return {
    applyMin: 0.60, applyMinTool: 0.8, applyMinExp: 0.7,
    applyMaxConn: 14, applyMaxProp: 35, applyMaxAge: 21,
    holdMin: 0.50, holdMaxConn: 16, yieldTarget: 8
  };
}

function initSettingsSheet_(ss, data, thresholds) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]).setFontWeight('bold');
  } else {
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }

  var rows = [
    ['Freelancer_Name',          data.name],
    ['Freelancer_Background',    data.background],
    ['Primary_Tools',            data.primaryTools],
    ['Contracts_Completed',      data.contractsCompleted  || 0],
    ['Reviews_Count',            data.reviewsCount        || 0],
    ['Job_Success_Score',        data.jobSuccessScore     || 0],
    ['Freelancer_Experience_Level', data.experienceLevel],
    ['Journey_Stage',            'New'],
    ['Proposal_Tone',            'Direct'],
    ['Proposal_Length',          'Medium'],
    ['Scoring_Profile',          data.scoringProfile],
    ['Apply_Min_Score',          thresholds.applyMin],
    ['Apply_Min_Tool_Score',     thresholds.applyMinTool],
    ['Apply_Min_Exp_Score',      thresholds.applyMinExp],
    ['Apply_Max_Connects',       thresholds.applyMaxConn],
    ['Apply_Max_Proposals',      thresholds.applyMaxProp],
    ['Apply_Max_Age_Days',       thresholds.applyMaxAge],
    ['Hold_Min_Score',           thresholds.holdMin],
    ['Hold_Max_Connects',        thresholds.holdMaxConn],
    ['Session_Yield_Target',     thresholds.yieldTarget]
  ];

  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  applySettingsValidation_(sheet, rows.map(function (r) { return r[0]; }));
}

// Dropdown validation for Journey_Stage/Proposal_Tone/Proposal_Length --
// Journey_Stage and Proposal_Tone allow a custom typed value alongside the
// dropdown (Journey_Stage's >20-char literal-override escape hatch in
// buildJourneyStage_, Lib_AIContext.gs; Proposal_Tone since any short tone
// word works fine in that prompt). Proposal_Length is strict -- Short/
// Medium/Long are the only values resolveProposalLengthSpec_
// (Lib_ProposalGenerator.gs) knows how to map to a character-range/
// max_tokens spec, so an unrecognized custom value would just silently
// fall back to Medium instead of doing what the user typed.
function applySettingsValidation_(sheet, settingNames) {
  var rowOf = function (name) {
    var idx = settingNames.indexOf(name);
    return idx === -1 ? 0 : idx + 2; // +2: header row, then 0-index -> 1-index
  };

  var journeyStageRow = rowOf('Journey_Stage');
  if (journeyStageRow) {
    sheet.getRange(journeyStageRow, 2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['New', 'Growing', 'Established'], true)
        .setAllowInvalid(true)
        .build()
    );
  }

  var proposalToneRow = rowOf('Proposal_Tone');
  if (proposalToneRow) {
    sheet.getRange(proposalToneRow, 2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Direct', 'Warm', 'Confident'], true)
        .setAllowInvalid(true)
        .build()
    );
  }

  var proposalLengthRow = rowOf('Proposal_Length');
  if (proposalLengthRow) {
    sheet.getRange(proposalLengthRow, 2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Short', 'Medium', 'Long'], true)
        .setAllowInvalid(false)
        .build()
    );
  }
}

function initConnectsHelper_(ss, startBalance) {
  var sheet = ss.getSheetByName('Connects_Helper');
  if (!sheet) {
    sheet = ss.insertSheet('Connects_Helper');
    sheet.getRange(1, 1, 1, 2).setValues([['Metric', 'Value']]).setFontWeight('bold');
  } else {
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }

  var metrics = [
    ['Current_Connect_Balance',     startBalance],
    ['Total_Connects_Purchased',    startBalance],
    ['Total_Connects_Used',         0],
    ['Total_Connects_Returned',     0],
    ['Connect_Replenishment',       ''],
    ['Connect_Replenishment_Date',  ''],
    ['Connect_Returned',            ''],
    ['Connect_Returned_Date',       ''],
    ['MTD_Sessions',                0],
    ['MTD_Jobs_Logged',             0],
    ['MTD_Proposals_Sent',         0],
    ['MTD_Connects_Used',           0],
    ['MTD_Replies',                 0],
    ['MTD_Interviews',              0],
    ['MTD_Hires',                   0],
    ['MTD_Revenue',                 0],
    ['Monthly_Revenue',             0],
    ['Monthly_Cost',                0],
    ['Monthly_ROI',                 0],
    ['Monthly_ROI_Dollar',          0],
    ['Total_Proposal_Cost',         0],
    ['Expected_Value_per_Proposal', 0],
    ['Revenue_per_Connect',         0],
    ['Net_Value_per_Connect',       0],
    ['Cost_per_Reply',              0],
    ['Cost_per_Interview',          0],
    ['Cost_per_Hire',               0]
  ];
  sheet.getRange(2, 1, metrics.length, 2).setValues(metrics);
}

function initProposalTemplates_(ss, portfolio, generatedTemplates) {
  var sheet = ss.getSheetByName('Proposal_Templates');
  if (!sheet) sheet = ss.insertSheet('Proposal_Templates');

  var headers = ['Template_ID','Job_Type','Hook_Version','CTA_Version','Angle','Credential_Hint','Tone','CTA_Style','Example_Output','Notes','Is_Default'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (sheet.getLastRow() > 1) return;

  var defaultCred = (portfolio && portfolio.length > 0) ? portfolio[0].name : 'Your Portfolio Project';

  var rows = (generatedTemplates || []).map(function (t, i) {
    return [
      'T' + (i + 1),
      t.jobType || 'General',
      'A', 'A',
      t.angle || '',
      defaultCred,
      t.tone || 'Direct',
      t.ctaStyle || 'Question',
      t.exampleOutput || '',
      t.notes || '',
      i === 0 ? 'Yes' : 'No'
    ];
  });

  // Niche-specific templates come from Step 3's AI assist (Lib_WizardAI's
  // generateNicheTemplates). If that call never ran or failed, this generic
  // row keeps the pipeline usable rather than shipping an empty sheet --
  // niche-agnostic on purpose, not a stand-in for the old hardcoded samples.
  if (rows.length === 0) {
    rows = [[
      'T1','General','A','A',
      'Lead with the most specific detail from the job post',
      defaultCred,'Direct','Question',
      'Open with one sentence referencing a specific job detail. Connect your relevant experience to their need. End with one direct question.',
      'Fallback template -- AI generation was unavailable during setup','Yes'
    ]];
  }

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// Projects is the single source of truth for portfolio data -- the wizard
// writes it once here, but it's a plain sheet meant to be edited afterward
// (add/remove/rename projects) same as Keyword_Search_List. Everything that
// used to read Portfolio_N / Portfolio_N_Keywords off Settings now reads
// this sheet instead (getPortfolioMapFromProjects_, getSettings_'s
// Portfolio_All), so edits here actually take effect on later formula runs.
function initProjectsSheet_(ss, portfolio) {
  var sheet = ss.getSheetByName('Projects');
  if (!sheet) {
    sheet = ss.insertSheet('Projects');
    sheet.getRange(1, 1, 1, 3).setValues([['Project_Name', 'Description', 'Keywords']]).setFontWeight('bold');
  }
  if (sheet.getLastRow() > 1) return;

  var rows = (portfolio || []).map(function (p) {
    return [p.name, p.description || '', p.keywords || ''];
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

function ensurePipelineSheets_(ss, starterKeywords) {
  var defs = [
    { name: 'Job_Discovery', headers: ['Discovery_ID','Date_Found','Session_ID','Job_Title','Description','Additional_Questions','Client_Name','Keyword_Search','Tool_Detected','Experience_Level','Minutes_Since_Posted','Hours_Since_Posted','Days_Since_Posted','Current_Age_Days','Proposal_Count','Payment_Verified','Client_Hires','Budget_Type','Budget','Hourly_Rate','Job_Link','Connects_Required','AI_Fit_Notes','Discovery_Status','Keyword_Fit_Score','Tool_Score','Experience_Score','Freshness_Score','Competition_Score','Verification_Score','Client_History_Score','Budget_Quick_Score','Discovery_Priority_Score','Discovery_Action'] },
    { name: 'Job_Scoring',   headers: ['Discovery_ID','Date_Scored','Job_Title','Description','Additional_Questions','Client_Name','Keyword_Search','Tool_Detected','Experience_Level','Hours_Since_Posted','Days_Since_Posted','Proposal_Count','Payment_Verified','Client_Hires','Budget_Type','Budget','Hourly_Rate','Connects_Required','Job_Link','Effort_Level','Estimated_Hours','Estimated_Hourly_Rate','Budget_Score','Keyword_Score','Tool_Score','Experience_Score','Freshness_Score','Competition_Score','Client_History_Score','Scope_Rating','Scope_Score','Portfolio_Match','Portfolio_Score','Connects_Affordability','Total_Score','Score_Per_Connect','Final_Decision','Proposal_Generator_Date','AI_Fit_Notes','Current_Age_Days'] },
    { name: 'Proposal_Generator', headers: ['Discovery_ID','Date','Job_Title','Client_Name','Description','Job_Link','Keyword_Search','Tool_Detected','Job_Type','Connects_Required','Proposal_Count','Budget','Portfolio_Project','Recommended_Template','Hook_Version','CTA_Version','Bid_1st','Bid_2nd','Bid_3rd','Bid_4th','Boost_Connects','Total_Connects_Spent','Bid_Recommendation','Additional_Questions','AI_Generated_Proposal','Additional_Answers','Proposal_Status','Proposal_Sent_Date','Proposal_Skip_Date','Notes'] },
    { name: 'Proposal_Tracker',  headers: ['Discovery_ID','Date_Applied','Job_Title','Client_Name','Keyword_Search','Tool_Requested','Days_Since_Posted','Proposal_Count','Total_Score','Template_Used','Hook_Version','CTA_Version','Viewed','Interview','Hired','Revenue','Notes','Age_Days','Current_Age_Days','Connects_Used','Boost_Connects','Proposal_Cost','Job_Link','HubSpot_Contact_ID','HubSpot_Deal_ID'] },
    { name: 'Client_Chat_Log',  headers: ['Discovery_ID','Job_Title','Client_Name','Message_Number','Sender_Name','Direction','Message_Time','Message_Text','Last_Synced'] },
    { name: 'Contract_Tracker', headers: ['Discovery_ID','Job_Title','Client_Name','Contract_Type','Contract_Value','Hourly_Rate','Start_Date','Status','Total_Released','Notes'] },
    { name: 'Milestone_Tracker', headers: ['Discovery_ID','Job_Title','Milestone_Number','Description','Amount','Status','Funded_Date','Delivered_Date','Released_Date','Notes'] },
    { name: 'Hourly_Log',       headers: ['Discovery_ID','Job_Title','Log_Date','Hours_Logged','Amount','Status','Notes'] },
    { name: 'Session_Log',       headers: ['Session_ID','Date','Start_Time','End_Time','Duration','Keywords_Searched','Jobs_Logged','Jobs_Moved_To_Scoring','Jobs_Review_Later','Duplicates_Skipped','Session_Yield','Saturation_Flag','Proposal_Trigger','Proposals_Sent','Proposals_Skipped','Connects_Spent','Notes'] },
    { name: 'Keyword_Search_List', headers: ['Tool','Business_Area','Intent','Search_Query','Last_Searched','Session_Yield'] },
    { name: 'Keyword_Strategy',  headers: ['Keyword','Recommended_Action','Actual_Count','Target_Count','Notes','Drop','Scale'] },
    // No fixed headers -- BUILD_KEYWORD_INTELLIGENCE (26_Keyword_Intelligence.gs)
    // clears and rewrites this sheet's full layout on every refresh, same
    // pattern as Dashboard below.
    { name: 'Keyword_Intelligence', headers: [] },
    { name: 'Monthly_Performance', headers: ['Month','Year','Total_Sessions','Jobs_Logged','Proposals_Sent','Connects_Used','Proposal_Cost','Replies','Interviews','Hires','Reply_Rate_Pct','Interview_Rate_Pct','Hire_Rate_Pct','Revenue','Cost','ROI','Cost_per_Reply','Cost_per_Interview','Cost_per_Hire','Monthly_ROI_Dollar','Expected_Value_per_Proposal','Revenue_per_Connect','Net_Value_per_Connect'] },
    // No fixed headers -- BUILD_DASHBOARD (25_Dashboard.gs) clears and
    // rewrites this sheet's full layout on every refresh, so a header row
    // here would just get wiped on the first run anyway.
    { name: 'Dashboard', headers: [] }
  ];

  for (var i = 0; i < defs.length; i++) {
    var def   = defs[i];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      if (def.headers.length > 0) {
        sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      }

      if (def.name === 'Job_Discovery') {
        applyJobDiscoveryFormulas_(sheet, def.headers);
        applyJobDiscoveryValidation_(sheet, def.headers);
        applyJobDiscoveryConditionalFormatting_(sheet, def.headers);
      } else if (def.name === 'Job_Scoring') {
        applyJobScoringPullFormula_(sheet, def.headers);
        applyJobScoringFormulas_(sheet, def.headers);
        applyJobScoringValidation_(sheet, def.headers);
        applyJobScoringConditionalFormatting_(sheet, def.headers);
      } else if (def.name === 'Proposal_Generator') {
        applyProposalGeneratorLookupFormulas_(sheet, def.headers);
        applyProposalGeneratorFormulas_(sheet, def.headers);
        applyProposalGeneratorValidation_(sheet, def.headers);
      } else if (def.name === 'Proposal_Tracker') {
        applyProposalTrackerValidation_(sheet, def.headers);
      } else if (def.name === 'Contract_Tracker') {
        applyContractTrackerValidation_(sheet, def.headers);
      } else if (def.name === 'Milestone_Tracker') {
        applyMilestoneTrackerValidation_(sheet, def.headers);
      } else if (def.name === 'Hourly_Log') {
        applyHourlyLogValidation_(sheet, def.headers);
      } else if (def.name === 'Keyword_Strategy') {
        applyKeywordStrategyValidation_(sheet, def.headers);
      }
    }
  }

  // Additional_Questions flows backwards (Proposal_Generator -> Job_Scoring
  // -> Job_Discovery, see Lib_WizardFormulas.gs's buildDiscoveryIdLookup_)
  // so it has to be wired up here, after all three sheets are guaranteed to
  // exist -- Job_Scoring's lookup formula reads Proposal_Generator's headers,
  // which don't exist yet during Job_Scoring's own creation step above.
  var jdSheetForLookup = ss.getSheetByName('Job_Discovery');
  var jsSheetForLookup = ss.getSheetByName('Job_Scoring');
  var pgSheetForLookup = ss.getSheetByName('Proposal_Generator');
  if (jdSheetForLookup && jsSheetForLookup && pgSheetForLookup) {
    var jsHeadersForLookup = jsSheetForLookup.getRange(1, 1, 1, jsSheetForLookup.getLastColumn()).getValues()[0];
    applyJobScoringAdditionalQuestionsLookup_(jsSheetForLookup, jsHeadersForLookup);
    // Proposal_Generator_Date reads Proposal_Generator's static Date, so it
    // needs that sheet's headers too.
    applyJobScoringProposalDateLookup_(jsSheetForLookup, jsHeadersForLookup);
    // Final_Decision's "already in Proposal_Generator" exemption from the
    // age/proposal caps needs Proposal_Generator's headers, which don't
    // exist yet when Job_Scoring is first created above.
    applyJobScoringFinalDecision_(jsSheetForLookup, jsHeadersForLookup);
    applyJobDiscoveryAdditionalQuestionsLookup_(jdSheetForLookup,
      jdSheetForLookup.getRange(1, 1, 1, jdSheetForLookup.getLastColumn()).getValues()[0]);
  }

  var kwSheet = ss.getSheetByName('Keyword_Search_List');
  if (kwSheet && kwSheet.getLastRow() <= 1 && starterKeywords.length > 0) {
    var rows = starterKeywords.map(function(kw) {
      return [kw, '', '', kw];
    });
    kwSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

// Sheet creation order above is driven by cross-sheet formula dependencies
// (e.g. Job_Scoring's Additional_Questions lookup needs Proposal_Generator's
// headers already in place), which doesn't match the tab order a customer
// actually wants to see. This runs once at the end of setup to move every
// tab into the intended reading order, independent of creation order.
// Deletes the default "Sheet1" left over from a brand-new spreadsheet, since
// by this point every real sheet has been created and it's just clutter.
function reorderPipelineTabs_(ss) {
  var order = [
    'Dashboard', 'Job_Discovery', 'Job_Scoring', 'Proposal_Generator', 'Proposal_Tracker',
    'Client_Chat_Log', 'Contract_Tracker', 'Milestone_Tracker', 'Hourly_Log',
    'Proposal_Templates', 'Keyword_Search_List', 'Keyword_Strategy', 'Keyword_Intelligence',
    'Connects_Helper', 'Session_Log', 'Projects', 'Settings', 'Tool_Aliases', 'Monthly_Performance'
  ];

  for (var i = 0; i < order.length; i++) {
    var sheet = ss.getSheetByName(order[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  var landingSheet = ss.getSheetByName('Job_Discovery');
  if (landingSheet) ss.setActiveSheet(landingSheet);
}

function registerEditTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'onEdit' || handler === 'handleEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}


// ---- Formula helpers -----------------------------------------------------------
//
// build*_ functions are pure -- headers/portfolioMap in, formula-string out, no
// SpreadsheetApp access. apply*_ functions are the sheet-I/O wrappers that call
// them and write the results. This split is what lets the build*_ functions move
// into the Apps Script Library later while apply*_ stays in the thin client, and
// it's also what REPAIR_FORMULAS() (15_Formula_Fixes.gs) reuses to re-apply
// formulas to an already-created sheet, not just at creation time.

function applyJobScoringFormulas_(sheet, headers) {
  var formulas = FFLib.buildJobScoringFormulas(headers, getProposalGeneratorHeaders_(sheet.getParent()));

  var fields = [
    'currentAgeDays', 'effortLevel', 'scopeRating', 'portfolioMatch', 'estimatedHours',
    'estimatedHourlyRate', 'budgetScore', 'keywordScore', 'toolScore', 'experienceScore',
    'freshnessScore', 'competitionScore', 'clientHistoryScore', 'scopeScore', 'portfolioScore',
    'connectsAffordability', 'totalScore', 'scorePerConnect', 'finalDecision'
  ];

  fields.forEach(function (field) {
    var col     = formulas[field + 'Col'];
    var formula = formulas[field + 'Formula'];
    if (col && formula) {
      sheet.getRange(2, col, FORMULA_PREFILL_ROWS, 1).setFormula(formula);
    }
  });
}

// Job_Scoring's raw-data columns (Job_Title, Description, Budget, etc.)
// auto-populate from Job_Discovery via a single spilling FILTER formula --
// one cell per contiguous column group, never copied down per row like the
// score formulas above (a FILTER re-spilling from every row would collide
// with itself). Clears the group's range first so stale hand-typed data
// (or a prior FILTER) can't block the new spill with a #REF! error.
function applyJobScoringPullFormula_(sheet, headers) {
  var jdSheet = sheet.getParent().getSheetByName('Job_Discovery');
  if (!jdSheet) return;
  var jdHeaders = jdSheet.getRange(1, 1, 1, jdSheet.getLastColumn()).getValues()[0];

  var groups = FFLib.buildJobScoringPullFormulas(jdHeaders, headers);
  if (groups.length === 0) return;

  var lastRow = Math.max(sheet.getLastRow(), FORMULA_PREFILL_ROWS + 1);
  groups.forEach(function (g) {
    sheet.getRange(2, g.col, lastRow - 1, g.width).clearContent();
    sheet.getRange(2, g.col).setFormula(g.formula);
  });
  forceDiscoveryIdNumberFormat_(sheet, headers, lastRow);
}

// Discovery_ID is a plain integer, but a column inserted next to a date
// column can inherit date formatting from its neighbor -- forcing the
// format here means the row number displays correctly (1, 2, 3...) no
// matter how the column was created or which sheet it's pulled into.
function forceDiscoveryIdNumberFormat_(sheet, headers, lastRow) {
  var col = headers.indexOf('Discovery_ID') + 1;
  if (col > 0) {
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('0');
  }
}

// Proposal_Generator's raw-data columns are per-row lookups keyed on each
// row's static Discovery_ID (see FFLib.buildProposalGeneratorLookupFormulas
// and 28_Proposal_Sync.gs). A sheet still on the old FILTER spill is
// migrated to static rows first, keeping today's row order.
function applyProposalGeneratorLookupFormulas_(sheet, headers) {
  var jsSheet = sheet.getParent().getSheetByName('Job_Scoring');
  var idCol   = headers.indexOf('Discovery_ID') + 1;
  if (!jsSheet || idCol <= 0) return 0;
  var jsHeaders = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];

  var migrated = migrateProposalGeneratorToStaticRows_(sheet, headers);

  var lookups = FFLib.buildProposalGeneratorLookupFormulas(jsHeaders, headers);
  var lastRow = Math.max(getLastProposalGeneratorRow_(sheet, idCol), FORMULA_PREFILL_ROWS + 1);
  lookups.forEach(function (l) {
    sheet.getRange(2, l.col, lastRow - 1, 1).setFormula(l.formula);
  });
  forceDiscoveryIdNumberFormat_(sheet, headers, lastRow);
  return migrated;
}

// Additional_Questions flows backwards -- a per-row lookup formula keyed on
// Discovery_ID, prefilled down FORMULA_PREFILL_ROWS same as any other
// per-row formula (not a FILTER spill, so no clearContent/single-cell
// pattern needed here).
function applyJobScoringAdditionalQuestionsLookup_(sheet, headers) {
  var pgSheet = sheet.getParent().getSheetByName('Proposal_Generator');
  if (!pgSheet) return;
  var pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];

  var result = FFLib.buildJobScoringAdditionalQuestionsLookup(pgHeaders, headers);
  if (result) {
    sheet.getRange(2, result.col, FORMULA_PREFILL_ROWS, 1).setFormula(result.formula);
  }
}

// Proposal_Generator's headers, or null while it's still on the old FILTER
// layout. That FILTER reads Final_Decision, so pointing Final_Decision back
// at Proposal_Generator before migration would form a formula loop.
function getProposalGeneratorHeaders_(ss) {
  var pgSheet = ss.getSheetByName('Proposal_Generator');
  if (!pgSheet || pgSheet.getLastColumn() === 0) return null;
  var headers = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
  var idCol   = headers.indexOf('Discovery_ID') + 1;
  if (idCol > 0 && pgSheet.getMaxRows() >= 2 && pgSheet.getRange(2, idCol).getFormula()) return null;
  return headers;
}

// Re-applies Final_Decision once Proposal_Generator exists. Skipped while
// Proposal_Generator is unmigrated: the new age/proposal caps would shrink
// the old FILTER before REPAIR_FORMULAS migrates it, moving typed data onto
// the wrong jobs.
function applyJobScoringFinalDecision_(sheet, headers) {
  var pgHeaders = getProposalGeneratorHeaders_(sheet.getParent());
  if (!pgHeaders) return;
  var formulas = FFLib.buildJobScoringFormulas(headers, pgHeaders);
  if (formulas.finalDecisionCol && formulas.finalDecisionFormula) {
    var rows = Math.max(sheet.getLastRow() - 1, FORMULA_PREFILL_ROWS);
    sheet.getRange(2, formulas.finalDecisionCol, rows, 1).setFormula(formulas.finalDecisionFormula);
  }
}

function applyJobScoringProposalDateLookup_(sheet, headers) {
  var pgSheet = sheet.getParent().getSheetByName('Proposal_Generator');
  if (!pgSheet) return;
  var pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];

  var result = FFLib.buildJobScoringProposalDateLookup(pgHeaders, headers);
  if (result) {
    sheet.getRange(2, result.col, FORMULA_PREFILL_ROWS, 1)
      .setFormula(result.formula)
      .setNumberFormat('m/d/yyyy');
  }
}

function applyJobDiscoveryAdditionalQuestionsLookup_(sheet, headers) {
  var jsSheet = sheet.getParent().getSheetByName('Job_Scoring');
  if (!jsSheet) return;
  var jsHeaders = jsSheet.getRange(1, 1, 1, jsSheet.getLastColumn()).getValues()[0];

  var result = FFLib.buildJobDiscoveryAdditionalQuestionsLookup(jsHeaders, headers);
  if (result) {
    sheet.getRange(2, result.col, FORMULA_PREFILL_ROWS, 1).setFormula(result.formula);
  }
}

function applyProposalGeneratorValidation_(sheet, headers) {
  var statusCol = headers.indexOf('Proposal_Status') + 1;
  if (statusCol > 0) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ready', 'Sent', 'Skip'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(statusRule);
  }

  // Additional_Answers is free-text (AI-drafted), but a column inserted next
  // to Proposal_Status's dropdown can inherit its validation rule the same
  // way Discovery_ID inherited date formatting -- clear it explicitly so it
  // can never show a stray dropdown arrow, no matter how the column was created.
  var answersCol = headers.indexOf('Additional_Answers') + 1;
  if (answersCol > 0) {
    sheet.getRange(2, answersCol, FORMULA_PREFILL_ROWS, 1).clearDataValidations();
  }
}

function applyJobScoringValidation_(sheet, headers) {
  var budgetTypeCol = headers.indexOf('Budget_Type') + 1;
  var payVerCol      = headers.indexOf('Payment_Verified') + 1;
  var expLevelCol    = headers.indexOf('Experience_Level') + 1;
  var propCountCol   = headers.indexOf('Proposal_Count') + 1;

  if (budgetTypeCol > 0) {
    var budgetTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Fixed', 'Hourly_Range', 'Hourly_Unknown'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, budgetTypeCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(budgetTypeRule);
  }

  if (payVerCol > 0) {
    var payVerRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, payVerCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(payVerRule);
  }

  if (expLevelCol > 0) {
    var expLevelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Entry Level', 'Intermediate', 'Expert'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, expLevelCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(expLevelRule);
  }

  if (propCountCol > 0) {
    var propCountRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Fewer than 5', '5 to 10', '10 to 15', '15 to 20', '20 to 50', '50+'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, propCountCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(propCountRule);
  }
}

// Viewed/Interview/Hired are read as strict "Y" string matches by both the
// Workflow Analyzer and handleEdit's MTD accumulation below -- a free-typed
// "yes"/"Yes" would silently fail to count, so this is a correctness fix,
// not cosmetic polish. Revenue gets currency formatting since it's the one
// freeform manual-entry number field on this sheet.
function applyProposalTrackerValidation_(sheet, headers) {
  var ynRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'N'], true)
    .setAllowInvalid(false)
    .build();

  ['Viewed', 'Interview', 'Hired'].forEach(function (name) {
    var col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, FORMULA_PREFILL_ROWS, 1).setDataValidation(ynRule);
    }
  });

  var revenueCol = headers.indexOf('Revenue') + 1;
  if (revenueCol > 0) {
    sheet.getRange(2, revenueCol, FORMULA_PREFILL_ROWS, 1).setNumberFormat('$#,##0.00');
  }
}

function applyContractTrackerValidation_(sheet, headers) {
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Fixed', 'Hourly'], true)
    .setAllowInvalid(false)
    .build();
  var typeCol = headers.indexOf('Contract_Type') + 1;
  if (typeCol > 0) {
    sheet.getRange(2, typeCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(typeRule);
  }

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Completed', 'Ended Early'], true)
    .setAllowInvalid(false)
    .build();
  var statusCol = headers.indexOf('Status') + 1;
  if (statusCol > 0) {
    sheet.getRange(2, statusCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(statusRule);
  }

  ['Contract_Value', 'Hourly_Rate', 'Total_Released'].forEach(function (name) {
    var col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.getRange(2, col, FORMULA_PREFILL_ROWS, 1).setNumberFormat('$#,##0.00');
    }
  });
}

function applyMilestoneTrackerValidation_(sheet, headers) {
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Funded', 'Delivered', 'Released'], true)
    .setAllowInvalid(false)
    .build();
  var statusCol = headers.indexOf('Status') + 1;
  if (statusCol > 0) {
    sheet.getRange(2, statusCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(statusRule);
  }

  var amountCol = headers.indexOf('Amount') + 1;
  if (amountCol > 0) {
    sheet.getRange(2, amountCol, FORMULA_PREFILL_ROWS, 1).setNumberFormat('$#,##0.00');
  }
}

// Amount is script-computed on Hours_Logged edit (see 14_Edit_Trigger.gs's
// HOURLY_LOG block) -- this formats the two number columns so entries look
// right from the first row, plus a Status dropdown that mirrors
// Contract_Tracker's Status and propagates to it on edit (applyHourlyLogStatusEffects_
// in 14_Edit_Trigger.gs) -- Hourly contracts have no milestones to
// auto-complete off of, so this is the equivalent signal for them.
function applyHourlyLogValidation_(sheet, headers) {
  var amountCol = headers.indexOf('Amount') + 1;
  if (amountCol > 0) {
    sheet.getRange(2, amountCol, FORMULA_PREFILL_ROWS, 1).setNumberFormat('$#,##0.00');
  }

  var hoursCol = headers.indexOf('Hours_Logged') + 1;
  if (hoursCol > 0) {
    sheet.getRange(2, hoursCol, FORMULA_PREFILL_ROWS, 1).setNumberFormat('0.00');
  }

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Completed'], true)
    .setAllowInvalid(false)
    .build();
  var statusCol = headers.indexOf('Status') + 1;
  if (statusCol > 0) {
    sheet.getRange(2, statusCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(statusRule);
  }
}

// Drop is a manual override, separate from the formula-driven
// Recommended_Action column -- select "Drop" on a keyword, then run
// System Tools > Drop Keywords (18_Keyword_Strategy.gs) to remove it from
// both Keyword_Search_List and this row's own Keyword_Strategy entry.
//
// Scale is the same kind of manual override, opposite direction -- select
// "Scale" on a keyword, then run System Tools > Scale Keywords
// (18_Keyword_Strategy.gs) to generate related search-phrase variations off
// it. Independent of Drop -- both columns can be set on the same row, though
// doing so is a contradictory combination: Drop still deletes the row
// regardless of the Scale flag on it.
function applyKeywordStrategyValidation_(sheet, headers) {
  var dropRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Drop'], true)
    .setAllowInvalid(false)
    .build();

  var dropCol = headers.indexOf('Drop') + 1;
  if (dropCol > 0) {
    sheet.getRange(2, dropCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(dropRule);
  }

  var scaleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Scale'], true)
    .setAllowInvalid(false)
    .build();

  var scaleCol = headers.indexOf('Scale') + 1;
  if (scaleCol > 0) {
    sheet.getRange(2, scaleCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(scaleRule);
  }
}

// Portfolio_Project is filled by RUN_JOB_CLASSIFICATION (11_Job_Classifier.gs)
// instead of a formula here -- see pickPortfolioProject (Lib_WizardFormulas.gs)
// for why.
function applyProposalGeneratorFormulas_(sheet, headers) {
  var primaryTools = (getSettings_()['Primary_Tools'] || '');
  var aliasMap     = getToolAliasMap_(sheet.getParent());
  var formulas     = FFLib.buildProposalGeneratorFormulas(headers, primaryTools, aliasMap);

  if (formulas.toolDetectedFormula) {
    sheet.getRange(2, formulas.toolDetectedCol, FORMULA_PREFILL_ROWS, 1).setFormula(formulas.toolDetectedFormula);
  }
}

function applyJobDiscoveryFormulas_(sheet, headers) {
  var settings      = getSettings_();
  var primaryTools  = settings['Primary_Tools'] || '';
  var aliasMap      = getToolAliasMap_(sheet.getParent());
  var formulas      = FFLib.buildJobDiscoveryFormulas(headers, primaryTools, aliasMap);

  var fields = [
    'discoveryId', 'currentAgeDays', 'keywordFitScore', 'toolDetected', 'toolScore', 'experienceScore',
    'freshnessScore', 'competitionScore', 'verificationScore', 'clientHistoryScore',
    'budgetQuickScore', 'discoveryPriorityScore', 'discoveryAction', 'discoveryStatus'
  ];

  fields.forEach(function (field) {
    var col     = formulas[field + 'Col'];
    var formula = formulas[field + 'Formula'];
    if (col && formula) {
      var range = sheet.getRange(2, col, FORMULA_PREFILL_ROWS, 1);
      range.setFormula(formula);
      // Discovery_ID is a plain integer, but a column inserted next to a
      // date column (Date_Found) can inherit date formatting from its
      // neighbor -- forcing the format here means the row number displays
      // correctly (1, 2, 3...) no matter how the column was created.
      if (field === 'discoveryId') range.setNumberFormat('0');
    }
  });
}

function applyJobDiscoveryValidation_(sheet, headers) {
  var budgetTypeCol = headers.indexOf('Budget_Type') + 1;
  var payVerCol      = headers.indexOf('Payment_Verified') + 1;
  var expLevelCol    = headers.indexOf('Experience_Level') + 1;
  var propCountCol   = headers.indexOf('Proposal_Count') + 1;

  if (budgetTypeCol > 0) {
    var budgetTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Fixed', 'Hourly_Range', 'Hourly_Unknown'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, budgetTypeCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(budgetTypeRule);
  }

  if (payVerCol > 0) {
    var payVerRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, payVerCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(payVerRule);
  }

  if (expLevelCol > 0) {
    var expLevelRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Entry Level', 'Intermediate', 'Expert'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, expLevelCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(expLevelRule);
  }

  if (propCountCol > 0) {
    var propCountRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Fewer than 5', '5 to 10', '10 to 15', '15 to 20', '20 to 50', '50+'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, propCountCol, FORMULA_PREFILL_ROWS, 1).setDataValidation(propCountRule);
  }
}

function applyJobDiscoveryConditionalFormatting_(sheet, headers) {
  var actionCol = headers.indexOf('Discovery_Action') + 1;
  if (actionCol <= 0) return;

  var actionLetter = colLetterClient_(actionCol);
  var fullRange    = sheet.getRange(2, 1, FORMULA_PREFILL_ROWS, headers.length);

  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + actionLetter + '2="Move to Scoring"')
      .setBackground('#d9ead3')
      .setRanges([fullRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + actionLetter + '2="Review Later"')
      .setBackground('#fff2cc')
      .setRanges([fullRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + actionLetter + '2="Skip"')
      .setBackground('#f4cccc')
      .setRanges([fullRange])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}

function applyJobScoringConditionalFormatting_(sheet, headers) {
  var decisionCol = headers.indexOf('Final_Decision') + 1;
  if (decisionCol <= 0) return;

  var decisionLetter = colLetterClient_(decisionCol);
  var fullRange       = sheet.getRange(2, 1, FORMULA_PREFILL_ROWS, headers.length);

  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + decisionLetter + '2="APPLY"')
      .setBackground('#d9ead3')
      .setRanges([fullRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + decisionLetter + '2="HOLD"')
      .setBackground('#fff2cc')
      .setRanges([fullRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + decisionLetter + '2="SKIP"')
      .setBackground('#f4cccc')
      .setRanges([fullRange])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}

// Client-side column-letter helper for conditional formatting ranges.
// (Library's colLetter_ is private to the Library and not callable
// cross-project -- this is the same trivial A1/A2.../AA logic.)
function colLetterClient_(n) {
  var s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function getPortfolioMapFromProjects_() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var sheet        = ss.getSheetByName('Projects');
  var portfolioMap = {};

  if (!sheet || sheet.getLastRow() < 2) return portfolioMap;

  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameCol  = headers.indexOf('Project_Name');
  var descCol  = headers.indexOf('Description');
  var kwCol    = headers.indexOf('Keywords');
  var data     = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  for (var i = 0; i < data.length; i++) {
    var name = nameCol  >= 0 ? String(data[i][nameCol]).trim() : '';
    if (!name) continue;

    var desc = descCol >= 0 ? String(data[i][descCol]).trim() : '';
    var kwRaw = kwCol   >= 0 ? String(data[i][kwCol]).trim()  : '';
    var keywords = kwRaw
      ? kwRaw.split(',').map(function(k) { return k.trim().toLowerCase(); }).filter(function(k) { return k; })
      : [];

    portfolioMap[i + 1] = { name: name, description: desc, keywords: keywords };
  }

  return portfolioMap;
}

// getPortfolioMapFromProjects_'s output feeds FFLib.pickPortfolioProject
// (called from RUN_JOB_CLASSIFICATION, 11_Job_Classifier.gs) and, before that,
// FFLib.generateKeywordStrategy's niche/portfolio summary (18_Keyword_Strategy.gs).
// colLetter_ moved to the Apps Script Library (private helper used internally
// by the formula builders) -- nothing in the client calls it directly.
