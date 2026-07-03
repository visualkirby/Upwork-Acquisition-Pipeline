/**
 * ============================================================
 * 00. SETUP WIZARD
 * Defines onOpen for the entire spreadsheet.
 * Do NOT define onOpen in any other module.
 * ============================================================
 */
function onOpen() {
  buildSystemMenu_();
  var prop = PropertiesService.getScriptProperties();
  if (prop.getProperty('FF_SETUP_COMPLETE') !== 'true') {
    Utilities.sleep(2000);
    openSetupWizard_();
  }
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

  var prompt =
    'You are helping a freelancer set up a job-bidding pipeline tool on Upwork. ' +
    'Based on their specialty description below, return a JSON object with exactly these three keys: ' +
    '"tools": a comma-separated string of their primary tools or software (max 8 tools, most relevant first), ' +
    '"background": a single professional sentence (under 25 words) describing their experience and focus for use in AI proposals, ' +
    '"keywords": an array of 10-14 Upwork search keyword strings relevant to their niche ' +
    '(format each as "[Tool] [Domain] [Type]" or "[Domain] [Type]", e.g. "Power BI Sales Dashboard", "Logistics Data Analysis"). ' +
    'Return only valid JSON. No explanation. No markdown. ' +
    'Specialty description: ' + description.substring(0, 500);

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + key },
      payload: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 400,
        temperature: 0.3
      }),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return { ok: false, message: 'API error: ' + data.error.message };

    var content = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : '';

    content = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    var parsed = JSON.parse(content);
    return {
      ok:         true,
      tools:      parsed.tools      || '',
      background: parsed.background || '',
      keywords:   Array.isArray(parsed.keywords) ? parsed.keywords : []
    };
  } catch (e) {
    return { ok: false, message: 'Could not parse AI response. Fill in manually.' };
  }
}


// ---- Main initialization (Step 7 button) ---------------------------------------

function wizard_initialize(data) {
  var prop = PropertiesService.getScriptProperties();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();

  prop.setProperty('UPWORK_OPENAI_API_KEY', data.apiKey.trim());

  var thresholds = getScoringThresholds_(data.scoringProfile);

  initSettingsSheet_(ss, data, thresholds);
  initConnectsHelper_(ss, Number(data.connectBalance) || 0);
  initProposalTemplates_(ss, data.portfolio);
  initFollowupTemplates_(ss);
  ensurePipelineSheets_(ss, data.starterKeywords || []);
  registerEditTrigger_();

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
    ['Journey_Stage',            ''],
    ['Proposal_Tone',            'Direct'],
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

  var portfolio = data.portfolio || [];
  for (var i = 0; i < Math.min(portfolio.length, 6); i++) {
    rows.push(['Portfolio_' + (i + 1),               portfolio[i].name]);
    rows.push(['Portfolio_' + (i + 1) + '_Keywords',  portfolio[i].keywords]);
  }

  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
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

function initProposalTemplates_(ss, portfolio) {
  var sheet = ss.getSheetByName('Proposal_Templates');
  if (!sheet) sheet = ss.insertSheet('Proposal_Templates');

  var headers = ['Template_ID','Job_Type','Hook_Version','CTA_Version','Angle','Credential_Hint','Tone','CTA_Style','Example_Output','Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (sheet.getLastRow() > 1) return;

  var defaultCred = (portfolio && portfolio.length > 0) ? portfolio[0].name : 'Your Portfolio Project';

  var samples = [
    ['T1','Dashboard Build','A','A',
     'Lead with the specific industry or data problem in the job post',
     defaultCred,'Direct','Question',
     'Open with one sentence referencing a specific job detail. Connect your portfolio project to their need. End with one direct question.',
     'Primary template -- use for most new dashboard builds'],
    ['T1','Dashboard Fix','B','A',
     'Acknowledge the existing issue before offering the fix',
     defaultCred,'Confident','Question',
     'Name the specific problem (slow, unclear, missing metric). Reference a fix you have done before. One targeted question.',
     'Use when the client mentions an existing dashboard needs work'],
    ['T2','Data to Dashboard','A','B',
     'Focus on the data cleaning step clients underestimate',
     defaultCred,'Direct','Offer',
     'Name the data format or problem. Reference your clean-to-visual pipeline experience. End with a scope question.',
     'Use for raw data or CSV to dashboard jobs']
  ];
  sheet.getRange(2, 1, samples.length, headers.length).setValues(samples);
}

function initFollowupTemplates_(ss) {
  var sheet = ss.getSheetByName('Followup_Templates');
  if (!sheet) sheet = ss.insertSheet('Followup_Templates');

  var headers = ['Template_ID','Followup_Number','Days_After_Send','Subject','Body','Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (sheet.getLastRow() > 1) return;

  var samples = [
    ['F1',1,5,'Re: [Job Title]',
     'Still available if you are moving forward with this project. Happy to answer any questions before you decide.',
     'Send 5 days after proposal if no reply'],
    ['F2',2,10,'Re: [Job Title]',
     'Following up one more time -- still interested if the timeline shifted. Let me know either way.',
     'Send 10 days after proposal'],
    ['F3',3,18,'Re: [Job Title]',
     'Last follow-up -- if this project is still open and you need this work done quickly, I can start this week.',
     'Final follow-up at 18 days']
  ];
  sheet.getRange(2, 1, samples.length, headers.length).setValues(samples);
}

function ensurePipelineSheets_(ss, starterKeywords) {
  var defs = [
    { name: 'Job_Discovery', headers: ['Session_ID','Date_Found','Job_Title','Client_Name','Description','Job_Link','Keyword_Search','Tool_Detected','Experience_Level','Hours_Since_Posted','Days_Since_Posted','Proposal_Count','Payment_Verified','Client_Hires','Budget_Type','Budget','Hourly_Rate','Connects_Required','Quick_Notes','Discovery_Action'] },
    { name: 'Job_Scoring',   headers: ['Job_Title','Client_Name','Description','Job_Link','Keyword_Search','Tool_Detected','Experience_Level','Hours_Since_Posted','Days_Since_Posted','Proposal_Count','Payment_Verified','Client_Hires','Budget_Type','Budget','Hourly_Rate','Connects_Required','Quick_Notes','Date_Scored','Effort_Level','Scope_Rating','Portfolio_Match','Tool_Score','Experience_Score','Freshness_Score','Competition_Score','Budget_Score','Verification_Score','Client_History_Score','Keyword_Fit_Score','Scope_Score','Tool_Match_Score','Connects_Affordability','Total_Score','Score_Per_Connect','Final_Decision','Proposal_Generator_Date'] },
    { name: 'Proposal_Generator', headers: ['Date','Job_Title','Client_Name','Description','Job_Link','Keyword_Search','Tool_Detected','Job_Type','Connects_Required','Proposal_Count','Budget','Portfolio_Project','Recommended_Template','Hook_Version','CTA_Version','Bid_1st','Bid_2nd','Bid_3rd','Boost_Connects','Total_Connects_Spent','Bid_Recommendation','AI_Generated_Proposal','Proposal_Status','Proposal_Sent_Date','Proposal_Skip_Date','Notes'] },
    { name: 'Proposal_Tracker',  headers: ['Date_Applied','Job_Title','Client_Name','Keyword_Search','Tool_Requested','Days_Since_Posted','Proposal_Count','Total_Score','Template_Used','Hook_Version','CTA_Version','Client_Replied','Interview','Hired','Revenue','Notes','Age_Days','Current_Age_Days','Connects_Used','Boost_Connects','Proposal_Cost','Job_Link'] },
    { name: 'Followup_Tracker',  headers: ['Date_Applied','Job_Title','Client_Name','Template_Used','Followup1_Sent','Followup1_Template','Followup2_Sent','Followup2_Template','Followup3_Sent','Followup3_Template','Client_Replied','Interview','Hired','Notes'] },
    { name: 'Session_Log',       headers: ['Session_ID','Date','Start_Time','End_Time','Duration','Keywords_Searched','Jobs_Logged','Jobs_Moved_To_Scoring','Jobs_Review_Later','Duplicates_Skipped','Session_Yield','Saturation_Flag','Proposal_Trigger','Proposals_Sent','Proposals_Skipped','Connects_Spent','Notes'] },
    { name: 'Keyword_Search_List', headers: ['Tool','Business_Area','Intent','Search_Query','Last_Searched','Session_Yield'] },
    { name: 'Keyword_Strategy',  headers: ['Keyword','Recommended_Action','Actual_Count','Target_Count','Notes'] },
    { name: 'Monthly_Performance', headers: ['Month','Year','Total_Sessions','Jobs_Logged','Proposals_Sent','Connects_Used','Proposal_Cost','Replies','Interviews','Hires','Reply_Rate_Pct','Interview_Rate_Pct','Hire_Rate_Pct','Revenue','Cost','ROI','Cost_per_Reply','Cost_per_Interview','Cost_per_Hire','Monthly_ROI_Dollar','Expected_Value_per_Proposal','Revenue_per_Connect','Net_Value_per_Connect'] }
  ];

  for (var i = 0; i < defs.length; i++) {
    var def   = defs[i];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');

      if (def.name === 'Job_Scoring') {
        applyJobScoringFormulas_(sheet, def.headers);
      } else if (def.name === 'Proposal_Generator') {
        applyProposalGeneratorFormulas_(sheet, def.headers);
      }
    }
  }

  var kwSheet = ss.getSheetByName('Keyword_Search_List');
  if (kwSheet && kwSheet.getLastRow() <= 1 && starterKeywords.length > 0) {
    var rows = starterKeywords.map(function(kw) {
      return [kw, '', '', kw];
    });
    kwSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

function registerEditTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}


// ---- Formula helpers -----------------------------------------------------------

function applyJobScoringFormulas_(sheet, headers) {
  var connReqIdx  = headers.indexOf('Connects_Required');
  var connAffIdx  = headers.indexOf('Connects_Affordability');
  var toolScoIdx  = headers.indexOf('Tool_Score');
  var expScoIdx   = headers.indexOf('Experience_Score');
  var totScoIdx   = headers.indexOf('Total_Score');
  var finalDecIdx = headers.indexOf('Final_Decision');

  if (connAffIdx >= 0 && connReqIdx >= 0) {
    var crL = colLetter_(connReqIdx + 1);
    sheet.getRange(2, connAffIdx + 1).setFormula(
      '=IF(' + crL + '2="","",IF(' + crL + '2<=Connects_Helper!$B$2,"Can Afford","Cannot Afford"))'
    );
  }

  if (finalDecIdx >= 0 && totScoIdx >= 0 && connAffIdx >= 0 && toolScoIdx >= 0 && expScoIdx >= 0 && connReqIdx >= 0) {
    var totL  = colLetter_(totScoIdx + 1);
    var affL  = colLetter_(connAffIdx + 1);
    var toolL = colLetter_(toolScoIdx + 1);
    var expL  = colLetter_(expScoIdx + 1);
    var crL   = colLetter_(connReqIdx + 1);
    sheet.getRange(2, finalDecIdx + 1).setFormula(
      '=IF(' + totL + '2="","",IF(' + affL + '2="Cannot Afford","SKIP",' +
      'IFS(' +
        'AND(' +
          totL  + '2>=VLOOKUP("Apply_Min_Score",Settings!$A:$B,2,0),' +
          toolL + '2>=VLOOKUP("Apply_Min_Tool_Score",Settings!$A:$B,2,0),' +
          expL  + '2>=VLOOKUP("Apply_Min_Exp_Score",Settings!$A:$B,2,0),' +
          crL   + '2<=VLOOKUP("Apply_Max_Connects",Settings!$A:$B,2,0)' +
        '),"APPLY",' +
        'AND(' +
          totL + '2>=VLOOKUP("Hold_Min_Score",Settings!$A:$B,2,0),' +
          crL  + '2<=VLOOKUP("Hold_Max_Connects",Settings!$A:$B,2,0)' +
        '),"HOLD",' +
        'TRUE,"SKIP"' +
      ')))'
    );
  }
}

function applyProposalGeneratorFormulas_(sheet, headers) {
  var jobTitleIdx  = headers.indexOf('Job_Title');
  var descIdx      = headers.indexOf('Description');
  var toolDetIdx   = headers.indexOf('Tool_Detected');
  var portfolioIdx = headers.indexOf('Portfolio_Project');

  if (jobTitleIdx < 0 || descIdx < 0) return;

  var jtL = colLetter_(jobTitleIdx + 1);
  var dcL = colLetter_(descIdx + 1);

  if (toolDetIdx >= 0) {
    sheet.getRange(2, toolDetIdx + 1).setFormula(
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
      '))'
    );
  }

  if (portfolioIdx >= 0) {
    var formula = buildPortfolioFormula_(jtL, dcL);
    sheet.getRange(2, portfolioIdx + 1).setFormula(formula);
  }
}

function buildPortfolioFormula_(jobTitleLetter, descLetter) {
  var jtL = jobTitleLetter || 'B';
  var dcL = descLetter     || 'D';

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName('Settings');

  if (!settings || settings.getLastRow() < 2) {
    return '=IF(' + jtL + '2="","","Add portfolio projects via FreelanceFlow Setup")';
  }

  var data = settings.getRange(2, 1, settings.getLastRow() - 1, 2).getValues();
  var portfolioMap = {};

  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = String(data[i][1]).trim();
    if (!val) continue;

    var nameMatch = key.match(/^Portfolio_(\d+)$/);
    if (nameMatch) {
      var n = nameMatch[1];
      portfolioMap[n] = portfolioMap[n] || { name: '', keywords: [] };
      portfolioMap[n].name = val;
    }

    var kwMatch = key.match(/^Portfolio_(\d+)_Keywords$/);
    if (kwMatch) {
      var n = kwMatch[1];
      portfolioMap[n] = portfolioMap[n] || { name: '', keywords: [] };
      portfolioMap[n].keywords = val.split(',').map(function(k) { return k.trim().toLowerCase(); }).filter(function(k) { return k; });
    }
  }

  var clauses     = [];
  var defaultName = '';
  var pNums       = Object.keys(portfolioMap).sort();

  for (var p = 0; p < pNums.length; p++) {
    var proj = portfolioMap[pNums[p]];
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
