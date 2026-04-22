/**
 * ============================================================
 * UPWORK ACQUISITION SYSTEM — APPS SCRIPT
 * ============================================================
 * Sections:
 *   1. MENU
 *   2. API KEY SETUP
 *   3. HELPERS
 *   4. SMART RESET
 *   5. AI HELPER FUNCTIONS
 *   6. KEYWORD MINING
 *   7. DUPLICATE JOB LINK COLORING
 *   8. BID RECOMMENDATION ENGINE
 *   9. AI PROPOSAL GENERATOR
 *  10. ANALYZE JOB WORKFLOW
 *  11. AI PROPOSAL GENERATION (AUTO)
 *  12. AI JOB TYPE CLASSIFIER
 *  13. RUN AI PROPOSALS (BATCH)
 *  14. SESSION MANAGEMENT
 *  15. SNAPSHOT MONTH END
 *  16. MAIN EDIT TRIGGER
 * ============================================================
 */


/**
 * ============================================================
 * 1. MENU
 * ============================================================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("System Tools")
    .addItem("Reset System",          "RESET_SYSTEM")
    .addItem("Mine Keywords",         "MINE_KEYWORDS")
    .addSeparator()
    .addItem("Start Session",         "START_SESSION")
    .addItem("End Session",           "END_SESSION")
    .addSeparator()
    .addItem("Analyze Job Workflow",  "ANALYZE_JOB_WORKFLOW")
    .addItem("Run Job Classification", "RUN_JOB_CLASSIFICATION")
    .addItem("Run AI Proposals",       "RUN_AI_PROPOSALS")
    .addSeparator()
    .addItem("Snapshot Month End",    "SNAPSHOT_MONTH_END")
    .addSeparator()
    .addItem("Setup API Key",         "SETUP_API_KEY")
    .addToUi();
}


/**
 * ============================================================
 * 2. API KEY SETUP
 *
 * Run once from the Apps Script editor to store your OpenAI
 * API key securely in PropertiesService.
 *
 * Steps:
 *   1. Paste your key between the quotes below
 *   2. Click Run → SETUP_API_KEY
 *   3. When the alert says "stored", delete your key from
 *      this function and save — it is now stored securely
 *
 * To switch to Anthropic later: paste your Anthropic key,
 * change the property name to "ANTHROPIC_API_KEY", run once.
 * ============================================================
 */
function TEST_JOB_TYPE_API() {
  var ui     = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    ui.alert("No API key found. Run Setup API Key first.");
    return;
  }

  ui.alert("API key found. Length: " + apiKey.length + "\nFirst 8 chars: " + apiKey.substring(0, 8));

  var testDesc  = "We are looking for a skilled developer to build a new Power BI dashboard tracking sales KPIs.";
  var testTitle = "Power BI Dashboard Developer";
  var result    = getJobType_(testDesc, testTitle);

  ui.alert("Test classification result: \"" + result + "\"\n\n" +
    (result ? "API call working correctly." : "API call returned empty — check key or quota."));
}


function DIAGNOSE_PROPOSAL_GENERATOR() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet) { ui.alert("Proposal_Generator not found."); return; }

  var map        = getHeaderMap_(sheet);
  var descCol    = getCol_(map, ["Description"]);
  var jobTypeCol = getCol_(map, ["Job_Type"]);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var lastRow    = sheet.getLastRow();

  var actualRows  = 0;
  var lastDescRow = 1;
  if (descCol && lastRow > 1) {
    var descVals = sheet.getRange(2, descCol, lastRow - 1, 1).getValues();
    for (var t = 0; t < descVals.length; t++) {
      if (String(descVals[t][0]).trim() !== "") {
        actualRows++;
        lastDescRow = t + 2;
      }
    }
  }

  var msg = "Proposal_Generator Diagnostic\n";
  msg += "getLastRow(): " + lastRow + "\n";
  msg += "Actual rows with Description: " + actualRows + "\n";
  msg += "Last description row: " + lastDescRow + "\n";
  msg += "Description col: " + descCol + "\n";
  msg += "Job_Type col: " + jobTypeCol + "\n";
  msg += "Job_Title col: " + titleCol + "\n\n";

  if (lastRow < 2) { ui.alert(msg + "No data rows found."); return; }

  var data = sheet.getRange(2, 1, Math.min(lastRow - 1, 5), sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var title = titleCol   ? String(data[i][titleCol   - 1]).substring(0, 30) : "N/A";
    var desc  = descCol    ? String(data[i][descCol    - 1]).trim()            : "N/A";
    var jtype = jobTypeCol ? String(data[i][jobTypeCol - 1]).trim()            : "N/A";
    msg += "Row " + (i + 2) + ": title=" + title + "\n";
    msg += "  Job_Type=\"" + jtype + "\" desc_length=" + desc.length + "\n";
    msg += "  desc_preview=\"" + desc.substring(0, 60) + "\"\n\n";
  }

  ui.alert(msg);
}


function SETUP_API_KEY() {
  var key = "paste-your-openai-key-here";

  if (key === "paste-your-openai-key-here" || key.trim() === "") {
    SpreadsheetApp.getUi().alert(
      "No key found.\n\nPaste your OpenAI API key between the quotes in SETUP_API_KEY, then run again."
    );
    return;
  }

  PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
  SpreadsheetApp.getUi().alert(
    "✓ API key stored securely.\n\nNow delete your key from the SETUP_API_KEY function and save the script."
  );
}


/**
 * ============================================================
 * 3. HELPERS
 * ============================================================
 */
function normalizeHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    map[normalizeHeader_(h)] = i + 1;
  });
  return map;
}

function getCol_(headerMap, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var key = normalizeHeader_(possibleNames[i]);
    if (headerMap[key]) return headerMap[key];
  }
  return null;
}

function getCellValue_(sheet, row, headerMap, possibleNames) {
  var col = getCol_(headerMap, possibleNames);
  if (!col) return "";
  return sheet.getRange(row, col).getValue();
}

function setCellValue_(sheet, row, headerMap, possibleNames, value) {
  var col = getCol_(headerMap, possibleNames);
  if (!col) return;
  sheet.getRange(row, col).setValue(value);
}

function cleanJobText_(text) {
  if (text === null || text === undefined) return "";
  return String(text)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n+/g, " ")
    .replace(/\t+/g, " ")
    .replace(/\u00A0/g, " ")
    .replace(/[•▪■●◦]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function getLastRealRow_(sheet) {
  var maxRows  = sheet.getMaxRows();
  var lastReal = 1;
  for (var col = 1; col <= 3; col++) {
    var vals = sheet.getRange(2, col, maxRows - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]).trim() !== "") {
        var rowNum = i + 2;
        if (rowNum > lastReal) lastReal = rowNum;
        break;
      }
    }
  }
  return lastReal;
}

function findFirstEmptyRowByColumn_(sheet, col) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values  = sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) return i + 2;
  }
  return lastRow + 1;
}

function formatDuration_(ms) {
  var totalSeconds = Math.floor(ms / 1000);
  var hours        = Math.floor(totalSeconds / 3600);
  var minutes      = Math.floor((totalSeconds % 3600) / 60);
  var seconds      = totalSeconds % 60;
  if (hours > 0) {
    return hours + "h " + minutes + "m " + seconds + "s";
  }
  return minutes + "m " + seconds + "s";
}

function formatTime_(date) {
  var h    = date.getHours();
  var m    = date.getMinutes();
  var s    = date.getSeconds();
  var ampm = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return h + ":" + (m < 10 ? "0" + m : m) + ":" + (s < 10 ? "0" + s : s) + " " + ampm;
}


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


/**
 * ============================================================
 * 5. AI HELPER FUNCTIONS
 * ============================================================
 */
function getQuickNotes_(description) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    return getQuickNotesRegex_(description);
  }

  var prompt =
    "You are analyzing an Upwork job description for a data analytics freelancer. " +
    "Read the description carefully and return exactly one line in this format: " +
    "[Complexity], [Scope] & [Tool Match]. " +
    "Rules: " +
    "Complexity: Large=ETL/pipelines/APIs/data modeling/warehouse/Azure/automation/integrations. " +
    "Complex=SQL/multiple dashboards/joins/merges/transformations/multiple BI tools. " +
    "Normal=single dashboard/report/KPI tracker/spreadsheet/Excel/Looker Studio build. " +
    "Simple=minor update or very small scope. " +
    "Scope: Clear=step-by-step requirements/specific examples/exact deliverables. " +
    "Mostly Clear=focused scope with some gaps. " +
    "Vague=general ask/no clear deliverable. " +
    "Very Vague=no clear scope at all. " +
    "Tool Match: Exact=explicitly names Tableau/Power BI/Looker Studio/Excel dashboard. " +
    "Strong=mentions a specific BI tool or KPI dashboard. " +
    "Partial=mentions dashboard/reporting/analytics without naming a tool. " +
    "Weak=tangentially related. " +
    "None=org chart/presentation/graphic design. " +
    "Return ONLY the formatted result. No explanation. No extra text. Example: Normal, Mostly Clear & Strong. " +
    "Job description: " + description.substring(0, 1500);

  var payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 30,
    temperature: 0.1
  };

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return getQuickNotesRegex_(description);

    var result = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "";

    return result || getQuickNotesRegex_(description);

  } catch (err) {
    return getQuickNotesRegex_(description);
  }
}


function getQuickNotesRegex_(description) {
  if (!description) return "";
  var d = description;

  var complexity =
    /azure|etl|pipeline|api|data model|warehouse|automation|integrat/i.test(d) ? "Large" :
    /sql|multiple dashboards|power bi|tableau|looker|ga4|join|merge|transform/i.test(d) ? "Complex" :
    /dashboard|report|kpi|spreadsheet|excel|looker studio/i.test(d) ? "Normal" : "Simple";

  var scope =
    /step-by-step|clearly defined|specific requirements|example outputs|exactly/i.test(d) ? "Clear" :
    /focused|scoped|improve|update|redesign/i.test(d) ? "Mostly Clear" :
    /help us|looking for|need someone|would like/i.test(d) ? "Vague" : "Very Vague";

  var toolMatch =
    /tableau dashboard|power bi dashboard|looker studio dashboard|excel dashboard/i.test(d) ? "Exact" :
    /tableau|power bi|looker studio|excel|kpi dashboard/i.test(d) ? "Strong" :
    /dashboard|reporting|analytics/i.test(d) ? "Partial" :
    /org chart|presentation|graphic design/i.test(d) ? "None" : "Weak";

  return complexity + ", " + scope + " & " + toolMatch;
}


function getWorkflowAnalysis_(jobTitle, description) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    return {
      detailed:  "API key not set. Run System Tools > Setup API Key first.",
      condensed: "API key not set."
    };
  }

  var prompt =
    "You are a senior data analytics consultant reviewing an Upwork job posting. " +
    "Analyze the job and provide a structured breakdown for a freelancer deciding whether to apply " +
    "and how to approach the work if hired. " +
    "Job Title: " + jobTitle + ". " +
    "Description: " + description.substring(0, 2000) + " " +
    "Provide your analysis in exactly this structure: " +
    "WHAT THE CLIENT NEEDS: (1-2 sentences on the real underlying need, not just the surface ask) | " +
    "TOOLS & SKILLS REQUIRED: (bullet list, be specific) | " +
    "SUGGESTED DELIVERY APPROACH: (step-by-step, 3-5 steps max) | " +
    "COMPLEXITY ASSESSMENT: (one line: Simple / Normal / Complex / Large and why) | " +
    "RED FLAGS: (anything vague, unrealistic, or risky, or write None) | " +
    "KEY QUESTIONS TO ASK CLIENT: (2-3 questions you would want answered before starting) " +
    "Keep each section concise. Total response under 250 words. Separate sections with a blank line.";

  var payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 400,
    temperature: 0.3
  };

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());

    if (data.error) {
      return {
        detailed:  "API error: " + data.error.message,
        condensed: "API error."
      };
    }

    var detailed = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "No response returned.";

    var condensed = detailed
      .replace(/WHAT THE CLIENT NEEDS:\s*/i, "")
      .split(/\n\n/)[0]
      .trim()
      .substring(0, 200);

    return { detailed: detailed, condensed: condensed };

  } catch (err) {
    return {
      detailed:  "Request failed: " + err.message,
      condensed: "Request failed."
    };
  }
}


/**
 * ============================================================
 * 6. KEYWORD MINING
 * ============================================================
 */
function MINE_KEYWORDS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    "Mine Keywords",
    "How many new keyword combinations do you want to add?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var requestedCount = parseInt(response.getResponseText(), 10);
  if (isNaN(requestedCount) || requestedCount < 1) {
    ui.alert("Please enter a whole number greater than 0.");
    return;
  }

  var discoverySheet  = ss.getSheetByName("Job_Discovery");
  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var strategySheet   = ss.getSheetByName("Keyword_Strategy");

  if (!discoverySheet || !searchListSheet || !strategySheet) {
    ui.alert("Missing sheet. Confirm Job_Discovery, Keyword_Search_List, and Keyword_Strategy all exist.");
    return;
  }

  var stratLastRow = strategySheet.getLastRow();
  var stratMap     = getHeaderMap_(strategySheet);
  var dropSet      = {};

  if (stratLastRow > 1) {
    var stratData = strategySheet
      .getRange(2, 1, stratLastRow - 1, strategySheet.getLastColumn())
      .getValues();

    var sKeywordCol = getCol_(stratMap, ["Keyword"]);
    var sActionCol  = getCol_(stratMap, ["Recommended_Action"]);
    var sActualCol  = getCol_(stratMap, ["Actual_Count"]);
    var sTargetCol  = getCol_(stratMap, ["Target_Count"]);

    for (var i = 0; i < stratData.length; i++) {
      var kw     = sKeywordCol ? String(stratData[i][sKeywordCol - 1]).trim().toLowerCase() : "";
      var action = sActionCol  ? String(stratData[i][sActionCol  - 1]).trim() : "";
      var actual = sActualCol  ? Number(stratData[i][sActualCol  - 1]) : 0;
      var target = sTargetCol  ? Number(stratData[i][sTargetCol  - 1]) : 0;

      if (action === "Drop" && actual >= target && kw !== "") {
        dropSet[kw] = true;
      }
    }
  }

  var slMap      = getHeaderMap_(searchListSheet);
  var slQueryCol = getCol_(slMap, ["Search_Query"]);
  var slLastRow  = searchListSheet.getLastRow();
  var purgedCount = 0;

  if (slLastRow > 1 && slQueryCol && Object.keys(dropSet).length > 0) {
    for (var r = slLastRow; r >= 2; r--) {
      var cellQuery = String(
        searchListSheet.getRange(r, slQueryCol).getValue()
      ).trim().toLowerCase();
      if (dropSet[cellQuery]) {
        searchListSheet.deleteRow(r);
        purgedCount++;
      }
    }
  }

  slLastRow = searchListSheet.getLastRow();
  var existingKeys = {};
  var slToolCol    = getCol_(slMap, ["Tool"]);
  var slBizCol     = getCol_(slMap, ["Business_Area"]);
  var slIntCol     = getCol_(slMap, ["Intent"]);

  if (slLastRow > 1 && slToolCol && slBizCol && slIntCol) {
    var slData = searchListSheet
      .getRange(2, 1, slLastRow - 1, searchListSheet.getLastColumn())
      .getValues();

    for (var i = 0; i < slData.length; i++) {
      var eTool = String(slData[i][slToolCol - 1]).trim().toLowerCase();
      var eBiz  = String(slData[i][slBizCol  - 1]).trim().toLowerCase();
      var eInt  = String(slData[i][slIntCol  - 1]).trim().toLowerCase();

      if (eTool !== "" && eBiz !== "" && eInt !== "") {
        existingKeys[eTool + "|" + eBiz + "|" + eInt] = true;
      }
    }
  }

  var TOOLS = [
    "Excel", "Power BI", "Looker Studio", "Tableau", "SQL",
    "Google Sheets", "Python", "BigQuery", "Power Query",
    "Google Analytics", "R Studio", "Snowflake", "dbt"
  ];

  var BUSINESS_AREAS = [
    "Sales", "HR", "Marketing", "Operations", "Finance",
    "Inventory", "Revenue", "Retail", "Healthcare", "Logistics",
    "Ecommerce", "Supply Chain", "Construction", "Real Estate",
    "Hospitality", "Procurement", "Manufacturing"
  ];

  var INTENTS = [
    "Dashboard", "Reporting Dashboard", "Dashboard Developer",
    "Dashboard Build", "Dashboard Creation", "Automation",
    "Data Analysis", "Visualization", "Pipeline", "Integration",
    "KPI Dashboard", "Analytics Dashboard"
  ];

  var discLastRow = discoverySheet.getLastRow();
  if (discLastRow < 2) {
    ui.alert("Job_Discovery has no data rows to mine.");
    return;
  }

  var discMap      = getHeaderMap_(discoverySheet);
  var discTitleCol = getCol_(discMap, ["Job_Title"]);
  var discDescCol  = getCol_(discMap, ["Description"]);
  var discData     = discoverySheet
    .getRange(2, 1, discLastRow - 1, discoverySheet.getLastColumn())
    .getValues();

  var tripletCounts = {};

  for (var r = 0; r < discData.length; r++) {
    var title    = discTitleCol ? String(discData[r][discTitleCol - 1]).toLowerCase() : "";
    var desc     = discDescCol  ? String(discData[r][discDescCol  - 1]).toLowerCase() : "";
    var combined = title + " " + desc;

    var foundTools = TOOLS.filter(function (t) {
      return combined.indexOf(t.toLowerCase()) !== -1;
    });
    var foundBiz = BUSINESS_AREAS.filter(function (b) {
      return combined.indexOf(b.toLowerCase()) !== -1;
    });
    var foundIntents = INTENTS.filter(function (n) {
      return combined.indexOf(n.toLowerCase()) !== -1;
    });

    if (foundTools.length === 0) foundTools = ["Other"];

    foundTools.forEach(function (tool) {
      foundBiz.forEach(function (biz) {
        foundIntents.forEach(function (intent) {
          var key = tool.toLowerCase() + "|" + biz.toLowerCase() + "|" + intent.toLowerCase();
          tripletCounts[key] = (tripletCounts[key] || 0) + 1;
        });
      });
    });
  }

  var MIN_FREQUENCY = 2;
  var candidates    = [];

  Object.keys(tripletCounts).forEach(function (key) {
    if (tripletCounts[key] < MIN_FREQUENCY) return;
    if (existingKeys[key]) return;

    var parts  = key.split("|");
    var tool   = TOOLS.find(function (t) { return t.toLowerCase() === parts[0]; }) || parts[0];
    var biz    = BUSINESS_AREAS.find(function (b) { return b.toLowerCase() === parts[1]; }) || parts[1];
    var intent = INTENTS.find(function (n) { return n.toLowerCase() === parts[2]; }) || parts[2];

    candidates.push({ tool: tool, biz: biz, intent: intent, freq: tripletCounts[key] });
  });

  candidates.sort(function (a, b) { return b.freq - a.freq; });

  if (candidates.length === 0) {
    ui.alert(
      "No new combinations found above the minimum frequency of " + MIN_FREQUENCY + ".\n" +
      "All qualifying combinations are already in Keyword_Search_List."
    );
    return;
  }

  var toWrite       = candidates.slice(0, requestedCount);
  var writeStartRow = getLastRealRow_(searchListSheet) + 1;
  var outputData    = toWrite.map(function (row) {
    return [row.tool, row.biz, row.intent];
  });

  searchListSheet
    .getRange(writeStartRow, 1, outputData.length, 3)
    .setValues(outputData);

  var summary = "Done.\n\n" +
    "✓ " + toWrite.length + " new keyword combinations added to Keyword_Search_List.\n";

  if (toWrite.length < requestedCount) {
    summary += "⚠ Only " + toWrite.length + " qualifying combinations were available " +
               "(you requested " + requestedCount + "). Run again after collecting more jobs.\n";
  }
  if (candidates.length > toWrite.length) {
    summary += "→ " + (candidates.length - toWrite.length) +
               " additional combinations ready for your next run.\n";
  }
  if (purgedCount > 0) {
    summary += "✓ " + purgedCount + " dropped + target-met rows removed before writing.";
  }

  ui.alert(summary);
}


/**
 * ============================================================
 * 7. DUPLICATE JOB LINK COLORING
 * ============================================================
 */
function colorDuplicateJobLinks() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Job_Discovery");
  if (!sheet) return;

  var map     = getHeaderMap_(sheet);
  var linkCol = getCol_(map, ["Job_Link"]);
  var skipCol = getCol_(map, ["Discovery_Action"]);
  if (!linkCol) return;

  var lastRow  = sheet.getLastRow();
  var lastCol  = sheet.getLastColumn();
  var startRow = 2;
  if (lastRow < startRow) return;

  var numRows    = lastRow - startRow + 1;
  var linkValues = sheet.getRange(startRow, linkCol, numRows, 1).getValues();

  var urlCount = {};
  linkValues.forEach(function (row) {
    var val = String(row[0]).trim();
    if (val === "" || val === "undefined") return;
    urlCount[val] = (urlCount[val] || 0) + 1;
  });

  var urlColor   = {};
  var colorIndex = 0;
  var palette    = [
    "#E63946", "#2A9D8F", "#E9C46A", "#4361EE",
    "#F4845F", "#52B788", "#9B5DE5", "#F15BB5",
    "#00B4D8", "#FF6B6B", "#06D6A0", "#FFB703"
  ];

  Object.keys(urlCount).forEach(function (url) {
    if (urlCount[url] > 1) {
      urlColor[url] = palette[colorIndex % palette.length];
      colorIndex++;
    }
  });

  var backgrounds = [];
  for (var i = 0; i < numRows; i++) {
    var val   = String(linkValues[i][0]).trim();
    var color = (val && val !== "undefined" && urlColor[val]) ? urlColor[val] : null;
    var rowColors = [];
    for (var c = 1; c <= lastCol; c++) {
      rowColors.push(c === skipCol ? null : color);
    }
    backgrounds.push(rowColors);
  }

  sheet.getRange(startRow, 1, numRows, lastCol).setBackgrounds(backgrounds);
}


/**
 * ============================================================
 * 8. BID RECOMMENDATION ENGINE
 * ============================================================
 */
var JOURNEY_STAGE_ =
  "The freelancer is Sawandi, an early-stage Upwork freelancer. " +
  "He currently has 0 completed contracts and 0 reviews on his profile. " +
  "He is budget-conscious with limited connects. " +
  "His strength is dashboard development across Tableau, Power BI, Looker Studio, Excel, and SQL. " +
  "Because he has no reviews yet, winning depends heavily on proposal quality and job fit, " +
  "not just bid position. Outbidding aggressively at this stage has lower ROI than targeting " +
  "lower-competition jobs where his proposal can stand out on merit. " +
  "He should be conservative with boost spending until he has at least 1-2 reviews.";

function getBidRecommendation_(jobTitle, baseConnects, proposalCount,
                                totalScore, bid1, bid2, bid3) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    return "API key not set. Run System Tools > Setup API Key first.";
  }

  var prompt =
    "You are a bid strategy advisor for an Upwork freelancer. " +
    "Give a structured bid recommendation using EXACTLY this format with no extra text: " +
    "DECISION: [NO BOOST / BOOST TO 3RD / BOOST TO 2ND / BOOST TO 1ST] " +
    "BID: [exact number of connects to spend, or just base connects if no boost] " +
    "REASON: [one sentence max] " +
    "FREELANCER CONTEXT: " + JOURNEY_STAGE_ + " " +
    "JOB DATA: " +
    "Title: " + jobTitle + ". " +
    "Base connects to submit: " + baseConnects + ". " +
    "Current proposal count: " + proposalCount + " (if this is text like Less than 5 treat it as 3). " +
    "Job quality score: " + totalScore + " out of 1. " +
    "Current bids — 1st: " + bid1 + " connects, 2nd: " + bid2 + " connects, 3rd: " + bid3 + " connects. " +
    "DECISION RULES (apply in order): " +
    "If proposal count is above 25: DECISION must be NO BOOST. " +
    "If score is below 0.65: DECISION must be NO BOOST. " +
    "If proposal count is under 10 AND score is above 0.70: seriously consider boosting. " +
    "BOOST TO 1ST only if proposals under 10 AND score above 0.75 AND gap between 1st and 2nd bid is 5 connects or fewer. " +
    "BOOST TO 2ND if proposals under 15 AND score above 0.70. " +
    "BOOST TO 3RD if proposals under 20 AND score above 0.65. " +
    "Otherwise NO BOOST. " +
    "Be direct. No preamble. Use the exact format above.";

  var payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 80,
    temperature: 0.1
  };

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return "API error: " + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "No response returned.";

  } catch (err) {
    return "Request failed: " + err.message;
  }
}


/**
 * ============================================================
 * 9. AI PROPOSAL GENERATOR (TEMPLATE-DRIVEN)
 * ============================================================
 */
function generateAIProposal_(jobTitle, description, toolDetected,
                               jobType, templateId, hookVersion, ctaVersion) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) return "API key not set. Run System Tools > Setup API Key first.";

  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var tmplSheet      = ss.getSheetByName("Proposal_Templates");
  var angle          = "";
  var credentialHint = "";
  var tone           = "";
  var ctaStyle       = "";
  var exampleOutput  = "";

  if (tmplSheet && tmplSheet.getLastRow() > 1) {
    var tmplMap     = getHeaderMap_(tmplSheet);
    var tmplIdCol   = getCol_(tmplMap, ["Template_ID"]);
    var tmplHookCol = getCol_(tmplMap, ["Hook_Version"]);
    var tmplCtaCol  = getCol_(tmplMap, ["CTA_Version"]);
    var angleCol    = getCol_(tmplMap, ["Angle"]);
    var credCol     = getCol_(tmplMap, ["Credential_Hint"]);
    var toneCol     = getCol_(tmplMap, ["Tone"]);
    var ctaStyleCol = getCol_(tmplMap, ["CTA_Style"]);
    var exampleCol  = getCol_(tmplMap, ["Example_Output"]);

    var tmplData = tmplSheet
      .getRange(2, 1, tmplSheet.getLastRow() - 1, tmplSheet.getLastColumn())
      .getValues();

    for (var i = 0; i < tmplData.length; i++) {
      var rowId   = tmplIdCol   ? String(tmplData[i][tmplIdCol   - 1]).trim() : "";
      var rowHook = tmplHookCol ? String(tmplData[i][tmplHookCol - 1]).trim() : "";
      var rowCta  = tmplCtaCol  ? String(tmplData[i][tmplCtaCol  - 1]).trim() : "";

      if (rowId === templateId && rowHook === hookVersion && rowCta === ctaVersion) {
        angle          = angleCol     ? String(tmplData[i][angleCol     - 1]).trim() : "";
        credentialHint = credCol      ? String(tmplData[i][credCol      - 1]).trim() : "";
        tone           = toneCol      ? String(tmplData[i][toneCol      - 1]).trim() : "";
        ctaStyle       = ctaStyleCol  ? String(tmplData[i][ctaStyleCol  - 1]).trim() : "";
        exampleOutput  = exampleCol   ? String(tmplData[i][exampleCol   - 1]).trim() : "";
        break;
      }
    }
  }

  var cred = credentialHint || "Operations Performance & KPI Monitoring Dashboard";

  var prompt =
    "You are writing an Upwork proposal for a data analytics freelancer named Sawandi. " +
    "STRICT RULES — violating any rule makes the proposal unusable: " +
    "1. Under 100 words total. " +
    "2. Do NOT start with Hi, Hello, or any greeting. " +
    "3. Do NOT use bullet points or numbered lists. " +
    "4. Do NOT list skills or tools generically. " +
    "5. First sentence MUST reference a specific detail from the job description — not a generic observation. " +
    "6. You MUST reference this exact portfolio project by name in the proposal: " + cred + " — do not substitute a different project. " +
    "7. End with exactly one direct question. No offers to help. " +
    "STRATEGIC ANGLE: " + (angle || "Lead with the specific client problem, not credentials.") + " " +
    "TONE: " + (tone || "Direct") + ". " +
    "CTA STYLE: " + (ctaStyle || "Question") + ". " +
    (exampleOutput ? "VOICE EXAMPLE (match this structure and directness, write new content): " + exampleOutput.substring(0, 250) + " " : "") +
    "JOB TITLE: " + jobTitle + ". " +
    "TOOL REQUESTED: " + (toolDetected || "not specified") + ". " +
    "JOB TYPE: " + (jobType || "dashboard project") + ". " +
    "JOB DESCRIPTION: " + description.substring(0, 1200);

  var payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 180,
    temperature: 0.5
  };

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return "API error: " + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "No response returned.";

  } catch (err) {
    return "Request failed: " + err.message;
  }
}


/**
 * ============================================================
 * 10. ANALYZE JOB WORKFLOW
 * ============================================================
 */
function ANALYZE_JOB_WORKFLOW() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Job_Discovery");

  if (!sheet) {
    ui.alert("Job_Discovery sheet not found.");
    return;
  }

  var lastRow = sheet.getLastRow();

  var rowResponse = ui.prompt(
    "Analyze Job Workflow",
    "Enter the row number to analyze (rows 2 to " + lastRow + "):",
    ui.ButtonSet.OK_CANCEL
  );
  if (rowResponse.getSelectedButton() !== ui.Button.OK) return;

  var row = parseInt(rowResponse.getResponseText(), 10);
  if (isNaN(row) || row < 2 || row > lastRow) {
    ui.alert("Invalid row number. Please enter a number between 2 and " + lastRow + ".");
    return;
  }

  var map         = getHeaderMap_(sheet);
  var titleCol    = getCol_(map, ["Job_Title"]);
  var descCol     = getCol_(map, ["Description"]);
  var workflowCol = getCol_(map, ["Job_Workflow_Advisor"]);

  var jobTitle    = titleCol ? sheet.getRange(row, titleCol).getValue() : "";
  var description = descCol  ? sheet.getRange(row, descCol).getValue()  : "";

  if (!description || String(description).trim() === "") {
    ui.alert("No description found in this row. Paste the job description first.");
    return;
  }

  if (workflowCol) {
    sheet.getRange(row, workflowCol).setValue("Analyzing...");
  }

  var result = getWorkflowAnalysis_(jobTitle, description);

  if (workflowCol) {
    sheet.getRange(row, workflowCol).setValue(result.condensed);
  }

  var popupText =
    "JOB WORKFLOW ANALYSIS\n" +
    "════════════════════════════════\n" +
    "Job: " + jobTitle + "\n" +
    "════════════════════════════════\n\n" +
    result.detailed +
    "\n\n────────────────────────────────\n" +
    "Condensed version saved to Job_Workflow_Advisor column.";

  ui.alert(popupText);
}


/**
 * ============================================================
 * 11. AI PROPOSAL GENERATION (AUTO — triggered by APPLY)
 * ============================================================
 */
var PORTFOLIO_CONTEXT_ =
  "Sawandi's completed portfolio projects: " +
  "(1) Inventory Optimization & Revenue Strategy Dashboard — tracks stock levels, revenue trends, and reorder signals across product lines. " +
  "(2) 3PL Logistics Cost & Performance Analytics Dashboard — monitors carrier performance, cost per shipment, and on-time delivery KPIs for a logistics operation. " +
  "(3) Operations Performance & KPI Monitoring Dashboard — tracks operational throughput, team performance metrics, and process efficiency KPIs. " +
  "Tools used across projects: Tableau, Power BI, Looker Studio, Excel, SQL, Google Sheets. " +
  "Background: 8 years operations experience at UPS before transitioning to data analytics.";

function generateAiProposal_(jobTitle, description, toolDetected, jobType, proposalCount, budget, keywordSearch) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) return "API key not set. Run System Tools > Setup API Key.";

  var competitionNote = proposalCount > 30
    ? "This job already has " + proposalCount + " proposals so the opener must immediately stand out."
    : proposalCount > 15
    ? "This job has " + proposalCount + " proposals — be specific and direct."
    : "This job has few proposals — a clear, confident proposal will stand out easily.";

  var prompt =
    "You are writing an Upwork proposal for Sawandi, a dashboard developer and data analyst. " +
    "Write a complete proposal in exactly 3 short paragraphs, under 120 words total. " +
    "Rules you must follow: " +
    "Do NOT start with Hi or the client's name. " +
    "Do NOT open with I or My or a statement about Sawandi. " +
    "Open with something specific from the job description that shows you read it carefully — reference the actual problem or tool or industry. " +
    "Second paragraph: connect one of Sawandi's portfolio projects or specific experience directly to what this client needs. Be concrete, not vague. " +
    "Third paragraph: end with ONE specific question that invites a reply. Not an offer to do free work. A question that shows you understand the project. " +
    "No bullet points. No sign-off. No filler phrases like I would love to or I am confident. Sound like a practitioner, not an applicant. " +
    competitionNote + " " +
    "PORTFOLIO AND BACKGROUND: " + PORTFOLIO_CONTEXT_ + " " +
    "JOB DETAILS: " +
    "Title: " + jobTitle + ". " +
    "Tool requested: " + (toolDetected || "not specified") + ". " +
    "Job type: " + (jobType || "dashboard") + ". " +
    "Budget: " + (budget || "not listed") + ". " +
    "Found via keyword: " + (keywordSearch || "not noted") + ". " +
    "Description: " + String(description).substring(0, 1800);

  var payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 200,
    temperature: 0.7
  };

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return "API error: " + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "No response returned.";

  } catch (err) {
    return "Request failed: " + err.message;
  }
}


/**
 * ============================================================
 * 12. AI JOB TYPE CLASSIFIER
 * ============================================================
 */
function getJobType_(description, jobTitle) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  function regexFallback_() {
    var t = (jobTitle + " " + description).toLowerCase();
    if (/fix|improve|update|modify|redesign|optimize|existing/.test(t)) return "Dashboard Fix";
    if (/clean|spreadsheet|raw data|data cleaning|csv|excel file|export/.test(t)) return "Data to Dashboard";
    if (/report|analysis|analytics(?! dashboard)|insight/.test(t) && !/dashboard/.test(t)) return "Reporting";
    return "Dashboard Build";
  }

  if (!apiKey) return regexFallback_();

  var prompt =
    "Classify this Upwork job into exactly one of these four categories: " +
    "Dashboard Build, Dashboard Fix, Data to Dashboard, Reporting. " +
    "Dashboard Build = new dashboard needed from scratch. " +
    "Dashboard Fix = existing dashboard needs fixing or improving. " +
    "Data to Dashboard = raw data needs cleaning then turned into a dashboard. " +
    "Reporting = data analysis or reporting without a dashboard. " +
    "Reply with only the category name and nothing else. " +
    "Job: " + jobTitle + ". " + description.substring(0, 800);

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        max_tokens: 10,
        temperature: 0
      }),
      muteHttpExceptions: true
    });

    var parsed = JSON.parse(response.getContentText());
    if (!parsed.choices || !parsed.choices[0]) return regexFallback_();

    var text = parsed.choices[0].message.content.trim();

    if (text.indexOf("Data to Dashboard") !== -1) return "Data to Dashboard";
    if (text.indexOf("Dashboard Fix")     !== -1) return "Dashboard Fix";
    if (text.indexOf("Dashboard Build")   !== -1) return "Dashboard Build";
    if (text.indexOf("Reporting")         !== -1) return "Reporting";
    return regexFallback_();

  } catch (err) {
    return regexFallback_();
  }
}


function RUN_JOB_CLASSIFICATION() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No rows found in Proposal_Generator.");
    return;
  }

  var map        = getHeaderMap_(sheet);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var descCol    = getCol_(map, ["Description"]);
  var jobTypeCol = getCol_(map, ["Job_Type"]);

  if (!descCol || !jobTypeCol || !titleCol) {
    ui.alert("Required columns not found. Confirm Job_Title, Description, and Job_Type columns exist.");
    return;
  }

  var validTypes = ["Dashboard Build", "Dashboard Fix", "Data to Dashboard", "Reporting"];
  var lastRow    = sheet.getLastRow();
  var jtValues   = sheet.getRange(2, jobTypeCol, lastRow - 1, 1).getValues();
  var filled     = 0;
  var skipped    = 0;

  for (var i = 0; i < jtValues.length; i++) {
    var current = String(jtValues[i][0]).trim();
    var done    = false;
    for (var v = 0; v < validTypes.length; v++) {
      if (current === validTypes[v]) { done = true; break; }
    }
    if (done) { skipped++; continue; }

    var r     = i + 2;
    var desc  = String(sheet.getRange(r, descCol).getValue()).trim();
    var title = String(sheet.getRange(r, titleCol).getValue()).trim();

    if (!desc && !title) { continue; }

    var result = getJobType_(desc, title);
    sheet.getRange(r, jobTypeCol).setValue(result);
    filled++;

    if (filled % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert("Done.\n\n✓ " + filled + " rows classified.\n→ " + skipped + " already had a Job_Type.");
}


/**
 * ============================================================
 * 13. RUN AI PROPOSALS (BATCH)
 * ============================================================
 */
function RUN_AI_PROPOSALS() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No rows found in Proposal_Generator.");
    return;
  }

  var map        = getHeaderMap_(sheet);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var descCol    = getCol_(map, ["Description"]);
  var toolCol    = getCol_(map, ["Tool_Detected"]);
  var jobTypeCol = getCol_(map, ["Job_Type"]);
  var tmplCol    = getCol_(map, ["Recommended_Template"]);
  var hookCol    = getCol_(map, ["Hook_Version"]);
  var ctaCol     = getCol_(map, ["CTA_Version"]);
  var aiPropCol  = getCol_(map, ["AI_Generated_Proposal"]);

  if (!descCol || !aiPropCol) {
    ui.alert("Required columns not found. Make sure Description and AI_Generated_Proposal columns exist.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var count   = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var desc     = descCol    ? String(data[i][descCol    - 1]).trim() : "";
    var aiProp   = aiPropCol  ? String(data[i][aiPropCol  - 1]).trim() : "";
    var jobTitle = titleCol   ? String(data[i][titleCol   - 1]).trim() : "";
    var tool     = toolCol    ? String(data[i][toolCol    - 1]).trim() : "";
    var jobType  = jobTypeCol ? String(data[i][jobTypeCol - 1]).trim() : "";
    var tmplId   = tmplCol    ? String(data[i][tmplCol    - 1]).trim().substring(0, 2) : "T1";
    var hookVer  = hookCol    ? String(data[i][hookCol    - 1]).trim() : "A";
    var ctaVer   = ctaCol     ? String(data[i][ctaCol     - 1]).trim() : "A";

    if (!desc || (aiProp && aiProp !== "" && aiProp !== "Drafting proposal...")) {
      skipped++;
      continue;
    }

    var dataRow = i + 2;
    sheet.getRange(dataRow, aiPropCol).setValue("Drafting proposal...");

    var result = generateAIProposal_(jobTitle, desc, tool, jobType,
                                      tmplId || "T1", hookVer || "A", ctaVer || "A");
    sheet.getRange(dataRow, aiPropCol).setValue(result);
    count++;

    if (count % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert(
    "Done.\n\n" +
    "✓ " + count + " AI proposals generated.\n" +
    "→ " + skipped + " rows skipped (no description or already had a proposal)."
  );
}


/**
 * ============================================================
 * 14. SESSION MANAGEMENT
 * ============================================================
 */
function START_SESSION() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var prop = PropertiesService.getScriptProperties();

  if (prop.getProperty("SESSION_ACTIVE") === "true") {
    ui.alert(
      "A session is already active (Session " + prop.getProperty("SESSION_ID") + ").\n" +
      "Use System Tools > End Session to close it before starting a new one."
    );
    return;
  }

  var idResponse = ui.prompt(
    "Start Session — Step 1 of 2",
    "Enter your Session ID (e.g. S001):",
    ui.ButtonSet.OK_CANCEL
  );
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;

  var sessionId = idResponse.getResponseText().trim().toUpperCase();
  if (!sessionId) {
    ui.alert("Session ID cannot be blank.");
    return;
  }

  var kwResponse = ui.prompt(
    "Start Session — Step 2 of 2",
    "Which keywords are you searching this session?\n" +
    "(Enter comma-separated, e.g. Excel Finance Dashboard Creation, Power BI Sales Dashboard)",
    ui.ButtonSet.OK_CANCEL
  );
  if (kwResponse.getSelectedButton() !== ui.Button.OK) return;

  var keywords = kwResponse.getResponseText().trim();
  if (!keywords) {
    ui.alert("Please enter at least one keyword.");
    return;
  }

  var discoverySheet = ss.getSheetByName("Job_Discovery");
  var startRowCount  = discoverySheet ? Math.max(discoverySheet.getLastRow() - 1, 0) : 0;

  prop.setProperties({
    "SESSION_ACTIVE":          "true",
    "SESSION_ID":              sessionId,
    "SESSION_START_TIME":      new Date().toISOString(),
    "SESSION_KEYWORDS":        keywords,
    "SESSION_START_ROW_COUNT": String(startRowCount),
    "SESSION_DUPE_COUNT":      "0"
  });

  ui.alert(
    "✓ Session " + sessionId + " started.\n\n" +
    "Keywords: " + keywords + "\n" +
    "Start time: " + formatTime_(new Date()) + "\n\n" +
    "Session target: 8 unique new jobs.\n" +
    "Go search Upwork — every job you log will be tracked automatically."
  );
}


function END_SESSION() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var prop = PropertiesService.getScriptProperties();

  if (prop.getProperty("SESSION_ACTIVE") !== "true") {
    ui.alert("No active session found.\nUse System Tools > Start Session to begin one.");
    return;
  }

  var sessionId    = prop.getProperty("SESSION_ID")           || "";
  var startTimeStr = prop.getProperty("SESSION_START_TIME")   || "";
  var keywords     = prop.getProperty("SESSION_KEYWORDS")     || "";
  var startCount   = parseInt(prop.getProperty("SESSION_START_ROW_COUNT") || "0", 10);
  var dupeCount    = parseInt(prop.getProperty("SESSION_DUPE_COUNT")      || "0", 10);

  var endTime    = new Date();
  var startTime  = startTimeStr ? new Date(startTimeStr) : endTime;
  var durationMs = endTime - startTime;

  var discoverySheet = ss.getSheetByName("Job_Discovery");
  var jobsLogged     = 0;
  var movedToScoring = 0;
  var reviewLater    = 0;

  if (discoverySheet && discoverySheet.getLastRow() > 1) {
    var discMap      = getHeaderMap_(discoverySheet);
    var sessionIdCol = getCol_(discMap, ["Session_ID"]);
    var actionCol    = getCol_(discMap, ["Discovery_Action"]);
    var totalRows    = discoverySheet.getLastRow() - 1;

    if (sessionIdCol && actionCol && totalRows > 0) {
      var discData = discoverySheet
        .getRange(2, 1, totalRows, discoverySheet.getLastColumn())
        .getValues();

      for (var i = 0; i < discData.length; i++) {
        var rowSession = String(discData[i][sessionIdCol - 1]).trim().toUpperCase();
        var rowAction  = String(discData[i][actionCol   - 1]).trim();

        if (rowSession === sessionId) {
          jobsLogged++;
          if (rowAction === "Move to Scoring") movedToScoring++;
          else if (rowAction === "Review Later") reviewLater++;
        }
      }
    }
  }

  var sessionYield = jobsLogged;
  var YIELD_TARGET = 8;
  var saturating   = sessionYield < YIELD_TARGET;
  var satFlag      = saturating
    ? "YES — yield " + sessionYield + "/" + YIELD_TARGET + " (below target)"
    : "No — yield " + sessionYield + "/" + YIELD_TARGET;

  var scoringSheet    = ss.getSheetByName("Job_Scoring");
  var applyNoProposal = 0;

  if (scoringSheet && scoringSheet.getLastRow() > 1) {
    var jsMap         = getHeaderMap_(scoringSheet);
    var jsDecisionCol = getCol_(jsMap, ["Final_Decision"]);
    var jsPropDateCol = getCol_(jsMap, ["Proposal_Generator_Date"]);

    if (jsDecisionCol && jsPropDateCol) {
      var jsData = scoringSheet
        .getRange(2, 1, scoringSheet.getLastRow() - 1, scoringSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < jsData.length; i++) {
        var decision = String(jsData[i][jsDecisionCol - 1]).trim();
        var propDate = jsData[i][jsPropDateCol - 1];
        if (decision === "APPLY" && (propDate === "" || propDate === null)) {
          applyNoProposal++;
        }
      }
    }
  }

  var proposalTrigger = applyNoProposal >= 5
    ? "⚠ YES — " + applyNoProposal + " APPLY jobs have no proposal sent. Send at least 2 before next session."
    : "No — " + applyNoProposal + " APPLY jobs pending (" + (5 - applyNoProposal) + " more needed to trigger).";

  var pgSheet          = ss.getSheetByName("Proposal_Generator");
  var proposalsSent    = 0;
  var proposalsSkipped = 0;
  var connectsSpent    = 0;

  if (pgSheet && pgSheet.getLastRow() > 1) {
    var pgMap         = getHeaderMap_(pgSheet);
    var pgStatusCol   = getCol_(pgMap, ["Proposal_Status"]);
    var pgSentDateCol = getCol_(pgMap, ["Proposal_Sent_Date"]);
    var pgConnectsCol = getCol_(pgMap, ["Total_Connects_Spent", "Connects_Required"]);

    if (pgStatusCol && pgSentDateCol) {
      var pgData = pgSheet
        .getRange(2, 1, pgSheet.getLastRow() - 1, pgSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < pgData.length; i++) {
        var pgStatus   = String(pgData[i][pgStatusCol   - 1]).trim();
        var pgSentDate = pgData[i][pgSentDateCol - 1];

        if (pgSentDate) {
          var sentDateObj = new Date(pgSentDate);
          if (sentDateObj >= startTime && sentDateObj <= endTime) {
            if (pgStatus === "Sent") {
              proposalsSent++;
              if (pgConnectsCol) {
                connectsSpent += Number(pgData[i][pgConnectsCol - 1]) || 0;
              }
            } else if (pgStatus === "Skip") {
              proposalsSkipped++;
            }
          }
        }
      }
    }
  }

  var notesResponse = ui.prompt(
    "End Session — Notes",
    "Any notes for this session? (optional — press OK to skip)",
    ui.ButtonSet.OK_CANCEL
  );
  var sessionNotes = (notesResponse.getSelectedButton() === ui.Button.OK)
    ? notesResponse.getResponseText().trim()
    : "";

  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  if (searchListSheet && keywords) {
    var slMap           = getHeaderMap_(searchListSheet);
    var slQueryCol      = getCol_(slMap, ["Search_Query"]);
    var slLastSearchCol = getCol_(slMap, ["Last_Searched"]);
    var slYieldCol      = getCol_(slMap, ["Session_Yield"]);
    var slLastRow       = searchListSheet.getLastRow();

    var keywordList = keywords.split(",").map(function (k) {
      return k.trim().toLowerCase();
    });

    if (slQueryCol && slLastRow > 1) {
      var slData = searchListSheet
        .getRange(2, 1, slLastRow - 1, searchListSheet.getLastColumn())
        .getValues();

      for (var i = 0; i < slData.length; i++) {
        var rowQuery = String(slData[i][slQueryCol - 1]).trim().toLowerCase();
        if (keywordList.indexOf(rowQuery) !== -1) {
          var dataRow = i + 2;
          if (slLastSearchCol) {
            searchListSheet.getRange(dataRow, slLastSearchCol).setValue(endTime);
          }
          if (slYieldCol) {
            var perKeyword = Math.round(sessionYield / keywordList.length);
            searchListSheet.getRange(dataRow, slYieldCol).setValue(perKeyword);
          }
        }
      }
    }
  }

  var logSheet = ss.getSheetByName("Session_Log");
  if (logSheet) {
    var logMap     = getHeaderMap_(logSheet);
    var nextLogRow = findFirstEmptyRowByColumn_(logSheet, 1);
    if (logSheet.getLastRow() <= 1) nextLogRow = 2;

    setCellValue_(logSheet, nextLogRow, logMap, ["Session_ID"],            sessionId);
    setCellValue_(logSheet, nextLogRow, logMap, ["Date"],                  endTime);
    setCellValue_(logSheet, nextLogRow, logMap, ["Start_Time"],            formatTime_(startTime));
    setCellValue_(logSheet, nextLogRow, logMap, ["End_Time"],              formatTime_(endTime));
    setCellValue_(logSheet, nextLogRow, logMap, ["Duration"],              formatDuration_(durationMs));
    setCellValue_(logSheet, nextLogRow, logMap, ["Keywords_Searched"],     keywords);
    setCellValue_(logSheet, nextLogRow, logMap, ["Jobs_Logged"],           jobsLogged);
    setCellValue_(logSheet, nextLogRow, logMap, ["Jobs_Moved_To_Scoring"], movedToScoring);
    setCellValue_(logSheet, nextLogRow, logMap, ["Jobs_Review_Later"],     reviewLater);
    setCellValue_(logSheet, nextLogRow, logMap, ["Duplicates_Skipped"],    dupeCount);
    setCellValue_(logSheet, nextLogRow, logMap, ["Session_Yield"],         sessionYield);
    setCellValue_(logSheet, nextLogRow, logMap, ["Saturation_Flag"],       saturating ? "Saturating" : "Healthy");
    setCellValue_(logSheet, nextLogRow, logMap, ["Proposal_Trigger"],      applyNoProposal >= 5 ? "TRIGGERED" : "Clear");
    setCellValue_(logSheet, nextLogRow, logMap, ["Proposals_Sent"],        proposalsSent);
    setCellValue_(logSheet, nextLogRow, logMap, ["Proposals_Skipped"],     proposalsSkipped);
    setCellValue_(logSheet, nextLogRow, logMap, ["Connects_Spent"],        connectsSpent);
    setCellValue_(logSheet, nextLogRow, logMap, ["Notes"],                 sessionNotes);
  }

  prop.deleteAllProperties();

  var runLog =
    "════════════════════════════════\n" +
    "  SESSION RUN LOG — " + sessionId + "\n" +
    "════════════════════════════════\n\n" +
    "Date:           " + endTime.toLocaleDateString() + "\n" +
    "Start:          " + formatTime_(startTime) + "\n" +
    "End:            " + formatTime_(endTime) + "\n" +
    "Duration:       " + formatDuration_(durationMs) + "\n\n" +
    "Keywords:\n  " + keywords.split(",").join("\n  ") + "\n\n" +
    "── Discovery ──────────────────\n" +
    "Jobs logged:        " + jobsLogged + "\n" +
    "Moved to Scoring:   " + movedToScoring + "\n" +
    "Review Later:       " + reviewLater + "\n" +
    "Duplicates skipped: " + dupeCount + "\n" +
    "Session yield:      " + sessionYield + " / " + YIELD_TARGET + " target\n\n" +
    "── Health ─────────────────────\n" +
    "Saturation flag:    " + satFlag + "\n" +
    "Proposal trigger:   " + proposalTrigger + "\n\n" +
    "── Proposals ──────────────────\n" +
    "Sent this session:  " + proposalsSent + "\n" +
    "Skipped this session: " + proposalsSkipped + "\n" +
    (connectsSpent > 0 ? "Connects spent:     " + connectsSpent + "\n" : "") +
    "\n" +
    (sessionNotes ? "Notes: " + sessionNotes + "\n\n" : "") +
    (logSheet ? "✓ Log written to Session_Log." : "⚠ Session_Log sheet not found — log not saved.");

  ui.alert(runLog);
}


/**
 * ============================================================
 * 15. SNAPSHOT MONTH END
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
    "MONTH END SNAPSHOT — " + monthName + " " + year + "\n" +
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


/**
 * ============================================================
 * 16. MAIN EDIT TRIGGER
 * ============================================================
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var ss        = e.source;
  var sheet     = e.range.getSheet();
  var sheetName = sheet.getName();
  var row       = e.range.getRow();
  var col       = e.range.getColumn();

  if (row <= 1) return;

  var map = getHeaderMap_(sheet);

  // ----------------------------------------------------------
  // 0) JOB_DISCOVERY
  // ----------------------------------------------------------
  if (sheetName === "Job_Discovery") {
    var descColJD    = getCol_(map, ["Description"]);
    var dateFoundCol = getCol_(map, ["Date_Found"]);
    var linkColJD    = getCol_(map, ["Job_Link"]);
    var sessionIdCol = getCol_(map, ["Session_ID"]);

    if (descColJD && col === descColJD) {
      var descriptionCell = sheet.getRange(row, descColJD);
      var rawText         = descriptionCell.getValue();

      if (rawText !== "" && rawText !== null && rawText !== undefined) {
        var cleanedText = cleanJobText_(rawText);
        if (cleanedText !== rawText) {
          descriptionCell.setValue(cleanedText);
        }

        if (dateFoundCol) {
          var timestampCell = sheet.getRange(row, dateFoundCol);
          if (timestampCell.getValue() === "") {
            timestampCell.setValue(new Date());
          }
        }

        if (sessionIdCol) {
          var prop      = PropertiesService.getScriptProperties();
          var active    = prop.getProperty("SESSION_ACTIVE");
          var sessionId = prop.getProperty("SESSION_ID");
          if (active === "true" && sessionId) {
            var sessionCell = sheet.getRange(row, sessionIdCol);
            if (sessionCell.getValue() === "") {
              sessionCell.setValue(sessionId);
            }
          }
        }

        var quickNotesCol = getCol_(map, ["Quick_Notes"]);
        if (quickNotesCol) {
          sheet.getRange(row, quickNotesCol).setValue("Analyzing...");
          var quickResult = getQuickNotes_(cleanedText || rawText);
          sheet.getRange(row, quickNotesCol).setValue(quickResult);
        }
      }
    }

    if (linkColJD && col === linkColJD) {
      var pastedLink = sheet.getRange(row, linkColJD).getValue();
      var prop2      = PropertiesService.getScriptProperties();
      var active2    = prop2.getProperty("SESSION_ACTIVE");

      if (active2 === "true" && pastedLink) {
        var lastRow2   = sheet.getLastRow();
        var linkValues = sheet.getRange(2, linkColJD, lastRow2 - 1, 1).getValues();
        var linkStr    = String(pastedLink).trim();
        var matchCount = 0;

        for (var i = 0; i < linkValues.length; i++) {
          if (String(linkValues[i][0]).trim() === linkStr) matchCount++;
        }

        if (matchCount > 1) {
          var dupeCount = parseInt(prop2.getProperty("SESSION_DUPE_COUNT") || "0", 10);
          prop2.setProperty("SESSION_DUPE_COUNT", String(dupeCount + 1));

          SpreadsheetApp.getUi().alert(
            "⚠ Duplicate Detected — Session " + prop2.getProperty("SESSION_ID") + "\n\n" +
            "This job link already exists in Job_Discovery.\n" +
            "You can delete this row and skip to the next job.\n\n" +
            "Duplicate count this session: " + (dupeCount + 1)
          );
        }
      }

      colorDuplicateJobLinks();
    }

    return;
  }

  // ----------------------------------------------------------
  // 1) JOB_SCORING
  // ----------------------------------------------------------
  if (sheetName === "Job_Scoring") {
    var jobTitleColJS      = getCol_(map, ["Job_Title"]);
    var dateScoredCol      = getCol_(map, ["Date_Scored"]);
    var finalDecisionCol   = getCol_(map, ["Final_Decision"]);
    var proposalGenDateCol = getCol_(map, ["Proposal_Generator_Date"]);

    if (jobTitleColJS && col === jobTitleColJS && dateScoredCol) {
      var dateScoredCell = sheet.getRange(row, dateScoredCol);
      if (dateScoredCell.getValue() === "") {
        dateScoredCell.setValue(new Date());
      }
    }

    if (finalDecisionCol && proposalGenDateCol) {
      var finalDecision       = sheet.getRange(row, finalDecisionCol).getValue();
      var proposalGenDateCell = sheet.getRange(row, proposalGenDateCol);

      if (finalDecision === "APPLY" && proposalGenDateCell.getValue() === "") {
        proposalGenDateCell.setValue(new Date());

        var jsMap2         = getHeaderMap_(sheet);
        var jsTitleCol2    = getCol_(jsMap2, ["Job_Title"]);
        var jsDescCol2     = getCol_(jsMap2, ["Description"]);
        var jsToolCol2     = getCol_(jsMap2, ["Tool_Detected"]);
        var jsKeywordCol2  = getCol_(jsMap2, ["Keyword_Search"]);
        var jsProposalCol2 = getCol_(jsMap2, ["Proposal_Count"]);
        var jsBudgetCol2   = getCol_(jsMap2, ["Budget"]);
        var jsQuickCol2    = getCol_(jsMap2, ["Quick_Notes"]);

        var aiJobTitle      = jsTitleCol2    ? sheet.getRange(row, jsTitleCol2).getValue()    : "";
        var aiDescription   = jsDescCol2     ? sheet.getRange(row, jsDescCol2).getValue()     : "";
        var aiTool          = jsToolCol2     ? sheet.getRange(row, jsToolCol2).getValue()     : "";
        var aiKeyword       = jsKeywordCol2  ? sheet.getRange(row, jsKeywordCol2).getValue()  : "";
        var aiProposalCount = jsProposalCol2 ? sheet.getRange(row, jsProposalCol2).getValue() : "";
        var aiBudget        = jsBudgetCol2   ? sheet.getRange(row, jsBudgetCol2).getValue()   : "";
        var aiQuickNotes    = jsQuickCol2    ? sheet.getRange(row, jsQuickCol2).getValue()    : "";

        var aiJobType = "Dashboard Build";
        if (aiQuickNotes) {
          var qn = String(aiQuickNotes).toLowerCase();
          if (qn.indexOf("fix") !== -1 || qn.indexOf("update") !== -1 || qn.indexOf("improve") !== -1) {
            aiJobType = "Dashboard Fix";
          } else if (qn.indexOf("data") !== -1 && qn.indexOf("dashboard") === -1) {
            aiJobType = "Data to Dashboard";
          }
        }

        if (aiJobTitle && aiDescription) {
          var pgSheet = ss.getSheetByName("Proposal_Generator");
          if (pgSheet && pgSheet.getLastRow() > 1) {
            var pgMap2      = getHeaderMap_(pgSheet);
            var pgTitleCol2 = getCol_(pgMap2, ["Job_Title"]);
            var pgAiCol2    = getCol_(pgMap2, ["AI_Proposal"]);

            if (pgTitleCol2 && pgAiCol2) {
              var pgData = pgSheet
                .getRange(2, 1, pgSheet.getLastRow() - 1, pgSheet.getLastColumn())
                .getValues();

              for (var p = 0; p < pgData.length; p++) {
                if (String(pgData[p][pgTitleCol2 - 1]).trim() === String(aiJobTitle).trim()) {
                  var pgRow = p + 2;
                  pgSheet.getRange(pgRow, pgAiCol2).setValue("Generating proposal...");
                  var aiProposalText = generateAiProposal_(
                    aiJobTitle, aiDescription, aiTool,
                    aiJobType, aiProposalCount, aiBudget, aiKeyword
                  );
                  pgSheet.getRange(pgRow, pgAiCol2).setValue(aiProposalText);
                  break;
                }
              }
            }
          }
        }
      }
    }

    return;
  }

  // ----------------------------------------------------------
  // 2) CONNECTS_HELPER
  // ----------------------------------------------------------
  if (sheetName === "Connects_Helper") {
    var metricCol = getCol_(map, ["Metric"]);
    var valueCol  = getCol_(map, ["Value"]);

    if (metricCol && valueCol && col === valueCol) {
      var metricLabel = sheet.getRange(row, metricCol).getValue();

      if (String(metricLabel).trim() === "Connect_Replenishment") {
        var newVal = sheet.getRange(row, valueCol).getValue();
        if (newVal !== "" && newVal !== null) {
          var lastRow    = sheet.getLastRow();
          var allMetrics = sheet.getRange(2, metricCol, lastRow - 1, 1).getValues();

          for (var m = 0; m < allMetrics.length; m++) {
            var mLabel = String(allMetrics[m][0]).trim();

            // Stamp Connect_Replenishment_Date
            if (mLabel === "Connect_Replenishment_Date") {
              sheet.getRange(m + 2, valueCol).setValue(new Date());
            }

            // Accumulate Total_Connects_Purchased
            if (mLabel === "Total_Connects_Purchased") {
              var currentTotal    = sheet.getRange(m + 2, valueCol).getValue();
              var currentTotalNum = Number(currentTotal) || 0;
              sheet.getRange(m + 2, valueCol).setValue(currentTotalNum + Number(newVal));
            }
          }
        }
      }
    }
    return;
  }

  // ----------------------------------------------------------
  // 3) PROPOSAL_GENERATOR
  // ----------------------------------------------------------
  if (sheetName === "Proposal_Generator") {

    var bid3Col = getCol_(map, ["Bid_3rd"]);
    if (bid3Col && col === bid3Col) {
      var bid3Val = sheet.getRange(row, bid3Col).getValue();
      if (bid3Val !== "" && bid3Val !== null) {
        var pgMap2       = map;
        var bid1Col      = getCol_(pgMap2, ["Bid_1st"]);
        var bid2Col      = getCol_(pgMap2, ["Bid_2nd"]);
        var baseConCol   = getCol_(pgMap2, ["Connects_Required"]);
        var propCountCol = getCol_(pgMap2, ["Proposal_Count"]);
        var titleCol2    = getCol_(pgMap2, ["Job_Title"]);
        var recCol       = getCol_(pgMap2, ["Bid_Recommendation"]);

        var bid1Val      = bid1Col      ? sheet.getRange(row, bid1Col).getValue()      : "";
        var bid2Val      = bid2Col      ? sheet.getRange(row, bid2Col).getValue()      : "";
        var baseConVal   = baseConCol   ? sheet.getRange(row, baseConCol).getValue()   : "";
        var propCountVal = propCountCol ? sheet.getRange(row, propCountCol).getValue() : "";
        var titleVal     = titleCol2    ? sheet.getRange(row, titleCol2).getValue()    : "";

        var totalScoreVal  = "";
        var scoringSheet2  = ss.getSheetByName("Job_Scoring");
        if (scoringSheet2 && titleVal) {
          var jsMap2      = getHeaderMap_(scoringSheet2);
          var jsTitleCol2 = getCol_(jsMap2, ["Job_Title"]);
          var jsScoreCol2 = getCol_(jsMap2, ["Total_Score"]);
          if (jsTitleCol2 && jsScoreCol2 && scoringSheet2.getLastRow() > 1) {
            var jsRows = scoringSheet2
              .getRange(2, 1, scoringSheet2.getLastRow() - 1, scoringSheet2.getLastColumn())
              .getValues();
            for (var s = 0; s < jsRows.length; s++) {
              if (String(jsRows[s][jsTitleCol2 - 1]).trim() === String(titleVal).trim()) {
                totalScoreVal = jsRows[s][jsScoreCol2 - 1];
                break;
              }
            }
          }
        }

        if (recCol) {
          sheet.getRange(row, recCol).setValue("Analyzing...");
          var recommendation = getBidRecommendation_(
            titleVal, baseConVal, propCountVal,
            totalScoreVal, bid1Val, bid2Val, bid3Val
          );
          sheet.getRange(row, recCol).setValue(recommendation);
        }
      }
    }

    var boostColPG = getCol_(map, ["Boost_Connects"]);
    if (boostColPG && col === boostColPG) {
      var boostVal = sheet.getRange(row, boostColPG).getValue();
      if (boostVal !== "" && boostVal !== null) {
        var pgMap3       = map;
        var titleColPG   = getCol_(pgMap3, ["Job_Title"]);
        var descColPG    = getCol_(pgMap3, ["Description"]);
        var toolColPG    = getCol_(pgMap3, ["Tool_Detected"]);
        var jobTypeColPG = getCol_(pgMap3, ["Job_Type"]);
        var tmplColPG    = getCol_(pgMap3, ["Recommended_Template"]);
        var hookColPG    = getCol_(pgMap3, ["Hook_Version"]);
        var ctaColPG     = getCol_(pgMap3, ["CTA_Version"]);
        var aiPropColPG  = getCol_(pgMap3, ["AI_Generated_Proposal"]);

        var pgTitle   = titleColPG   ? sheet.getRange(row, titleColPG).getValue()   : "";
        var pgDesc    = descColPG    ? sheet.getRange(row, descColPG).getValue()    : "";
        var pgTool    = toolColPG    ? sheet.getRange(row, toolColPG).getValue()    : "";
        var pgJobType = jobTypeColPG ? sheet.getRange(row, jobTypeColPG).getValue() : "";
        var pgTemplate = tmplColPG   ? sheet.getRange(row, tmplColPG).getValue()    : "";
        var pgHook    = hookColPG    ? sheet.getRange(row, hookColPG).getValue()    : "";
        var pgCta     = ctaColPG     ? sheet.getRange(row, ctaColPG).getValue()     : "";

        if (pgDesc && aiPropColPG) {
          var tmplId  = String(pgTemplate).trim().substring(0, 2) || "T1";
          var hookVer = pgHook || "A";
          var ctaVer  = pgCta  || "A";

          sheet.getRange(row, aiPropColPG).setValue("Drafting proposal...");
          var aiProposal = generateAIProposal_(
            pgTitle, pgDesc, pgTool, pgJobType, tmplId, hookVer, ctaVer
          );
          sheet.getRange(row, aiPropColPG).setValue(aiProposal);
        }
      }
    }

    var proposalStatusCol   = getCol_(map, ["Proposal_Status"]);
    var proposalSentDateCol = getCol_(map, ["Proposal_Sent_Date"]);

    if (!proposalStatusCol || !proposalSentDateCol) return;
    if (col !== proposalStatusCol) return;

    var proposalStatus = sheet.getRange(row, proposalStatusCol).getValue();
    if (proposalStatus !== "Sent") return;

    var proposalSentDateCell = sheet.getRange(row, proposalSentDateCol);
    if (proposalSentDateCell.getValue() !== "") return;

    var tracker  = ss.getSheetByName("Proposal_Tracker");
    var followup = ss.getSheetByName("Followup_Tracker");
    var scoring  = ss.getSheetByName("Job_Scoring");

    if (!tracker || !followup || !scoring) return;

    var pgMap = map;
    var ptMap = getHeaderMap_(tracker);
    var ftMap = getHeaderMap_(followup);
    var jsMap = getHeaderMap_(scoring);

    var dateInGenerator    = getCellValue_(sheet, row, pgMap, ["Date"]);
    var jobTitle           = getCellValue_(sheet, row, pgMap, ["Job_Title"]);
    var clientName         = getCellValue_(sheet, row, pgMap, ["Client_Name", "Client Name"]);
    var toolRequested      = getCellValue_(sheet, row, pgMap, ["Tool_Detected", "Tool_Requested"]);
    var templateUsed       = getCellValue_(sheet, row, pgMap, ["Recommended_Template", "Template_Used"]);
    var hookVersion        = getCellValue_(sheet, row, pgMap, ["Hook_Version"]);
    var ctaVersion         = getCellValue_(sheet, row, pgMap, ["CTA_Version"]);
    var notes              = getCellValue_(sheet, row, pgMap, ["Notes"]);
    var jobLink            = getCellValue_(sheet, row, pgMap, ["Job_Link"]);
    var boostConnects      = getCellValue_(sheet, row, pgMap, ["Boost_Connects"]);
    var totalConnectsSpent = getCellValue_(sheet, row, pgMap, ["Total_Connects_Spent"]);

    var scoringLastRow = scoring.getLastRow();
    var scoringLastCol = scoring.getLastColumn();
    var scoringData    = [];
    if (scoringLastRow > 1 && scoringLastCol > 0) {
      scoringData = scoring.getRange(2, 1, scoringLastRow - 1, scoringLastCol).getValues();
    }

    var keywordSearch    = "";
    var daysSincePosted  = "";
    var hoursSincePosted = "";
    var proposalCount    = "";
    var totalScore       = "";
    var ageDays          = "";
    var currentAgeDays   = "";
    var connectsUsed     = totalConnectsSpent !== "" ? totalConnectsSpent
                           : (boostConnects !== "" ? (Number(boostConnects) + 0) : "");

    var jsJobTitleCol      = getCol_(jsMap, ["Job_Title"]);
    var jsClientCol        = getCol_(jsMap, ["Client_Name", "Client Name"]);
    var jsKeywordCol       = getCol_(jsMap, ["Keyword_Search"]);
    var jsDaysCol          = getCol_(jsMap, ["Days_Since_Posted"]);
    var jsHoursCol         = getCol_(jsMap, ["Hours_Since_Posted"]);
    var jsProposalCountCol = getCol_(jsMap, ["Proposal_Count"]);
    var jsConnectsCol      = getCol_(jsMap, ["Connects_Required"]);
    var jsJobLinkCol       = getCol_(jsMap, ["Job_Link"]);
    var jsTotalScoreCol    = getCol_(jsMap, ["Total_Score"]);
    var jsDateScoredCol    = getCol_(jsMap, ["Date_Scored"]);

    for (var i = 0; i < scoringData.length; i++) {
      var rowTitle  = jsJobTitleCol ? scoringData[i][jsJobTitleCol - 1] : "";
      var rowClient = jsClientCol   ? scoringData[i][jsClientCol   - 1] : "";

      var titleMatch  = rowTitle === jobTitle;
      var clientMatch = !clientName || !rowClient || rowClient === clientName;

      if (titleMatch && clientMatch) {
        keywordSearch    = jsKeywordCol       ? scoringData[i][jsKeywordCol       - 1] : "";
        daysSincePosted  = jsDaysCol          ? scoringData[i][jsDaysCol          - 1] : "";
        hoursSincePosted = jsHoursCol         ? scoringData[i][jsHoursCol         - 1] : "";
        proposalCount    = jsProposalCountCol ? scoringData[i][jsProposalCountCol - 1] : "";
        if (connectsUsed === "") {
          connectsUsed = jsConnectsCol ? scoringData[i][jsConnectsCol - 1] : "";
        }
        if (!jobLink) {
          jobLink = jsJobLinkCol ? scoringData[i][jsJobLinkCol - 1] : "";
        }
        totalScore = jsTotalScoreCol ? scoringData[i][jsTotalScoreCol - 1] : "";
        if (jsDateScoredCol && scoringData[i][jsDateScoredCol - 1]) {
          var scoredDate = new Date(scoringData[i][jsDateScoredCol - 1]);
          var today      = new Date();
          ageDays = Math.floor((today - scoredDate) / (1000 * 60 * 60 * 24));
        }
        if (jsDateScoredCol && scoringData[i][jsDateScoredCol - 1]) {
          var anchorDate  = new Date(scoringData[i][jsDateScoredCol - 1]);
          var now         = new Date();
          var elapsedDays = (now - anchorDate) / (1000 * 60 * 60 * 24);
          var hVal        = parseFloat(hoursSincePosted) || 0;
          var dVal        = parseFloat(daysSincePosted)  || 0;
          if (hVal > 0) {
            currentAgeDays = Math.round(((hVal / 24) + elapsedDays) * 10) / 10;
          } else if (dVal > 0) {
            currentAgeDays = Math.round((dVal + elapsedDays) * 10) / 10;
          }
        }
        break;
      }
    }

    var sentDate    = new Date();
    var appliedDate = dateInGenerator || sentDate;

    function existsInProposalTracker_() {
      var tJobCol      = getCol_(ptMap, ["Job_Title"]);
      var tClientCol   = getCol_(ptMap, ["Client_Name", "Client Name"]);
      var tTemplateCol = getCol_(ptMap, ["Template_Used", "Recommended_Template"]);
      var tHookCol     = getCol_(ptMap, ["Hook_Version"]);
      var tCtaCol      = getCol_(ptMap, ["CTA_Version"]);

      if (!tJobCol || !tClientCol || !tTemplateCol || !tHookCol || !tCtaCol) return false;

      var lastRealRow = tracker.getLastRow();
      if (lastRealRow < 2) return false;

      var data = tracker.getRange(2, 1, lastRealRow - 1, tracker.getLastColumn()).getValues();
      for (var i = 0; i < data.length; i++) {
        if (
          data[i][tJobCol      - 1] === jobTitle     &&
          data[i][tClientCol   - 1] === clientName   &&
          data[i][tTemplateCol - 1] === templateUsed &&
          data[i][tHookCol     - 1] === hookVersion  &&
          data[i][tCtaCol      - 1] === ctaVersion
        ) {
          return true;
        }
      }
      return false;
    }

    function existsInFollowupTracker_() {
      var fJobCol      = getCol_(ftMap, ["Job_Title"]);
      var fClientCol   = getCol_(ftMap, ["Client_Name", "Client Name"]);
      var fTemplateCol = getCol_(ftMap, ["Template_Used"]);

      if (!fJobCol || !fClientCol || !fTemplateCol) return false;

      var lastRealRow = followup.getLastRow();
      if (lastRealRow < 2) return false;

      var data = followup.getRange(2, 1, lastRealRow - 1, followup.getLastColumn()).getValues();
      for (var i = 0; i < data.length; i++) {
        if (
          data[i][fJobCol      - 1] === jobTitle     &&
          data[i][fClientCol   - 1] === clientName   &&
          data[i][fTemplateCol - 1] === templateUsed
        ) {
          return true;
        }
      }
      return false;
    }

    if (!existsInProposalTracker_()) {
      var ptJobTitleCol  = getCol_(ptMap, ["Job_Title"]);
      var nextTrackerRow = findFirstEmptyRowByColumn_(tracker, ptJobTitleCol);

      setCellValue_(tracker, nextTrackerRow, ptMap, ["Date_Applied"],                    appliedDate);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Job_Title"],                       jobTitle);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Client_Name", "Client Name"],      clientName);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Keyword_Search"],                  keywordSearch);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Tool_Requested", "Tool_Detected"], toolRequested);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Days_Since_Posted"],               daysSincePosted);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Proposal_Count"],                  proposalCount);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Total_Score"],                     totalScore);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Template_Used"],                   templateUsed);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Hook_Version"],                    hookVersion);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["CTA_Version"],                     ctaVersion);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Client_Replied"],                  "N");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Interview"],                       "N");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Hired"],                           "N");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Revenue"],                         "");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Notes"],                           notes);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Age_Days"],                        ageDays);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Current_Age_Days"],                currentAgeDays);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Connects_Used"],                   connectsUsed);
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Boost_Connects"],                  boostConnects !== "" ? boostConnects : 0);
      var totalForCost = connectsUsed !== "" ? Number(connectsUsed) : 0;
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Proposal_Cost"],                   totalForCost > 0 ? "$" + (totalForCost * 0.15).toFixed(2) : "");
      setCellValue_(tracker, nextTrackerRow, ptMap, ["Job_Link"],                        jobLink);
    }

    if (!existsInFollowupTracker_()) {
      var fJobTitleCol    = getCol_(ftMap, ["Job_Title"]);
      var nextFollowupRow = findFirstEmptyRowByColumn_(followup, fJobTitleCol);

      setCellValue_(followup, nextFollowupRow, ftMap, ["Date_Applied"],               appliedDate);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Job_Title"],                  jobTitle);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Client_Name", "Client Name"], clientName);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Template_Used"],              templateUsed);
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup1_Sent"],             "");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup1_Template"],         "F1");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup2_Sent"],             "");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup2_Template"],         "F2");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup3_Sent"],             "");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Followup3_Template"],         "F3");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Client_Replied"],             "N");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Interview"],                  "N");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Hired"],                      "N");
      setCellValue_(followup, nextFollowupRow, ftMap, ["Notes"],                      notes);
    }

    proposalSentDateCell.setValue(sentDate);
  }
}
