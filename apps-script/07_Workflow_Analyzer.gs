/**
 * ============================================================
 * 7. WORKFLOW ANALYZER
 * AI-powered structured breakdown of a job posting.
 * getWorkflowAnalysis_ returns { detailed, condensed }.
 * ANALYZE_JOB_WORKFLOW is the menu-triggered entry point.
 * ============================================================
 */
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
