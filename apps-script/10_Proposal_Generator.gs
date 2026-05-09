/**
 * ============================================================
 * 10. PROPOSAL GENERATOR
 *
 * generateAIProposal_: template-driven generator. Reads
 * Proposal_Templates sheet and Settings. Triggered by
 * Boost_Connects edit or RUN_AI_PROPOSALS batch function.
 *
 * PORTFOLIO_CONTEXT_ / generateAiProposal_: auto-triggered
 * generator fired when a job is marked APPLY in Job_Scoring.
 * Uses a static portfolio context string.
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

  var settings_     = getSettings_();
  var proposalTone  = settings_["Proposal_Tone"] || "Direct";
  var portfolioAll  = settings_["Portfolio_All"]  || "";
  var cred = credentialHint || "Operations Performance & KPI Monitoring Dashboard";

  var prompt =
    "You are writing an Upwork proposal for a data analytics freelancer named Sawandi. " +
    "FREELANCER PROFILE: " + getJourneyStage_() + " " +
    (portfolioAll ? "FULL PORTFOLIO (for context only): " + portfolioAll + ". " : "") +
    "OVERALL TONE GUIDANCE: " + proposalTone + ". " +
    "STRICT RULES -- violating any rule makes the proposal unusable: " +
    "1. Under 100 words total. " +
    "2. Do NOT start with Hi, Hello, or any greeting. " +
    "3. Do NOT use bullet points or numbered lists. " +
    "4. Do NOT list skills or tools generically. " +
    "5. First sentence MUST reference a specific detail from the job description -- not a generic observation. " +
    "6. You MUST reference this exact portfolio project by name in the proposal: " + cred + " -- do not substitute a different project. " +
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


// Static portfolio context used by the auto-triggered proposal generator.
// Update this string whenever new portfolio projects are added.
var PORTFOLIO_CONTEXT_ =
  "Sawandi's completed portfolio projects: " +
  "(1) Inventory Optimization & Revenue Strategy Dashboard -- tracks stock levels, revenue trends, and reorder signals across product lines. " +
  "(2) 3PL Logistics Cost & Performance Analytics Dashboard -- monitors carrier performance, cost per shipment, and on-time delivery KPIs for a logistics operation. " +
  "(3) Operations Performance & KPI Monitoring Dashboard -- tracks operational throughput, team performance metrics, and process efficiency KPIs. " +
  "Tools used across projects: Tableau, Power BI, Looker Studio, Excel, SQL, Google Sheets. " +
  "Background: 8 years operations experience at UPS before transitioning to data analytics.";


function generateAiProposal_(jobTitle, description, toolDetected, jobType, proposalCount, budget, keywordSearch) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) return "API key not set. Run System Tools > Setup API Key.";

  var competitionNote = proposalCount > 30
    ? "This job already has " + proposalCount + " proposals so the opener must immediately stand out."
    : proposalCount > 15
    ? "This job has " + proposalCount + " proposals -- be specific and direct."
    : "This job has few proposals -- a clear, confident proposal will stand out easily.";

  var prompt =
    "You are writing an Upwork proposal for Sawandi, a dashboard developer and data analyst. " +
    "Write a complete proposal in exactly 3 short paragraphs, under 120 words total. " +
    "Rules you must follow: " +
    "Do NOT start with Hi or the client's name. " +
    "Do NOT open with I or My or a statement about Sawandi. " +
    "Open with something specific from the job description that shows you read it carefully -- reference the actual problem or tool or industry. " +
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
