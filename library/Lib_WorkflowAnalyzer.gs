/**
 * ============================================================
 * FreelanceFlow Library -- Workflow Analyzer (per-job AI analysis)
 * Dormant/unwired in the thin client today, kept here ready to
 * activate later without further extraction work.
 * ============================================================
 */
function getWorkflowAnalysis(jobTitle, description, apiKey) {
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
