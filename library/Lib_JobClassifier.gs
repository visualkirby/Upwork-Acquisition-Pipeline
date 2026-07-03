/**
 * ============================================================
 * FreelanceFlow Library -- Job Classifier
 * Classifies a job into one of four categories, with a regex
 * fallback when no API key is available or the AI call fails.
 * ============================================================
 */
function getJobType(description, jobTitle, apiKey) {
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
