/**
 * ============================================================
 * 6. QUICK NOTES
 * AI-powered (with regex fallback) job complexity classifier.
 * Returns a formatted string: "Complexity, Scope & Tool Match"
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
