/**
 * ============================================================
 * FreelanceFlow Library -- AI Job-Fit Classifier (Quick Notes)
 * AI-powered (with regex fallback) job complexity classifier.
 * Returns a formatted string: "Effort, Scope & Portfolio Match"
 *
 * Niche-agnostic: judges Effort/Scope/Portfolio Match against
 * THIS user's own Primary_Tools, Portfolio projects, and
 * Freelancer_Background (passed in via settings), not a fixed
 * BI-tool vocabulary -- same Settings-driven precedent as
 * Job_Discovery's Tool_Detected/Tool_Score.
 *
 * getQuickNotes is the public entry point the thin client calls
 * (FFLib.getQuickNotes(...)). getQuickNotesRegex_ keeps its
 * trailing underscore on purpose -- Apps Script Libraries treat
 * underscore-suffixed functions as private, so this fallback
 * stays internal to the Library and is not part of its public API.
 * ============================================================
 */
function getQuickNotes(description, apiKey, settings) {
  var s             = settings || {};
  var primaryTools  = s['Primary_Tools'] || '';
  var portfolioCtx  = getPortfolioContext(s);

  if (!apiKey) {
    return getQuickNotesRegex_(description, primaryTools);
  }

  var prompt =
    "You are analyzing an Upwork job description for a freelancer. " +
    "Judge the job against THIS freelancer's own profile below, not any generic skillset:\n" +
    portfolioCtx + "\n\n" +
    "Read the job description carefully and return exactly one line in this format: " +
    "[Effort], [Scope] & [Portfolio Match]. " +
    "Rules: " +
    "Effort: Large=building an entire system/pipeline/workflow or multi-step automation from scratch. " +
    "Complex=multiple interconnected deliverables, several stages of work, or more tools/steps than the freelancer's usual single deliverable. " +
    "Normal=a single well-scoped deliverable that matches the freelancer's core listed tools/services. " +
    "Simple=a minor fix, small update, or very small scope. " +
    "Scope: Clear=step-by-step requirements/specific examples/exact deliverables. " +
    "Mostly Clear=focused scope with some gaps. " +
    "Vague=general ask/no clear deliverable. " +
    "Very Vague=no clear scope at all. " +
    "Portfolio Match: Exact=explicitly names one of the freelancer's own tools (" + primaryTools + ") or closely matches one of their listed portfolio projects. " +
    "Strong=clearly falls within the freelancer's general niche/background even without naming their exact tool. " +
    "Partial=related work but doesn't name or clearly imply the freelancer's tools or niche. " +
    "Weak=only tangentially related to the freelancer's listed skills. " +
    "None=unrelated to the freelancer's service line entirely. " +
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
    if (data.error) return getQuickNotesRegex_(description, primaryTools);

    var result = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "";

    return result || getQuickNotesRegex_(description, primaryTools);

  } catch (err) {
    return getQuickNotesRegex_(description, primaryTools);
  }
}


function getQuickNotesRegex_(description, primaryToolsCsv) {
  if (!description) return "";
  var d = description;

  var effort =
    /pipeline|workflow automation|end-to-end system|integrat.*(and|with).*multiple|from scratch/i.test(d) ? "Large" :
    /multiple deliverables|several (stages|steps|phases)|combine.*(and|with).*(clean|process)/i.test(d) ? "Complex" :
    /help us|looking for|need someone|would like|update|fix|small/i.test(d) ? "Normal" : "Simple";

  var scope =
    /step-by-step|clearly defined|specific requirements|example outputs|exactly/i.test(d) ? "Clear" :
    /focused|scoped|improve|update|redesign/i.test(d) ? "Mostly Clear" :
    /help us|looking for|need someone|would like/i.test(d) ? "Vague" : "Very Vague";

  var tools = (primaryToolsCsv || '').split(',')
    .map(function (t) { return t.trim(); })
    .filter(function (t) { return t; });

  var toolMatch = "None";
  var dLower = d.toLowerCase();
  for (var i = 0; i < tools.length; i++) {
    if (dLower.indexOf(tools[i].toLowerCase()) !== -1) {
      toolMatch = "Exact";
      break;
    }
  }
  if (toolMatch === "None" && /dashboard|report|reporting|analytics|automation|bookkeeping|clean ?up/i.test(d)) {
    toolMatch = "Partial";
  }

  return effort + ", " + scope + " & " + toolMatch;
}
