/**
 * ============================================================
 * FreelanceFlow Library -- Job Classifier
 * Classifies a job into one of the freelancer's own niche-derived
 * Job_Type categories (from Proposal_Templates, seeded at setup by
 * Lib_WizardAI's generateNicheTemplates) -- not a fixed universal
 * list. categories: [{name, notes}, ...], notes optional -- doubles
 * as classifier guidance and human documentation in the sheet.
 * Falls back to the first category with no API key or on failure.
 * ============================================================
 */
function getJobType(description, jobTitle, apiKey, categories) {
  var cats = (categories && categories.length > 0) ? categories : [{ name: 'General', notes: '' }];

  function fallback_() {
    return cats[0].name;
  }

  if (!apiKey) return fallback_();

  var categoryText = cats.map(function (c) {
    return c.notes ? (c.name + ' (' + c.notes + ')') : c.name;
  }).join('; ');

  var prompt =
    "Classify this Upwork job into exactly one of these categories: " + categoryText + ". " +
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
    if (!parsed.choices || !parsed.choices[0]) return fallback_();

    var text = parsed.choices[0].message.content.trim();

    for (var i = 0; i < cats.length; i++) {
      if (text.indexOf(cats[i].name) !== -1) return cats[i].name;
    }
    return fallback_();

  } catch (err) {
    return fallback_();
  }
}
