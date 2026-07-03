/**
 * ============================================================
 * FreelanceFlow Library -- Wizard AI Assist
 * Step 3 (optional) niche analysis: suggests tools/background/
 * keywords from the customer's own specialty description. The
 * thin client's wizard_analyzeNiche(description, apiKey) is the
 * google.script.run entry point the sidebar HTML calls (Library
 * functions can't be called directly from client-side HTML); it
 * delegates the actual prompt + parsing here.
 * ============================================================
 */
function analyzeNiche(description, apiKey) {
  if (!apiKey) return { ok: false, message: 'API key not found. Complete Step 2 first.' };

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
      headers: { 'Authorization': 'Bearer ' + apiKey },
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
