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


/**
 * Fires right after analyzeNiche succeeds -- defines however many distinct
 * JOB TYPE categories actually fit this freelancer's niche (3-6, AI's call,
 * not a fixed count) plus a proposal-writing strategy per category. Replaces
 * Proposal_Templates' old hardcoded BI/dashboard sample rows and Job
 * Classifier's old hardcoded 4-category list with something niche-derived --
 * same fix pattern as Tool_Detected earlier, applied to job typing.
 *
 * Each returned template's "notes" field doubles as human documentation AND
 * classifier guidance -- Lib_JobClassifier's getJobType reads it back at
 * classification time so the categories and their meaning never drift apart.
 */
function generateNicheTemplates(description, tools, background, apiKey) {
  if (!apiKey) return { ok: false, message: 'API key not found. Complete Step 2 first.' };

  var prompt =
    'You are helping a freelancer set up an Upwork proposal-writing system. ' +
    'Based on their specialty below, define 3 to 6 distinct JOB TYPE categories that jobs in ' +
    'their niche typically fall into (e.g. for a bookkeeper: Cleanup, Ongoing Bookkeeping, ' +
    'Reconciliation, Reporting). For EACH category, write a short proposal-writing strategy. ' +
    'Return only a valid JSON array, no markdown, no explanation. Each element must have exactly ' +
    'these keys: ' +
    '"jobType" (1-3 word category name), ' +
    '"notes" (one sentence describing what qualifies a job for this category -- used to help an AI classifier route jobs correctly), ' +
    '"angle" (one plain sentence telling the writer to open with the client\'s specific problem for this ' +
    'category, then name one fact from the portfolio project the proposal cites; never tell the writer to ' +
    'describe their own skills, expertise, or passion), ' +
    '"exampleOutput" (2 to 3 short sentences in a practitioner\'s plain voice showing how an opener and a ' +
    'closing question sound for this category, ending with one specific question; no claims of experience, ' +
    'no project facts, no client names). ' +
    'Specialty: ' + description.substring(0, 500) + '. ' +
    'Tools: ' + (tools || 'not specified') + '. ' +
    'Background: ' + (background || 'not specified') + '.' +
    VOICE_RULES_;

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: 'gpt-4.1-mini',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 900,
        temperature: 0.4
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
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return { ok: false, message: 'No categories returned.' };
    }

    // Tone and CTA are fixed, not the AI's call: on 2026-10-08 the generated
    // rows picked "Creative", "Confident" and "Offer", and the proposal
    // prompts carried those into drafts that read as AI-written.
    parsed.forEach(function (t) {
      t.tone     = 'Direct';
      t.ctaStyle = 'Question';
    });

    return { ok: true, templates: parsed };
  } catch (e) {
    return { ok: false, message: 'Could not parse AI response.' };
  }
}
