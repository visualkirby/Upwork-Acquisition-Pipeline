/**
 * ============================================================
 * FreelanceFlow Library -- Proposal Generator
 *
 * getPortfolioContext: builds the freelancer portfolio string
 * from an already-fetched Settings object -- no personal data
 * in code, no SpreadsheetApp access.
 *
 * generateAIProposal: template-driven generator. Caller resolves
 * the Proposal_Templates row and Settings client-side and passes
 * them in.
 *
 * generateAiProposal: auto-triggered generator fired when a job
 * is marked APPLY in Job_Scoring.
 *
 * Both generators take an optional mandatoryProject -- the row's
 * Portfolio_Project value, picked per-job by RUN_JOB_CLASSIFICATION
 * (FFLib.pickPortfolioProject). When set, the proposal MUST cite
 * that exact project. Before this was wired through, the column was
 * written and displayed but never read back, so every proposal fell
 * back to the template's Credential_Hint or the first project in the
 * portfolio list.
 *
 * Note: template lookup (reading the Proposal_Templates sheet)
 * stays in the thin client's lookupProposalTemplate_ -- it's
 * sheet I/O, not proprietary logic.
 * ============================================================
 */

/**
 * Given a classified Job_Type, all Proposal_Templates rows, and aggregated
 * Proposal_Tracker stats keyed by "TemplateID|Hook|CTA", picks a template
 * row for RUN_JOB_CLASSIFICATION to write into Proposal_Generator.
 *
 * When more than one row matches the Job_Type (different Hook/CTA variants),
 * selection is weighted by each variant's live View rate from Proposal_Tracker
 * -- Laplace-smoothed (views+1)/(sent+2) so an untested variant starts at a
 * neutral 50/50 instead of 0, and a proven variant gets picked more often
 * without ever fully losing its shot at more data. Falls back to the
 * Is_Default-flagged row when nothing matches the Job_Type at all.
 *
 * templateRows: [{templateId, jobType, hookVersion, ctaVersion, isDefault}, ...]
 * trackerStats: { "T1|A|A": {sent: 5, viewed: 3}, ... }
 */
function pickWeightedTemplate(jobType, templateRows, trackerStats) {
  var stats   = trackerStats || {};
  var matches = (templateRows || []).filter(function (r) { return r.jobType === jobType; });

  if (matches.length === 0) {
    var fallback = (templateRows || []).filter(function (r) { return r.isDefault; });
    return fallback.length > 0 ? fallback[0] : null;
  }

  if (matches.length === 1) return matches[0];

  var weights = matches.map(function (r) {
    var key = r.templateId + '|' + r.hookVersion + '|' + r.ctaVersion;
    var s   = stats[key] || { sent: 0, viewed: 0 };
    return (s.viewed + 1) / (s.sent + 2);
  });

  var total = weights.reduce(function (a, b) { return a + b; }, 0);
  var roll  = Math.random() * total;
  var acc   = 0;
  for (var i = 0; i < matches.length; i++) {
    acc += weights[i];
    if (roll <= acc) return matches[i];
  }
  return matches[matches.length - 1];
}


/**
 * Some Upwork jobs add extra client-specified application questions beyond
 * the main cover letter. questions is whatever raw text the freelancer
 * pasted into Additional_Questions (however many questions, however
 * formatted) -- returns each one repeated verbatim with its answer, so
 * there's no need to parse/count them into fixed columns.
 */
function generateAdditionalAnswers(questions, jobTitle, description, portfolioContext, freelancerName, apiKey) {
  if (!apiKey) return 'API key not set. Run System Tools > Setup API Key first.';
  if (!questions) return '';

  var prompt =
    'You are answering additional application questions for an Upwork job proposal -- the main cover letter is written separately. ' +
    'FREELANCER: ' + (freelancerName || 'the freelancer') + '. ' +
    'PORTFOLIO AND BACKGROUND: ' + (portfolioContext || '') + ' ' +
    'JOB TITLE: ' + jobTitle + '. ' +
    'JOB DESCRIPTION: ' + String(description || '').substring(0, 1000) + '. ' +
    'Answer each question below directly and specifically, in the order given. ' +
    'Keep each answer under 40 words -- no filler, no restating the question, no greeting. ' +
    'Format your response as each question repeated verbatim, followed by your answer on the next line, ' +
    'with a blank line between question/answer pairs. ' +
    'Answer ONLY the exact question(s) listed below -- do not invent, add, or answer any question ' +
    'that is not explicitly listed, even if it seems like a typical one for this kind of job. ' +
    'If only one question is listed, return only that one question and its answer. ' +
    'QUESTIONS:\n' + questions;

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: 400,
    temperature: 0.4
  };

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return 'API error: ' + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : 'No response returned.';

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}


// Proposal_Length's 3 preset values (Short/Medium/Long, set via a Settings
// dropdown -- see applySettingsValidation_, 00_Setup_Wizard.gs). Character
// ranges stay well under Upwork's 5,000-character proposal cap; maxTokens
// is set generously above what each tier's word range needs so the model
// isn't cut off mid-sentence (roughly 1.3 tokens/word for English, plus
// headroom).
var PROPOSAL_LENGTH_SPECS_ = {
  'Short':  { words: '90-140',  chars: '500-800 characters',   maxTokens: 220 },
  'Medium': { words: '140-260', chars: '800-1,500 characters', maxTokens: 380 },
  'Long':   { words: '260-430', chars: '1,500-2,500 characters', maxTokens: 600 }
};

function resolveProposalLengthSpec_(proposalLength) {
  return PROPOSAL_LENGTH_SPECS_[proposalLength] || PROPOSAL_LENGTH_SPECS_['Medium'];
}

function getPortfolioContext(settings) {
  var s     = settings || {};
  var name  = s['Freelancer_Name']       || 'the freelancer';
  var tools = s['Primary_Tools']         || 'data analytics tools';
  var bg    = s['Freelancer_Background'] || '';

  var portfolioParts = (s['Portfolio_All'] || '')
    .split(';')
    .map(function(p) { return p.trim(); })
    .filter(function(p) { return p; });

  var context = name + "'s completed portfolio projects: ";

  if (portfolioParts.length > 0) {
    context += portfolioParts.map(function(p, i) {
      return '(' + (i + 1) + ') ' + p;
    }).join(' ');
  } else {
    context += '(No portfolio projects configured. Add them to the Projects sheet.)';
  }

  context += ' Tools used: ' + tools + '.';
  if (bg) context += ' Background: ' + bg;

  return context;
}


function generateAIProposal(jobTitle, description, toolDetected, jobType, template,
                             apiKey, journeyContext, portfolioAll, proposalTone, freelancerName, proposalLength,
                             mandatoryProject) {
  if (!apiKey) return 'API key not set. Run System Tools > Setup API Key first.';

  var lengthSpec     = resolveProposalLengthSpec_(proposalLength);
  var angle          = (template && template.angle)          || '';
  var credentialHint = (template && template.credentialHint) || '';
  var tone           = (template && template.tone)           || '';
  var ctaStyle       = (template && template.ctaStyle)       || '';
  var exampleOutput  = (template && template.exampleOutput)  || '';

  var portfolioParts = portfolioAll
    ? portfolioAll.split(';').map(function(p) { return p.trim(); }).filter(function(p) { return p; })
    : [];

  // The per-job Portfolio_Project pick (mandatoryProject) wins over the
  // template's fixed Credential_Hint, which wins over "just use the first
  // one". Rule 6 below forces the proposal to name whichever this resolves
  // to and no other.
  var forcedProject = mandatoryProject ? String(mandatoryProject).trim() : '';
  var cred = forcedProject || credentialHint || (portfolioParts.length > 0 ? portfolioParts[0] : 'a portfolio project');

  var prompt =
    'You are writing an Upwork proposal for a freelancer named ' + (freelancerName || 'the freelancer') + '. ' +
    'FREELANCER PROFILE: ' + journeyContext + ' ' +
    (portfolioAll ? 'FULL PORTFOLIO (for context only): ' + portfolioAll + '. ' : '') +
    'OVERALL TONE GUIDANCE: ' + (proposalTone || 'Direct') + '. ' +
    'STRICT RULES -- violating any rule makes the proposal unusable: ' +
    '1. Between ' + lengthSpec.words + ' words total (roughly ' + lengthSpec.chars + '). ' +
    '2. Do NOT start with Hi, Hello, or any greeting. ' +
    '3. Do NOT use bullet points or numbered lists. ' +
    '4. Do NOT list skills or tools generically. ' +
    '5. First sentence MUST reference a specific detail from the job description -- not a generic observation. ' +
    'Write the proposal in full; do not begin mid-sentence and capitalize the first word. ' +
    '6. You MUST reference this exact portfolio project by name in the proposal: ' + cred + ' -- do not substitute or add a different project. ' +
    '7. End with exactly one direct question. No offers to help. ' +
    'STRATEGIC ANGLE: ' + (angle || 'Lead with the specific client problem, not credentials.') + ' ' +
    'TONE: ' + (tone || 'Direct') + '. ' +
    'CTA STYLE: ' + (ctaStyle || 'Question') + '. ' +
    (exampleOutput ? 'VOICE EXAMPLE (match this structure and directness, write new content): ' + exampleOutput.substring(0, 250) + ' ' : '') +
    'JOB TITLE: ' + jobTitle + '. ' +
    'TOOL REQUESTED: ' + (toolDetected || 'not specified') + '. ' +
    'JOB TYPE: ' + (jobType || 'dashboard project') + '. ' +
    'JOB DESCRIPTION: ' + description.substring(0, 1200);

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: lengthSpec.maxTokens,
    temperature: 0.5
  };

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return 'API error: ' + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : 'No response returned.';

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}


function generateAiProposal(jobTitle, description, toolDetected, jobType, proposalCount, budget, keywordSearch,
                             apiKey, portfolioContext, freelancerName, proposalLength, mandatoryProject) {
  if (!apiKey) return 'API key not set. Run System Tools > Setup API Key.';

  var lengthSpec = resolveProposalLengthSpec_(proposalLength);
  var name = freelancerName || 'the freelancer';

  var competitionNote = proposalCount > 30
    ? 'This job already has ' + proposalCount + ' proposals so the opener must immediately stand out.'
    : proposalCount > 15
    ? 'This job has ' + proposalCount + ' proposals -- be specific and direct.'
    : 'This job has few proposals -- a clear, confident proposal will stand out easily.';

  var forcedProject = mandatoryProject ? String(mandatoryProject).trim() : '';
  var secondParagraphRule = forcedProject
    ? 'Second paragraph: connect this specific portfolio project, by name -- ' + forcedProject +
      ' -- directly to what this client needs. Reference that project and no other. Be concrete. '
    : 'Second paragraph: connect one of the freelancer\'s portfolio projects or specific experience directly to what this client needs. Be concrete, not vague. ';

  var prompt =
    'You are writing an Upwork proposal for ' + name + ', a freelancer. ' +
    'Write a complete proposal in exactly 3 paragraphs, between ' + lengthSpec.words +
    ' words total (roughly ' + lengthSpec.chars + '). ' +
    'Rules you must follow: ' +
    'Do NOT start with Hi or the client\'s name. ' +
    'Do NOT open with I or My or a statement about the freelancer. ' +
    'Write the proposal in full -- do not begin mid-sentence, and capitalize the first word. ' +
    'Open with something specific from the job description that shows you read it carefully -- reference the actual problem or tool or industry. ' +
    secondParagraphRule +
    'Third paragraph: end with ONE specific question that invites a reply. Not an offer to do free work. A question that shows you understand the project. ' +
    'No bullet points. No sign-off. No filler phrases like I would love to or I am confident. Sound like a practitioner, not an applicant. ' +
    competitionNote + ' ' +
    'PORTFOLIO AND BACKGROUND: ' + portfolioContext + ' ' +
    'JOB DETAILS: ' +
    'Title: ' + jobTitle + '. ' +
    'Tool requested: ' + (toolDetected || 'not specified') + '. ' +
    'Job type: ' + (jobType || 'dashboard') + '. ' +
    'Budget: ' + (budget || 'not listed') + '. ' +
    'Found via keyword: ' + (keywordSearch || 'not noted') + '. ' +
    'Description: ' + String(description).substring(0, 1800);

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: lengthSpec.maxTokens,
    temperature: 0.7
  };

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return 'API error: ' + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : 'No response returned.';

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}
