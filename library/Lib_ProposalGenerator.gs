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


// Shared by every generator below. A real draft once answered "Yes, I set
// up Google Ads for a healthcare provider" and quoted a $500 price, neither
// of which the freelancer had ever said. Placed last in each prompt so it
// overrides any earlier "be specific" instruction.
var GROUNDING_RULES_ =
  ' GROUNDING RULES (these override every instruction above): ' +
  'Only claim experience, clients, industries, tools, results, and numbers that appear in the ' +
  'portfolio/background or the job post given here. Never invent a client, employer, industry, ' +
  'metric, percentage, count of years, or project. If the job asks about experience that is not ' +
  'in the portfolio/background, say so plainly and point to the closest related work instead. ' +
  'Never state a price, rate, budget, timeline, or availability -- write [YOUR RATE], [TIMELINE], ' +
  'or [AVAILABILITY] where one is needed so the freelancer fills it in.';

// Questions and instructions a client wrote into the job description for
// applicants to answer in the proposal itself ("please answer these in your
// proposal", "start your proposal with the word banana"). Job 119 on
// 2026-10-01 had four of them and the draft answered none. Plain text
// matching, not AI: sentences ending in "?" plus a list of instruction
// phrases. Shown to the freelancer in Log Proposal Bid's Step 3 and handed
// to both cover-letter generators as an explicit list, since a long
// description gets cut before its closing questions reach the prompt.
var DESCRIPTION_INSTRUCTION_RE_ = new RegExp(
  '\\b(in your (proposal|cover letter|application|reply|response|bid)' +
  '|please (answer|include|mention|describe|share|explain|list|tell)' +
  '|(start|begin) (your )?(proposal|application|cover letter|reply) with' +
  '|include the (word|phrase)' +
  '|answer (the following|these|this|each)' +
  '|when (you )?apply)\\b', 'i');

function extractDescriptionQuestions(description) {
  // Numbered/bulleted list markers ("1.", "2)", "- ") become breaks so
  // "answer these: 1. X? 2. Y?" splits into the lead-in plus each question.
  var text = String(description || '')
    .replace(/\s+/g, ' ')
    .replace(/(^|\s)(\d{1,2}[.)]|[-*•])\s+/g, '\n')
    .trim();
  if (!text) return [];

  var seen = {};
  return text.split(/(?<=[.?!:])\s+|\n+/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) {
      if (s.length < 12) return false;
      var hit = /\?$/.test(s) || DESCRIPTION_INSTRUCTION_RE_.test(s);
      var key = s.toLowerCase();
      if (!hit || seen[key]) return false;
      seen[key] = true;
      return true;
    })
    .map(function (s) { return s.length > 300 ? s.substring(0, 300).trim() + '...' : s; })
    .slice(0, 8);
}

// Prompt block for the cover-letter generators; empty when the description
// asks nothing.
function descriptionQuestionsBlock_(description) {
  var qs = extractDescriptionQuestions(description);
  if (qs.length === 0) return '';
  // A softer "answer these inside the body" version lost to each template's
  // fixed shape (3 paragraphs, no lists, end with one question) in a live
  // test on job 119: the draft talked around all four questions and turned
  // one of them back into its closing question. A labelled block after the
  // letter is also how clients expect numbered questions answered.
  return ' QUESTIONS AND INSTRUCTIONS IN THE JOB POST: ' +
    qs.map(function (q, i) { return '(' + (i + 1) + ') ' + q; }).join(' ') +
    ' The client asked applicants to respond to these. This overrides the length, no-list, and ' +
    'closing-question rules above: after the proposal, add a blank line, then the heading ' +
    '"Your questions:", then a numbered answer for each question that is a real request to ' +
    'applicants, in the client\'s order, one or two sentences each. Skip lead-in lines like ' +
    '"please answer these" and anything rhetorical. If an instruction is about the proposal ' +
    'itself (such as a word to start with), follow it in the proposal instead of answering it. ' +
    'Do not repeat a client question as your own closing question.';
}

// Deterministic check after generation: any dollar amount or number in the
// draft that doesn't appear in the material the AI was given gets listed on
// a warning line above the draft. Nothing is removed -- the freelancer
// decides. Numbers glued to letters (GA4, O365) aren't standalone numbers
// and are skipped. Error strings from a failed call pass through untouched.
function flagUngroundedNumbers_(text, sources) {
  var draft = String(text || '');
  if (!draft || /^(API error|Request failed|API key not set|No response)/.test(draft)) return draft;

  var numberRe = /\$?\b\d[\d,]*(?:\.\d+)?\b(?:\s?(?:k|K|%|percent|years?|yrs?|months?|weeks?|days?|hours?|hrs?))?/g;
  var digitsOf = function (s) { return String(s).replace(/[^\d.]/g, '').replace(/\.$/, ''); };

  var known = {};
  (String(sources.join(' ')).match(numberRe) || []).forEach(function (m) { known[digitsOf(m)] = true; });

  var seen   = {};
  var flags  = [];
  (draft.match(numberRe) || []).forEach(function (m) {
    var d = digitsOf(m);
    if (!d || known[d] || seen[m]) return;
    seen[m] = true;
    flags.push('"' + m.trim() + '"');
  });

  if (flags.length === 0) return draft;
  return '⚠ CHECK BEFORE SENDING: ' + flags.join(', ') +
    ' not found in your portfolio or the job post. Delete this line before sending.\n\n' + draft;
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
    'QUESTIONS:\n' + questions + '\n' +
    GROUNDING_RULES_;

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: 400,
    temperature: 0.2
  };
  var sources = [portfolioContext, jobTitle, description, questions];

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
      ? flagUngroundedNumbers_(data.choices[0].message.content.trim(), sources)
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
    'JOB DESCRIPTION: ' + description.substring(0, 1200) +
    descriptionQuestionsBlock_(description) +
    GROUNDING_RULES_;

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: lengthSpec.maxTokens,
    temperature: 0.5
  };
  var sources = [journeyContext, portfolioAll, cred, jobTitle, toolDetected, jobType, description];

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
      ? flagUngroundedNumbers_(data.choices[0].message.content.trim(), sources)
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
    'Description: ' + String(description).substring(0, 1800) +
    descriptionQuestionsBlock_(description) +
    GROUNDING_RULES_;

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: lengthSpec.maxTokens,
    temperature: 0.7
  };
  var sources = [portfolioContext, forcedProject, jobTitle, toolDetected, jobType, proposalCount, budget, keywordSearch, description];

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
      ? flagUngroundedNumbers_(data.choices[0].message.content.trim(), sources)
      : 'No response returned.';

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}
