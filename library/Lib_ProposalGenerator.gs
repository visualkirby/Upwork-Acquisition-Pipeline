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
// overrides any earlier "be specific" instruction. Job 126 on 2026-10-05
// added the kind-of-work line: the draft opened with "My experience in
// reorganizing SharePoint environments" when no project shows that work.
var GROUNDING_RULES_ =
  ' GROUNDING RULES (these override every instruction above): ' +
  'Only claim experience, clients, industries, tools, results, and numbers that appear in the ' +
  'portfolio/background or the job post given here. Never invent a client, employer, industry, ' +
  'metric, percentage, count of years, or project. If the job asks about experience that is not ' +
  'in the portfolio/background, say so plainly and point to the closest related work instead. ' +
  'Never claim experience with a kind of work ("my experience reorganizing SharePoint ' +
  'environments") unless a project or the background shows that exact work -- describe what the ' +
  'named project actually did instead. ' +
  'Never say or imply familiarity with a platform, tool, or industry the job names (for example ' +
  'Azure or Copilot Studio) unless it appears in the tools list or a project. If the job\'s main ' +
  'tool is not there, say so in one plain sentence and name the closest real work. ' +
  'Never state a price, rate, budget, timeline, or availability -- write [YOUR RATE], [TIMELINE], ' +
  'or [AVAILABILITY] where one is needed so the freelancer fills it in.';

// Phrases that made real drafts read as AI-written (jobs 151 and 153 on
// 2026-10-08: "keen eye for detail", "aligns perfectly", "robust system",
// "dynamic environment"). The same list feeds the prompt and the
// after-generation check, so a phrase the model slips in anyway gets flagged.
var BANNED_PHRASES_ = [
  'keen eye', 'aligns perfectly', 'robust', 'seamless', 'leverage', 'honed', 'well-versed',
  'actionable insights', 'data-driven decision', 'i am eager', "i'm eager", 'i am excited',
  "i'm excited", 'dynamic environment', 'tailor', 'deep understanding', 'passionate',
  'look no further', 'perfect fit', 'delve'
];

var VOICE_RULES_ =
  ' VOICE RULES: Write like a practitioner typing a quick reply. Plain words, short sentences, ' +
  'varied sentence length. Never use these words or phrases: ' + BANNED_PHRASES_.join(', ') + '. ' +
  'No em dashes. No lists of three adjectives. Do not restate the client\'s request back to them.';

// What to pull from the named project. Job 151 (debug a Power Automate flow)
// got the tracker's feature list when the project's own description held a
// real debugging story that fit the job.
var TASK_MATCH_RULE_ =
  ' From the named project\'s details, use the one fact that matches the kind of work this job ' +
  'is: for a fix, debug, or troubleshooting job, a problem the project hit and how it was found ' +
  'and fixed; for a build job, what the project does. Use only facts written in those details.';

// The named project's own description, cut out of the full portfolio text
// so it sits next to the rule that names it. Portfolio text comes in two
// shapes: "Name: desc; Name: desc" (Portfolio_All) and "(1) Name: desc (2)
// ... Tools used: ..." (getPortfolioContext). '' when the name isn't found.
function namedProjectDetails_(projectName, portfolioText) {
  var name = String(projectName || '').trim();
  var text = String(portfolioText || '');
  if (!name || !text) return '';
  var start = text.indexOf(name + ': ');
  if (start < 0) return '';
  var rest = text.substring(start + name.length + 2);
  var end  = rest.length;
  [/;\s/, /\s\(\d+\)\s/, /\sTools used:/].forEach(function (re) {
    var m = rest.search(re);
    if (m >= 0 && m < end) end = m;
  });
  return rest.substring(0, end).trim();
}

function namedProjectBlock_(projectName, portfolioText) {
  var details = namedProjectDetails_(projectName, portfolioText);
  return details ? ' NAMED PROJECT DETAILS (' + projectName + '): ' + details + '.' : '';
}

// Deterministic check after generation, same shape as the number check:
// a banned phrase that got through goes on a warning line above the draft.
function flagBannedPhrases_(text) {
  var draft = String(text || '');
  if (!draft || /^(API error|Request failed|API key not set|No response)/.test(draft)) return draft;
  var lower = draft.toLowerCase();
  var hits  = BANNED_PHRASES_.filter(function (p) { return lower.indexOf(p) >= 0; });
  if (draft.indexOf('—') >= 0) hits.push('em dash');
  if (hits.length === 0) return draft;
  return '⚠ CHECK BEFORE SENDING: reads as AI-written ("' + hits.join('", "') +
    '"). Reword, then delete this line.\n\n' + draft;
}

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

// Job 126 on 2026-10-05 ended with "Please answer in your application 1. In a
// paragraph or two, explain ... 2. ... 3. Briefly describe ...". None of the
// three items ends in "?" or contains an instruction phrase, so only the
// lead-in line came back and the draft made up its own questions. A list
// that directly follows a lead-in is now taken whole, item by item, in place
// of the lead-in.
var LIST_ITEM_MARK_       = '\u0001';
var DESCRIPTION_ITEM_MAX_ = 300;

function extractDescriptionQuestions(description) {
  // Numbered/bulleted list markers ("1.", "2)", "- ") start a new chunk,
  // flagged so list items can be told apart from prose.
  var text = String(description || '')
    .replace(/\s+/g, ' ')
    .replace(/(^|\s)(\d{1,2}[.)]|[-*•])\s+/g, '\n' + LIST_ITEM_MARK_)
    .trim();
  if (!text) return [];

  var chunks = text.split('\n').filter(function (c) { return c.replace(LIST_ITEM_MARK_, '').trim(); });
  var isItem = function (c) { return c && c.charAt(0) === LIST_ITEM_MARK_; };
  var bodyOf = function (c) { return isItem(c) ? c.substring(1) : c; };
  var sentencesOf = function (s) {
    return s.split(/(?<=[.?!:])\s+/).map(function (x) { return x.trim(); }).filter(function (x) { return x; });
  };
  var isHit = function (s) { return /\?$/.test(s) || DESCRIPTION_INSTRUCTION_RE_.test(s); };

  var out  = [];
  var seen = {};
  var add  = function (s) {
    s = String(s || '').trim();
    var key = s.toLowerCase();
    if (s.length < 12 || seen[key]) return;
    seen[key] = true;
    out.push(s.length > DESCRIPTION_ITEM_MAX_ ? s.substring(0, DESCRIPTION_ITEM_MAX_).trim() + '...' : s);
  };

  for (var i = 0; i < chunks.length; i++) {
    var sentences = sentencesOf(bodyOf(chunks[i]));
    var lastSentence = sentences.length ? sentences[sentences.length - 1] : '';

    if (isItem(chunks[i + 1]) && DESCRIPTION_INSTRUCTION_RE_.test(lastSentence)) {
      sentences.slice(0, -1).forEach(function (s) { if (isHit(s)) add(s); });

      var j = i + 1;
      for (; j < chunks.length && isItem(chunks[j]); j++) {
        var itemSentences = sentencesOf(bodyOf(chunks[j]));
        // Line breaks are gone by this point, so the last item runs into
        // whatever prose follows the list. Keep only its first sentence, plus
        // what follows a colon ("address this hypothetical: A department...").
        if (isItem(chunks[j + 1])) {
          add(itemSentences.join(' '));
        } else {
          var kept = itemSentences[0] || '';
          for (var k = 1; k < itemSentences.length && /:$/.test(kept); k++) {
            kept += ' ' + itemSentences[k];
          }
          add(kept);
        }
      }
      i = j - 1;
      continue;
    }

    sentences.forEach(function (s) { if (isHit(s)) add(s); });
  }

  return out.slice(0, 8);
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
  // Job 126 (2026-10-05) exposed two more failures once its questions came
  // through: the draft answered each question twice (in the letter body and
  // again under "Your questions:"), ran out of tokens mid-answer, and claimed
  // "I have reorganized a tenant..." when no portfolio project says so.
  return ' QUESTIONS AND INSTRUCTIONS IN THE JOB POST: ' +
    qs.map(function (q, i) { return '(' + (i + 1) + ') ' + q; }).join(' ') +
    ' The client asked applicants to respond to these. Write the proposal body exactly as the ' +
    'rules above say and do not answer these questions inside it. Then, overriding the length, ' +
    'no-list, and closing-question rules: add a blank line, the heading "Your questions:", and a ' +
    'numbered answer for each question that is a real request to applicants, in the client\'s ' +
    'order, one to three sentences each. Skip lead-in lines like "please answer these" and ' +
    'anything rhetorical. If an instruction is about the proposal itself (such as a word to start ' +
    'with), follow it in the proposal instead of answering it. Do not repeat a client question as ' +
    'your own closing question. When a question asks about work the freelancer has done ("describe ' +
    'a tenant you reorganized", "tell me about a flow you built"), answer only from a project or ' +
    'background item given above and name it. If none of them is that kind of work, say plainly ' +
    'that the freelancer has not done that exact work yet, then name the closest real project. ' +
    'Never describe a past project, client, or result that is not given above.';
}

// Room for the "Your questions:" block on top of the letter itself, so the
// last answer isn't cut off mid-sentence.
var TOKENS_PER_DESCRIPTION_QUESTION_ = 110;

function questionTokenAllowance_(description) {
  return extractDescriptionQuestions(description).length * TOKENS_PER_DESCRIPTION_QUESTION_;
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
    GROUNDING_RULES_ +
    VOICE_RULES_;

  var payload = {
    model: 'gpt-4.1-mini',
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
      ? flagBannedPhrases_(flagUngroundedNumbers_(data.choices[0].message.content.trim(), sources))
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

  // Once the named project's own description is found, the rest of the
  // portfolio stays out of the prompt. With all of it in, the job 151 draft
  // (2026-10-08) credited the Support-Team tracker with the Looker Studio
  // project's "performance scoring engine". The full portfolio is only the
  // fallback when the name isn't found in it.
  var namedBlock = namedProjectBlock_(cred, portfolioAll);

  var prompt =
    'You are writing an Upwork proposal for a freelancer named ' + (freelancerName || 'the freelancer') + '. ' +
    'FREELANCER PROFILE: ' + journeyContext + ' ' +
    (portfolioAll && !namedBlock ? 'FULL PORTFOLIO (for context only): ' + portfolioAll + '. ' : '') +
    'OVERALL TONE GUIDANCE: ' + (proposalTone || 'Direct') + '. ' +
    'STRICT RULES -- violating any rule makes the proposal unusable: ' +
    '1. Between ' + lengthSpec.words + ' words total (roughly ' + lengthSpec.chars + '). ' +
    '2. Do NOT start with Hi, Hello, or any greeting. ' +
    '3. Do NOT use bullet points or numbered lists. ' +
    '4. Do NOT list skills or tools generically. ' +
    '5. First sentence MUST reference a specific detail from the job description -- not a generic observation. ' +
    'Write the proposal in full; do not begin mid-sentence and capitalize the first word. ' +
    '6. You MUST reference this exact portfolio project by name in the proposal: ' + cred + ' -- do not substitute or add a different project. ' +
    namedBlock + TASK_MATCH_RULE_ + ' ' +
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
    GROUNDING_RULES_ +
    VOICE_RULES_;

  // 0.3, down from 0.5: at 0.5 the job 151 draft (2026-10-08) turned the
  // Zendesk client into "a Freshdesk client" and merged two separate MFA
  // breaks into one event, both from facts given correctly in the prompt.
  // gpt-4.1-mini, up from gpt-4o-mini (2026-10-08): with every fact correct
  // in the prompt, 4o-mini still invented how a bug was found ("by testing
  // the connection"). The 151/153 redrafts on 4.1-mini had no invented
  // details, so all three generators in this file now use it.
  var payload = {
    model: 'gpt-4.1-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: lengthSpec.maxTokens + questionTokenAllowance_(description),
    temperature: 0.3
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
      ? flagBannedPhrases_(flagUngroundedNumbers_(data.choices[0].message.content.trim(), sources))
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
      ' -- directly to what this client needs. Reference that project and no other. Be concrete.' +
      namedProjectBlock_(forcedProject, portfolioContext) + TASK_MATCH_RULE_ + ' '
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
    GROUNDING_RULES_ +
    VOICE_RULES_;

  var payload = {
    model: 'gpt-4.1-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: lengthSpec.maxTokens + questionTokenAllowance_(description),
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
      ? flagBannedPhrases_(flagUngroundedNumbers_(data.choices[0].message.content.trim(), sources))
      : 'No response returned.';

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}
