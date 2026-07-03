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
 * Note: template lookup (reading the Proposal_Templates sheet)
 * stays in the thin client's lookupProposalTemplate_ -- it's
 * sheet I/O, not proprietary logic.
 * ============================================================
 */

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
    context += '(No portfolio projects configured. Add them in the Settings sheet under Portfolio_1 through Portfolio_5.)';
  }

  context += ' Tools used: ' + tools + '.';
  if (bg) context += ' Background: ' + bg;

  return context;
}


function generateAIProposal(jobTitle, description, toolDetected, jobType, template,
                             apiKey, journeyContext, portfolioAll, proposalTone, freelancerName) {
  if (!apiKey) return 'API key not set. Run System Tools > Setup API Key first.';

  var angle          = (template && template.angle)          || '';
  var credentialHint = (template && template.credentialHint) || '';
  var tone           = (template && template.tone)           || '';
  var ctaStyle       = (template && template.ctaStyle)       || '';
  var exampleOutput  = (template && template.exampleOutput)  || '';

  var portfolioParts = portfolioAll
    ? portfolioAll.split(';').map(function(p) { return p.trim(); }).filter(function(p) { return p; })
    : [];
  var cred = credentialHint || (portfolioParts.length > 0 ? portfolioParts[0] : 'a portfolio project');

  var prompt =
    'You are writing an Upwork proposal for a freelancer named ' + (freelancerName || 'the freelancer') + '. ' +
    'FREELANCER PROFILE: ' + journeyContext + ' ' +
    (portfolioAll ? 'FULL PORTFOLIO (for context only): ' + portfolioAll + '. ' : '') +
    'OVERALL TONE GUIDANCE: ' + (proposalTone || 'Direct') + '. ' +
    'STRICT RULES -- violating any rule makes the proposal unusable: ' +
    '1. Under 100 words total. ' +
    '2. Do NOT start with Hi, Hello, or any greeting. ' +
    '3. Do NOT use bullet points or numbered lists. ' +
    '4. Do NOT list skills or tools generically. ' +
    '5. First sentence MUST reference a specific detail from the job description -- not a generic observation. ' +
    '6. You MUST reference this exact portfolio project by name in the proposal: ' + cred + ' -- do not substitute a different project. ' +
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
    max_tokens: 180,
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
                             apiKey, portfolioContext, freelancerName) {
  if (!apiKey) return 'API key not set. Run System Tools > Setup API Key.';

  var name = freelancerName || 'the freelancer';

  var competitionNote = proposalCount > 30
    ? 'This job already has ' + proposalCount + ' proposals so the opener must immediately stand out.'
    : proposalCount > 15
    ? 'This job has ' + proposalCount + ' proposals -- be specific and direct.'
    : 'This job has few proposals -- a clear, confident proposal will stand out easily.';

  var prompt =
    'You are writing an Upwork proposal for ' + name + ', a freelancer. ' +
    'Write a complete proposal in exactly 3 short paragraphs, under 120 words total. ' +
    'Rules you must follow: ' +
    'Do NOT start with Hi or the client\'s name. ' +
    'Do NOT open with I or My or a statement about the freelancer. ' +
    'Open with something specific from the job description that shows you read it carefully -- reference the actual problem or tool or industry. ' +
    'Second paragraph: connect one of the freelancer\'s portfolio projects or specific experience directly to what this client needs. Be concrete, not vague. ' +
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
    max_tokens: 200,
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
