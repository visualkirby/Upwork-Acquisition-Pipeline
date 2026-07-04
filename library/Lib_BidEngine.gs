/**
 * ============================================================
 * FreelanceFlow Library -- Bid Engine
 * AI-powered bid strategy advisor -- the most concentrated
 * proprietary decision logic in the codebase (8 ordered
 * thresholds baked into the prompt).
 * ============================================================
 */
function getBidRecommendation(jobTitle, baseConnects, proposalCount,
                               totalScore, bid1, bid2, bid3, bid4,
                               apiKey, journeyContext, noBoostMaxProp, noBoostMinScore) {
  if (!apiKey) {
    return 'API key not set. Run System Tools > Setup API Key first.';
  }

  var boost1stMaxProp  = 10;
  var boost1stMinScore = 0.75;
  var boost2ndMaxProp  = 15;
  var boost2ndMinScore = noBoostMinScore;
  var boost3rdMaxProp  = 20;
  var boost3rdMinScore = noBoostMinScore;
  var boost4thMaxProp  = 25;
  var boost4thMinScore = noBoostMinScore;

  var prompt =
    'You are a bid strategy advisor for an Upwork freelancer. Upwork shows the top 4 competing bids. ' +
    'Give a structured decision using EXACTLY this format with no extra text: ' +
    'DECISION: [NO BOOST / BOOST TO 4TH / BOOST TO 3RD / BOOST TO 2ND / BOOST TO 1ST] ' +
    'REASON: [one sentence max] ' +
    'FREELANCER CONTEXT: ' + journeyContext + ' ' +
    'JOB DATA: ' +
    'Title: ' + jobTitle + '. ' +
    'Base connects to submit: ' + baseConnects + '. ' +
    'Current proposal count: ' + proposalCount + ' (if text like Less than 5 treat as 3). ' +
    'Job quality score: ' + totalScore + ' out of 1. ' +
    'Current top bids -- 1st: ' + bid1 + ' connects, 2nd: ' + bid2 + ' connects, 3rd: ' + bid3 + ' connects, 4th: ' + bid4 + ' connects. ' +
    'DECISION RULES (apply in order): ' +
    'If proposal count is above ' + noBoostMaxProp + ': DECISION must be NO BOOST. ' +
    'If score is below ' + noBoostMinScore + ': DECISION must be NO BOOST. ' +
    'If proposal count is under ' + boost1stMaxProp + ' AND score is above ' + boost1stMinScore + ' AND gap between 1st and 2nd bid is 5 connects or fewer: BOOST TO 1ST. ' +
    'If proposal count is under ' + boost2ndMaxProp + ' AND score is above ' + boost2ndMinScore + ': BOOST TO 2ND. ' +
    'If proposal count is under ' + boost3rdMaxProp + ' AND score is above ' + boost3rdMinScore + ': BOOST TO 3RD. ' +
    'If proposal count is under ' + boost4thMaxProp + ' AND score is above ' + boost4thMinScore + ': BOOST TO 4TH. ' +
    'Otherwise NO BOOST. ' +
    'Be direct. No preamble. Use the exact format above -- no BID line, that number is computed separately.';

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: 60,
    temperature: 0.1
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

    var text = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : '';
    if (!text) return 'No response returned.';

    return applyBidMath_(text, baseConnects, bid1, bid2, bid3, bid4);

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}

/**
 * Never trust the AI to do the bidding arithmetic -- 1 connect above the
 * target position's current bid is what actually outranks them on Upwork.
 * The AI only picks the strategic DECISION (which position to chase, if
 * any); this fills in the exact BID number deterministically from the real
 * bid1-4 values instead of letting the model guess a number.
 */
function applyBidMath_(aiText, baseConnects, bid1, bid2, bid3, bid4) {
  var decisionMatch = aiText.match(/DECISION:\s*(NO BOOST|BOOST TO (?:1ST|2ND|3RD|4TH))/i);
  var reasonMatch    = aiText.match(/REASON:\s*([\s\S]*)/i);

  if (!decisionMatch) return aiText;

  var decision = decisionMatch[1].toUpperCase();
  var reason   = reasonMatch ? reasonMatch[1].trim() : '';

  var targets = { '1ST': bid1, '2ND': bid2, '3RD': bid3, '4TH': bid4 };
  var bidLine;

  if (decision === 'NO BOOST') {
    bidLine = baseConnects + ' (no boost)';
  } else {
    var place  = decision.replace('BOOST TO ', '');
    var target = Number(targets[place]);
    bidLine = isNaN(target)
      ? 'unable to compute -- ' + place + ' place bid not entered'
      : (target + 1) + ' connects (1 more than current ' + place + ' place bid of ' + target + ')';
  }

  return 'DECISION: ' + decision + '\nBID: ' + bidLine + (reason ? '\nREASON: ' + reason : '');
}
