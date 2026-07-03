/**
 * ============================================================
 * FreelanceFlow Library -- Bid Engine
 * AI-powered bid strategy advisor -- the most concentrated
 * proprietary decision logic in the codebase (8 ordered
 * thresholds baked into the prompt).
 * ============================================================
 */
function getBidRecommendation(jobTitle, baseConnects, proposalCount,
                               totalScore, bid1, bid2, bid3,
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

  var prompt =
    'You are a bid strategy advisor for an Upwork freelancer. ' +
    'Give a structured bid recommendation using EXACTLY this format with no extra text: ' +
    'DECISION: [NO BOOST / BOOST TO 3RD / BOOST TO 2ND / BOOST TO 1ST] ' +
    'BID: [exact number of connects to spend, or just base connects if no boost] ' +
    'REASON: [one sentence max] ' +
    'FREELANCER CONTEXT: ' + journeyContext + ' ' +
    'JOB DATA: ' +
    'Title: ' + jobTitle + '. ' +
    'Base connects to submit: ' + baseConnects + '. ' +
    'Current proposal count: ' + proposalCount + ' (if text like Less than 5 treat as 3). ' +
    'Job quality score: ' + totalScore + ' out of 1. ' +
    'Current bids -- 1st: ' + bid1 + ' connects, 2nd: ' + bid2 + ' connects, 3rd: ' + bid3 + ' connects. ' +
    'DECISION RULES (apply in order): ' +
    'If proposal count is above ' + noBoostMaxProp + ': DECISION must be NO BOOST. ' +
    'If score is below ' + noBoostMinScore + ': DECISION must be NO BOOST. ' +
    'If proposal count is under ' + boost1stMaxProp + ' AND score is above ' + boost1stMinScore + ' AND gap between 1st and 2nd bid is 5 connects or fewer: BOOST TO 1ST. ' +
    'If proposal count is under ' + boost2ndMaxProp + ' AND score is above ' + boost2ndMinScore + ': BOOST TO 2ND. ' +
    'If proposal count is under ' + boost3rdMaxProp + ' AND score is above ' + boost3rdMinScore + ': BOOST TO 3RD. ' +
    'Otherwise NO BOOST. ' +
    'Be direct. No preamble. Use the exact format above.';

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: 80,
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

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : 'No response returned.';

  } catch (err) {
    return 'Request failed: ' + err.message;
  }
}
