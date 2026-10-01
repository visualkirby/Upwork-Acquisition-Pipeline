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
  // The freelancer enters 0 for all four bids when the Apply page has no
  // boost table, so the job can't be boosted at all. Answered without the
  // AI, which would otherwise recommend boosting a strong job for 1 Connect.
  // Some zeros but not all is a real table with empty spots (fewer than 4
  // people boosted), so that still goes through the rules below.
  var bids = [bid1, bid2, bid3, bid4];
  var allZero = bids.every(function (b) {
    return b !== '' && b !== null && b !== undefined && Number(b) === 0;
  });
  if (allZero) {
    return 'DECISION: NO BOOST\nBID: ' + baseConnects + ' (no boost)\n' +
      'REASON: This job has no boost option on Upwork (all four bids are 0).';
  }

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

  // Context and job data come first, clearly fenced off from the output
  // spec at the bottom. An earlier version ran all of this into one "use
  // EXACTLY this format: DECISION ... REASON ... FREELANCER CONTEXT: <text>
  // JOB DATA: ..." string, and the model read the FREELANCER CONTEXT and
  // JOB DATA labels as part of the format and echoed them back into the
  // answer -- where applyBidMath_'s greedy REASON capture then stored the
  // echo (itself truncated by max_tokens) as the reason text.
  var prompt =
    'You are a bid strategy advisor for an Upwork freelancer. Upwork shows the top 4 competing bids.\n\n' +
    'FREELANCER CONTEXT: ' + journeyContext + '\n\n' +
    'JOB DATA:\n' +
    'Title: ' + jobTitle + '\n' +
    'Base connects to submit: ' + baseConnects + '\n' +
    'Current proposal count: ' + proposalCount + ' (if text like "Less than 5", treat as 3)\n' +
    'Job quality score: ' + totalScore + ' out of 1\n' +
    'Current top bids -- 1st: ' + bid1 + ' connects, 2nd: ' + bid2 + ' connects, 3rd: ' + bid3 + ' connects, 4th: ' + bid4 + ' connects\n\n' +
    'DECISION RULES (apply in order, stop at the first that matches):\n' +
    '1. If proposal count is above ' + noBoostMaxProp + ': NO BOOST.\n' +
    '2. If score is below ' + noBoostMinScore + ': NO BOOST.\n' +
    '3. If proposal count is under ' + boost1stMaxProp + ' AND score is above ' + boost1stMinScore +
      ' AND the gap between the 1st and 2nd bid is 5 connects or fewer: BOOST TO 1ST.\n' +
    '4. If proposal count is under ' + boost2ndMaxProp + ' AND score is above ' + boost2ndMinScore + ': BOOST TO 2ND.\n' +
    '5. If proposal count is under ' + boost3rdMaxProp + ' AND score is above ' + boost3rdMinScore + ': BOOST TO 3RD.\n' +
    '6. If proposal count is under ' + boost4thMaxProp + ' AND score is above ' + boost4thMinScore + ': BOOST TO 4TH.\n' +
    '7. Otherwise: NO BOOST.\n\n' +
    'Return ONLY these two lines, with nothing before, after, or in between them:\n' +
    'DECISION: <one of: NO BOOST | BOOST TO 4TH | BOOST TO 3RD | BOOST TO 2ND | BOOST TO 1ST>\n' +
    'REASON: <one sentence, 20 words or fewer, no line breaks>\n' +
    'Do not add a BID line -- that number is computed separately. Do not restate the job data or this context.';

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
  if (!decisionMatch) return aiText;

  var decision = decisionMatch[1].toUpperCase();

  // Take only the REASON sentence: the first line after "REASON:", with any
  // echoed section label (and everything after it) stripped, then a hard
  // length cap as a last-resort guard. Was /REASON:\s*([\s\S]*)/i -- a
  // greedy grab that swallowed everything to the end of the response,
  // including an echoed, token-truncated copy of the context block.
  var reason = '';
  var reasonMatch = aiText.match(/REASON:\s*([\s\S]*)/i);
  if (reasonMatch) {
    reason = reasonMatch[1]
      .split('\n')[0]
      .replace(/\s*\b(FREELANCER CONTEXT|JOB DATA|DECISION RULES|DECISION)\b[\s\S]*$/i, '')
      .trim();
    if (reason.length > 300) reason = reason.substring(0, 300).trim();
  }

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
