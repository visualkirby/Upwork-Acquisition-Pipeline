/**
 * ============================================================
 * FreelanceFlow Library -- Keyword Strategy Builder
 * AI-generated Primary/Secondary/Negative keyword strategy from
 * the customer's niche + portfolio, mirroring PipelineIQ's
 * generate_keywords() service. The thin client's
 * GENERATE_KEYWORD_STRATEGY() (18_Keyword_Strategy.gs) is the
 * sheet I/O wrapper that calls this and writes the results into
 * Keyword_Search_List and Keyword_Strategy.
 * ============================================================
 */
function generateKeywordStrategy(niche, portfolioSummary, apiKey) {
  if (!apiKey) return { ok: false, message: 'API key not found. Run Setup API Key first.' };

  var prompt =
    'A freelancer wants a keyword strategy for Upwork job searches.\n\n' +
    'Niche: ' + (niche || 'not specified') + '\n' +
    'Portfolio summary: ' + (portfolioSummary || 'not specified') + '\n\n' +
    'Generate a keyword strategy with exactly these fields:\n' +
    '{\n' +
    '  "primary": [string] (5-8 high-intent Upwork search phrases to use regularly),\n' +
    '  "secondary": [string] (5-8 broader search phrases to rotate in),\n' +
    '  "negative": [string] (3-5 terms that indicate bad-fit jobs to avoid)\n' +
    '}\n\n' +
    'Return only valid JSON. No explanation. No markdown.';

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 500,
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
      ok:        true,
      primary:   Array.isArray(parsed.primary)   ? parsed.primary   : [],
      secondary: Array.isArray(parsed.secondary) ? parsed.secondary : [],
      negative:  Array.isArray(parsed.negative)  ? parsed.negative  : []
    };
  } catch (e) {
    return { ok: false, message: 'Could not parse AI response. Try again.' };
  }
}


// generateKeywordVariations mirrors generateKeywordStrategy above, scoped to
// a single already-proven keyword instead of the freelancer's whole niche --
// the AI backing SCALE_KEYWORDS() (18_Keyword_Strategy.gs), called once per
// keyword flagged "Scale" in Keyword_Strategy.
function generateKeywordVariations(keyword, niche, apiKey) {
  if (!apiKey) return { ok: false, message: 'API key not found. Run Setup API Key first.' };

  var prompt =
    'A freelancer found an Upwork search keyword that performs well and wants to scale it up ' +
    'with closely related search phrases.\n\n' +
    'High-performing keyword: ' + keyword + '\n' +
    'Niche: ' + (niche || 'not specified') + '\n\n' +
    'Generate 5-8 closely related Upwork search phrases that would surface similar jobs to this ' +
    'keyword -- synonyms, adjacent tool names, and alternate phrasing a client might use for the ' +
    'same type of work.\n\n' +
    'Return only valid JSON with exactly this field:\n' +
    '{\n' +
    '  "variations": [string]\n' +
    '}\n\n' +
    'Return only valid JSON. No explanation. No markdown.';

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 500,
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
      variations: Array.isArray(parsed.variations) ? parsed.variations : []
    };
  } catch (e) {
    return { ok: false, message: 'Could not parse AI response. Try again.' };
  }
}

// buildKeywordStrategyFormula is pure -- headers/row in, formula-string
// out, no SpreadsheetApp access. GENERATE_KEYWORD_STRATEGY() (thin
// client, 18_Keyword_Strategy.gs) calls this once per new row it writes,
// since each row's Recommended_Action formula must reference that row's
// own Actual_Count/Target_Count cells. colLetter_ is the shared private
// helper already defined in Lib_WizardFormulas.gs.
function buildKeywordStrategyFormula(headers, row) {
  var kwIdx     = headers.indexOf('Keyword');
  var actionIdx = headers.indexOf('Recommended_Action');
  var actualIdx = headers.indexOf('Actual_Count');
  var targetIdx = headers.indexOf('Target_Count');

  var result = {};
  if (kwIdx < 0 || actionIdx < 0 || actualIdx < 0 || targetIdx < 0) return result;

  var kwL     = colLetter_(kwIdx + 1);
  var actualL = colLetter_(actualIdx + 1);
  var targetL = colLetter_(targetIdx + 1);

  result.recommendedActionCol = actionIdx + 1;
  result.recommendedActionFormula =
    '=IF(' + kwL + row + '="","",IF(' + targetL + row + '>=1,' +
    'IF(' + actualL + row + '>=' + targetL + row + ',"Complete","Keep Testing"),' +
    '"Avoid"))';

  return result;
}
