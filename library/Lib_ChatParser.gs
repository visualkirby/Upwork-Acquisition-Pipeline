/**
 * ============================================================
 * FreelanceFlow Library -- Chat Transcript Parser
 * Client_Chat_Log's Chat Import sidebar lets the user paste a raw
 * copied Upwork chat thread (alternating Sender / Time / Message
 * lines). parseChatTranscript turns that freeform paste into a
 * JSON array of {sender, time, message} objects -- one per
 * message, in the order given -- for the thin client to write as
 * one row per message. Same JSON-array-return pattern as
 * Lib_WizardAI.gs's generateNicheTemplates: strict "return ONLY
 * valid JSON, don't invent anything" prompt, markdown-fence
 * stripping, array/type validation, {ok, ...}/{ok:false, message}
 * return shape.
 * ============================================================
 */
function parseChatTranscript(rawText, freelancerName, apiKey) {
  if (!apiKey) return { ok: false, message: 'API key not found. Complete Step 2 first.' };
  if (!rawText || !rawText.trim()) return { ok: false, message: 'Paste a chat transcript first.' };

  var prompt =
    'You are parsing a raw copy-pasted Upwork chat transcript into structured data. ' +
    'The freelancer is ' + (freelancerName || 'the freelancer') + '; the other participant is the client. ' +
    'The transcript is a repeating pattern of a sender name, a time, then one or more lines of message text, ' +
    'in chronological order. Return ONLY a valid JSON array, no markdown, no explanation. ' +
    'Each element must have exactly these keys: ' +
    '"sender" (the name exactly as it appears before that message), ' +
    '"time" (the time exactly as it appears, e.g. "2:32 PM"), ' +
    '"message" (the full message text for that turn, trimmed). ' +
    'Preserve every message verbatim -- do not summarize, merge separate messages together, split one message ' +
    'into multiple, invent a sender/time/message that is not in the text, or reorder anything. ' +
    'If a message spans multiple lines before the next sender/time pair appears, keep it as one single message. ' +
    'TRANSCRIPT:\n' + rawText;

  var payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'user', content: prompt }],
    max_tokens: 2000,
    temperature: 0.2
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
    if (data.error) return { ok: false, message: 'API error: ' + data.error.message };

    var content = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : '';

    content = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    var parsed = JSON.parse(content);
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return { ok: false, message: 'No messages could be parsed from that transcript.' };
    }

    return { ok: true, messages: parsed };
  } catch (e) {
    return { ok: false, message: 'Could not parse AI response.' };
  }
}
