/**
 * ============================================================
 * 9. BID ENGINE
 * colorDuplicateJobLinks: highlights duplicate Job_Link cells
 * getBidRecommendation_: AI-powered bid strategy advisor
 * ============================================================
 */
function colorDuplicateJobLinks() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Job_Discovery");
  if (!sheet) return;

  var map     = getHeaderMap_(sheet);
  var linkCol = getCol_(map, ["Job_Link"]);
  var skipCol = getCol_(map, ["Discovery_Action"]);
  if (!linkCol) return;

  var lastRow  = sheet.getLastRow();
  var lastCol  = sheet.getLastColumn();
  var startRow = 2;
  if (lastRow < startRow) return;

  var numRows    = lastRow - startRow + 1;
  var linkValues = sheet.getRange(startRow, linkCol, numRows, 1).getValues();

  var urlCount = {};
  linkValues.forEach(function (row) {
    var val = String(row[0]).trim();
    if (val === "" || val === "undefined") return;
    urlCount[val] = (urlCount[val] || 0) + 1;
  });

  var urlColor   = {};
  var colorIndex = 0;
  var palette    = [
    "#E63946", "#2A9D8F", "#E9C46A", "#4361EE",
    "#F4845F", "#52B788", "#9B5DE5", "#F15BB5",
    "#00B4D8", "#FF6B6B", "#06D6A0", "#FFB703"
  ];

  Object.keys(urlCount).forEach(function (url) {
    if (urlCount[url] > 1) {
      urlColor[url] = palette[colorIndex % palette.length];
      colorIndex++;
    }
  });

  var backgrounds = [];
  for (var i = 0; i < numRows; i++) {
    var val   = String(linkValues[i][0]).trim();
    var color = (val && val !== "undefined" && urlColor[val]) ? urlColor[val] : null;
    var rowColors = [];
    for (var c = 1; c <= lastCol; c++) {
      rowColors.push(c === skipCol ? null : color);
    }
    backgrounds.push(rowColors);
  }

  sheet.getRange(startRow, 1, numRows, lastCol).setBackgrounds(backgrounds);
}


function getBidRecommendation_(jobTitle, baseConnects, proposalCount,
                                totalScore, bid1, bid2, bid3) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    return "API key not set. Run System Tools > Setup API Key first.";
  }

  var prompt =
    "You are a bid strategy advisor for an Upwork freelancer. " +
    "Give a structured bid recommendation using EXACTLY this format with no extra text: " +
    "DECISION: [NO BOOST / BOOST TO 3RD / BOOST TO 2ND / BOOST TO 1ST] " +
    "BID: [exact number of connects to spend, or just base connects if no boost] " +
    "REASON: [one sentence max] " +
    "FREELANCER CONTEXT: " + getJourneyStage_() + " " +
    "JOB DATA: " +
    "Title: " + jobTitle + ". " +
    "Base connects to submit: " + baseConnects + ". " +
    "Current proposal count: " + proposalCount + " (if this is text like Less than 5 treat it as 3). " +
    "Job quality score: " + totalScore + " out of 1. " +
    "Current bids -- 1st: " + bid1 + " connects, 2nd: " + bid2 + " connects, 3rd: " + bid3 + " connects. " +
    "DECISION RULES (apply in order): " +
    "If proposal count is above 25: DECISION must be NO BOOST. " +
    "If score is below 0.65: DECISION must be NO BOOST. " +
    "If proposal count is under 10 AND score is above 0.70: seriously consider boosting. " +
    "BOOST TO 1ST only if proposals under 10 AND score above 0.75 AND gap between 1st and 2nd bid is 5 connects or fewer. " +
    "BOOST TO 2ND if proposals under 15 AND score above 0.70. " +
    "BOOST TO 3RD if proposals under 20 AND score above 0.65. " +
    "Otherwise NO BOOST. " +
    "Be direct. No preamble. Use the exact format above.";

  var payload = {
    model: "gpt-4o-mini",
    messages: [{ role: "user", content: prompt }],
    max_tokens: 80,
    temperature: 0.1
  };

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return "API error: " + data.error.message;

    return data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : "No response returned.";

  } catch (err) {
    return "Request failed: " + err.message;
  }
}
