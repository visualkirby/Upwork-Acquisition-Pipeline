/**
 * ============================================================
 * 9. BID ENGINE
 * colorDuplicateJobLinks: highlights duplicate Job_Link cells
 * getBidRecommendation_: AI-powered bid strategy advisor
 * ============================================================
 */
function colorDuplicateJobLinks() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Job_Discovery');
  if (!sheet) return;

  var map     = getHeaderMap_(sheet);
  var linkCol = getCol_(map, ['Job_Link']);
  var skipCol = getCol_(map, ['Discovery_Action']);
  if (!linkCol) return;

  var lastRow  = getLastRealRow_(sheet);
  var lastCol  = sheet.getLastColumn();
  var startRow = 2;
  if (lastRow < startRow) return;

  var numRows    = lastRow - startRow + 1;
  var linkValues = sheet.getRange(startRow, linkCol, numRows, 1).getValues();

  var urlCount = {};
  linkValues.forEach(function(row) {
    var val = String(row[0]).trim();
    if (val === '' || val === 'undefined') return;
    urlCount[val] = (urlCount[val] || 0) + 1;
  });

  var urlColor   = {};
  var colorIndex = 0;
  var palette    = [
    '#E63946', '#2A9D8F', '#E9C46A', '#4361EE',
    '#F4845F', '#52B788', '#9B5DE5', '#F15BB5',
    '#00B4D8', '#FF6B6B', '#06D6A0', '#FFB703'
  ];

  Object.keys(urlCount).forEach(function(url) {
    if (urlCount[url] > 1) {
      urlColor[url] = palette[colorIndex % palette.length];
      colorIndex++;
    }
  });

  var backgrounds = [];
  for (var i = 0; i < numRows; i++) {
    var val   = String(linkValues[i][0]).trim();
    var color = (val && val !== 'undefined' && urlColor[val]) ? urlColor[val] : null;
    var rowColors = [];
    for (var c = 1; c <= lastCol; c++) {
      rowColors.push(c === skipCol ? null : color);
    }
    backgrounds.push(rowColors);
  }

  sheet.getRange(startRow, 1, numRows, lastCol).setBackgrounds(backgrounds);
}


// Bid recommendation engine moved to the Apps Script Library:
// FFLib.getBidRecommendation(jobTitle, baseConnects, proposalCount,
// totalScore, bid1, bid2, bid3, apiKey, journeyContext, noBoostMaxProp,
// noBoostMinScore). Called from 14_Edit_Trigger.gs.
