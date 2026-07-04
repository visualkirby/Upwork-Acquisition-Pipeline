/**
 * ============================================================
 * 3. HELPERS
 * ============================================================
 */

function normalizeHeader_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    map[normalizeHeader_(h)] = i + 1;
  });
  return map;
}

function getCol_(headerMap, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var key = normalizeHeader_(possibleNames[i]);
    if (headerMap[key]) return headerMap[key];
  }
  return null;
}

function getCellValue_(sheet, row, headerMap, possibleNames) {
  var col = getCol_(headerMap, possibleNames);
  if (!col) return "";
  return sheet.getRange(row, col).getValue();
}

function setCellValue_(sheet, row, headerMap, possibleNames, value) {
  var col = getCol_(headerMap, possibleNames);
  if (!col) return;
  sheet.getRange(row, col).setValue(value);
}

function cleanJobText_(text) {
  if (text === null || text === undefined) return "";
  return String(text)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n+/g, " ")
    .replace(/\t+/g, " ")
    .replace(/ /g, " ")
    .replace(/[•▪■●◦]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function getLastRealRow_(sheet) {
  var maxRows  = sheet.getMaxRows();
  var lastReal = 1;
  for (var col = 1; col <= 3; col++) {
    var vals = sheet.getRange(2, col, maxRows - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]).trim() !== "") {
        var rowNum = i + 2;
        if (rowNum > lastReal) lastReal = rowNum;
        break;
      }
    }
  }
  return lastReal;
}

function findFirstEmptyRowByColumn_(sheet, col) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values  = sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) return i + 2;
  }
  return lastRow + 1;
}

function formatDuration_(ms) {
  var totalSeconds = Math.floor(ms / 1000);
  var hours        = Math.floor(totalSeconds / 3600);
  var minutes      = Math.floor((totalSeconds % 3600) / 60);
  var seconds      = totalSeconds % 60;
  if (hours > 0) {
    return hours + "h " + minutes + "m " + seconds + "s";
  }
  return minutes + "m " + seconds + "s";
}

function formatTime_(date) {
  var h    = date.getHours();
  var m    = date.getMinutes();
  var s    = date.getSeconds();
  var ampm = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return h + ":" + (m < 10 ? "0" + m : m) + ":" + (s < 10 ? "0" + s : s) + " " + ampm;
}

// Connects_Helper is metric-row structured (Metric name in column A, Value in
// column B) rather than header-column structured like the pipeline sheets, so
// updating one of its running totals from another sheet's edit trigger means
// scanning column A for the matching label rather than a header lookup.
function incrementConnectsHelperMetric_(ss, metricName, amount) {
  if (!amount) return;
  var sheet = ss.getSheetByName('Connects_Helper');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === metricName) {
      var current = Number(data[i][1]) || 0;
      sheet.getRange(i + 2, 2).setValue(current + amount);
      return;
    }
  }
}

// Looks up an Hourly contract's rate from Contract_Tracker by Discovery_ID --
// used by the Hourly_Log handler to compute Amount from Hours_Logged without
// relying on a live formula (see 14_Edit_Trigger.gs's HOURLY_LOG block for why).
function getContractHourlyRate_(ss, discoveryId) {
  var sheet = ss.getSheetByName('Contract_Tracker');
  if (!sheet || sheet.getLastRow() < 2) return 0;

  var map    = getHeaderMap_(sheet);
  var idCol  = getCol_(map, ['Discovery_ID']);
  var rateCol = getCol_(map, ['Hourly_Rate']);
  if (!idCol || !rateCol) return 0;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idCol - 1]) === String(discoveryId)) {
      return Number(data[i][rateCol - 1]) || 0;
    }
  }
  return 0;
}

// First-session walkthrough plumbing. Checks+marks a one-time script-property
// flag (same FF_-prefixed convention as FF_SETUP_COMPLETE) and returns whether
// it had ALREADY been seen before this call -- so callers can gate on "was
// this the first time" with a single call, no separate check-then-set.
function showWalkthroughSeen_(key) {
  var prop = PropertiesService.getScriptProperties();
  var alreadySeen = prop.getProperty(key) === 'true';
  prop.setProperty(key, 'true');
  return alreadySeen;
}

// Alert-based walkthrough popup for handleEdit hook points -- fires once,
// the first time the given stage is ever reached, then never again.
function showWalkthroughOnce_(key, title, message) {
  if (showWalkthroughSeen_(key)) return;
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}
