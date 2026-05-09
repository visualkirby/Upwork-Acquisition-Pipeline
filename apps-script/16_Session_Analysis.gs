/**
 * ============================================================
 * 16. SESSION ANALYSIS (S026 / S027)
 * Reads all rows in Session_Log and produces two aggregate
 * reports: performance by time of day, and by weekday.
 * Works from existing Start_Time / Date columns -- no new
 * columns required in the sheet.
 * ============================================================
 */
function ANALYZE_SESSION_PATTERNS() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var log = ss.getSheetByName("Session_Log");

  if (!log || log.getLastRow() < 2) {
    ui.alert("Session_Log is empty or missing.");
    return;
  }

  var map         = getHeaderMap_(log);
  var dateCol     = getCol_(map, ["Date"]);
  var startCol    = getCol_(map, ["Start_Time"]);
  var durationCol = getCol_(map, ["Duration"]);
  var yieldCol    = getCol_(map, ["Session_Yield"]);
  var propsSentCol= getCol_(map, ["Proposals_Sent"]);
  var connCol     = getCol_(map, ["Connects_Spent"]);
  var satCol      = getCol_(map, ["Saturation_Flag"]);

  if (!dateCol || !startCol || !yieldCol) {
    ui.alert("Session_Log is missing required columns (Date, Start_Time, or Session_Yield).");
    return;
  }

  var lastRow = log.getLastRow();
  var data    = log.getRange(2, 1, lastRow - 1, log.getLastColumn()).getValues();

  // bucket definitions for time of day
  var buckets = ["Morning", "Afternoon", "Evening", "Night"];
  var days    = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  var timeBuckets = {};
  var weekdayMap  = {};
  buckets.forEach(function (b) { timeBuckets[b] = newBucket_(); });
  days.forEach(function (d)    { weekdayMap[d]  = newBucket_(); });

  var totalParsed = 0;

  for (var i = 0; i < data.length; i++) {
    var row       = data[i];
    var dateVal   = dateCol     ? row[dateCol     - 1] : null;
    var startVal  = startCol    ? row[startCol    - 1] : null;
    var durVal    = durationCol ? row[durationCol - 1] : null;
    var yieldVal  = yieldCol    ? Number(row[yieldCol    - 1]) || 0 : 0;
    var propsVal  = propsSentCol? Number(row[propsSentCol- 1]) || 0 : 0;
    var connVal   = connCol     ? Number(row[connCol     - 1]) || 0 : 0;
    var satVal    = satCol      ? String(row[satCol      - 1]).trim() : "";

    if (!dateVal && !startVal) continue;

    var bucket  = parseTimeBucket_(startVal);
    var weekday = parseWeekday_(dateVal);
    var durMs   = parseDurationMs_(durVal);
    var sat     = satVal === "Saturating" ? 1 : 0;

    accumulate_(timeBuckets[bucket], yieldVal, propsVal, connVal, durMs, sat);
    accumulate_(weekdayMap[weekday], yieldVal, propsVal, connVal, durMs, sat);
    totalParsed++;
  }

  if (totalParsed === 0) {
    ui.alert("No parseable rows found in Session_Log.");
    return;
  }

  var report =
    "SESSION ANALYSIS -- " + totalParsed + " sessions\n" +
    "════════════════════════════════\n\n" +
    "BY TIME OF DAY\n" +
    "Morning   (6AM-12PM):  " + fmtBucket_(timeBuckets["Morning"])   + "\n" +
    "Afternoon (12PM-5PM):  " + fmtBucket_(timeBuckets["Afternoon"]) + "\n" +
    "Evening   (5PM-9PM):   " + fmtBucket_(timeBuckets["Evening"])   + "\n" +
    "Night     (9PM+):      " + fmtBucket_(timeBuckets["Night"])      + "\n\n" +
    "BY WEEKDAY\n" +
    fmtWeekday_("Mon", weekdayMap["Monday"])    +
    fmtWeekday_("Tue", weekdayMap["Tuesday"])   +
    fmtWeekday_("Wed", weekdayMap["Wednesday"]) +
    fmtWeekday_("Thu", weekdayMap["Thursday"])  +
    fmtWeekday_("Fri", weekdayMap["Friday"])    +
    fmtWeekday_("Sat", weekdayMap["Saturday"])  +
    fmtWeekday_("Sun", weekdayMap["Sunday"]);

  ui.alert("Session Patterns", report, ui.ButtonSet.OK);
}


// ---- helpers -------------------------------------------------------

function newBucket_() {
  return { count: 0, yield: 0, props: 0, conn: 0, durMs: 0, sat: 0 };
}

function accumulate_(bucket, yld, props, conn, durMs, sat) {
  bucket.count++;
  bucket.yield += yld;
  bucket.props += props;
  bucket.conn  += conn;
  bucket.durMs += durMs;
  bucket.sat   += sat;
}

function fmtBucket_(b) {
  if (b.count === 0) return "0 sessions";
  var avgYield = (b.yield / b.count).toFixed(1);
  var avgProps = (b.props / b.count).toFixed(1);
  var avgConn  = (b.conn  / b.count).toFixed(1);
  var avgDur   = b.durMs > 0 ? " | " + formatDuration_(b.durMs / b.count) + " avg" : "";
  var satRate  = b.sat > 0 ? " | " + b.sat + "/" + b.count + " sat" : "";
  return b.count + " sessions | Yield " + avgYield +
         " | Props " + avgProps + " | Conn " + avgConn + avgDur + satRate;
}

function fmtWeekday_(label, b) {
  if (b.count === 0) return "";
  var avgYield = (b.yield / b.count).toFixed(1);
  var avgProps = (b.props / b.count).toFixed(1);
  var avgConn  = (b.conn  / b.count).toFixed(1);
  return label + ": " +
    pad_(String(b.count) + " sess", 9) +
    " | Yield " + pad_(avgYield, 4) +
    " | Props " + pad_(avgProps, 3) +
    " | Conn " + avgConn + "\n";
}

function pad_(str, len) {
  str = String(str);
  while (str.length < len) str = str + " ";
  return str;
}

function parseTimeBucket_(timeVal) {
  if (!timeVal) return "Unknown";
  var str = String(timeVal).trim();

  // Handle Date objects serialized as strings or actual Date values
  var hour = NaN;
  if (timeVal instanceof Date) {
    hour = timeVal.getHours();
  } else {
    // "h:mm:ss a" format (e.g. "10:31:48 AM" or "5:27:50 PM")
    var match = str.match(/^(\d+):(\d+):\d+\s*(AM|PM)$/i);
    if (match) {
      hour = parseInt(match[1], 10);
      var period = match[3].toUpperCase();
      if (period === "PM" && hour !== 12) hour += 12;
      if (period === "AM" && hour === 12) hour = 0;
    }
  }

  if (isNaN(hour)) return "Unknown";
  if (hour >= 6  && hour < 12) return "Morning";
  if (hour >= 12 && hour < 17) return "Afternoon";
  if (hour >= 17 && hour < 21) return "Evening";
  return "Night";
}

function parseWeekday_(dateVal) {
  if (!dateVal) return "Unknown";
  var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var d;
  if (dateVal instanceof Date) {
    d = dateVal;
  } else {
    d = new Date(String(dateVal));
  }
  if (isNaN(d.getTime())) return "Unknown";
  return days[d.getDay()];
}

function parseDurationMs_(durVal) {
  if (!durVal) return 0;
  var str    = String(durVal);
  var hours  = 0, mins = 0, secs = 0;
  var hMatch = str.match(/(\d+)h/);
  var mMatch = str.match(/(\d+)m/);
  var sMatch = str.match(/(\d+)s/);
  if (hMatch) hours = parseInt(hMatch[1], 10);
  if (mMatch) mins  = parseInt(mMatch[1], 10);
  if (sMatch) secs  = parseInt(sMatch[1], 10);
  return (hours * 3600 + mins * 60 + secs) * 1000;
}
