/**
 * ============================================================
 * FreelanceFlow Library -- Bid Analysis
 * Thresholds (based on S022-S033 data, adjustable here):
 *   EXTREME_THRESHOLD  -- Bid_1st >= this -> extreme competition
 *   HIGH_THRESHOLD     -- Bid_1st >= this -> high competition
 *   MEDIUM_THRESHOLD   -- Bid_1st >= this -> medium competition
 *   BOOST_HEAVY_RATIO  -- Boost/Connects_Required >= this -> heavy
 *   BOOST_EXCESSIVE_RATIO -- Boost/Connects_Required >= this -> excessive
 *
 * bucketBidPatterns and fmtBucketLine are both called directly by
 * the thin client's ANALYZE_BID_PATTERNS(), so both are public
 * (no trailing underscore).
 * ============================================================
 */
var EXTREME_THRESHOLD_      = 50;
var HIGH_THRESHOLD_         = 30;
var MEDIUM_THRESHOLD_       = 10;
var BOOST_HEAVY_RATIO_      = 0.75;
var BOOST_EXCESSIVE_RATIO_  = 1.0;

function bucketBidPatterns(rows) {
  var buckets = {
    low:     { label: "Low     (0-9)",   count: 0, sent: 0, skip: 0, other: 0 },
    medium:  { label: "Medium  (10-29)", count: 0, sent: 0, skip: 0, other: 0 },
    high:    { label: "High    (30-49)", count: 0, sent: 0, skip: 0, other: 0 },
    extreme: { label: "Extreme (50+)",   count: 0, sent: 0, skip: 0, other: 0 }
  };

  var bidRowCount      = 0;
  var extremeSentFlags = [];
  var boostRows        = [];

  rows.forEach(function (row) {
    var b1     = Number(row.bid1) || 0;
    var b2     = Number(row.bid2) || 0;
    var b3     = Number(row.bid3) || 0;
    var boost  = Number(row.boost) || 0;
    var conn   = Number(row.connectsRequired) || 0;
    var total  = Number(row.totalConnectsSpent) || 0;
    var status = String(row.status || "").trim();
    var title  = String(row.title || "").trim();

    var hasBid   = (b1 > 0 || b2 > 0 || b3 > 0);
    var hasBoost = (boost > 0);

    if (!hasBid && !hasBoost) return;

    bidRowCount++;

    var bucket;
    if      (b1 >= EXTREME_THRESHOLD_) bucket = buckets.extreme;
    else if (b1 >= HIGH_THRESHOLD_)    bucket = buckets.high;
    else if (b1 >= MEDIUM_THRESHOLD_)  bucket = buckets.medium;
    else                               bucket = buckets.low;

    bucket.count++;
    if      (status === "Sent") bucket.sent++;
    else if (status === "Skip") bucket.skip++;
    else                        bucket.other++;

    if (b1 >= EXTREME_THRESHOLD_ && status === "Sent") {
      extremeSentFlags.push({
        title:  title.substring(0, 44),
        b1:     b1,
        conn:   conn,
        boost:  boost,
        total:  total
      });
    }

    if (hasBoost) {
      var ratio = conn > 0 ? boost / conn : 0;
      var level = ratio >= BOOST_EXCESSIVE_RATIO_ ? "EXCESSIVE"
                : ratio >= BOOST_HEAVY_RATIO_      ? "HEAVY"
                : "";
      boostRows.push({
        title:  title.substring(0, 44),
        boost:  boost,
        conn:   conn,
        total:  total,
        ratio:  ratio,
        level:  level,
        status: status,
        b1:     b1
      });
    }
  });

  var totalSent = buckets.low.sent + buckets.medium.sent + buckets.high.sent + buckets.extreme.sent;
  var totalSkip = buckets.low.skip + buckets.medium.skip + buckets.high.skip + buckets.extreme.skip;
  var sendAccuracy = totalSent > 0
    ? Math.round(((buckets.low.sent + buckets.medium.sent) / totalSent) * 100) + "% of sends were low/medium competition"
    : "N/A";

  return {
    buckets:          buckets,
    bidRowCount:      bidRowCount,
    extremeSentFlags: extremeSentFlags,
    boostRows:        boostRows,
    totalSent:        totalSent,
    totalSkip:        totalSkip,
    sendAccuracy:     sendAccuracy
  };
}


function fmtBucketLine(b) {
  if (b.count === 0) return b.label + ": 0 jobs";
  var sendRate = Math.round((b.sent / b.count) * 100);
  return b.label + ": " + b.count + " jobs | " +
    "Sent=" + b.sent + " Skip=" + b.skip +
    (b.other > 0 ? " Other=" + b.other : "") +
    " | Send rate " + sendRate + "%";
}
