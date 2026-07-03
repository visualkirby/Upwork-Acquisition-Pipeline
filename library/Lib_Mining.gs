/**
 * ============================================================
 * FreelanceFlow Library -- Keyword Mining
 * Pure taxonomy + mining algorithm -- data in, candidates out.
 * No SpreadsheetApp access; the thin client's MINE_KEYWORDS()
 * handles all sheet I/O and calls this.
 * ============================================================
 */
function mineTriplets(discoveryRows, existingKeys) {
  var TOOLS = [
    "Excel", "Power BI", "Looker Studio", "Tableau", "SQL",
    "Google Sheets", "Python", "BigQuery", "Power Query",
    "Google Analytics", "R Studio", "Snowflake", "dbt"
  ];

  var BUSINESS_AREAS = [
    "Sales", "HR", "Marketing", "Operations", "Finance",
    "Inventory", "Revenue", "Retail", "Healthcare", "Logistics",
    "Ecommerce", "Supply Chain", "Construction", "Real Estate",
    "Hospitality", "Procurement", "Manufacturing"
  ];

  var INTENTS = [
    "Dashboard", "Reporting Dashboard", "Dashboard Developer",
    "Dashboard Build", "Dashboard Creation", "Automation",
    "Data Analysis", "Visualization", "Pipeline", "Integration",
    "KPI Dashboard", "Analytics Dashboard"
  ];

  var MIN_FREQUENCY = 2;
  var tripletCounts = {};

  discoveryRows.forEach(function (row) {
    var combined = (String(row.title || "") + " " + String(row.description || "")).toLowerCase();

    var foundTools = TOOLS.filter(function (t) {
      return combined.indexOf(t.toLowerCase()) !== -1;
    });
    var foundBiz = BUSINESS_AREAS.filter(function (b) {
      return combined.indexOf(b.toLowerCase()) !== -1;
    });
    var foundIntents = INTENTS.filter(function (n) {
      return combined.indexOf(n.toLowerCase()) !== -1;
    });

    if (foundTools.length === 0) foundTools = ["Other"];

    foundTools.forEach(function (tool) {
      foundBiz.forEach(function (biz) {
        foundIntents.forEach(function (intent) {
          var key = tool.toLowerCase() + "|" + biz.toLowerCase() + "|" + intent.toLowerCase();
          tripletCounts[key] = (tripletCounts[key] || 0) + 1;
        });
      });
    });
  });

  var candidates = [];

  Object.keys(tripletCounts).forEach(function (key) {
    if (tripletCounts[key] < MIN_FREQUENCY) return;
    if (existingKeys[key]) return;

    var parts  = key.split("|");
    var tool   = TOOLS.find(function (t) { return t.toLowerCase() === parts[0]; }) || parts[0];
    var biz    = BUSINESS_AREAS.find(function (b) { return b.toLowerCase() === parts[1]; }) || parts[1];
    var intent = INTENTS.find(function (n) { return n.toLowerCase() === parts[2]; }) || parts[2];

    candidates.push({ tool: tool, biz: biz, intent: intent, freq: tripletCounts[key] });
  });

  candidates.sort(function (a, b) { return b.freq - a.freq; });

  return candidates;
}
