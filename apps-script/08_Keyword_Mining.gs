/**
 * ============================================================
 * 8. KEYWORD MINING
 * Mines Job_Discovery titles and descriptions to surface new
 * Tool + Business_Area + Intent triplets above MIN_FREQUENCY.
 * Purges "Drop" keywords from Keyword_Search_List first.
 * ============================================================
 */
function MINE_KEYWORDS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    "Mine Keywords",
    "How many new keyword combinations do you want to add?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var requestedCount = parseInt(response.getResponseText(), 10);
  if (isNaN(requestedCount) || requestedCount < 1) {
    ui.alert("Please enter a whole number greater than 0.");
    return;
  }

  var discoverySheet  = ss.getSheetByName("Job_Discovery");
  var searchListSheet = ss.getSheetByName("Keyword_Search_List");
  var strategySheet   = ss.getSheetByName("Keyword_Strategy");

  if (!discoverySheet || !searchListSheet || !strategySheet) {
    ui.alert("Missing sheet. Confirm Job_Discovery, Keyword_Search_List, and Keyword_Strategy all exist.");
    return;
  }

  var stratLastRow = strategySheet.getLastRow();
  var stratMap     = getHeaderMap_(strategySheet);
  var dropSet      = {};

  if (stratLastRow > 1) {
    var stratData = strategySheet
      .getRange(2, 1, stratLastRow - 1, strategySheet.getLastColumn())
      .getValues();

    var sKeywordCol = getCol_(stratMap, ["Keyword"]);
    var sActionCol  = getCol_(stratMap, ["Recommended_Action"]);
    var sActualCol  = getCol_(stratMap, ["Actual_Count"]);
    var sTargetCol  = getCol_(stratMap, ["Target_Count"]);

    for (var i = 0; i < stratData.length; i++) {
      var kw     = sKeywordCol ? String(stratData[i][sKeywordCol - 1]).trim().toLowerCase() : "";
      var action = sActionCol  ? String(stratData[i][sActionCol  - 1]).trim() : "";
      var actual = sActualCol  ? Number(stratData[i][sActualCol  - 1]) : 0;
      var target = sTargetCol  ? Number(stratData[i][sTargetCol  - 1]) : 0;

      if (action === "Drop" && actual >= target && kw !== "") {
        dropSet[kw] = true;
      }
    }
  }

  var slMap      = getHeaderMap_(searchListSheet);
  var slQueryCol = getCol_(slMap, ["Search_Query"]);
  var slLastRow  = searchListSheet.getLastRow();
  var purgedCount = 0;

  if (slLastRow > 1 && slQueryCol && Object.keys(dropSet).length > 0) {
    for (var r = slLastRow; r >= 2; r--) {
      var cellQuery = String(
        searchListSheet.getRange(r, slQueryCol).getValue()
      ).trim().toLowerCase();
      if (dropSet[cellQuery]) {
        searchListSheet.deleteRow(r);
        purgedCount++;
      }
    }
  }

  slLastRow = searchListSheet.getLastRow();
  var existingKeys = {};
  var slToolCol    = getCol_(slMap, ["Tool"]);
  var slBizCol     = getCol_(slMap, ["Business_Area"]);
  var slIntCol     = getCol_(slMap, ["Intent"]);

  if (slLastRow > 1 && slToolCol && slBizCol && slIntCol) {
    var slData = searchListSheet
      .getRange(2, 1, slLastRow - 1, searchListSheet.getLastColumn())
      .getValues();

    for (var i = 0; i < slData.length; i++) {
      var eTool = String(slData[i][slToolCol - 1]).trim().toLowerCase();
      var eBiz  = String(slData[i][slBizCol  - 1]).trim().toLowerCase();
      var eInt  = String(slData[i][slIntCol  - 1]).trim().toLowerCase();

      if (eTool !== "" && eBiz !== "" && eInt !== "") {
        existingKeys[eTool + "|" + eBiz + "|" + eInt] = true;
      }
    }
  }

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

  var discLastRow = discoverySheet.getLastRow();
  if (discLastRow < 2) {
    ui.alert("Job_Discovery has no data rows to mine.");
    return;
  }

  var discMap      = getHeaderMap_(discoverySheet);
  var discTitleCol = getCol_(discMap, ["Job_Title"]);
  var discDescCol  = getCol_(discMap, ["Description"]);
  var discData     = discoverySheet
    .getRange(2, 1, discLastRow - 1, discoverySheet.getLastColumn())
    .getValues();

  var tripletCounts = {};

  for (var r = 0; r < discData.length; r++) {
    var title    = discTitleCol ? String(discData[r][discTitleCol - 1]).toLowerCase() : "";
    var desc     = discDescCol  ? String(discData[r][discDescCol  - 1]).toLowerCase() : "";
    var combined = title + " " + desc;

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
  }

  var MIN_FREQUENCY = 2;
  var candidates    = [];

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

  if (candidates.length === 0) {
    ui.alert(
      "No new combinations found above the minimum frequency of " + MIN_FREQUENCY + ".\n" +
      "All qualifying combinations are already in Keyword_Search_List."
    );
    return;
  }

  var toWrite       = candidates.slice(0, requestedCount);
  var writeStartRow = getLastRealRow_(searchListSheet) + 1;
  var outputData    = toWrite.map(function (row) {
    return [row.tool, row.biz, row.intent];
  });

  searchListSheet
    .getRange(writeStartRow, 1, outputData.length, 3)
    .setValues(outputData);

  var summary = "Done.\n\n" +
    "✓ " + toWrite.length + " new keyword combinations added to Keyword_Search_List.\n";

  if (toWrite.length < requestedCount) {
    summary += "⚠ Only " + toWrite.length + " qualifying combinations were available " +
               "(you requested " + requestedCount + "). Run again after collecting more jobs.\n";
  }
  if (candidates.length > toWrite.length) {
    summary += "-> " + (candidates.length - toWrite.length) +
               " additional combinations ready for your next run.\n";
  }
  if (purgedCount > 0) {
    summary += "✓ " + purgedCount + " dropped + target-met rows removed before writing.";
  }

  ui.alert(summary);
}
