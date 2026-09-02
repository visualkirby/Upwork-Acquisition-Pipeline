/**
 * ============================================================
 * 8. KEYWORD MINING
 * Mines Job_Discovery titles and descriptions to surface new
 * Tool + Business_Area + Intent triplets above the Library's
 * minimum mining frequency. Purges "Drop" keywords from
 * Keyword_Search_List first.
 *
 * The taxonomy + mining algorithm (FFLib.mineTriplets) lives in
 * the Apps Script Library. This file is the UI + sheet I/O
 * wrapper that calls it.
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

  var purged = purgeDroppedKeywords_(ss);

  var slMap        = getHeaderMap_(searchListSheet);
  var slLastRow    = searchListSheet.getLastRow();
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

  var discLastRow = getLastRealRow_(discoverySheet);
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

  var discoveryRows = discData.map(function (row) {
    return {
      title:       discTitleCol ? String(row[discTitleCol - 1]) : "",
      description: discDescCol  ? String(row[discDescCol  - 1]) : ""
    };
  });

  var candidates = FFLib.mineTriplets(discoveryRows, existingKeys);

  if (candidates.length === 0) {
    ui.alert(
      "No new combinations found above the minimum mining frequency.\n" +
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
  if (purged.strategyRemoved > 0) {
    summary += "✓ " + purged.strategyRemoved + " dropped keyword(s) removed from Keyword_Strategy and " +
      purged.searchListRemoved + " row(s) removed from Keyword_Search_List before writing.";
  }

  ui.alert(summary);
}
