/**
 * ============================================================
 * 29. TOOL ALIASES
 * Tool_Detected matches each Settings > Primary_Tools entry by its own name
 * OR any of its aliases ("Data Studio" -> Looker Studio, "Microsoft Flow" ->
 * Power Automate), and always reports the customer's own tool name.
 *
 * Aliases live in a visible Tool_Aliases tab (Tool | Aliases | Source).
 * A tool the Library's built-in map knows (FFLib.builtinToolAliases) is
 * re-read from that map on every rebuild. Anything else gets a single
 * GPT-4o-mini call (FFLib.generateToolAliases) and is cached. Nobody has to
 * touch the tab. Editing a row's Aliases marks it "You" and the edit wins
 * from then on.
 *
 * rebuildToolDetection_ rewrites the Tool_Detected formulas on
 * Job_Discovery and Proposal_Generator. handleEdit runs it whenever
 * Primary_Tools or the Tool_Aliases tab changes, so a new tool takes
 * effect immediately and every existing job re-tags against it.
 * ============================================================
 */
var TOOL_ALIASES_SHEET_   = 'Tool_Aliases';
var TOOL_ALIASES_HEADERS_ = ['Tool', 'Aliases', 'Source'];
var TOOL_ALIAS_SOURCE_AI_FAILED_ = 'Name only (AI unavailable, retries next rebuild)';

function getPrimaryToolsList_() {
  return String(getSettings_()['Primary_Tools'] || '').split(',')
    .map(function (t) { return t.trim(); })
    .filter(function (t) { return t; });
}

// { toolName: [aliases] } for every Primary_Tools entry. Adds a
// Tool_Aliases row for any tool that doesn't have one yet; existing rows
// (including ones the customer edited) are read as-is. A row whose AI call
// failed is retried on the next call.
function getToolAliasMap_(ss) {
  var tools = getPrimaryToolsList_();
  var map   = {};
  if (tools.length === 0) return map;

  var sheet = ss.getSheetByName(TOOL_ALIASES_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(TOOL_ALIASES_SHEET_);
    sheet.getRange(1, 1, 1, TOOL_ALIASES_HEADERS_.length)
      .setValues([TOOL_ALIASES_HEADERS_]).setFontWeight('bold');
  }

  var sMap      = getHeaderMap_(sheet);
  var toolCol   = getCol_(sMap, ['Tool']);
  var aliasCol  = getCol_(sMap, ['Aliases']);
  var sourceCol = getCol_(sMap, ['Source']);
  if (!toolCol || !aliasCol || !sourceCol) return map;

  var rowByKey = {};
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    data.forEach(function (r, i) {
      var key = String(r[toolCol - 1]).trim().toLowerCase();
      if (key) {
        rowByKey[key] = {
          row:     i + 2,
          aliases: String(r[aliasCol - 1]),
          source:  String(r[sourceCol - 1]).trim()
        };
      }
    });
  }

  var apiKey  = PropertiesService.getScriptProperties().getProperty('UPWORK_OPENAI_API_KEY');
  var aiCalls = 0;

  tools.forEach(function (tool) {
    var existing = rowByKey[tool.toLowerCase()];
    var aliases  = FFLib.builtinToolAliases(tool);
    var source   = 'Built-in';

    // "You" and "AI" rows are kept as-is. "Built-in" rows are re-read from
    // the Library every time, so a fix to the built-in map reaches existing
    // copies without anyone editing the tab.
    if (existing && existing.source !== 'Built-in' && existing.source !== TOOL_ALIAS_SOURCE_AI_FAILED_) {
      map[tool] = splitAliases_(existing.aliases);
      return;
    }
    if (existing && existing.source === 'Built-in' && aliases &&
        existing.aliases === aliases.join(', ')) {
      map[tool] = aliases;
      return;
    }

    if (!aliases) {
      var ai = FFLib.generateToolAliases(tool, apiKey);
      aliases = ai.ok ? ai.aliases : [];
      source  = ai.ok ? 'AI' : TOOL_ALIAS_SOURCE_AI_FAILED_;
      aiCalls++;
      if (aiCalls % 5 === 0) Utilities.sleep(1000);
    }

    var targetRow = existing ? existing.row : sheet.getLastRow() + 1;
    var rowValues = [];
    rowValues[toolCol - 1]   = tool;
    rowValues[aliasCol - 1]  = aliases.join(', ');
    rowValues[sourceCol - 1] = source;
    for (var c = 0; c < rowValues.length; c++) if (rowValues[c] === undefined) rowValues[c] = '';
    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    if (!existing) rowByKey[tool.toLowerCase()] = { row: targetRow, aliases: '', source: source };

    map[tool] = aliases;
  });

  return map;
}

function splitAliases_(text) {
  return String(text || '').split(',')
    .map(function (a) { return a.trim(); })
    .filter(function (a) { return a; });
}

// Rewrites Tool_Detected on Job_Discovery and Proposal_Generator with the
// current Primary_Tools + aliases, down every existing data row, then
// re-syncs Proposal_Generator since re-tagged jobs can change
// Final_Decision.
function rebuildToolDetection_(ss) {
  var tools    = String(getSettings_()['Primary_Tools'] || '');
  var aliasMap = getToolAliasMap_(ss);

  var jd = ss.getSheetByName('Job_Discovery');
  if (jd) {
    var jdHeaders = jd.getRange(1, 1, 1, jd.getLastColumn()).getValues()[0];
    var jdF = FFLib.buildJobDiscoveryFormulas(jdHeaders, tools, aliasMap);
    if (jdF.toolDetectedFormula) {
      var jdRows = Math.max(jd.getLastRow() - 1, FORMULA_PREFILL_ROWS);
      jd.getRange(2, jdF.toolDetectedCol, jdRows, 1).setFormula(jdF.toolDetectedFormula);
    }
  }

  var pg = ss.getSheetByName('Proposal_Generator');
  if (pg) {
    var pgHeaders = pg.getRange(1, 1, 1, pg.getLastColumn()).getValues()[0];
    var pgF = FFLib.buildProposalGeneratorFormulas(pgHeaders, tools, aliasMap);
    if (pgF.toolDetectedFormula) {
      var pgRows = Math.max(pg.getLastRow() - 1, FORMULA_PREFILL_ROWS);
      pg.getRange(2, pgF.toolDetectedCol, pgRows, 1).setFormula(pgF.toolDetectedFormula);
    }
  }

  syncProposalGenerator_(ss);
}
