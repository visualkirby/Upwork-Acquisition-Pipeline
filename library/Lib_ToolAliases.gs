/**
 * ============================================================
 * FreelanceFlow Library -- Tool Aliases
 * Lets a customer list tools the way they think of them ("Power BI,
 * Looker Studio, GA4") while Tool_Detected still catches the other names
 * job posts use ("PowerBI", "Data Studio", "Google Analytics").
 *
 * builtinToolAliases looks a tool up in a fixed map of common BI,
 * automation, and analytics tools. generateToolAliases is the GPT-4o-mini
 * fallback for anything outside the map. The thin client caches both in a
 * visible Tool_Aliases tab (getToolAliasMap_, 29_Tool_Aliases.gs), so the
 * AI runs once per new tool, never per job.
 *
 * Every alias list is deliberately conservative: short bare words that
 * appear in ordinary text ("GA", "Make", "Monday", "Planner", "Flow") are
 * left out, because Tool_Detected matches them anywhere in a description.
 * ============================================================
 */
var BUILTIN_TOOL_ALIASES_ = [
  ['Power BI',                ['PowerBI', 'PBI', 'Microsoft Power BI', 'MS Power BI']],
  ['Power Automate',          ['Microsoft Flow', 'MS Flow', 'PowerAutomate', 'Power Automate Desktop']],
  ['Power Apps',              ['PowerApps', 'Microsoft Power Apps']],
  ['Power Platform',          ['Microsoft Power Platform']],
  ['Power Query',             ['PowerQuery']],
  ['Excel',                   ['Microsoft Excel', 'MS Excel', 'Excel 365']],
  ['Google Sheets',           ['GSheets', 'G Sheets', 'Google Spreadsheet', 'Google Spreadsheets']],
  ['Apps Script',             ['Google Apps Script']],
  ['Looker Studio',           ['Data Studio', 'Google Data Studio', 'Looker Studio Pro']],
  ['GA4',                     ['Google Analytics 4', 'Google Analytics', 'GA 4']],
  // "GTM" is left out on purpose: job posts use it for "go-to-market" too.
  ['Google Tag Manager',      ['Tag Manager']],
  ['Google Search Console',   ['Search Console', 'GSC']],
  ['PageSpeed Insights',      ['PageSpeed', 'Page Speed Insights', 'Core Web Vitals']],
  ['Google Ads',              ['AdWords', 'Google AdWords']],
  ['Meta Ads',                ['Facebook Ads', 'Instagram Ads', 'Facebook Ads Manager']],
  ['BigQuery',                ['Big Query', 'Google BigQuery', 'GBQ']],
  ['SQL',                     ['MySQL', 'PostgreSQL', 'Postgres', 'T-SQL', 'SQL Server', 'SQLite']],
  ['SharePoint',              ['Share Point', 'SharePoint Online']],
  ['Microsoft Teams',         ['MS Teams']],
  ['Microsoft 365',           ['M365', 'Office 365', 'O365', 'MS 365']],
  ['Outlook',                 ['Microsoft Outlook', 'MS Outlook']],
  ['Microsoft Planner',       ['MS Planner', 'Planner Premium']],
  ['Microsoft Project',       ['MS Project', 'Project for the Web']],
  ['Dynamics 365',            ['D365', 'Microsoft Dynamics', 'Dynamics CRM']],
  ['Copilot',                 ['Microsoft Copilot', 'Copilot Studio', 'M365 Copilot']],
  ['Salesforce',              ['SFDC', 'Sales Cloud']],
  ['HubSpot',                 ['Hub Spot', 'HubSpot CRM']],
  ['Tableau',                 ['Tableau Desktop', 'Tableau Public', 'Tableau Server']],
  ['Make',                    ['Make.com', 'Integromat']],
  ['Monday.com',              ['monday dot com', 'Monday Work Management']],
  ['Airtable',                ['Air Table']],
  ['Smartsheet',              ['Smart Sheet']],
  ['QuickBooks',              ['QBO', 'QuickBooks Online', 'QuickBooks Desktop', 'Intuit QuickBooks']],
  ['Odoo',                    ['Odoo ERP']],
  ['WordPress',               ['Word Press']],
  ['dbt',                     ['data build tool']]
];

function normalizeToolName_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Other names for toolName from the built-in map, or null if the map has no
// entry. Matches loosely ("PowerBI", "power bi" and "Power BI" all find the
// same entry) and never returns a name equal to toolName itself.
function builtinToolAliases(toolName) {
  var key = normalizeToolName_(toolName);
  if (!key) return null;

  for (var i = 0; i < BUILTIN_TOOL_ALIASES_.length; i++) {
    var names = [BUILTIN_TOOL_ALIASES_[i][0]].concat(BUILTIN_TOOL_ALIASES_[i][1]);
    var hit = names.some(function (n) { return normalizeToolName_(n) === key; });
    if (hit) {
      // Drop only the exact name (case-insensitive). Spacing variants like
      // "PowerBI" for "Power BI" stay, since the regex needs them spelled out.
      var own = String(toolName).trim().toLowerCase();
      return names.filter(function (n) { return n.toLowerCase() !== own; });
    }
  }
  return null;
}

// GPT-4o-mini fallback for a tool the built-in map doesn't cover. Returns
// { ok, aliases } -- aliases filtered to drop anything under 3 characters
// or equal to the tool's own name.
function generateToolAliases(toolName, apiKey) {
  if (!apiKey) return { ok: false, message: 'API key not found.' };

  var prompt =
    'A freelancer lists "' + toolName + '" as a software tool they use. Job posts on Upwork ' +
    'sometimes refer to the same tool by a different name.\n\n' +
    'List up to 6 other names job posts use for exactly this tool: abbreviations, former ' +
    'product names, spacing variants, and the vendor-prefixed name. Do not include different ' +
    'products, generic words (like "dashboard" or "automation"), or any name shorter than 3 ' +
    'characters. Return an empty list if there are none.\n\n' +
    'Return only valid JSON with exactly this field:\n' +
    '{\n' +
    '  "aliases": [string]\n' +
    '}\n\n' +
    'Return only valid JSON. No explanation. No markdown.';

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 200,
        temperature: 0
      }),
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    if (data.error) return { ok: false, message: 'API error: ' + data.error.message };

    var content = data.choices && data.choices[0]
      ? data.choices[0].message.content.trim()
      : '';
    content = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    var parsed = JSON.parse(content);
    var own    = String(toolName).trim().toLowerCase();
    var seen   = {};
    var aliases = (Array.isArray(parsed.aliases) ? parsed.aliases : [])
      .map(function (a) { return String(a).trim(); })
      .filter(function (a) {
        var k = a.toLowerCase();
        if (a.length < 3 || k === own || seen[k]) return false;
        seen[k] = true;
        return true;
      });
    return { ok: true, aliases: aliases };
  } catch (e) {
    return { ok: false, message: 'Could not parse AI response.' };
  }
}
