/**
 * ============================================================
 * 27. HUBSPOT CRM SYNC
 *
 * Pushes Proposal_Tracker activity into a HubSpot CRM pipeline so the
 * Upwork funnel (Sent -> Reply -> Interview -> Hired) stays visible as a
 * real deal pipeline instead of only living in this sheet.
 *
 * One-way sync, Sheet -> HubSpot. HubSpot is never read back into the
 * sheet except to store the Contact/Deal IDs this script itself creates,
 * so those IDs can be reused on the next stage update instead of
 * re-searching.
 *
 * Setup:
 *   1. In HubSpot: Settings > Integrations > Private Apps > create one
 *      with crm.objects.contacts.write/read and crm.objects.deals.write/read
 *      scopes. Copy the access token.
 *   2. Run System Tools > Setup HubSpot Access Token, paste it in.
 *   3. In HubSpot: create a pipeline (e.g. "Upwork Client Acquisition")
 *      with 4 stages matching the funnel below. Copy the pipeline ID and
 *      each stage ID (Settings > Objects > Deals > Pipelines, or via the
 *      CRM API's /crm/v3/pipelines/deals endpoint).
 *   4. Add these rows to the Settings sheet:
 *        HubSpot_Pipeline_ID
 *        HubSpot_Stage_Proposal_Sent
 *        HubSpot_Stage_Reply_Received
 *        HubSpot_Stage_Interview
 *        HubSpot_Stage_Hired
 *
 * Nothing in this file hardcodes a pipeline/stage ID -- all 5 come from
 * the Settings sheet via getSettings_(), same convention as the rest of
 * this project's config.
 * ============================================================
 */

function SETUP_HUBSPOT_ACCESS_TOKEN() {
  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt("Setup HubSpot Access Token", "Paste your HubSpot Private App access token below:", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var token = response.getResponseText().trim();

  if (!token) {
    ui.alert("No token entered. Setup cancelled.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("UPWORK_HUBSPOT_ACCESS_TOKEN", token);
  ui.alert("HubSpot access token stored. You do not need to run this again unless you rotate it.");
}


function CHECK_HUBSPOT_ACCESS_TOKEN() {
  var ui    = SpreadsheetApp.getUi();
  var token = PropertiesService.getScriptProperties().getProperty("UPWORK_HUBSPOT_ACCESS_TOKEN");

  if (!token) {
    ui.alert("No HubSpot access token found. Run System Tools > Setup HubSpot Access Token.");
    return;
  }

  ui.alert("HubSpot access token is set.\nLength: " + token.length + " chars\nPrefix: " + token.substring(0, 8) + "...");
}


function getHubSpotToken_() {
  var token = PropertiesService.getScriptProperties().getProperty("UPWORK_HUBSPOT_ACCESS_TOKEN");
  if (!token) throw new Error("HubSpot access token not set. Run System Tools > Setup HubSpot Access Token first.");
  return token;
}


// Prints every deal pipeline and its stage IDs so the values can be pasted
// into the Settings sheet. Run once after creating the pipeline in HubSpot.
function LIST_HUBSPOT_IDS() {
  var ui = SpreadsheetApp.getUi();
  var data;
  try {
    data = hubspotRequest_("GET", "/crm/v3/pipelines/deals", null);
  } catch (err) {
    ui.alert(err.message);
    return;
  }

  var lines = [];
  (data.results || []).forEach(function (p) {
    lines.push("PIPELINE: " + p.label);
    lines.push("  HubSpot_Pipeline_ID = " + p.id);
    (p.stages || [])
      .sort(function (a, b) { return a.displayOrder - b.displayOrder; })
      .forEach(function (st) {
        lines.push("  stage \"" + st.label + "\" = " + st.id);
      });
    lines.push("");
  });

  ui.alert("HubSpot Deal Pipelines", lines.join("\n"), ui.ButtonSet.OK);
}


// Shared request wrapper -- bearer auth, JSON in/out, throws with the
// response body on any non-2xx so a failed sync is loud in the alert/log
// rather than silently doing nothing.
function hubspotRequest_(method, path, payload) {
  var options = {
    method:             method,
    contentType:        "application/json",
    headers:            { Authorization: "Bearer " + getHubSpotToken_() },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);

  var response = UrlFetchApp.fetch("https://api.hubapi.com" + path, options);
  var code     = response.getResponseCode();
  var body     = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("HubSpot API " + method + " " + path + " returned " + code + ": " + body);
  }
  return body ? JSON.parse(body) : {};
}


// Adds HubSpot_Contact_ID / HubSpot_Deal_ID to Proposal_Tracker's header
// row if they aren't there yet -- lets an existing, already-populated
// sheet pick up this feature without a Setup Wizard re-run. Dynamic
// column placement (next open header cell), no fixed column number.
function ensureHubSpotColumns_(sheet) {
  var map     = getHeaderMap_(sheet);
  var lastCol = sheet.getLastColumn();
  var toAdd   = [];

  if (!getCol_(map, ["HubSpot_Contact_ID"])) toAdd.push("HubSpot_Contact_ID");
  if (!getCol_(map, ["HubSpot_Deal_ID"]))    toAdd.push("HubSpot_Deal_ID");

  if (toAdd.length) {
    sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  }
}


// Finds a Contact by exact firstname match (Upwork proposals don't carry
// a client email, so there's no cleaner dedup key available) or creates
// one. Returns the Contact ID.
function hubspotFindOrCreateContact_(clientName) {
  var searchResult = hubspotRequest_("POST", "/crm/v3/objects/contacts/search", {
    filterGroups: [{
      filters: [{ propertyName: "firstname", operator: "EQ", value: clientName }]
    }],
    limit: 1
  });

  if (searchResult.results && searchResult.results.length) {
    return searchResult.results[0].id;
  }

  var created = hubspotRequest_("POST", "/crm/v3/objects/contacts", {
    properties: { firstname: clientName }
  });
  return created.id;
}


function hubspotAssociateDealToContact_(dealId, contactId) {
  hubspotRequest_("PUT", "/crm/v4/objects/deal/" + dealId + "/associations/default/contact/" + contactId, null);
}


// Fires once, right after a Proposal_Tracker row is created for a "Sent"
// proposal (see handleProposalStatusChange_ in 14_Edit_Trigger.gs). Creates
// the Contact + Deal, drops the Deal into the Proposal Sent stage, and
// writes both IDs back onto the row so later stage moves can PATCH by ID
// instead of re-searching. Wrapped in try/catch by the caller -- a HubSpot
// outage or missing Settings config shouldn't block the sheet-side sync
// this function rides along with.
function hubspotSyncProposalSent_(tracker, row, jobTitle, clientName) {
  var settings   = getSettings_();
  var pipelineId = settings["HubSpot_Pipeline_ID"];
  var stageId    = settings["HubSpot_Stage_Proposal_Sent"];
  if (!pipelineId || !stageId) {
    throw new Error("HubSpot_Pipeline_ID / HubSpot_Stage_Proposal_Sent not set on the Settings sheet.");
  }

  ensureHubSpotColumns_(tracker);
  var freshMap = getHeaderMap_(tracker);

  var contactId = hubspotFindOrCreateContact_(clientName || "Unknown Client");

  var deal = hubspotRequest_("POST", "/crm/v3/objects/deals", {
    properties: {
      dealname:  jobTitle + (clientName ? " (" + clientName + ")" : ""),
      pipeline:  pipelineId,
      dealstage: stageId
    }
  });

  hubspotAssociateDealToContact_(deal.id, contactId);

  setCellValue_(tracker, row, freshMap, ["HubSpot_Contact_ID"], contactId);
  setCellValue_(tracker, row, freshMap, ["HubSpot_Deal_ID"],    deal.id);
}


// Moves an existing Deal to a new stage -- called from the Viewed/
// Interview/Hired blocks in 14_Edit_Trigger.gs's PROPOSAL_TRACKER section.
// No-ops quietly if the row has no HubSpot_Deal_ID yet (proposal predates
// this feature, or the initial sync failed) rather than throwing, since a
// missing ID here isn't a real error condition worth surfacing every time.
function hubspotUpdateDealStage_(tracker, row, ptMap, settingsStageKey) {
  var dealId = getCellValue_(tracker, row, ptMap, ["HubSpot_Deal_ID"]);
  if (!dealId) return;

  var stageId = getSettings_()[settingsStageKey];
  if (!stageId) throw new Error(settingsStageKey + " not set on the Settings sheet.");

  hubspotRequest_("PATCH", "/crm/v3/objects/deals/" + dealId, {
    properties: { dealstage: stageId }
  });
}


// Guard wrapper for the edit-trigger call sites -- skips entirely when no
// token is configured, swallows any HubSpot error to the log so a CRM
// outage never breaks the Proposal_Tracker edit flow.
function hubspotMoveDealStageSafe_(tracker, row, ptMap, settingsStageKey) {
  if (!PropertiesService.getScriptProperties().getProperty("UPWORK_HUBSPOT_ACCESS_TOKEN")) return;
  try {
    hubspotUpdateDealStage_(tracker, row, ptMap, settingsStageKey);
  } catch (err) {
    Logger.log("HubSpot sync (" + settingsStageKey + ") failed: " + err.message);
  }
}
