/**
 * Module: Lib_WooCommerceGrant.gs
 * Purpose: Web App (doPost) that receives WooCommerce's order webhook for FreelanceFlow.
 *          On a completed order: verifies it, creates a private per-buyer copy of the
 *          Master Template, shares it with the buyer as Editor, and logs the grant.
 *          On a refund: revokes the buyer's Editor access on their existing copy (the
 *          file itself is kept, not deleted, for audit/record purposes) and logs the
 *          revocation.
 *
 * Kept as a standalone project, separate from the Gumroad license-grant script, on
 * purpose -- a public webhook endpoint for one sales channel should never be able to
 * touch another channel's config by accident. Both write to the same Grants_Log
 * spreadsheet (see the Source column) so there's one unified audit trail across
 * channels, and both create copies from the same Master Template into the same
 * Sold_Copies folder.
 *
 * Why the webhook body itself is never trusted: WooCommerce normally proves a webhook's
 * authenticity via an X-WC-Webhook-Signature header, but Apps Script Web Apps cannot
 * read custom request headers in doPost(e) at all. Instead, the payload's order ID is
 * used only to look the order up directly via WooCommerce's authenticated REST API
 * (Consumer Key/Secret, Read-only) -- the API response is the source of truth, not the
 * webhook body's own status/email/line-items.
 *
 * Reads:   Script Properties (config only, no sheet reads)
 * Writes:  Grants_Log spreadsheet (same sheet the Gumroad script writes to)
 *
 * Setup required before deploying -- set these directly via Project Settings > Script
 * Properties in the Apps Script editor (no code execution needed, no secrets in source):
 *   MASTER_TEMPLATE_ID  - Sheet ID of the FreelanceFlow file buyers' copies are made from
 *                          (same value as the Gumroad project's MASTER_TEMPLATE_ID)
 *   GRANTS_LOG_ID        - Sheet ID of the Grants_Log spreadsheet (same as Gumroad's)
 *   SOLD_COPIES_FOLDER_ID - Drive folder ID where buyer copies get filed (same as Gumroad's)
 *   WC_STORE_URL         - e.g. https://benchlineanalytics.com
 *   WC_CONSUMER_KEY      - WooCommerce REST API Consumer Key (Settings > Advanced > REST API, Read permission)
 *   WC_CONSUMER_SECRET   - WooCommerce REST API Consumer Secret
 *   WC_PRODUCT_ID        - WooCommerce product ID for the FreelanceFlow listing
 */

function getColIndex(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx === -1) throw new Error(`Header "${headerName}" not found in sheet "${sheet.getName()}"`);
  return idx + 1;
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput('ignored: unparseable JSON');
  }

  const orderId = payload && payload.id;
  if (!orderId) {
    return ContentService.createTextOutput('ignored: no order id in payload');
  }

  const props = PropertiesService.getScriptProperties();
  const order = fetchWooCommerceOrder_(
    props.getProperty('WC_STORE_URL'),
    props.getProperty('WC_CONSUMER_KEY'),
    props.getProperty('WC_CONSUMER_SECRET'),
    orderId
  );

  if (!order) {
    logGrant_({
      email: '', tier: 'order#' + orderId, timestamp: new Date(), copyUrl: '',
      status: 'failed: could not verify order via WooCommerce REST API'
    });
    return ContentService.createTextOutput('failed: could not verify order');
  }

  const productId = props.getProperty('WC_PRODUCT_ID');
  const hasProduct = (order.line_items || []).some(function (li) {
    return String(li.product_id) === String(productId);
  });
  if (!hasProduct) {
    return ContentService.createTextOutput('ignored: order does not contain FreelanceFlow');
  }

  const buyerEmail = order.billing && order.billing.email;
  if (!buyerEmail) {
    logGrant_({
      email: '', tier: 'FreelanceFlow', timestamp: new Date(), copyUrl: '',
      status: 'failed: no buyer email on order #' + orderId
    });
    return ContentService.createTextOutput('failed: no email on order');
  }

  if (order.status === 'refunded') {
    return handleRefund_(buyerEmail, 'FreelanceFlow');
  }
  if (order.status !== 'completed') {
    return ContentService.createTextOutput('ignored: order status is ' + order.status);
  }

  return grantAccess_(buyerEmail, 'FreelanceFlow');
}

/**
 * Fetches the order directly from WooCommerce's REST API rather than trusting the webhook
 * body -- this IS the authenticity check, standing in for the header-signature verification
 * Apps Script can't do (see module header comment).
 */
function fetchWooCommerceOrder_(storeUrl, consumerKey, consumerSecret, orderId) {
  if (!storeUrl || !consumerKey || !consumerSecret) return null;
  try {
    const url = storeUrl.replace(/\/$/, '') + '/wp-json/wc/v3/orders/' + encodeURIComponent(orderId);
    const auth = Utilities.base64Encode(consumerKey + ':' + consumerSecret);
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Basic ' + auth },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return null;
    return JSON.parse(resp.getContentText());
  } catch (err) {
    return null;
  }
}

function grantAccess_(buyerEmail, tier) {
  if (alreadyGranted_(buyerEmail, tier)) {
    return ContentService.createTextOutput('skipped: already granted');
  }

  const props = PropertiesService.getScriptProperties();
  try {
    const masterId = props.getProperty('MASTER_TEMPLATE_ID');
    const folderId = props.getProperty('SOLD_COPIES_FOLDER_ID');
    const masterFile = DriveApp.getFileById(masterId);
    const destFolder = folderId ? DriveApp.getFolderById(folderId) : masterFile.getParents().next();

    const copy = masterFile.makeCopy('FreelanceFlow -- ' + buyerEmail, destFolder);
    copy.addEditor(buyerEmail);

    logGrant_({
      email: buyerEmail, tier: tier, timestamp: new Date(),
      copyUrl: copy.getUrl(), status: 'success'
    });
    return ContentService.createTextOutput('granted');
  } catch (err) {
    logGrant_({
      email: buyerEmail, tier: tier, timestamp: new Date(),
      copyUrl: '', status: 'failed: ' + err.message
    });
    return ContentService.createTextOutput('failed: ' + err.message);
  }
}

/**
 * Revokes a refunded buyer's Editor access on their existing copy. The copy itself is
 * kept (not deleted) so there's still a record if the refund is later reversed/disputed.
 * Idempotent: a webhook retry after a successful revoke is a no-op, not a repeat attempt.
 */
function handleRefund_(buyerEmail, tier) {
  const grant = findGrantForRevoke_(buyerEmail, tier);

  if (!grant.copyUrl) {
    logGrant_({
      email: buyerEmail, tier: tier, timestamp: new Date(), copyUrl: '',
      status: 'revoke_failed: no prior grant found'
    });
    return ContentService.createTextOutput('revoke_failed: no prior grant found');
  }

  if (grant.alreadyRevoked) {
    return ContentService.createTextOutput('skipped: already revoked');
  }

  try {
    const fileId = extractFileIdFromUrl_(grant.copyUrl);
    if (!fileId) throw new Error('could not parse file ID from stored Copy URL');

    const file = DriveApp.getFileById(fileId);
    file.removeEditor(buyerEmail);

    logGrant_({
      email: buyerEmail, tier: tier, timestamp: new Date(),
      copyUrl: grant.copyUrl, status: 'revoked'
    });
    return ContentService.createTextOutput('revoked');
  } catch (err) {
    logGrant_({
      email: buyerEmail, tier: tier, timestamp: new Date(),
      copyUrl: grant.copyUrl, status: 'revoke_failed: ' + err.message
    });
    return ContentService.createTextOutput('revoke_failed: ' + err.message);
  }
}

/**
 * Scoped to rows with Source = 'WooCommerce' only -- a refund on this channel should
 * never touch a grant that actually came from Gumroad for the same email+tier.
 */
function alreadyGranted_(email, tier) {
  const sheet = getGrantsLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const emailCol = getColIndex(sheet, 'Email');
  const tierCol = getColIndex(sheet, 'Tier');
  const statusCol = getColIndex(sheet, 'Status');
  const sourceCol = getColIndex(sheet, 'Source');
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return data.some(function (row) {
    return row[emailCol - 1] === email && row[tierCol - 1] === tier &&
      row[statusCol - 1] === 'success' && row[sourceCol - 1] === 'WooCommerce';
  });
}

/**
 * Scans Grants_Log in order for this email+tier (WooCommerce rows only) and returns the
 * most recent successful grant's Copy URL, plus whether a 'revoked' row has already
 * logged after that grant (so a refund webhook retry doesn't attempt a second revoke).
 */
function findGrantForRevoke_(email, tier) {
  const sheet = getGrantsLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { copyUrl: null, alreadyRevoked: false };

  const emailCol = getColIndex(sheet, 'Email');
  const tierCol = getColIndex(sheet, 'Tier');
  const statusCol = getColIndex(sheet, 'Status');
  const copyUrlCol = getColIndex(sheet, 'Copy URL');
  const sourceCol = getColIndex(sheet, 'Source');
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let latestCopyUrl = null;
  let revokedSinceLastGrant = false;

  data.forEach(function (row) {
    if (row[emailCol - 1] !== email || row[tierCol - 1] !== tier || row[sourceCol - 1] !== 'WooCommerce') return;
    const status = row[statusCol - 1];
    if (status === 'success') {
      latestCopyUrl = row[copyUrlCol - 1];
      revokedSinceLastGrant = false;
    } else if (status === 'revoked') {
      revokedSinceLastGrant = true;
    }
  });

  return { copyUrl: latestCopyUrl, alreadyRevoked: revokedSinceLastGrant };
}

function extractFileIdFromUrl_(url) {
  const match = /\/d\/([a-zA-Z0-9_-]+)/.exec(url || '');
  return match ? match[1] : null;
}

function logGrant_(row) {
  const sheet = getGrantsLogSheet_();
  const emailCol = getColIndex(sheet, 'Email');
  const tierCol = getColIndex(sheet, 'Tier');
  const timestampCol = getColIndex(sheet, 'Timestamp');
  const copyUrlCol = getColIndex(sheet, 'Copy URL');
  const statusCol = getColIndex(sheet, 'Status');
  const sourceCol = getColIndex(sheet, 'Source');

  const newRow = new Array(sheet.getLastColumn()).fill('');
  newRow[emailCol - 1] = row.email;
  newRow[tierCol - 1] = row.tier;
  newRow[timestampCol - 1] = row.timestamp;
  newRow[copyUrlCol - 1] = row.copyUrl;
  newRow[statusCol - 1] = row.status;
  newRow[sourceCol - 1] = 'WooCommerce';

  sheet.appendRow(newRow);
}

/**
 * Points at the SAME Grants_Log spreadsheet the Gumroad project writes to (via the
 * GRANTS_LOG_ID Script Property, set to the same Sheet ID). Adds the Source column if
 * it's somehow still missing (defensive only -- the Gumroad project already added it).
 */
function getGrantsLogSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('GRANTS_LOG_ID');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('Grants_Log') || ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Email', 'Tier', 'Timestamp', 'Copy URL', 'Status', 'Source']);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('Source') === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Source');
    }
  }
  return sheet;
}
