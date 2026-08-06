/**
 * Module: Lib_LicenseGrant.gs
 * Purpose: Web App (doPost) that receives Gumroad's Ping webhook for FreelanceFlow.
 *          On a sale: verifies it, creates a private per-buyer copy of the Master Template,
 *          shares it with the buyer as Editor, and logs the grant.
 *          On a refund (same Ping resource, `refunded: true`): revokes the buyer's Editor
 *          access on their existing copy (the file itself is kept, not deleted, for
 *          audit/record purposes) and logs the revocation.
 * Reads:   Script Properties (config only, no sheet reads)
 * Writes:  Grants_Log spreadsheet (one row per webhook call, success/failure/revoked)
 *
 * Setup required before deploying (run setLicenseGrantProperties_ once with real values,
 * then delete/clear the values from the function body -- never leave secrets in source):
 *   MASTER_TEMPLATE_ID        - Sheet ID of the FreelanceFlow file buyers' copies are made from
 *   GRANTS_LOG_ID              - Sheet ID of the Grants_Log spreadsheet
 *   SOLD_COPIES_FOLDER_ID      - Drive folder ID where buyer copies get filed
 *   ALLOWED_PRODUCT_PERMALINKS - comma-separated Gumroad permalinks this should act on
 *   GUMROAD_SELLER_ID          - Sawandi's Gumroad seller_id, for lightweight payload verification
 */

function getColIndex(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx === -1) throw new Error(`Header "${headerName}" not found in sheet "${sheet.getName()}"`);
  return idx + 1;
}

/**
 * Run this once manually from the Apps Script editor to store config, then clear the
 * literal values below so they don't sit in source control. Never call this from doPost.
 */
function setLicenseGrantProperties_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'MASTER_TEMPLATE_ID': 'PASTE_MASTER_TEMPLATE_SHEET_ID',
    'GRANTS_LOG_ID': 'PASTE_GRANTS_LOG_SHEET_ID',
    'SOLD_COPIES_FOLDER_ID': 'PASTE_DRIVE_FOLDER_ID',
    'ALLOWED_PRODUCT_PERMALINKS': 'PASTE_COMMA_SEPARATED_PERMALINKS',
    'GUMROAD_SELLER_ID': 'PASTE_GUMROAD_SELLER_ID'
  });
}

function doPost(e) {
  const params = e.parameter;
  logGrant_(buildFailureRow_(params, 'received (seller_id=' + params.seller_id + ')'));

  const props = PropertiesService.getScriptProperties();

  if (params.resource_name && params.resource_name !== 'sale') {
    return ContentService.createTextOutput('ignored: not a sale event');
  }
  if (params.test === 'true' || params.test === true) {
    return ContentService.createTextOutput('ignored: test purchase');
  }

  const expectedSellerId = props.getProperty('GUMROAD_SELLER_ID');
  if (!expectedSellerId || params.seller_id !== expectedSellerId) {
    logGrant_(buildFailureRow_(params, 'rejected: seller_id mismatch'));
    return ContentService.createTextOutput('rejected');
  }

  const allowedPermalinks = (props.getProperty('ALLOWED_PRODUCT_PERMALINKS') || '')
    .split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const permalink = params.product_permalink || params.permalink || '';
  if (allowedPermalinks.length && allowedPermalinks.indexOf(permalink) === -1) {
    return ContentService.createTextOutput('ignored: not a FreelanceFlow product');
  }

  const buyerEmail = params.email;
  if (!buyerEmail) {
    logGrant_(buildFailureRow_(params, 'failed: no buyer email in payload'));
    return ContentService.createTextOutput('failed: no email');
  }

  const isRefund = params.refunded === 'true' || params.refunded === true;
  if (isRefund) {
    return handleRefund_(buyerEmail, permalink);
  }

  if (alreadyGranted_(buyerEmail, permalink)) {
    return ContentService.createTextOutput('skipped: already granted');
  }

  try {
    const masterId = props.getProperty('MASTER_TEMPLATE_ID');
    const folderId = props.getProperty('SOLD_COPIES_FOLDER_ID');
    const masterFile = DriveApp.getFileById(masterId);
    const destFolder = folderId ? DriveApp.getFolderById(folderId) : masterFile.getParents().next();

    const copy = masterFile.makeCopy('FreelanceFlow -- ' + buyerEmail, destFolder);
    copy.addEditor(buyerEmail);

    logGrant_({
      email: buyerEmail,
      tier: permalink,
      timestamp: new Date(),
      copyUrl: copy.getUrl(),
      status: 'success'
    });

    return ContentService.createTextOutput('granted');
  } catch (err) {
    logGrant_({
      email: buyerEmail,
      tier: permalink,
      timestamp: new Date(),
      copyUrl: '',
      status: 'failed: ' + err.message
    });
    return ContentService.createTextOutput('failed: ' + err.message);
  }
}

/**
 * Revokes a refunded buyer's Editor access on their existing copy. The copy itself is
 * kept (not deleted) so there's still a record if the refund is later reversed/disputed.
 * Idempotent: a webhook retry after a successful revoke is a no-op, not a repeat attempt.
 */
function handleRefund_(buyerEmail, permalink) {
  const grant = findGrantForRevoke_(buyerEmail, permalink);

  if (!grant.copyUrl) {
    logGrant_({
      email: buyerEmail,
      tier: permalink,
      timestamp: new Date(),
      copyUrl: '',
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
      email: buyerEmail,
      tier: permalink,
      timestamp: new Date(),
      copyUrl: grant.copyUrl,
      status: 'revoked'
    });
    return ContentService.createTextOutput('revoked');
  } catch (err) {
    logGrant_({
      email: buyerEmail,
      tier: permalink,
      timestamp: new Date(),
      copyUrl: grant.copyUrl,
      status: 'revoke_failed: ' + err.message
    });
    return ContentService.createTextOutput('revoke_failed: ' + err.message);
  }
}

function alreadyGranted_(email, tier) {
  const sheet = getGrantsLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const emailCol = getColIndex(sheet, 'Email');
  const tierCol = getColIndex(sheet, 'Tier');
  const statusCol = getColIndex(sheet, 'Status');
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return data.some(function (row) {
    return row[emailCol - 1] === email && row[tierCol - 1] === tier && row[statusCol - 1] === 'success';
  });
}

/**
 * Scans Grants_Log in order for this email+tier and returns the most recent successful
 * grant's Copy URL, plus whether a 'revoked' row has already logged after that grant
 * (so a refund webhook retry doesn't attempt a second revoke).
 */
function findGrantForRevoke_(email, tier) {
  const sheet = getGrantsLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { copyUrl: null, alreadyRevoked: false };

  const emailCol = getColIndex(sheet, 'Email');
  const tierCol = getColIndex(sheet, 'Tier');
  const statusCol = getColIndex(sheet, 'Status');
  const copyUrlCol = getColIndex(sheet, 'Copy URL');
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let latestCopyUrl = null;
  let revokedSinceLastGrant = false;

  data.forEach(function (row) {
    if (row[emailCol - 1] !== email || row[tierCol - 1] !== tier) return;
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

function buildFailureRow_(params, note) {
  return {
    email: params.email || '',
    tier: params.product_permalink || params.permalink || '',
    timestamp: new Date(),
    copyUrl: '',
    status: note
  };
}

function logGrant_(row) {
  const sheet = getGrantsLogSheet_();
  const emailCol = getColIndex(sheet, 'Email');
  const tierCol = getColIndex(sheet, 'Tier');
  const timestampCol = getColIndex(sheet, 'Timestamp');
  const copyUrlCol = getColIndex(sheet, 'Copy URL');
  const statusCol = getColIndex(sheet, 'Status');

  const newRow = new Array(sheet.getLastColumn()).fill('');
  newRow[emailCol - 1] = row.email;
  newRow[tierCol - 1] = row.tier;
  newRow[timestampCol - 1] = row.timestamp;
  newRow[copyUrlCol - 1] = row.copyUrl;
  newRow[statusCol - 1] = row.status;

  sheet.appendRow(newRow);
}

function getGrantsLogSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('GRANTS_LOG_ID');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('Grants_Log') || ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Email', 'Tier', 'Timestamp', 'Copy URL', 'Status']);
  }
  return sheet;
}
