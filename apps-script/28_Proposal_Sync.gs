/**
 * ============================================================
 * 28. PROPOSAL_GENERATOR SYNC
 * Proposal_Generator holds one stable row per job, keyed by Discovery_ID.
 * syncProposalGenerator_ appends a row (static Discovery_ID + Date) for
 * every Job_Scoring APPLY job that doesn't have one yet. Rows are never
 * reordered or removed, so the typed columns beside them (Job_Type, bids,
 * Proposal_Status, AI_Generated_Proposal) always stay with their job. The
 * rest of each row is per-row lookup formulas keyed on its Discovery_ID
 * (FFLib.buildProposalGeneratorLookupFormulas).
 *
 * A job that drops off APPLY after its row exists keeps the row. Unless it's
 * already Sent or Skip, its Notes get a "[Not APPLY now: ...]" tag naming the
 * current decision, and the tag clears itself if the job returns to APPLY.
 * Proposal_Status is never set for the freelancer.
 *
 * Final_Decision mostly changes through recalculation (a Job_Discovery
 * entry, the Connects balance, Settings thresholds), which never fires
 * onEdit. So this runs from every path that can cause one: handleEdit on
 * Job_Discovery/Job_Scoring/Connects_Helper/Settings, Log New Job, a
 * proposal marked Sent, Log Proposal Bid opening, and Repair Formulas.
 * ============================================================
 */
var NOT_APPLY_TAG_RE_     = /\s*\[Not APPLY now:[^\]]*\]/g;
var NOT_APPLY_TAG_PREFIX_ = '[Not APPLY now:';

// Returns { added: [Discovery_ID strings appended this run], flagged: rows
// currently carrying a Not-APPLY tag }.
function syncProposalGenerator_(ss) {
  var result  = { added: [], flagged: 0 };
  var pgSheet = ss.getSheetByName('Proposal_Generator');
  var jsSheet = ss.getSheetByName('Job_Scoring');
  if (!pgSheet || !jsSheet || jsSheet.getLastRow() < 2) return result;

  var pgMap       = getHeaderMap_(pgSheet);
  var pgIdCol     = getCol_(pgMap, ['Discovery_ID']);
  var pgDateCol   = getCol_(pgMap, ['Date']);
  var pgStatusCol = getCol_(pgMap, ['Proposal_Status']);
  var pgNotesCol  = getCol_(pgMap, ['Notes']);
  if (!pgIdCol) return result;

  // Still on the old FILTER layout. Writing static IDs under a live spill
  // would break it with #REF!, so this waits for Repair Formulas to migrate.
  if (pgSheet.getRange(2, pgIdCol).getFormula()) return result;

  var jsMap       = getHeaderMap_(jsSheet);
  var jsIdCol     = getCol_(jsMap, ['Discovery_ID']);
  var jsDecCol    = getCol_(jsMap, ['Final_Decision']);
  var jsAffordCol = getCol_(jsMap, ['Connects_Affordability']);
  var jsAgeCol    = getCol_(jsMap, ['Current_Age_Days']);
  var jsPropCol   = getCol_(jsMap, ['Proposal_Count']);
  if (!jsIdCol || !jsDecCol) return result;

  var settings  = getSettings_();
  var capNumber = function (v) {
    var n = parseFloat(v);
    return isNaN(n) ? null : n;
  };
  var caps = {
    maxAge:   capNumber(settings['Apply_Max_Age_Days']),
    maxProps: capNumber(settings['Apply_Max_Proposals'])
  };

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    SpreadsheetApp.flush();

    var jsData = jsSheet.getRange(2, 1, jsSheet.getLastRow() - 1, jsSheet.getLastColumn()).getValues();
    var decisionById = {};
    var applyIds     = [];
    jsData.forEach(function (r) {
      var id = String(r[jsIdCol - 1]).trim();
      if (!id) return;
      var decision = String(r[jsDecCol - 1]).trim();
      decisionById[id] = {
        decision: decision,
        afford:   jsAffordCol ? String(r[jsAffordCol - 1]).trim() : '',
        age:      jsAgeCol ? r[jsAgeCol - 1] : '',
        props:    jsPropCol ? r[jsPropCol - 1] : ''
      };
      if (decision === 'APPLY') applyIds.push(r[jsIdCol - 1]);
    });

    var lastIdRow  = getLastProposalGeneratorRow_(pgSheet, pgIdCol);
    var pgRowCount = lastIdRow - 1;
    var pgIds      = pgRowCount > 0 ? pgSheet.getRange(2, pgIdCol, pgRowCount, 1).getValues() : [];
    var existing   = {};
    pgIds.forEach(function (r) {
      var id = String(r[0]).trim();
      if (id) existing[id] = true;
    });

    if (pgNotesCol && pgRowCount > 0) {
      result.flagged = tagNotApplyRows_(pgSheet, pgIds, pgStatusCol, pgNotesCol, decisionById, caps);
    }

    var toAdd = applyIds.filter(function (id) { return !existing[String(id).trim()]; });
    if (toAdd.length > 0) {
      var startRow = getProposalGeneratorAppendRow_(pgSheet, lastIdRow);
      pgSheet.getRange(startRow, pgIdCol, toAdd.length, 1)
        .setValues(toAdd.map(function (id) { return [id]; }))
        .setNumberFormat('0');
      if (pgDateCol) {
        pgSheet.getRange(startRow, pgDateCol, toAdd.length, 1).setValue(new Date());
      }
      extendProposalGeneratorRowFormulas_(pgSheet, pgIdCol, pgStatusCol, startRow, toAdd.length);
      SpreadsheetApp.flush();
      result.added = toAdd.map(function (id) { return String(id).trim(); });
    }
  } finally {
    lock.releaseLock();
  }

  return result;
}

// Adds, updates, or clears each row's Not-APPLY tag in Notes, keeping
// whatever the freelancer typed there. A blank Final_Decision (score still
// recalculating) leaves the row alone. Returns how many rows carry a tag.
//
// A job that already has a row keeps Final_Decision = APPLY past the age
// and proposal caps (FFLib.buildJobScoringFormulas exempts it), so an
// undecided row past a cap is tagged here instead.
function tagNotApplyRows_(pgSheet, pgIds, pgStatusCol, pgNotesCol, decisionById, caps) {
  var count      = pgIds.length;
  var statuses   = pgStatusCol ? pgSheet.getRange(2, pgStatusCol, count, 1).getValues() : null;
  var notesRange = pgSheet.getRange(2, pgNotesCol, count, 1);
  var notes      = notesRange.getValues();
  var changed    = false;
  var flagged    = 0;

  for (var i = 0; i < count; i++) {
    var id = String(pgIds[i][0]).trim();
    if (!id) continue;

    var status  = statuses ? String(statuses[i][0]).trim() : '';
    var current = String(notes[i][0]);
    var hadTag  = current.indexOf(NOT_APPLY_TAG_PREFIX_) >= 0;
    var info    = decisionById[id];
    var tag     = '';

    if (status !== 'Sent' && status !== 'Skip') {
      if (!info) {
        tag = NOT_APPLY_TAG_PREFIX_ + ' removed from Job_Scoring]';
      } else if (!info.decision) {
        continue;
      } else if (info.decision !== 'APPLY') {
        tag = NOT_APPLY_TAG_PREFIX_ + ' ' + info.decision +
          (info.afford === 'Cannot Afford' ? ', Cannot Afford' : '') + ']';
      } else {
        var reasons = [];
        var age     = typeof info.age === 'number' ? info.age : parseFloat(info.age);
        if (caps.maxAge !== null && !isNaN(age) && age > caps.maxAge) {
          reasons.push('older than ' + caps.maxAge + ' days');
        }
        if (caps.maxProps !== null && FFLib.proposalCountFloor(info.props) >= caps.maxProps) {
          reasons.push(info.props + ' proposals, cap ' + caps.maxProps);
        }
        if (reasons.length > 0) tag = NOT_APPLY_TAG_PREFIX_ + ' ' + reasons.join('; ') + ']';
      }
    }

    if (tag) flagged++;
    if (!tag && !hadTag) continue;

    var base = current.replace(NOT_APPLY_TAG_RE_, '').trim();
    var next = tag ? (base ? base + ' ' + tag : tag) : base;
    if (next !== current) {
      notes[i][0] = next;
      changed = true;
    }
  }

  if (changed) notesRange.setValues(notes);
  return flagged;
}

// Formula columns are prefilled FORMULA_PREFILL_ROWS deep. A row appended
// past that copies row 2's formulas (lookups, Tool_Detected) and the
// Proposal_Status dropdown down to it.
function extendProposalGeneratorRowFormulas_(pgSheet, pgIdCol, pgStatusCol, startRow, numRows) {
  var lastCol     = pgSheet.getLastColumn();
  var row2        = pgSheet.getRange(2, 1, 1, lastCol).getFormulas()[0];
  var lastNewRow  = startRow + numRows - 1;
  var lastRowFormulas = pgSheet.getRange(lastNewRow, 1, 1, lastCol).getFormulas()[0];

  for (var c = 1; c <= lastCol; c++) {
    if (c === pgIdCol || !row2[c - 1] || lastRowFormulas[c - 1]) continue;
    pgSheet.getRange(2, c).copyTo(pgSheet.getRange(startRow, c, numRows, 1));
  }

  if (pgStatusCol && !pgSheet.getRange(lastNewRow, pgStatusCol).getDataValidation()) {
    pgSheet.getRange(2, pgStatusCol).copyTo(
      pgSheet.getRange(startRow, pgStatusCol, numRows, 1),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );
  }
}

// Last row with a Discovery_ID, or 1 if none. getLastRow() alone would
// return the bottom of the prefilled formula range.
function getLastProposalGeneratorRow_(pgSheet, pgIdCol) {
  var lastRow = pgSheet.getLastRow();
  if (lastRow < 2) return 1;
  var ids = pgSheet.getRange(2, pgIdCol, lastRow - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

// First row below both the last Discovery_ID and the last typed value.
// An orphaned row (typed data, no ID) left over from the old FILTER layout
// must never be handed to a new job.
function getProposalGeneratorAppendRow_(pgSheet, lastIdRow) {
  var lastRow = pgSheet.getLastRow();
  if (lastRow <= lastIdRow) return lastIdRow + 1;

  var lastCol = pgSheet.getLastColumn();
  var row2    = pgSheet.getRange(2, 1, 1, lastCol).getFormulas()[0];
  var values  = pgSheet.getRange(lastIdRow + 1, 1, lastRow - lastIdRow, lastCol).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    for (var c = 0; c < lastCol; c++) {
      if (!row2[c] && String(values[i][c]).trim() !== '') return lastIdRow + i + 2;
    }
  }
  return lastIdRow + 1;
}

// One-time move from the old FILTER(Final_Decision="APPLY") spill to static
// rows. Reads the IDs the spill currently shows, clears every column the
// spill owned, then writes those IDs back as static values in the same
// order, so each row's typed data stays with the job it sits beside today.
// The old Date column showed TODAY(), so each migrated row's Date becomes
// its Proposal_Sent_Date or Proposal_Skip_Date, or today if it has neither.
// No-ops (returns 0) once Discovery_ID is already static.
function migrateProposalGeneratorToStaticRows_(pgSheet, headers) {
  var idCol = headers.indexOf('Discovery_ID') + 1;
  if (idCol <= 0 || !pgSheet.getRange(2, idCol).getFormula()) return 0;

  // The old spill's Date column reads Job_Scoring's Proposal_Generator_Date,
  // and the new Proposal_Generator_Date lookup reads this sheet's Date. If
  // that lookup is already in place (ensurePipelineSheets_ applies it), the
  // two form a loop and the whole spill shows #REF!. Clearing it breaks the
  // loop; REPAIR_FORMULAS re-applies it once this migration is done.
  var jsSheet = pgSheet.getParent().getSheetByName('Job_Scoring');
  if (jsSheet) {
    var jsDateCol = getCol_(getHeaderMap_(jsSheet), ['Proposal_Generator_Date']);
    if (jsDateCol && jsSheet.getLastRow() >= 2) {
      jsSheet.getRange(2, jsDateCol, jsSheet.getLastRow() - 1, 1).clearContent();
      SpreadsheetApp.flush();
    }
  }

  var lastRow = pgSheet.getLastRow();
  if (lastRow < 2) return 0;
  var numRows = lastRow - 1;
  var data    = pgSheet.getRange(2, 1, numRows, pgSheet.getLastColumn()).getValues();

  var firstId = String(data[0][idCol - 1]);
  if (firstId.charAt(0) === '#') {
    throw new Error('Proposal_Generator\'s Discovery_ID column shows ' + firstId +
      '. Clear whatever is blocking the FILTER in column ' + idCol + ', then run Repair Formulas again.');
  }

  var sentCol = headers.indexOf('Proposal_Sent_Date') + 1;
  var skipCol = headers.indexOf('Proposal_Skip_Date') + 1;
  var today   = new Date();
  var count   = 0;
  for (var i = 0; i < numRows; i++) {
    if (String(data[i][idCol - 1]).trim() !== '') count = i + 1;
  }

  var ids   = [];
  var dates = [];
  for (var j = 0; j < count; j++) {
    var sent = sentCol ? data[j][sentCol - 1] : '';
    var skip = skipCol ? data[j][skipCol - 1] : '';
    ids.push([data[j][idCol - 1]]);
    dates.push([sent || skip || today]);
  }

  ['Discovery_ID', 'Date', 'Job_Title', 'Client_Name', 'Description', 'Job_Link',
   'Keyword_Search', 'Connects_Required', 'Proposal_Count', 'Budget'].forEach(function (name) {
    var col = headers.indexOf(name) + 1;
    if (col > 0) pgSheet.getRange(2, col, numRows, 1).clearContent();
  });

  if (count > 0) {
    pgSheet.getRange(2, idCol, count, 1).setValues(ids).setNumberFormat('0');
    var dateCol = headers.indexOf('Date') + 1;
    if (dateCol > 0) pgSheet.getRange(2, dateCol, count, 1).setValues(dates);
  }
  return count;
}
