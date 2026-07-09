/**
 * ============================================================
 * 24. MANAGE PROJECTS SIDEBAR
 * Add/edit portfolio projects after initial setup -- the Setup Wizard's
 * Step 6 only collects 2-5 projects at first-run time, but Upwork profiles
 * support up to 10. This sidebar is the add-more/edit-existing path.
 *
 * Projects sheet has no ID column (00_Setup_Wizard.gs's initProjectsSheet_),
 * so rows are matched by Project_Name string, same as getPortfolioMapFromProjects_
 * already does.
 * ============================================================
 */
var PROJECTS_MAX_COUNT_ = 10;

function openProjectsSidebar_() {
  var html = HtmlService.createHtmlOutputFromFile('ProjectsSidebar')
    .setTitle('Manage Projects')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function MANAGE_PROJECTS() {
  openProjectsSidebar_();
}

// Dropdown source -- every existing project name.
function projects_getProjects() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var map     = getHeaderMap_(sheet);
  var nameCol = getCol_(map, ['Project_Name']);
  if (!nameCol) return [];

  var values = sheet.getRange(2, nameCol, sheet.getLastRow() - 1, 1).getValues();
  var names  = [];
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0]).trim();
    if (name) names.push(name);
  }
  return names;
}

// Prefill for a picked project.
function projects_getProjectDetails(projectName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Projects');
  if (!sheet) return { ok: false, message: 'Projects sheet not found.' };

  var map = getHeaderMap_(sheet);
  var row = findProjectRowByName_(sheet, map, projectName);
  if (!row) return { ok: false, message: 'Could not find that project. Refresh and try again.' };

  return {
    ok:          true,
    name:        getCellValue_(sheet, row, map, ['Project_Name']),
    description: getCellValue_(sheet, row, map, ['Description']),
    keywords:    getCellValue_(sheet, row, map, ['Keywords'])
  };
}

// Creates a new project (enforcing the 10-project Upwork cap) or updates an
// existing one (matched by originalName, blank for a new project). Either
// way, rebuilds Proposal_Generator's Portfolio_Project formula immediately
// afterward -- that formula is only built once at setup time otherwise, so
// without this a new/edited project wouldn't affect job matching until the
// next manual Repair Formulas run.
function projects_saveProject(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Projects');
  if (!sheet) return { ok: false, message: 'Projects sheet not found.' };

  var name = String(data.name || '').trim();
  if (!name) return { ok: false, message: 'Project name is required.' };

  var map     = getHeaderMap_(sheet);
  var nameCol = getCol_(map, ['Project_Name']);
  if (!nameCol) return { ok: false, message: 'Project_Name column not found.' };

  var originalName = String(data.originalName || '').trim();
  var row = originalName ? findProjectRowByName_(sheet, map, originalName) : null;

  if (!row) {
    var existingCount = 0;
    if (sheet.getLastRow() > 1) {
      var values = sheet.getRange(2, nameCol, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim()) existingCount++;
      }
    }
    if (existingCount >= PROJECTS_MAX_COUNT_) {
      return { ok: false, message: 'Upwork allows a maximum of ' + PROJECTS_MAX_COUNT_ + ' portfolio projects.' };
    }
    row = findFirstEmptyRowByColumn_(sheet, nameCol);
  }

  setCellValue_(sheet, row, map, ['Project_Name'], name);
  setCellValue_(sheet, row, map, ['Description'], data.description || '');
  setCellValue_(sheet, row, map, ['Keywords'], data.keywords || '');

  var pgSheet = ss.getSheetByName('Proposal_Generator');
  if (pgSheet && pgSheet.getLastColumn() > 0) {
    var pgHeaders = pgSheet.getRange(1, 1, 1, pgSheet.getLastColumn()).getValues()[0];
    applyProposalGeneratorFormulas_(pgSheet, pgHeaders);
  }

  return { ok: true };
}

function findProjectRowByName_(sheet, map, projectName) {
  var nameCol = getCol_(map, ['Project_Name']);
  if (!nameCol || sheet.getLastRow() < 2) return null;

  var values = sheet.getRange(2, nameCol, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(projectName).trim()) return i + 2;
  }
  return null;
}
