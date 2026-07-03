/**
 * ============================================================
 * 10. PROPOSAL GENERATOR
 *
 * The two proposal-writing prompt engines (FFLib.generateAIProposal,
 * FFLib.generateAiProposal) and the portfolio-context builder
 * (FFLib.getPortfolioContext) moved to the Apps Script Library --
 * they're the product's headline "secret sauce". Called from
 * 11_Job_Classifier.gs (RUN_AI_PROPOSALS) and 14_Edit_Trigger.gs.
 *
 * lookupProposalTemplate_ stays here: it's sheet I/O (reads
 * Proposal_Templates), not proprietary logic.
 * ============================================================
 */

function lookupProposalTemplate_(templateId, hookVersion, ctaVersion) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var tmplSheet = ss.getSheetByName('Proposal_Templates');
  if (!tmplSheet || tmplSheet.getLastRow() < 2) return null;

  var tmplMap     = getHeaderMap_(tmplSheet);
  var tmplIdCol   = getCol_(tmplMap, ['Template_ID']);
  var tmplHookCol = getCol_(tmplMap, ['Hook_Version']);
  var tmplCtaCol  = getCol_(tmplMap, ['CTA_Version']);
  var angleCol    = getCol_(tmplMap, ['Angle']);
  var credCol     = getCol_(tmplMap, ['Credential_Hint']);
  var toneCol     = getCol_(tmplMap, ['Tone']);
  var ctaStyleCol = getCol_(tmplMap, ['CTA_Style']);
  var exampleCol  = getCol_(tmplMap, ['Example_Output']);

  var tmplData = tmplSheet
    .getRange(2, 1, tmplSheet.getLastRow() - 1, tmplSheet.getLastColumn())
    .getValues();

  for (var i = 0; i < tmplData.length; i++) {
    var rowId   = tmplIdCol   ? String(tmplData[i][tmplIdCol   - 1]).trim() : '';
    var rowHook = tmplHookCol ? String(tmplData[i][tmplHookCol - 1]).trim() : '';
    var rowCta  = tmplCtaCol  ? String(tmplData[i][tmplCtaCol  - 1]).trim() : '';

    if (rowId === templateId && rowHook === hookVersion && rowCta === ctaVersion) {
      return {
        angle:          angleCol     ? String(tmplData[i][angleCol     - 1]).trim() : '',
        credentialHint: credCol      ? String(tmplData[i][credCol      - 1]).trim() : '',
        tone:           toneCol      ? String(tmplData[i][toneCol      - 1]).trim() : '',
        ctaStyle:       ctaStyleCol  ? String(tmplData[i][ctaStyleCol  - 1]).trim() : '',
        exampleOutput:  exampleCol   ? String(tmplData[i][exampleCol   - 1]).trim() : ''
      };
    }
  }
  return null;
}
