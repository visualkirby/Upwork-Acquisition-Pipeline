/**
 * ============================================================
 * FreelanceFlow Library -- AI Context
 * Journey-stage narrative logic, moved from the thin client's
 * 05_AI_Context.gs. Takes the already-fetched Settings object as
 * a parameter -- never touches SpreadsheetApp or PropertiesService.
 * ============================================================
 */
// Journey_Stage's 3 preset values (New/Growing/Established, set via a
// Settings dropdown -- see applySettingsValidation_, 00_Setup_Wizard.gs)
// explicitly pick one of these narrative tiers, replacing the original
// auto-detect-from-Contracts_Completed behavior -- a preset short label
// wasn't distinguishable from "not set" under the old design, since only a
// value longer than 20 characters was ever read as anything but blank.
var JOURNEY_STAGE_TIERS_ = {
  'New':
    'Because they have no reviews yet, winning depends heavily on proposal quality and job fit, ' +
    'not just bid position. They should be conservative with boost spending until they have at least 1-2 reviews. ' +
    'Lower-competition jobs where the proposal can stand out on merit are higher priority than aggressive outbidding.',
  'Growing':
    'They have some early contracts and are building their reputation. ' +
    'Moderate boost spending is acceptable on strong-fit jobs. ' +
    'Focus on maintaining high review scores and job completion rate.',
  'Established':
    'They have an established track record. ' +
    'Strategic boosting is appropriate on high-value jobs. ' +
    'Prioritize quality clients and long-term relationships over volume.'
};

function buildJourneyStage(settings) {
  var s = settings || {};

  var name      = s['Freelancer_Name']     || 'the freelancer';
  var contracts = s['Contracts_Completed'] || '0';
  var reviews   = s['Reviews_Count']       || '0';
  var score     = s['Job_Success_Score']   || '0';
  var tools     = s['Primary_Tools']       || 'data analytics tools';
  var stage     = String(s['Journey_Stage'] || '').trim();

  // A value longer than any preset label is a literal custom narrative
  // override, typed in directly instead of picked from the dropdown --
  // still supported for anyone who wants to write their own, same escape
  // hatch the original design had.
  if (stage.length > 20) return stage;

  var intro =
    'The freelancer is ' + name + ', a freelancer on Upwork. ' +
    'They currently have ' + contracts + ' completed contract(s) and ' + reviews + ' review(s). ' +
    'Their Job Success Score is ' + (score !== '0' ? score + '%.' : 'not yet established.') + ' ' +
    'Their core tools are: ' + tools + '. ';

  if (JOURNEY_STAGE_TIERS_[stage]) return intro + JOURNEY_STAGE_TIERS_[stage];

  // Blank or unrecognized (e.g. a copy that hasn't run Repair Formulas since
  // Journey_Stage became a preset dropdown) -- fall back to the original
  // auto-detect-from-Contracts_Completed behavior so nothing breaks mid-upgrade.
  if (contracts === '0') return intro + JOURNEY_STAGE_TIERS_['New'];
  if (parseInt(contracts) < 5) return intro + JOURNEY_STAGE_TIERS_['Growing'];
  return intro + JOURNEY_STAGE_TIERS_['Established'];
}
