/**
 * ============================================================
 * FreelanceFlow Library -- AI Context
 * Journey-stage narrative logic, moved from the thin client's
 * 05_AI_Context.gs. Takes the already-fetched Settings object as
 * a parameter -- never touches SpreadsheetApp or PropertiesService.
 * ============================================================
 */
function buildJourneyStage(settings) {
  var s = settings || {};

  var name      = s['Freelancer_Name']     || 'the freelancer';
  var contracts = s['Contracts_Completed'] || '0';
  var reviews   = s['Reviews_Count']       || '0';
  var score     = s['Job_Success_Score']   || '0';
  var tools     = s['Primary_Tools']       || 'data analytics tools';
  var stage     = s['Journey_Stage']       || '';

  if (stage && stage.length > 20) return stage;

  var context =
    'The freelancer is ' + name + ', a freelancer on Upwork. ' +
    'They currently have ' + contracts + ' completed contract(s) and ' + reviews + ' review(s). ' +
    'Their Job Success Score is ' + (score !== '0' ? score + '%.' : 'not yet established.') + ' ' +
    'Their core tools are: ' + tools + '. ';

  if (contracts === '0') {
    context +=
      'Because they have no reviews yet, winning depends heavily on proposal quality and job fit, ' +
      'not just bid position. They should be conservative with boost spending until they have at least 1-2 reviews. ' +
      'Lower-competition jobs where the proposal can stand out on merit are higher priority than aggressive outbidding.';
  } else if (parseInt(contracts) < 5) {
    context +=
      'They have some early contracts and are building their reputation. ' +
      'Moderate boost spending is acceptable on strong-fit jobs. ' +
      'Focus on maintaining high review scores and job completion rate.';
  } else {
    context +=
      'They have an established track record. ' +
      'Strategic boosting is appropriate on high-value jobs. ' +
      'Prioritize quality clients and long-term relationships over volume.';
  }

  return context;
}
