/**
 * ============================================================
 * FreelanceFlow Library -- Wizard Formula Builders
 * Pure formula-string builders -- headers/portfolioMap in,
 * formula-string out, no SpreadsheetApp access. The thin client
 * (00_Setup_Wizard.gs's apply*_ functions at setup time, and
 * 15_Formula_Fixes.gs's REPAIR_FORMULAS() for an existing sheet)
 * calls these and writes the returned formula strings via
 * setFormula(). This is the actual "how do we build a working
 * pipeline automatically" mechanism -- the highest-value piece
 * of IP to keep out of customers' hands.
 *
 * buildJobScoringFormulas moved to Lib_JobScoringFormulas.gs (the
 * full 9-score chain). buildProposalGeneratorFormulas is the other
 * entry point the client calls, so it's public (no trailing
 * underscore). buildPortfolioFormula_ and colLetter_ are only ever
 * called internally, so they keep the underscore and stay private
 * to the Library.
 * ============================================================
 */
function buildProposalGeneratorFormulas(headers, portfolioMap, primaryToolsCsv) {
  var jobTitleIdx  = headers.indexOf('Job_Title');
  var descIdx      = headers.indexOf('Description');
  var toolDetIdx   = headers.indexOf('Tool_Detected');
  var portfolioIdx = headers.indexOf('Portfolio_Project');

  var result = {};
  if (jobTitleIdx < 0 || descIdx < 0) return result;

  var jtL = colLetter_(jobTitleIdx + 1);
  var dcL = colLetter_(descIdx + 1);

  // Tool_Detected -- dynamic, one clause per tool the user listed in Settings.
  // Same precedent as Job_Discovery/Job_Scoring's Tool_Detected: built fresh
  // per user instead of hardcoded to a fixed BI-tool list.
  if (toolDetIdx >= 0) {
    var tools = (primaryToolsCsv || '').split(',')
      .map(function (t) { return t.trim(); })
      .filter(function (t) { return t; });

    if (tools.length > 0) {
      var ifsArgs = tools.map(function (t) {
        var safe = t.replace(/"/g, '""');
        return 'ISNUMBER(SEARCH("' + safe.toLowerCase() + '",LOWER(' + dcL + '2))),"' + safe + '"';
      });
      result.toolDetectedCol = toolDetIdx + 1;
      result.toolDetectedFormula =
        '=IF(' + jtL + '2="","",IFS(' + ifsArgs.join(',') + ',TRUE,"Other"))';
    }
  }

  if (portfolioIdx >= 0) {
    result.portfolioProjectCol     = portfolioIdx + 1;
    result.portfolioProjectFormula = buildPortfolioFormula_(jtL, dcL, portfolioMap);
  }

  return result;
}


function buildPortfolioFormula_(jobTitleLetter, descLetter, portfolioMap) {
  var jtL = jobTitleLetter || 'B';
  var dcL = descLetter     || 'D';
  var map = portfolioMap   || {};

  var clauses     = [];
  var defaultName = '';
  var pNums       = Object.keys(map).sort();

  if (pNums.length === 0) {
    return '=IF(' + jtL + '2="","","Add portfolio projects via FreelanceFlow Setup")';
  }

  for (var p = 0; p < pNums.length; p++) {
    var proj = map[pNums[p]];
    if (!proj.name) continue;
    if (!defaultName) defaultName = proj.name;
    var safeName = proj.name.replace(/"/g, '""');
    for (var k = 0; k < proj.keywords.length; k++) {
      var kw = proj.keywords[k].replace(/"/g, '""');
      if (kw) clauses.push('ISNUMBER(SEARCH("' + kw + '",' + dcL + '2)),"' + safeName + '"');
    }
  }

  if (clauses.length === 0) {
    return '=IF(' + jtL + '2="","","Add portfolio keyword mappings in Settings sheet")';
  }

  var safeDefault = (defaultName || 'Portfolio Project').replace(/"/g, '""');
  return '=IF(' + jtL + '2="","",IFS(' + clauses.join(',') + ',TRUE,"' + safeDefault + '"))';
}


/**
 * Experience_Score -- relative to the freelancer's own Upwork profile level
 * (Settings!Freelancer_Experience_Level), not a fixed universal preference for
 * "Entry Level" jobs. Exact tier match scores highest; each tier of distance
 * (Entry Level / Intermediate / Expert) penalizes further. Shared by
 * Lib_JobDiscoveryFormulas and Lib_JobScoringFormulas -- same job-side
 * Experience_Level dropdown, same freelancer-side Settings lookup.
 */
function buildExperienceScoreFormula_(expL) {
  var tiers   = '{"Entry Level","Intermediate","Expert"}';
  var jobTier = 'MATCH(' + expL + '2,' + tiers + ',0)';
  var meTier  = 'MATCH(VLOOKUP("Freelancer_Experience_Level",Settings!$A:$B,2,0),' + tiers + ',0)';
  var dist    = 'ABS(' + jobTier + '-' + meTier + ')';
  return '=IF(' + expL + '2="","",IFS(' + dist + '=0,1,' + dist + '=1,0.7,TRUE,0.4))';
}

function colLetter_(n) {
  var s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

/**
 * Proposal_Generator's raw-data columns (Job_Title, Description, Budget,
 * etc.) auto-populate from Job_Scoring wherever Final_Decision="APPLY" --
 * same FILTER-pull precedent as buildJobScoringPullFormulas, just one
 * pipeline stage further along. Tool_Detected/Portfolio_Project are excluded
 * here on purpose -- they're computed fresh in this sheet by
 * buildProposalGeneratorFormulas above, not pulled. Job_Type and the
 * template/bid fields are excluded too -- those are filled later by the
 * RUN_JOB_CLASSIFICATION / RUN_AI_PROPOSALS batch actions, not by this pull.
 * Additional_Questions is excluded too -- it flows the OPPOSITE direction
 * from everything else here (see buildDiscoveryIdLookup_ below): it's only
 * knowable at this stage (Upwork only shows a job's extra application
 * questions on the submission page), so it's a manual field in this sheet,
 * looked up backwards into Job_Scoring/Job_Discovery instead of pulled
 * forward into it.
 */
function buildProposalGeneratorPullFormulas(jobScoringHeaders, proposalGeneratorHeaders) {
  var fieldPairs = [
    ['Discovery_ID', 'Discovery_ID'],
    ['Date', 'Proposal_Generator_Date'],
    ['Job_Title', 'Job_Title'],
    ['Client_Name', 'Client_Name'],
    ['Description', 'Description'],
    ['Job_Link', 'Job_Link'],
    ['Keyword_Search', 'Keyword_Search'],
    ['Connects_Required', 'Connects_Required'],
    ['Proposal_Count', 'Proposal_Count'],
    ['Budget', 'Budget']
  ];

  return buildFilterPull_(
    fieldPairs, 'Job_Scoring', jobScoringHeaders, proposalGeneratorHeaders,
    'Final_Decision', 'APPLY'
  );
}

/**
 * Additional_Questions travels backwards through the pipeline: it's a manual
 * field in Proposal_Generator (only knowable once the freelancer is on
 * Upwork's submission page), looked up into Job_Scoring and Job_Discovery so
 * the earlier stages can still show it for reference. A VLOOKUP keyed on
 * Discovery_ID, not a FILTER -- this is a one-row-to-one-row match, not a
 * filtered subset, and Discovery_ID is guaranteed column A in all three
 * sheets. IFERROR covers rows with no matching Discovery_ID yet (job hasn't
 * reached that stage) or no value entered.
 */
function buildDiscoveryIdLookup_(ownDiscoveryIdCol, sourceSheetName, sourceHeaders, sourceValueHeader) {
  var idIdx  = sourceHeaders.indexOf('Discovery_ID');
  var valIdx = sourceHeaders.indexOf(sourceValueHeader);
  if (idIdx < 0 || valIdx < 0 || ownDiscoveryIdCol <= 0) return null;

  var ownIdL   = colLetter_(ownDiscoveryIdCol);
  var lastColL = colLetter_(sourceHeaders.length);
  var offset   = valIdx - idIdx + 1;

  return '=IF(' + ownIdL + '2="","",IFERROR(VLOOKUP(' + ownIdL + '2,' +
    sourceSheetName + '!$A:$' + lastColL + ',' + offset + ',FALSE),""))';
}

function buildJobScoringAdditionalQuestionsLookup(proposalGeneratorHeaders, jobScoringHeaders) {
  var ownCol = jobScoringHeaders.indexOf('Discovery_ID') + 1;
  var aqCol  = jobScoringHeaders.indexOf('Additional_Questions') + 1;
  if (ownCol <= 0 || aqCol <= 0) return null;

  var formula = buildDiscoveryIdLookup_(ownCol, 'Proposal_Generator', proposalGeneratorHeaders, 'Additional_Questions');
  return formula ? { col: aqCol, formula: formula } : null;
}

function buildJobDiscoveryAdditionalQuestionsLookup(jobScoringHeaders, jobDiscoveryHeaders) {
  var ownCol = jobDiscoveryHeaders.indexOf('Discovery_ID') + 1;
  var aqCol  = jobDiscoveryHeaders.indexOf('Additional_Questions') + 1;
  if (ownCol <= 0 || aqCol <= 0) return null;

  var formula = buildDiscoveryIdLookup_(ownCol, 'Job_Scoring', jobScoringHeaders, 'Additional_Questions');
  return formula ? { col: aqCol, formula: formula } : null;
}

/**
 * Generic FILTER-pull builder shared by buildJobScoringPullFormulas and
 * buildProposalGeneratorPullFormulas. fieldPairs is [targetHeader,
 * sourceHeader] tuples in target-sheet column order. Contiguous runs of
 * target columns become one combined-array FILTER (matches the proven
 * sheet's own pattern of several separate FILTER calls rather than one
 * giant array spanning the whole row) -- keeps each formula's column
 * references dynamic instead of hardcoded, so a header reorder can't
 * silently desync it the way the original hand-typed FILTER did.
 */
function buildFilterPull_(fieldPairs, sourceSheetName, sourceHeaders, targetHeaders, conditionHeader, conditionValue) {
  var condIdx = sourceHeaders.indexOf(conditionHeader);
  if (condIdx < 0) return [];
  var condL = colLetter_(condIdx + 1);

  var entries = fieldPairs
    .map(function (pair) {
      var targetIdx = targetHeaders.indexOf(pair[0]);
      var sourceIdx = sourceHeaders.indexOf(pair[1]);
      if (targetIdx < 0 || sourceIdx < 0) return null;
      return { targetCol: targetIdx + 1, sourceL: colLetter_(sourceIdx + 1) };
    })
    .filter(function (e) { return e; })
    .sort(function (a, b) { return a.targetCol - b.targetCol; });

  var groups = [];
  entries.forEach(function (e) {
    var current = groups.length > 0 ? groups[groups.length - 1] : null;
    if (current && e.targetCol === current.entries[current.entries.length - 1].targetCol + 1) {
      current.entries.push(e);
    } else {
      groups.push({ entries: [e] });
    }
  });

  var conditionRef = sourceSheetName + '!' + condL + '2:' + condL + '="' + conditionValue + '"';

  return groups.map(function (g) {
    var refs = g.entries.map(function (e) {
      return sourceSheetName + '!' + e.sourceL + '2:' + e.sourceL;
    });
    var formula = refs.length === 1
      ? '=FILTER(' + refs[0] + ',' + conditionRef + ')'
      : '=FILTER({' + refs.join(',') + '},' + conditionRef + ')';
    return { col: g.entries[0].targetCol, width: g.entries.length, formula: formula };
  });
}
