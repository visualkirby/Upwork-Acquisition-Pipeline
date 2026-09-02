/**
 * ============================================================
 * FreelanceFlow Library -- Wizard Formula Builders
 * Pure formula-string builders -- headers in, formula-string out,
 * no SpreadsheetApp access. The thin client (00_Setup_Wizard.gs's
 * apply*_ functions at setup time, and 15_Formula_Fixes.gs's
 * REPAIR_FORMULAS() for an existing sheet) calls these and writes
 * the returned formula strings via setFormula(). This is the actual
 * "how do we build a working pipeline automatically" mechanism --
 * the highest-value piece of IP to keep out of customers' hands.
 *
 * buildJobScoringFormulas moved to Lib_JobScoringFormulas.gs (the
 * full 9-score chain). buildProposalGeneratorFormulas and
 * pickPortfolioProject are the other entry points the client calls,
 * so they're public (no trailing underscore). buildToolDetectedIfsArgs_
 * and colLetter_ are only ever called internally, so they keep the
 * underscore and stay private to the Library.
 * ============================================================
 */
// Builds the IFS clause list for a Tool_Detected formula -- shared by
// Job_Discovery (buildJobDiscoveryFormulas, Lib_JobDiscoveryFormulas.gs) and
// Proposal_Generator (buildProposalGeneratorFormulas below), so the two
// sheets can't drift into different matching behavior for the same tool
// list again.
//
// Two things fixed here vs. the earlier per-sheet versions: (1) every tool
// name is word-boundary-anchored (\b...\b), so "Excel" can't match inside
// "excellent" and "SQL" can't match inside "MySQL"/"NoSQL" -- a real job in
// testing ("Implementation Analyst") was wrongly tagged Excel purely because
// its description used the word "excellent" four times; (2) Job_Title
// clauses are listed before Description clauses, so a tool explicitly named
// in the title outranks one only mentioned in passing in the body -- another
// real job ("...Advanced Excel & Interactive Dashboard Expert") was wrongly
// tagged Tableau because the description had one aside about "tools like
// Tableau" while Excel, the actual ask, was only checked against the body
// text and lost to Tableau's earlier position in Primary_Tools.
//
// A tool that only appears as part of a compound word (MySQL, PostgreSQL)
// won't match under this stricter check unless that exact variant is also
// listed in Primary_Tools -- deliberate, since there's no regex-only way to
// allow "MySQL" without also reopening the "NoSQL" false-positive it's meant
// to close. Users who want a variant tracked list it themselves (Setup
// Wizard's Primary_Tools field has a hint for this now).
//
// Returns null if primaryToolsCsv has no usable tools (caller leaves
// Tool_Detected formula-free, same as before).
function buildToolDetectedIfsArgs_(titleL, descL, primaryToolsCsv) {
  var tools = (primaryToolsCsv || '').split(',')
    .map(function (t) { return t.trim(); })
    .filter(function (t) { return t; });
  if (tools.length === 0) return null;

  function regexEscape(t) {
    return t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  function quoteEscape(s) {
    return s.replace(/"/g, '""');
  }

  function clausesFor(colL) {
    if (!colL) return [];
    return tools.map(function (t) {
      var pattern = '\\b' + regexEscape(t) + '\\b';
      return 'REGEXMATCH(LOWER(' + colL + '2),LOWER("' + quoteEscape(pattern) + '")),"' + quoteEscape(t) + '"';
    });
  }

  return clausesFor(titleL).concat(clausesFor(descL));
}

function buildProposalGeneratorFormulas(headers, primaryToolsCsv) {
  var jobTitleIdx  = headers.indexOf('Job_Title');
  var descIdx      = headers.indexOf('Description');
  var toolDetIdx   = headers.indexOf('Tool_Detected');

  var result = {};
  if (jobTitleIdx < 0 || descIdx < 0) return result;

  var jtL = colLetter_(jobTitleIdx + 1);
  var dcL = colLetter_(descIdx + 1);

  // Tool_Detected -- dynamic, one clause per tool the user listed in Settings.
  // Same precedent as Job_Discovery/Job_Scoring's Tool_Detected: built fresh
  // per user instead of hardcoded to a fixed BI-tool list. See
  // buildToolDetectedIfsArgs_ above for the shared title-priority,
  // word-boundary matching logic.
  if (toolDetIdx >= 0) {
    var toolArgs = buildToolDetectedIfsArgs_(jtL, dcL, primaryToolsCsv);
    if (toolArgs) {
      result.toolDetectedCol = toolDetIdx + 1;
      result.toolDetectedFormula =
        '=IF(' + jtL + '2="","",IFS(' + toolArgs.join(',') + ',TRUE,"Other"))';
    }
  }

  // Portfolio_Project is no longer built here -- a formula can only express
  // "first matching keyword wins", and a typical portfolio's keywords overlap
  // too much for that to mean anything (a real test found "Dashboard" listed
  // on 5 of 7 projects, so whichever project happened to be listed first won
  // 10 of 13 real jobs regardless of actual fit). A later keyword-count-scoring
  // version fixed the overlap problem but still put the matching burden on
  // customers correctly anticipating job phrasing in a hand-authored keyword
  // list. See pickPortfolioProject below -- now an AI classification off each
  // project's own Name + Description instead, since FreelanceFlow customers
  // name/describe their portfolio projects however they want and no fixed
  // keyword scheme can generalize across that. Called from the thin client
  // (RUN_JOB_CLASSIFICATION, 11_Job_Classifier.gs) instead of live in-sheet,
  // since Proposal_Generator's Job_Title/Description arrive via a FILTER
  // pull that never fires an edit event to recompute a formula against.

  return result;
}

// AI-classifies which portfolio project best matches a job, off each
// project's own Project_Name + Description -- the same information a
// customer already has to write for their own Upwork portfolio, so there's
// no separate keyword-authoring step to skip or get wrong. Replaces the
// earlier keyword-regex version (scored projects by counting word-boundary
// keyword hits), which put the matching burden on the customer correctly
// anticipating job phrasing in a hand-maintained Keywords field -- FreelanceFlow
// customers name and describe their own projects however they want, so a
// fixed keyword list can't generalize across them the way an AI read of the
// actual description can. Mirrors FFLib.getJobType's classify-into-one-of-N
// pattern (Lib_JobClassifier.gs): constrained prompt, reply-with-name-only,
// substring match back to a known project name, fallback to the first
// project on any failure (no API key, parse error, no match in the reply) --
// same "always attribute something" default the old version used, just a
// different mechanism for getting there.
//
// Skips the AI call entirely for 0 or 1 real project (nothing to choose
// between), so no cost is added over the old version in the common
// single-project-portfolio case.
function pickPortfolioProject(jobTitle, description, portfolioMap, apiKey) {
  var map = portfolioMap || {};
  var pNums = Object.keys(map).map(Number).sort(function (a, b) { return a - b; });
  var projects = pNums.map(function (n) { return map[n]; }).filter(function (p) { return p && p.name; });

  if (projects.length === 0) {
    return 'Add portfolio projects via FreelanceFlow Setup';
  }

  function fallback_() {
    return projects[0].name;
  }

  if (projects.length === 1 || !apiKey) return fallback_();

  var projectText = projects.map(function (p) {
    return p.name + (p.description ? (' (' + p.description + ')') : '');
  }).join('; ');

  var prompt =
    "Pick exactly one portfolio project that best matches this Upwork job, based on the project " +
    "descriptions below. Reply with only the project name and nothing else.\n\n" +
    "Portfolio projects: " + projectText + "\n\n" +
    "Job: " + jobTitle + ". " + String(description || '').substring(0, 800);

  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + apiKey },
      payload: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        max_tokens: 20,
        temperature: 0
      }),
      muteHttpExceptions: true
    });

    var parsed = JSON.parse(response.getContentText());
    if (!parsed.choices || !parsed.choices[0]) return fallback_();

    var text = parsed.choices[0].message.content.trim();
    for (var i = 0; i < projects.length; i++) {
      if (text.indexOf(projects[i].name) !== -1) return projects[i].name;
    }
    return fallback_();
  } catch (err) {
    return fallback_();
  }
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
 * pipeline stage further along. Tool_Detected is excluded here on purpose --
 * it's computed fresh in this sheet by buildProposalGeneratorFormulas above,
 * not pulled. Job_Type, Portfolio_Project, and the template/bid fields are
 * excluded too -- those are filled later by the RUN_JOB_CLASSIFICATION /
 * RUN_AI_PROPOSALS batch actions, not by this pull.
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
