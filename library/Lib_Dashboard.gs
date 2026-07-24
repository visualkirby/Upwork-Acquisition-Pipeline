/**
 * ============================================================
 * FreelanceFlow Library -- Dashboard
 * computeFunnelStages: pure stage-counting for the 4-stage pipeline funnel,
 *   shared by ANALYZE_JOB_WORKFLOW's text report and the Dashboard sheet's
 *   funnel section (via readFunnelStagesFromSheets_ in 03_Helpers.gs), so
 *   both always agree on the same numbers instead of computing it twice.
 * buildDashboardLayoutSpec: the Dashboard sheet's content structure --
 *   which Connects_Helper metrics to show as KPI tiles and their display
 *   labels, funnel stage labels, and section titles. This is the "layout"
 *   half of the build/apply split every formula-builder in this Library
 *   already uses (see Lib_WizardFormulas.gs) -- data/structure out, no
 *   SpreadsheetApp access. 25_Dashboard.gs's BUILD_DASHBOARD applies it.
 * ============================================================
 */
function computeFunnelStages(discoveryActions, scoringDecisions, proposalStatuses, hiredValues, interviewValues, viewedValues) {
  var discovery = { total: 0, toScoring: 0, reviewLater: 0, other: 0 };
  (discoveryActions || []).forEach(function (raw) {
    var action = String(raw).trim();
    if (!action) return;
    discovery.total++;
    if (action === "Move to Scoring") discovery.toScoring++;
    else if (action === "Review Later") discovery.reviewLater++;
    else discovery.other++;
  });

  var scoring = { total: 0, apply: 0, hold: 0, skip: 0, other: 0 };
  (scoringDecisions || []).forEach(function (raw) {
    var dec = String(raw).trim();
    if (!dec) return;
    scoring.total++;
    if (dec === "APPLY") scoring.apply++;
    else if (dec === "HOLD") scoring.hold++;
    else if (dec === "SKIP") scoring.skip++;
    else scoring.other++;
  });

  var proposals = { total: 0, sent: 0, skip: 0, ready: 0, other: 0 };
  (proposalStatuses || []).forEach(function (raw) {
    var st = String(raw).trim();
    if (!st) return;
    proposals.total++;
    if (st === "Sent") proposals.sent++;
    else if (st === "Skip") proposals.skip++;
    else if (st === "Ready") proposals.ready++;
    else proposals.other++;
  });

  var outcomes = { total: (hiredValues || []).length, hired: 0, interview: 0, viewed: 0, notViewed: 0 };
  (hiredValues || []).forEach(function (raw, i) {
    if (String(raw).trim() === "Y") outcomes.hired++;
    if (interviewValues && String(interviewValues[i]).trim() === "Y") outcomes.interview++;
    if (viewedValues && String(viewedValues[i]).trim() === "Y") outcomes.viewed++;
  });
  outcomes.notViewed = outcomes.total - outcomes.viewed;

  return { discovery: discovery, scoring: scoring, proposals: proposals, outcomes: outcomes };
}

function buildDashboardLayoutSpec() {
  return {
    title: "FreelanceFlow Dashboard",
    generatedLabel: "Last refreshed",
    kpiSection: {
      title: "Key Metrics (Month-to-Date)",
      metrics: [
        { key: "Current_Connect_Balance", label: "Connects Balance", format: "number" },
        { key: "MTD_Revenue",             label: "MTD Revenue",      format: "currency" },
        { key: "Monthly_Revenue",         label: "Monthly Revenue",  format: "currency" },
        { key: "MTD_Proposals_Sent",      label: "Proposals Sent",   format: "number" },
        { key: "MTD_Replies",             label: "Views",            format: "number" },
        { key: "MTD_Interviews",          label: "Interviews",       format: "number" },
        { key: "MTD_Hires",               label: "Hires",            format: "number" },
        { key: "Monthly_ROI",             label: "Monthly ROI",      format: "percent" }
      ]
    },
    // Derived rates -- none of these are stored in Connects_Helper, all
    // computed fresh each refresh from its MTD_* counters (see
    // computeDashboardRates_ below, pure function, no SpreadsheetApp).
    ratesSection: {
      title: "Core Performance Rates (Month-to-Date)",
      metrics: [
        { key: "replyRate",         label: "Proposal View Rate", format: "percent" },
        { key: "interviewRate",     label: "Interview Rate",      format: "percent" },
        { key: "hireRate",          label: "Hire Rate",           format: "percent" },
        { key: "jobsToProposalRate", label: "Jobs -> Proposal Rate", format: "percent" }
      ]
    },
    funnelSection: {
      title: "Pipeline Funnel",
      stages: [
        { key: "discovery", label: "Discovery", countKey: "total" },
        { key: "scoring",   label: "Scoring",   countKey: "total" },
        { key: "proposals", label: "Proposals", countKey: "sent"  },
        { key: "outcomes",  label: "Hired",     countKey: "hired" }
      ]
    },
    // All 6 metrics here are already stored Connects_Helper values -- this
    // section is a display grouping, not a new data dependency. Weekly/
    // Daily connect budget targets and a Connects_Status flag exist in the
    // older Upwork_Acquisition_System this was ported from, but they need a
    // user-set budget Setting that doesn't exist in FreelanceFlow yet --
    // deliberately left out rather than invented.
    connectsCostSection: {
      title: "Connects / Cost",
      metrics: [
        { key: "Current_Connect_Balance", label: "Current Connect Balance", format: "number" },
        { key: "Total_Connects_Used",     label: "Total Connects Used",     format: "number" },
        { key: "costPerConnect",          label: "Cost per Connect",        format: "currency" },
        { key: "Total_Proposal_Cost",     label: "Total Proposal Cost",     format: "currency" },
        { key: "Cost_per_Reply",          label: "Cost per View",           format: "currency" },
        { key: "Cost_per_Interview",      label: "Cost per Interview",      format: "currency" },
        { key: "Cost_per_Hire",           label: "Cost per Hire",           format: "currency" },
        { key: "Monthly_Cost",            label: "Monthly Cost",            format: "currency" },
        { key: "Monthly_ROI",             label: "Monthly ROI",             format: "percent" }
      ]
    },
    // Expected_Value_per_Proposal/Revenue_per_Connect/Net_Value_per_Connect
    // are stored Connects_Helper values; the per-Reply/Interview/Hire EVs
    // are derived fresh from MTD_Revenue / MTD_Replies|Interviews|Hires
    // (computeDashboardRates_), same reasoning as the rates section above.
    expectedValueSection: {
      title: "Expected Value",
      metrics: [
        { key: "Expected_Value_per_Proposal", label: "EV per Proposal",  format: "currency" },
        { key: "evPerReply",                  label: "EV per View",      format: "currency" },
        { key: "evPerInterview",              label: "EV per Interview", format: "currency" },
        { key: "evPerHire",                   label: "EV per Hire",      format: "currency" },
        { key: "Revenue_per_Connect",         label: "Revenue per Connect", format: "currency" },
        { key: "Net_Value_per_Connect",       label: "Net Value per Connect", format: "currency" }
      ]
    },
    topPerformersSection: {
      title: "Top Performers"
    },
    keywordIntelligenceSection: {
      title: "Keyword Priority Score",
      scatterTitle: "Priority Score vs. Competition, by Status",
      maxKeywords: 12
    },
    toolMarketShareSection: {
      title: "Tool Market Share (Job_Discovery)",
      maxTools: 3
    },
    weeklyMetricsSection: {
      title: "Weekly Metrics (Last 7 Days)",
      metrics: [
        { key: "jobsFoundWeek",      label: "Jobs Found (Last 7 Days)" },
        { key: "proposalsSentWeek",  label: "Proposals Sent (Last 7 Days)" },
        { key: "repliesWeek",        label: "Views (Last 7 Days)" },
        { key: "interviewsWeek",     label: "Interviews (Last 7 Days)" },
        { key: "hiresWeek",          label: "Hires (Last 7 Days)" }
      ]
    },
    revenueTrendSection: {
      title: "Revenue Trend"
    }
  };
}

// Rates/EVs not stored anywhere -- computed fresh from Connects_Helper's
// MTD_* counters every refresh. Pure function (metrics in, numbers out), no
// SpreadsheetApp access, matching this Library's build/apply split. Guards
// every division since an early-stage copy will have mostly-zero counters.
function computeDashboardRates(metrics) {
  var div = function (numerator, denominator) {
    var d = Number(denominator) || 0;
    return d > 0 ? (Number(numerator) || 0) / d : 0;
  };

  var proposalsSent = Number(metrics['MTD_Proposals_Sent']) || 0;
  var jobsLogged    = Number(metrics['MTD_Jobs_Logged']) || 0;
  var replies       = Number(metrics['MTD_Replies']) || 0;
  var interviews    = Number(metrics['MTD_Interviews']) || 0;
  var hires         = Number(metrics['MTD_Hires']) || 0;
  var revenue       = Number(metrics['MTD_Revenue']) || 0;
  var totalCost     = Number(metrics['Total_Proposal_Cost']) || 0;
  var connectsUsed  = Number(metrics['Total_Connects_Used']) || 0;

  return {
    replyRate:          div(replies, proposalsSent),
    interviewRate:       div(interviews, proposalsSent),
    hireRate:            div(hires, proposalsSent),
    jobsToProposalRate:  div(proposalsSent, jobsLogged),
    costPerConnect:      div(totalCost, connectsUsed),
    evPerReply:          div(revenue, replies),
    evPerInterview:      div(revenue, interviews),
    evPerHire:           div(revenue, hires)
  };
}

// Validated palette subset (see the dataviz skill's references/palette.md)
// applied to Dashboard charts: a single mid-tone blue for single-series
// magnitude bars/rankings (Keyword Priority Score, Tool Market Share -- both
// are a "which is biggest" question, not a category-identity question), and
// the reserved status triad for Scale/Test More/Drop on the priority-vs-
// competition scatter (green/amber/red -- fixed meaning, never reused for a
// plain series).
function getDashboardColors() {
  return {
    singleSeries: '#256abf',
    status: { 'Scale': '#0ca30c', 'Test More': '#fab219', 'Drop': '#d03b3b' }
  };
}
