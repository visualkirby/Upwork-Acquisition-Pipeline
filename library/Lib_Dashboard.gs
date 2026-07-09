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
        { key: "MTD_Replies",             label: "Replies",          format: "number" },
        { key: "MTD_Interviews",          label: "Interviews",       format: "number" },
        { key: "MTD_Hires",               label: "Hires",            format: "number" },
        { key: "Monthly_ROI",             label: "Monthly ROI",      format: "percent" }
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
    revenueTrendSection: {
      title: "Revenue Trend"
    }
  };
}
