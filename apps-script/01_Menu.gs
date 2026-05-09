/**
 * ============================================================
 * 1. MENU
 * ============================================================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("System Tools")
    .addItem("Reset System",           "RESET_SYSTEM")
    .addItem("Mine Keywords",          "MINE_KEYWORDS")
    .addSeparator()
    .addItem("Start Session",          "START_SESSION")
    .addItem("End Session",            "END_SESSION")
    .addSeparator()
    .addItem("Analyze Job Workflow",   "ANALYZE_JOB_WORKFLOW")
    .addItem("Run Job Classification", "RUN_JOB_CLASSIFICATION")
    .addItem("Run AI Proposals",       "RUN_AI_PROPOSALS")
    .addSeparator()
    .addItem("Snapshot Month End",     "SNAPSHOT_MONTH_END")
    .addSeparator()
    .addItem("Apply Formula Fixes",    "APPLY_FORMULA_FIXES")
    .addSeparator()
    .addItem("Setup API Key",          "SETUP_API_KEY")
    .addToUi();
}
