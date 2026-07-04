/**
 * ============================================================
 * 1. MENU
 * onOpen is defined in 00_Setup_Wizard.gs -- do not redefine.
 * ============================================================
 */
function buildSystemMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('System Tools')
    .addItem('Reset System',             'RESET_SYSTEM')
    .addItem('Generate Keyword Strategy', 'GENERATE_KEYWORD_STRATEGY')
    .addItem('Mine Keywords',            'MINE_KEYWORDS')
    .addSeparator()
    .addItem('Start Session',            'START_SESSION')
    .addItem('End Session',              'END_SESSION')
    .addItem('Log New Job',              'LOG_NEW_JOB')
    .addSeparator()
    .addItem('Analyze Job Workflow',      'ANALYZE_JOB_WORKFLOW')
    .addItem('Analyze Session Patterns',  'ANALYZE_SESSION_PATTERNS')
    .addItem('Analyze Bid Patterns',      'ANALYZE_BID_PATTERNS')
    .addItem('Run Job Classification',    'RUN_JOB_CLASSIFICATION')
    .addItem('Run AI Proposals',          'RUN_AI_PROPOSALS')
    .addSeparator()
    .addItem('Import Client Chat',        'IMPORT_CLIENT_CHAT')
    .addItem('Log New Contract',          'LOG_NEW_CONTRACT')
    .addSeparator()
    .addItem('Snapshot Month End',        'SNAPSHOT_MONTH_END')
    .addSeparator()
    .addItem('Repair Formulas',           'REPAIR_FORMULAS')
    .addItem('Send Diagnostic Report',    'SEND_DIAGNOSTIC_REPORT')
    .addSeparator()
    .addItem('Setup API Key',             'SETUP_API_KEY')
    .addItem('Check API Key',             'CHECK_API_KEY')
    .addSeparator()
    .addItem('FreelanceFlow Setup',       'OPEN_SETUP_WIZARD')
    .addToUi();
}
