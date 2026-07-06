/**
 * ============================================================
 * 1. MENU
 * onOpen is defined in 00_Setup_Wizard.gs -- do not redefine.
 * ============================================================
 */
function buildSystemMenu_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('System Tools')
    .addItem('FreelanceFlow Setup',       'OPEN_SETUP_WIZARD')
    .addSeparator()
    .addItem('Start Session',            'START_SESSION')
    .addItem('End Session',              'END_SESSION')
    .addSeparator()
    .addItem('Log New Job',              'LOG_NEW_JOB')
    .addItem('Log Proposal Bid',         'LOG_PROPOSAL_BID')
    .addItem('Import Client Chat',        'IMPORT_CLIENT_CHAT')
    .addItem('Log New Contract',          'LOG_NEW_CONTRACT')
    .addSeparator()
    .addItem('Run Job Classification',    'RUN_JOB_CLASSIFICATION')
    .addItem('Run AI Proposals',          'RUN_AI_PROPOSALS')
    .addItem('Analyze Job Workflow',      'ANALYZE_JOB_WORKFLOW')
    .addItem('Analyze Session Patterns',  'ANALYZE_SESSION_PATTERNS')
    .addItem('Analyze Bid Patterns',      'ANALYZE_BID_PATTERNS')
    .addSeparator()
    .addItem('Generate Keyword Strategy', 'GENERATE_KEYWORD_STRATEGY')
    .addItem('Mine Keywords',            'MINE_KEYWORDS')
    .addItem('Drop Keywords',            'DROP_KEYWORDS')
    .addSeparator()
    .addItem('Snapshot Month End',        'SNAPSHOT_MONTH_END')
    .addSeparator()
    .addItem('Setup API Key',             'SETUP_API_KEY')
    .addItem('Check API Key',             'CHECK_API_KEY')
    .addSeparator()
    .addItem('Send Diagnostic Report',    'SEND_DIAGNOSTIC_REPORT')
    .addItem('Repair Formulas',           'REPAIR_FORMULAS')
    .addSeparator()
    .addSubMenu(ui.createMenu('Reset System')
      .addItem('Reset to Right After Setup',             'RESET_TO_AFTER_SETUP')
      .addItem('Reset to Before Setup Wizard (Factory)',  'RESET_TO_BEFORE_SETUP'))
    .addToUi();
}
