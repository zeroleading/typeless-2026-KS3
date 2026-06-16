/**
 * Controller.gs
 * Handles the user interface, custom menus, and authorisation routing.
 */

/**
 * Triggered automatically when the spreadsheet is opened.
 * @param {Object} e The event object.
 */
function onOpen(e) {
  buildDynamicMenu();
}

/**
 * Builds the custom menu based on the active user's permissions defined in Config.gs.
 */
function buildDynamicMenu() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();
  const menu = ui.createMenu('Typeless Reports');

  if (!email) {
    menu.addItem('Authorise Script', 'authoriseScript').addToUi();
    return;
  }

  const isSuperUser = CONFIG.AUTH.SUPER_USERS.includes(email);

  let menuHasItems = false;

  if (isSuperUser) {
    menu.addItem('Setup Subject Sheets', 'triggerSetup');
    menu.addItem('Freeze Import Data', 'triggerFreeze');
    menu.addItem('Thaw Import Data', 'triggerThaw');
    menu.addSeparator(); 
    menu.addItem('Run Progress Review', 'triggerProgressReview');
    menu.addItem('Run Next Steps Summary', 'triggerNextStepsSummary');
    menuHasItems = true;
  }

  if (menuHasItems) {
    menu.addToUi();
  }
}

function authoriseScript() {
  SpreadsheetApp.getUi().alert('Authorisation complete. Please refresh the page to see your custom menu.');
}

// --- Trigger Functions ---

function triggerSetup() {
  Setup.triggerCreateSubjectSheets();
}

function triggerFreeze() {
  Setup.freezeImportSheet();
}

function triggerThaw() {
  Setup.thawImportSheet();
}

function triggerProgressReview() {
  _runReportBatch(CONFIG.REPORTS.PROGRESS_REVIEW, 'Progress Reviews');
}

function triggerNextStepsSummary() {
  _runReportBatch(CONFIG.REPORTS.NEXT_STEPS_SUMMARY, 'Next Steps Summaries');
}

/**
 * Shared execution logic for all report types.
 * @private
 */
function _runReportBatch(reportConfig, reportFriendlyName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Pre-run validation: Check if the import sheet is safely frozen
  // The Why: We do not want to generate reports from data that is actively shifting/recalculating.
  const importSheet = ss.getSheetByName('import'); 
  if (importSheet) {
    const status = importSheet.getRange('A1').getValue();
    if (status !== '🥶') {
      ui.alert(
        'Validation Error', 
        'The import sheet must be frozen (🥶) before generating reports. Please use the menu: Typeless Reports > Freeze Import Data.', 
        ui.ButtonSet.OK
      );
      return;
    }
  }

  const response = ui.alert('Confirm', `Generate ${reportFriendlyName}?`, ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    ss.toast(`Gathering data for ${reportFriendlyName}...`, 'Typeless');
    
    // 1. Build the data payload
    const payload = DataService.buildStudentDataPayload(reportConfig);
    
    if (payload.length === 0) {
      ui.alert('Error', 'No student data found. Please check the master list and subject sheets.', ui.ButtonSet.OK);
      return;
    }

    ss.toast(`Generating documents for ${payload.length} students...`, 'Typeless');
    
    // 2. Generate the documents
    const folderId = DocumentBuilder.generateBatch(reportConfig, payload);
    
    // 3. Alert completion
    ui.alert('Merge Complete', `Documents generated successfully.\nFolder ID: ${folderId}`, ui.ButtonSet.OK);
  }
}