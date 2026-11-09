/**
 * @fileoverview Controller.gs (KS3)
 * Handles user interface lifecycles, dynamic menu rendering based on user roles,
 * modal dialog triggers, and RPC endpoints for chunked document batch processing.
 */

/**
 * Standard trigger executed whenever the spreadsheet is opened.
 * Builds role-tailored menus based on the active user's credentials.
 * @param {Object} e The open event object.
 */
function onOpen(e) {
  buildDynamicMenu();
}

/**
 * Constructs the custom Google Sheets menu.
 * Verifies super-user privileges before exposing administrative tools.
 */
function buildDynamicMenu() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();
  const menu = ui.createMenu('Typeless Reports');

  // If running in an unauthenticated context, prompt for initial authorisation
  if (!email) {
    menu.addItem('Authorise Script', 'authoriseScript').addToUi();
    return;
  }

  const isSuperUser = CONFIG.PERMISSIONS.SUPER_USERS.includes(email);
  let menuHasItems = false;

  if (isSuperUser) {
    menu.addItem('Setup Subject Sheets', 'triggerSetup');
    menu.addItem('Freeze Import Data', 'triggerFreeze');
    menu.addItem('Thaw Import Data', 'triggerThaw');
    menu.addItem('Create Copy with Permissions', 'createCopyWithPermissions');
    menu.addSeparator(); 
    menu.addItem('Run Progress Review', 'triggerProgressReview');
    menu.addItem('Run Next Steps Summary', 'triggerNextStepsSummary');
    menuHasItems = true;
  }

  if (menuHasItems) {
    menu.addToUi();
  }
}

/**
 * Feedback handler called when an unauthenticated user triggers initial permission grants.
 */
function authoriseScript() {
  SpreadsheetApp.getUi().alert('Authorisation complete. Please refresh the page to see your custom menu.');
}

/** Triggers subject sheet initialisation via Setup.gs */
function triggerSetup() { Setup.triggerCreateSubjectSheets(); }

/** Freezes dynamic formulas on the import sheet to preserve calculation state */
function triggerFreeze() { Setup.freezeImportSheet(); }

/** Restores dynamic formulas on the import sheet for MIS data refresh */
function triggerThaw() { Setup.thawImportSheet(); }

// --- KS3 REPORT TRIGGERS ---
function triggerProgressReview() { showBatchModal('PROGRESS_REVIEW', 'Progress Reviews'); }
function triggerNextStepsSummary() { showBatchModal('NEXT_STEPS_SUMMARY', 'Next Steps Summaries'); }

/**
 * Opens the chunking modal dialog for batch generation.
 * Validates that import data is frozen beforehand to avoid reading unstable formulas.
 * @param {string} configKey Key corresponding to CONFIG.REPORTS.
 * @param {string} friendlyName Human-readable report title for the UI modal.
 */
function showBatchModal(configKey, friendlyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import'); 
  
  // Guard against reading volatile array formulas while calculations are live
  if (importSheet) {
    const status = importSheet.getRange('A1').getValue();
    if (status !== '🥶') {
      SpreadsheetApp.getUi().alert(
        'Validation Error', 
        'The import sheet must be frozen (🥶) before generating reports. Please use the menu: Typeless Reports > Freeze Import Data.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
  }

  // Inject report configuration metadata into the HTML dialog template
  const template = HtmlService.createTemplateFromFile('-09- BatchGeneration');
  template.configKey = configKey;
  template.friendlyName = friendlyName;
  
  const html = template.evaluate()
      .setWidth(450)
      .setHeight(380)
      .setTitle('Batch Generator');
      
  SpreadsheetApp.getUi().showModalDialog(html, 'Report Engine');
}

/**
 * Server-side initialisation for modal chunked processing.
 * Conducts a pre-flight data audit and constructs the target Drive folder.
 * @param {string} configKey Report configuration identifier.
 * @param {string|boolean} auditMode 'ignore', 'drop', or false if pre-flight audit check.
 * @return {Object} Status payload containing issues or folder details.
 */
function server_initBatch(configKey, auditMode) {
  const reportConfig = CONFIG.REPORTS[configKey];
  const payload = DataService.buildStudentDataPayload(reportConfig);

  if (payload.length === 0) return { error: "No student data found." };

  // Conduct the audit check only if the user has not yet selected a resolution strategy
  if (!auditMode) {
    const studentsWithIssues = payload.filter(s => s.auditIssues && s.auditIssues.length > 0);
    if (studentsWithIssues.length > 0) {
      const issuesList = studentsWithIssues.map(s => `<b>${s.name}</b>: ${s.auditIssues.join(' | ')}`);
      return {
        status: 'audit_warning',
        issues: issuesList,
        totalStudents: payload.length
      };
    }
  }

  // Provision the timestamped Drive folder once the audit is cleared or bypassed
  const folderId = DocumentBuilder.createBatchFolder(reportConfig, payload[0]);
  
  return {
    status: 'ready',
    folderId: folderId,
    folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    totalStudents: payload.length
  };
}

/**
 * Server-side chunk execution called cyclically by the modal dialog.
 * Slices student data into manageable chunks to stay well beneath GAS timeout limits.
 * @param {string} configKey Report configuration key.
 * @param {string} folderId Destination folder ID.
 * @param {number} startIndex Zero-based start index of the chunk slice.
 * @param {number} chunkSize Number of students to process.
 * @param {string} auditMode Audit flag ('ignore' or 'drop').
 * @return {Object} Simple success confirmation.
 */
function server_processChunk(configKey, folderId, startIndex, chunkSize, auditMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportConfig = CONFIG.REPORTS[configKey];
  
  // Re-build student payload dynamically in memory to maintain lean execution state
  const payload = DataService.buildStudentDataPayload(reportConfig);
  const chunk = payload.slice(startIndex, startIndex + chunkSize);

  ss.toast(`Merging chunk: ${startIndex + 1} to ${startIndex + chunk.length}...`, 'Background Engine');
  
  // Delegate document creation to the builder module
  DocumentBuilder.generateChunk(reportConfig, chunk, folderId, auditMode);
  
  return { success: true };
}