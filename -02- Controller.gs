/**
 * Controller.gs
 * Handles the user interface, custom menus, and authorisation routing.
 */

function onOpen(e) {
  buildDynamicMenu();
}

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

function triggerSetup() { Setup.triggerCreateSubjectSheets(); }
function triggerFreeze() { Setup.freezeImportSheet(); }
function triggerThaw() { Setup.thawImportSheet(); }

// --- REPORT TRIGGERS ---
function triggerProgressReview() { showBatchModal('PROGRESS_REVIEW', 'Progress Reviews'); }
function triggerNextStepsSummary() { showBatchModal('NEXT_STEPS_SUMMARY', 'Next Steps Summaries'); }

/**
 * Opens the new Chunking Modal for heavy report generation.
 * @param {string} configKey The key in CONFIG.REPORTS to use.
 * @param {string} friendlyName The display name for the UI.
 */
function showBatchModal(configKey, friendlyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import'); 
  
  if (importSheet) {
    const status = importSheet.getRange('A1').getValue();
    if (status !== '🥶') {
      SpreadsheetApp.getUi().alert('Validation Error', 'The import sheet must be frozen (🥶) before generating reports. Please use the menu: Typeless Reports > Freeze Import Data.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  // Use HtmlTemplate to pass variables to the HTML file
  const template = HtmlService.createTemplateFromFile('09- BatchGeneration');
  template.configKey = configKey;
  template.friendlyName = friendlyName;
  
  const html = template.evaluate()
      .setWidth(450)
      .setHeight(380)
      .setTitle('Batch Generator');
      
  SpreadsheetApp.getUi().showModalDialog(html, 'Report Engine');
}

/**
 * Called by the Modal (Step 1): Prepares the folder and audits the data.
 * @param {string} configKey The report configuration key.
 * @param {boolean} forceProceed Whether to bypass audit warnings.
 * @returns {Object} Status payload containing issues or folder details.
 */
function server_initBatch(configKey, forceProceed) {
  const reportConfig = CONFIG.REPORTS[configKey];
  const payload = DataService.buildStudentDataPayload(reportConfig);

  if (payload.length === 0) return { error: "No student data found." };

  // 1. Audit Check
  if (!forceProceed) {
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

  // 2. Folder Creation
  const folderId = DocumentBuilder.createBatchFolder(reportConfig, payload[0]);
  
  return {
    status: 'ready',
    folderId: folderId,
    folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    totalStudents: payload.length
  };
}

/**
 * Called by the Modal (Step 3 Loop): Processes a specific chunk of students.
 * @param {string} configKey The report configuration key.
 * @param {string} folderId The Google Drive folder ID to save to.
 * @param {number} startIndex Where to begin slicing the array.
 * @param {number} chunkSize How many students to process in this run.
 * @returns {Object} Success flag.
 */
function server_processChunk(configKey, folderId, startIndex, chunkSize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportConfig = CONFIG.REPORTS[configKey];
  
  // Re-build payload dynamically (fast and keeps memory lean)
  const payload = DataService.buildStudentDataPayload(reportConfig);
  
  // Slice out just the 10 students requested
  const chunk = payload.slice(startIndex, startIndex + chunkSize);

  ss.toast(`Merging chunk: ${startIndex + 1} to ${startIndex + chunk.length}...`, 'Background Engine');
  
  // Send to builder
  DocumentBuilder.generateChunk(reportConfig, chunk, folderId);
  
  return { success: true };
}