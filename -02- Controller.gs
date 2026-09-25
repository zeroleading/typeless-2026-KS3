/**
 * Controller.gs
 * Handles the user interface, custom menus, authorisation routing, and batch coordination.
 * British English conventions are maintained across all internal commentary.
 */

/**
 * Builds the custom application menu on spreadsheet open.
 * @param {Object} e The open event payload.
 */
function onOpen(e) {
  buildDynamicMenu();
}

/**
 * Constructs the contextual menu depending on user permissions.
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

/**
 * Displays an informational dialogue after permission elevation.
 */
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
 * Validates sheet state and opens the chunking modal dialogue.
 * @param {string} configKey The key in CONFIG.REPORTS to use.
 * @param {string} friendlyName The display name for the dialogue header.
 */
function showBatchModal(configKey, friendlyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import'); 
  
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
 * Initialises batch generation by executing data extraction once and caching the result.
 * Sparing the Google Sheets storage engine from repeated matrix reads prevents Error code INTERNAL.
 * 
 * @param {string} configKey The report configuration key.
 * @param {boolean} forceProceed Whether to bypass audit warnings.
 * @returns {Object} Status payload containing audit issues or folder details.
 */
function server_initBatch(configKey, forceProceed) {
  const reportConfig = CONFIG.REPORTS[configKey];
  if (!reportConfig) {
    throw new Error(`Report configuration not found for key: ${configKey}`);
  }

  // Extract the complete student dataset once from the spreadsheet backend
  const payload = DataService.buildStudentDataPayload(reportConfig);
  if (!payload || payload.length === 0) {
    return { error: 'No student data found in the spreadsheet.' };
  }

  // Run audit checks to highlight missing values prior to document creation
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

  // Create the designated output folder in Google Drive
  const folderId = DocumentBuilder.createBatchFolder(reportConfig, payload[0]);

  // Serialise and store the extracted payload in CacheService
  // This completely eliminates subsequent spreadsheet reads during chunk iterations
  _storePayloadInCache(configKey, payload);

  return {
    status: 'ready',
    folderId: folderId,
    folderUrl: `https://drive.google.com/drive/folders/${folderId}`,
    totalStudents: payload.length
  };
}

/**
 * Processes a specific chunk of students retrieved entirely from cache.
 * 
 * @param {string} configKey The report configuration key.
 * @param {string} folderId The Google Drive folder ID to save documents to.
 * @param {number} startIndex Index where the current chunk slice begins.
 * @param {number} chunkSize Number of students to process in this slice.
 * @returns {Object} Execution confirmation payload.
 */
function server_processChunk(configKey, folderId, startIndex, chunkSize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportConfig = CONFIG.REPORTS[configKey];
  
  // Retrieve the full student dataset directly from cache
  const payload = _retrievePayloadFromCache(configKey);
  
  if (!payload || payload.length === 0) {
    throw new Error('Batch cache has expired or could not be found. Please restart the batch generation.');
  }

  // Slice out the exact cohort subset for this iteration
  const chunk = payload.slice(startIndex, startIndex + chunkSize);

  ss.toast(`Merging cohort: ${startIndex + 1} to ${startIndex + chunk.length}...`, 'Background Engine');

  // Dispatch slice directly to the document builder
  DocumentBuilder.generateChunk(reportConfig, chunk, folderId);

  return { success: true };
}

/**
 * Persists an array payload into CacheService, segmenting across multiple keys if required.
 * CacheService enforces a 100 KB limit per entry; segmenting guarantees safety for larger cohorts.
 * 
 * @private
 * @param {string} configKey The report configuration key.
 * @param {Array<Object>} payload The full array of student objects.
 */
function _storePayloadInCache(configKey, payload) {
  const cache = CacheService.getUserCache();
  const jsonString = JSON.stringify(payload);
  const maxChunkSize = 90000; // Keep safely below the 100 KB limit per key
  const totalChunks = Math.ceil(jsonString.length / maxChunkSize);
  
  const cacheEntries = {};
  cacheEntries[`BATCH_${configKey}_COUNT`] = String(totalChunks);

  for (let i = 0; i < totalChunks; i++) {
    const chunkContent = jsonString.substr(i * maxChunkSize, maxChunkSize);
    cacheEntries[`BATCH_${configKey}_PART_${i}`] = chunkContent;
  }

  // Set time-to-live to 1500 seconds (25 minutes), comfortably exceeding batch duration
  cache.putAll(cacheEntries, 1500);
}

/**
 * Retrieves and reassembles a segment-cached payload.
 * 
 * @private
 * @param {string} configKey The report configuration key.
 * @returns {Array<Object>|null} The parsed student array, or null if missing.
 */
function _retrievePayloadFromCache(configKey) {
  const cache = CacheService.getUserCache();
  const countStr = cache.get(`BATCH_${configKey}_COUNT`);
  
  if (!countStr) {
    return null;
  }

  const totalChunks = parseInt(countStr, 10);
  const chunkKeys = [];
  for (let i = 0; i < totalChunks; i++) {
    chunkKeys.push(`BATCH_${configKey}_PART_${i}`);
  }

  const cachedParts = cache.getAll(chunkKeys);
  let completeJson = '';

  for (let i = 0; i < totalChunks; i++) {
    const part = cachedParts[`BATCH_${configKey}_PART_${i}`];
    if (!part) {
      return null; // A cache fragment was prematurely evicted
    }
    completeJson += part;
  }

  return JSON.parse(completeJson);
}