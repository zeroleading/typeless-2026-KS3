/**
 * DocumentBuilder.gs
 * Handles the generation of Google Docs by cloning templates and injecting data.
 */

const DocumentBuilder = {

  generateBatch: function(reportConfig, studentPayload) {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    const batchFolder = outputFolder.createFolder(`${reportConfig.name} - ${dateStr}`);

    studentPayload.forEach(student => {
      if (student.subjects && student.subjects.length > 0) {
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });

    return batchFolder.getId(); 
  },

  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    const docName = `${student.adNo} - ${student.name} - ${reportName}`;
    const newFile = templateFile.makeCopy(docName, destinationFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();

    // 1. Replace Global Variables 
    this._replaceGlobalPlaceholders(body, student);
    if (header) this._replaceGlobalPlaceholders(header, student);
    if (footer) this._replaceGlobalPlaceholders(footer, student);

    // 2. Process the dynamic Subject Table
    this._populateSubjectTable(body, student.subjects);

    // 3. Post-Merge Polish: Make any raw URLs clickable
    this._makeUrlsClickable(body);

    doc.saveAndClose();
  },

  _replaceGlobalPlaceholders: function(element, student) {
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "MMMM yyyy");
    element.replaceText('_Name_', student.name);
    element.replaceText('_Reg_', student.reg);
    element.replaceText('_AdNo_', student.adNo);
    element.replaceText('_Tutor_', student.tutor);
    element.replaceText('_Date_', dateStr);
  },

  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    let targetTable = null;
    let templateRowIndex = -1;

    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      for (let r = 0; r < table.getNumRows(); r++) {
        if (table.getRow(r).getText().includes('{{subjectName}}')) {
          targetTable = table;
          templateRowIndex = r;
          break;
        }
      }
      if (targetTable) break;
    }

    if (!targetTable || templateRowIndex === -1) return;

    const templateRow = targetTable.getRow(templateRowIndex);

    subjects.forEach(subject => {
      const newRow = targetTable.appendTableRow(templateRow.copy());
      newRow.replaceText('{{subjectName}}', subject.subjectName || '-');
      newRow.replaceText('{{tg}}', subject.tg || '-');
      newRow.replaceText('{{eoy}}', subject.eoy || '-');
      newRow.replaceText('{{ucas}}', subject.ucas || '-');
      newRow.replaceText('{{rank}}', subject.rank || '-');
      newRow.replaceText('{{ucasRef}}', subject.ucasRef || '-');
      newRow.replaceText('{{crnt}}', subject.crnt || '-');
      newRow.replaceText('{{nextSteps}}', subject.nextSteps || '-');
      newRow.replaceText('{{att}}', subject.att || '-');
    });

    targetTable.removeRow(templateRowIndex);
  },

  /**
   * Helper: Sweeps the document body and converts raw text URLs into actual hyperlinks.
   * @private
   */
  _makeUrlsClickable: function(body) {
    const URL_PATTERN = 'http[s]?://[-a-zA-Z0-9@:%_+.~#?&//=]*';
    let foundElement = body.findText(URL_PATTERN);
    
    while (foundElement !== null) {
      const foundText = foundElement.getElement().asText();
      const start = foundElement.getStartOffset();
      const end = foundElement.getEndOffsetInclusive();
      
      // Extract the exact URL string
      const url = foundText.getText().substring(start, end + 1);
      
      // Set the link on that specific text range
      foundText.setLinkUrl(start, end, url);
      
      // Find the next occurrence
      foundElement = body.findText(URL_PATTERN, foundElement);
    }
  }

};