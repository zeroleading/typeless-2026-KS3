/**
 * DocumentBuilder.gs
 * Handles the generation of Google Docs by cloning templates and injecting data.
 */

const DocumentBuilder = {

  generateBatch: function(reportConfig, studentPayload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // Grab active sheet for progress toasts
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    
    // Extract globals from the first student payload for folder naming
    const academicYear = studentPayload[0]?.academicYear || '';
    const collection = studentPayload[0]?.collection || '';
    const yearGroup = studentPayload[0]?.yearGroup || '';
    
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      folderName += " next-steps";
    }
    
    const batchFolder = outputFolder.createFolder(folderName);

    const totalStudents = studentPayload.length;

    studentPayload.forEach((student, index) => {
      if (student.subjects && student.subjects.length > 0) {
        
        // Live Progress Indicator
        // Keeps the toast alive for 10 seconds, or until the next loop overwrites it
        ss.toast(`Merging document ${index + 1} of ${totalStudents}...\n(${student.name})`, 'Progress Tracker', 10);
        
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });

    return batchFolder.getId(); 
  },

  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    const paddedAdNo = String(student.adNo).padStart(5, '0');
    
    // Format: [reg] [name] [paddedAdno] [shortName]
    let fileName = `${student.reg} ${student.name} ${paddedAdNo} ${student.shortName || ''}`.trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += " next-steps";
    }

    const newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();
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
    element.replaceText('_YearGroup_', student.yearGroup || '');
    element.replaceText('_Collection_', student.collection || '');
    
    // Replace Tutor Placeholders
    if (student.tutorInfo) {
      element.replaceText('_AttPercent_', student.tutorInfo.percentAtt || '-');
      element.replaceText('_PossSessions_', student.tutorInfo.possibleSessions || '-');
      element.replaceText('_AuthAbs_', student.tutorInfo.authAbsences || '0');
      element.replaceText('_UnauthAbs_', student.tutorInfo.unauthAbsences || '0');
      element.replaceText('_Lates_', student.tutorInfo.lates || '0');
      element.replaceText('_PSHE_', student.tutorInfo.pshe || '-');
    }
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
      newRow.replaceText('{{teacher}}', subject.teacher || '-');
      newRow.replaceText('{{tg}}', subject.tg || '-');
      newRow.replaceText('{{crnt}}', subject.crnt || '-');
      newRow.replaceText('{{ci1}}', subject.ci1 || '-');
      newRow.replaceText('{{ci2}}', subject.ci2 || '-');
      newRow.replaceText('{{ci3}}', subject.ci3 || '-');
      newRow.replaceText('{{ci4}}', subject.ci4 || '-');
      newRow.replaceText('{{nextSteps1}}', subject.nextSteps1 || '-');
      newRow.replaceText('{{nextSteps2}}', subject.nextSteps2 || '-');
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