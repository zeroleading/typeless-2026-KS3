/**
 * DocumentBuilder.gs
 * Handles the generation of Google Docs from templates.
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
    
    // Format: [academicYear] [collection] [yearGroup] [datestamp]
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      folderName += " next-steps";
    }
    
    const batchFolder = outputFolder.createFolder(folderName);
    const totalStudents = studentPayload.length;

    studentPayload.forEach((student, index) => {
      if (student.subjects && student.subjects.length > 0) {
        
        // Live Progress Indicator
        ss.toast(`Merging document ${index + 1} of ${totalStudents}...\n(${student.name})`, 'Progress Tracker', 10);
        
        this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
      }
    });

    return batchFolder.getId(); 
  },

  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    const paddedAdNo = String(student.adNo).padStart(6, '0');
    
    // Format: [reg] [name] [paddedAdno] [shortName]
    let fileName = `${student.reg} ${student.name} ${paddedAdNo} ${student.shortName || ''}`.trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += " next-steps";
    }

    const newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    
    // Define the specific sections of the document
    const body = newDoc.getBody();
    const header = newDoc.getHeader(); // This defines 'header' to fix your ReferenceError!
    const footer = newDoc.getFooter();

    // Helper function to replace text safely in any document section
    const replaceGlobalsInSection = (section) => {
      if (!section) return; // Safely skip if the template has no header/footer
      
      section.replaceText('_Name_', student.name || '');
      section.replaceText('_Reg_', student.reg || '');
      section.replaceText('_AdNo_', paddedAdNo);
      section.replaceText('_Tutor_', student.tutor || '');
      section.replaceText('_Date_', Utilities.formatDate(new Date(), "Europe/London", "MMMM yyyy"));
      section.replaceText('_YearGroup_', student.yearGroup || '');
      section.replaceText('_Collection_', student.collection || '');
      section.replaceText('_Until_', student.until || '');
      
      if (student.tutorInfo) {
        section.replaceText('_AttPercent_', student.tutorInfo.percentAtt || '-');
        section.replaceText('_PossSessions_', student.tutorInfo.possibleSessions || '-');
        section.replaceText('_AuthAbs_', student.tutorInfo.authAbsences || '0');
        section.replaceText('_UnauthAbs_', student.tutorInfo.unauthAbsences || '0');
        section.replaceText('_Lates_', student.tutorInfo.lates || '0');
        section.replaceText('_PSHE_', student.tutorInfo.pshe || '-');
      }
    };

    // Apply the global replacements to the body, header, and footer automatically
    replaceGlobalsInSection(body);
    replaceGlobalsInSection(header);
    replaceGlobalsInSection(footer);

    // Finally, populate the dynamic subjects table in the main body
    this._populateSubjectTable(body, student.subjects);

    newDoc.saveAndClose();
  },

  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    if (tables.length === 0) return;

    // Find the table that contains our template tags
    let targetTable = null;
    let templateRow = null;
    let templateRowIndex = -1;

    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        if (row.getText().includes('{{subjectName}}')) {
          targetTable = table;
          templateRow = row.copy();
          templateRowIndex = r;
          // Remove the original template row
          table.removeRow(r);
          break;
        }
      }
      if (targetTable) break;
    }

    if (!targetTable || !templateRow) return;

    // Add a row for each subject and replace the specific tags
    subjects.forEach((subj, index) => {
      const newRow = templateRow.copy();
      
      newRow.replaceText('{{subjectName}}', subj.subjectName || '');
      newRow.replaceText('{{teacher}}', subj.teacher || '');
      newRow.replaceText('{{tg}}', subj.tg || '');
      newRow.replaceText('{{crnt}}', subj.crnt || '');
      newRow.replaceText('{{ci1}}', subj.ci1 || '');
      newRow.replaceText('{{ci2}}', subj.ci2 || '');
      newRow.replaceText('{{ci3}}', subj.ci3 || '');
      newRow.replaceText('{{ci4}}', subj.ci4 || '');
      newRow.replaceText('{{nextSteps1}}', subj.nextSteps1 || '');
      newRow.replaceText('{{nextSteps2}}', subj.nextSteps2 || '');
      
      targetTable.insertTableRow(templateRowIndex + index, newRow);
    });
  }

};