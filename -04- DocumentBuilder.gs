/**
 * DocumentBuilder.gs
 * Handles the generation of Google Docs from templates using a high-speed hybrid approach.
 */

const DocumentBuilder = {

  // --- CHUNKING ENGINE METHODS ---

  /**
   * Creates the destination folder in Google Drive.
   * @param {Object} reportConfig The configuration for the current report.
   * @param {Object} sampleStudent A single student record to extract global data from.
   * @returns {string} The ID of the newly created folder.
   */
  createBatchFolder: function(reportConfig, sampleStudent) {
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    
    // Extract globals from the sample student for folder naming
    const academicYear = sampleStudent?.academicYear || '';
    const collection = sampleStudent?.collection || '';
    const yearGroup = sampleStudent?.yearGroup || '';
    
    // Format: [academicYear] [collection] [yearGroup] [datestamp]
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      folderName += " next-steps";
    }
    
    const batchFolder = outputFolder.createFolder(folderName);
    return batchFolder.getId();
  },

  /**
   * Generates a single chunk of documents.
   * @param {Object} reportConfig The configuration for the current report.
   * @param {Array} chunkPayload The subset of students to process.
   * @param {string} folderId The ID of the destination folder.
   */
  generateChunk: function(reportConfig, chunkPayload, folderId) {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const batchFolder = DriveApp.getFolderById(folderId);
    
    chunkPayload.forEach((student) => {
      try {
        if (student.subjects && student.subjects.length > 0) {
          this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
        }
      } catch (error) {
        console.error(`Failed to generate document for ${student.name} (${student.adNo}): ${error.message}`);
        // Continuing to the next student in the chunk to prevent total failure
      }
    });
  },

  /**
   * Core generation logic combining DocumentApp (structural) and Docs API (text replacement).
   * @private
   */
  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    // Defensive check to ensure adNo exists before padding
    const safeAdNo = student.adNo ? String(student.adNo) : '000000';
    const paddedAdNo = safeAdNo.padStart(6, '0');
    
    // Format: [reg] [name] [paddedAdno] [shortName]
    let fileName = `${student.reg} ${student.name} ${paddedAdNo} ${student.shortName || ''}`.trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += " next-steps";
    }

    // 1. Create the physical file copy
    const newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    const docId = newDocFile.getId();
    
    // --- PHASE 1: Structural Table Building (DocumentApp) ---
    // We use DocumentApp here because cloning table rows structurally is easiest this way.
    // We inject the subject text directly into the row, which is highly scoped and fast.
    const newDoc = DocumentApp.openById(docId);
    const body = newDoc.getBody();
    
    this._populateSubjectTable(body, student.subjects);
    
    // We MUST save and close the document to flush the structural changes to Google Drive 
    // before the Docs API attempts to modify the text in Phase 2.
    newDoc.saveAndClose();

    // --- PHASE 2: High-Speed Global Text Replacement (Docs API) ---
    // We use the Advanced Google Docs API to replace all global tags (headers, footers, body)
    // in a single lightning-fast batch request.
    const requests = this._buildGlobalReplacementRequests(student, paddedAdNo);

    if (requests.length > 0) {
      Docs.Documents.batchUpdate({ requests: requests }, docId);
    }
  },

  /**
   * Constructs the payload required for the Google Docs API batchUpdate.
   * @private
   */
  _buildGlobalReplacementRequests: function(student, paddedAdNo) {
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "MMMM yyyy");
    
    // Map of all global tags to their target values
    const replacements = {
      '_Name_': student.name || '',
      '_Reg_': student.reg || '',
      '_AdNo_': paddedAdNo,
      '_Tutor_': student.tutor || '',
      '_Date_': dateStr,
      '_YearGroup_': student.yearGroup || '',
      '_Collection_': student.collection || '',
      '_Until_': student.until || ''
    };

    if (student.tutorInfo) {
      replacements['_AttPercent_'] = student.tutorInfo.percentAtt || '-';
      replacements['_PossSessions_'] = student.tutorInfo.possibleSessions || '-';
      replacements['_AuthAbs_'] = student.tutorInfo.authAbsences || '0';
      replacements['_UnauthAbs_'] = student.tutorInfo.unauthAbsences || '0';
      replacements['_Lates_'] = student.tutorInfo.lates || '0';
      replacements['_PSHE_'] = student.tutorInfo.pshe || '-';
    }

    // Convert the map into the specific array structure required by the Docs API
    return Object.keys(replacements).map(tag => ({
      replaceAllText: {
        containsText: { text: tag, matchCase: true },
        replaceText: String(replacements[tag]) // Ensure it is always cast as a string
      }
    }));
  },

  /**
   * Locates the subject template row, duplicates it, and cleans up the original.
   * @private
   */
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
          
          // Hygiene: Always remove the original template row so it doesn't linger 
          // if the student has no subjects.
          table.removeRow(r);
          break;
        }
      }
      if (targetTable) break;
    }

    if (!targetTable || !templateRow) return;

    // Add a row for each subject and replace the specific tags locally within that row object
    if (subjects && subjects.length > 0) {
      subjects.forEach((subj, index) => {
        const newRow = templateRow.copy();
        
        // Because these replacements are scoped to 'newRow', they execute extremely quickly
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
  }

};