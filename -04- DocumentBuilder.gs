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
    let outputFolder, batchFolder;
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
        const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
        
        // Extract globals from the sample student for folder naming
        const academicYear = sampleStudent?.academicYear || '';
        const collection = sampleStudent?.collection || '';
        const yearGroup = sampleStudent?.yearGroup || '';
        
        // Format: [academicYear] [collection] [yearGroup] [datestamp]
        let folderName = (academicYear + " " + collection + " " + yearGroup + " " + dateStr).trim();
        if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
          folderName += " next-steps";
        }
        
        batchFolder = outputFolder.createFolder(folderName);
        return batchFolder.getId();
      } catch (e) {
        if (attempt === 3) throw e;
        Utilities.sleep(1000 * attempt);
      }
    }
  },

  /**
   * Generates a single chunk of documents.
   * @param {Object} reportConfig The configuration for the current report.
   * @param {Array} chunkPayload The subset of students to process.
   * @param {string} folderId The ID of the destination folder.
   */
  generateChunk: function(reportConfig, chunkPayload, folderId) {
    let templateFile = null;
    let batchFolder = null;
    
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        templateFile = DriveApp.getFileById(reportConfig.templateId);
        batchFolder = DriveApp.getFolderById(folderId);
        break;
      } catch (e) {
        if (attempt === 3) throw e;
        Utilities.sleep(1000 * attempt);
      }
    }
    
    let lastError = null;
    let successCount = 0;

    chunkPayload.forEach((student) => {
      try {
        if (student.subjects && student.subjects.length > 0) {
          this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
          successCount++;
        }
      } catch (error) {
        lastError = error;
        console.error(`Failed to generate document for ${student.name} (${student.adNo}): ${error.message}`);
      }
    });

    if (successCount === 0 && chunkPayload.length > 0 && lastError) {
      throw lastError;
    }
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
    let fileName = (student.reg + " " + student.name + " " + paddedAdNo + " " + (student.shortName || '')).trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += " next-steps";
    }

    let lastError = null;

    // Enclose the full creation and editing pipeline in a transaction retry loop.
    // If a storage sync error occurs during makeCopy, getBody, saveAndClose, or batchUpdate,
    // the broken file copy is trashed and recreated cleanly.
    for (let attempt = 1; attempt <= 5; attempt++) {
      let newDocFile = null;
      let docId = '';

      try {
        // 1. Create physical copy from template
        newDocFile = templateFile.makeCopy(fileName, destinationFolder);
        docId = newDocFile.getId();
        
        // Incremental pause to allow Google Drive storage backend propagation
        Utilities.sleep(800 * attempt);

        // 2. Open document handle
        const newDoc = DocumentApp.openById(docId);
        const body = newDoc.getBody();

        // 3. Phase 1: Structural Table Building (DocumentApp)
        this._populateSubjectTable(body, student.subjects);

        // 4. Save and close to flush structural edits before Docs API operations
        newDoc.saveAndClose();
        Utilities.sleep(400);

        // 5. Phase 2: High-Speed Global Text Replacement (Docs API)
        const requests = this._buildGlobalReplacementRequests(student, paddedAdNo);
        if (requests.length > 0) {
          Docs.Documents.batchUpdate({ requests: requests }, docId);
        }

        // Execution succeeded
        return;

      } catch (e) {
        lastError = e;
        console.warn(`Attempt ${attempt} failed for ${student.name} (${student.adNo}): ${e.message}`);

        // Trash the failed document copy so un-synced locks are discarded
        if (newDocFile) {
          try {
            newDocFile.setTrashed(true);
          } catch (trashErr) {
            // Ignore cleanup errors
          }
        }

        if (attempt === 5) {
          throw new Error(`Storage sync error after 5 attempts: ${e.message}`);
        }

        Utilities.sleep(1000 * attempt);
      }
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
