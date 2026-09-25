/**
 * DocumentBuilder.gs [KS3]
 * High-speed hybrid document generation engine utilising DocumentApp for row cloning
 * and the Google Docs Advanced API for global token replacement.
 * British English conventions are maintained across all internal commentary.
 */

const DocumentBuilder = {

  // --- CHUNKING ENGINE METHODS ---

  /**
   * Creates the destination folder in Google Drive with exponential backoff.
   * @param {Object} reportConfig The configuration for the current report.
   * @param {Object} sampleStudent A single student record to extract global folder data from.
   * @returns {string} The ID of the newly created folder.
   */
  createBatchFolder: function(reportConfig, sampleStudent) {
    let outputFolder;
    let batchFolder;
    
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
        const dateStr = Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd');
        
        // Extract metadata for standardised folder naming
        const academicYear = sampleStudent?.academicYear || '';
        const collection = sampleStudent?.collection || '';
        const yearGroup = sampleStudent?.yearGroup || '';
        
        let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
        if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
          folderName += ' next-steps';
        }
        
        batchFolder = outputFolder.createFolder(folderName);
        return batchFolder.getId();
      } catch (e) {
        if (attempt === 3) {
          throw new Error(`[Folder Creation Error]: Unable to create batch destination folder: ${e.message}`);
        }
        Utilities.sleep(1000 * attempt);
      }
    }
  },

  /**
   * Generates a single chunk of documents from the pre-sliced payload.
   * @param {Object} reportConfig The configuration for the current report.
   * @param {Array<Object>} chunkPayload The subset of students to process.
   * @param {string} folderId The ID of the destination folder.
   */
  generateChunk: function(reportConfig, chunkPayload, folderId) {
    let templateFile = null;
    let batchFolder = null;
    
    // Acquire Drive resources once per chunk with retry protection
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        templateFile = DriveApp.getFileById(reportConfig.templateId);
        batchFolder = DriveApp.getFolderById(folderId);
        break;
      } catch (e) {
        if (attempt === 3) {
          throw new Error(`[Drive Acquisition Error]: Could not acquire template or target folder: ${e.message}`);
        }
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

    // If every student in this chunk failed, surface the error so the UI modal can flag it
    if (successCount === 0 && chunkPayload.length > 0 && lastError) {
      throw lastError;
    }
  },

  /**
   * Orchestrates physical copy, structural table duplication, and Docs API token substitution.
   * @private
   * @param {Object} student The student data record.
   * @param {GoogleAppsScript.Drive.File} templateFile The master template file object.
   * @param {GoogleAppsScript.Drive.Folder} destinationFolder The target folder destination.
   * @param {string} reportName The current report identifier.
   */
  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    const safeAdNo = student.adNo ? String(student.adNo) : '000000';
    const paddedAdNo = safeAdNo.padStart(6, '0');
    
    let fileName = `${student.reg} ${student.name} ${paddedAdNo} ${student.shortName || ''}`.trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += ' next-steps';
    }

    // Execute within a transactional retry loop to defend against storage sync blips
    for (let attempt = 1; attempt <= 3; attempt++) {
      let newDocFile = null;
      let docId = '';

      try {
        // 1. Duplicate template in destination folder
        newDocFile = templateFile.makeCopy(fileName, destinationFolder);
        docId = newDocFile.getId();
        
        // Brief pause allowing Drive metadata propagation across Google storage clusters
        Utilities.sleep(400 * attempt);

        // 2. Structural Phase: DocumentApp row cloning
        const newDoc = DocumentApp.openById(docId);
        const body = newDoc.getBody();
        this._populateSubjectTable(body, student.subjects);

        // Explicitly flush and close to unlock document before advanced API access
        newDoc.saveAndClose();
        Utilities.sleep(250);

        // 3. Text Replacement Phase: Advanced Docs API batchUpdate
        const requests = this._buildGlobalReplacementRequests(student, paddedAdNo);
        if (requests.length > 0) {
          Docs.Documents.batchUpdate({ requests: requests }, docId);
        }

        // Successfully created and populated
        return;

      } catch (e) {
        console.warn(`Attempt ${attempt} failed for ${student.name} (${student.adNo}): ${e.message}`);

        // Purge the incomplete document to keep destination folders tidy
        if (newDocFile) {
          try {
            newDocFile.setTrashed(true);
          } catch (trashErr) {
            // Silently ignore cleanup errors on failed references
          }
        }

        if (attempt === 3) {
          throw new Error(`[Document Construction Error] (${student.name}): ${e.message}`);
        }

        Utilities.sleep(750 * attempt);
      }
    }
  },

  /**
   * Constructs the payload required for the Google Docs API batchUpdate call.
   * @private
   * @param {Object} student The student record.
   * @param {string} paddedAdNo The zero-padded admission number.
   * @returns {Array<Object>} The array of replaceAllText request objects.
   */
  _buildGlobalReplacementRequests: function(student, paddedAdNo) {
    const dateStr = Utilities.formatDate(new Date(), 'Europe/London', 'MMMM yyyy');
    
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

    return Object.keys(replacements).map(tag => ({
      replaceAllText: {
        containsText: { text: tag, matchCase: true },
        replaceText: String(replacements[tag])
      }
    }));
  },

  /**
   * Locates the subject template row, duplicates it for each subject, and cleans up the placeholder.
   * Tailored for Next Steps Summaries by omitting unneeded target grade evaluations.
   * @private
   * @param {GoogleAppsScript.Document.Body} body The document body object.
   * @param {Array<Object>} subjects The list of subject records for this student.
   */
  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    if (tables.length === 0) return;

    let targetTable = null;
    let templateRow = null;
    let templateRowIndex = -1;

    // Identify the specific table containing our template tags
    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        if (row.getText().includes('{{subjectName}}')) {
          targetTable = table;
          templateRow = row.copy();
          templateRowIndex = r;
          
          // Remove the placeholder row before injecting student data
          table.removeRow(r);
          break;
        }
      }
      if (targetTable) break;
    }

    if (!targetTable || !templateRow) return;

    // Clone and populate a row for each active subject
    if (subjects && subjects.length > 0) {
      subjects.forEach((subj, index) => {
        const newRow = templateRow.copy();
        
        // Perform scoped replacements directly on the row element
        newRow.replaceText('{{subjectName}}', subj.subjectName || '');
        newRow.replaceText('{{teacher}}', subj.teacher || '');
        newRow.replaceText('{{crnt}}', subj.crnt || '');
        newRow.replaceText('{{ci1}}', subj.ci1 || '');
        newRow.replaceText('{{ci2}}', subj.ci2 || '');
        newRow.replaceText('{{ci3}}', subj.ci3 || '');
        newRow.replaceText('{{ci4}}', subj.ci4 || '');
        newRow.replaceText('{{nextSteps1}}', subj.nextSteps1 || '');
        newRow.replaceText('{{nextSteps2}}', subj.nextSteps2 || '');

        // Safe fallback in case the template also includes legacy target tags
        newRow.replaceText('{{tg}}', subj.tg || '');
        newRow.replaceText('{{stg}}', subj.stg || '');
        
        targetTable.insertTableRow(templateRowIndex + index, newRow);
      });
    }
  }

};