/**
 * @fileoverview DocumentBuilder.gs (KS3)
 * Handles the generation of Google Docs from templates using a high-speed hybrid approach:
 * DocumentApp for row duplication and the Advanced Google Docs API for global tag replacement.
 */

const DocumentBuilder = {

  // ---------------------------------------------------------------------------
  // 1. Chunking Engine Methods
  // ---------------------------------------------------------------------------

  /**
   * Creates the destination batch folder in Google Drive.
   * @param {Object} reportConfig Target report profile configuration.
   * @param {Object} sampleStudent A single student record used to derive folder naming tokens.
   * @return {string} The ID of the newly created batch folder.
   */
  createBatchFolder: function(reportConfig, sampleStudent) {
    const outputFolder = DriveApp.getFolderById(CONFIG.GLOBAL.OUTPUT_FOLDER_ID);
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    
    // Extract metadata tokens from sample student to structure clean archive names
    const academicYear = sampleStudent?.academicYear || '';
    const collection = sampleStudent?.collection || '';
    const yearGroup = sampleStudent?.yearGroup || '';
    
    let folderName = `${academicYear} ${collection} ${yearGroup} ${dateStr}`.trim();
    if (reportConfig.name === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      folderName += " next-steps";
    }
    
    const batchFolder = outputFolder.createFolder(folderName);
    return batchFolder.getId();
  },

  /**
   * Generates a single chunk of documents, honouring the audit pruning mode.
   * @param {Object} reportConfig Active report configuration profile.
   * @param {Array<Object>} chunkPayload Subset of student records to process in this execution.
   * @param {string} folderId Google Drive destination folder identifier.
   * @param {string} [auditMode='ignore'] Audit decision flag ('ignore' or 'drop').
   */
  generateChunk: function(reportConfig, chunkPayload, folderId, auditMode = 'ignore') {
    const templateFile = DriveApp.getFileById(reportConfig.templateId);
    const batchFolder = DriveApp.getFolderById(folderId);
    
    chunkPayload.forEach((student) => {
      try {
        // Filter incomplete subjects if the user elected to prune them during pre-flight review
        let validSubjects = student.subjects || [];
        if (auditMode === 'drop') {
          validSubjects = validSubjects.filter(subj => this._isSubjectComplete(subj, reportConfig.name));
        }

        // Generate the document only if at least one complete subject remains
        if (validSubjects.length > 0) {
          student.subjects = validSubjects;
          this._buildSingleDocument(student, templateFile, batchFolder, reportConfig.name);
        }
      } catch (error) {
        // Log individual document failures defensively to avoid aborting the entire chunk
        console.error(`Failed to generate document for ${student.name} (${student.adNo}): ${error.message}`);
      }
    });
  },

  /**
   * Evaluates whether an individual subject record satisfies completeness criteria.
   * Synchronised with the pre-flight checks in DataService.gs.
   * @private
   * @param {Object} subj Subject assessment record.
   * @param {string} [reportName] Canonical report title.
   * @return {boolean} True if all mandatory indicators are present, false otherwise.
   */
  _isSubjectComplete: function(subj, reportName) {
    const hasCrnt = subj.crnt !== '' && subj.crnt !== undefined;
    const hasCi1 = subj.ci1 !== '' && subj.ci1 !== undefined;
    const hasCi2 = subj.ci2 !== '' && subj.ci2 !== undefined;
    const hasCi3 = subj.ci3 !== '' && subj.ci3 !== undefined;
    const hasCi4 = subj.ci4 !== '' && subj.ci4 !== undefined;
    
    // Ensure at least one Next Step column has been selected
    const hasNextSteps = (subj.nextSteps1 && String(subj.nextSteps1).trim() !== '') ||
                         (subj.nextSteps2 && String(subj.nextSteps2).trim() !== '');

    return hasCrnt && hasCi1 && hasCi2 && hasCi3 && hasCi4 && hasNextSteps;
  },

  // ---------------------------------------------------------------------------
  // 2. Document Assembly Implementation
  // ---------------------------------------------------------------------------

  /**
   * Core generation logic combining DocumentApp (structural row duplication)
   * and Docs API batchUpdate (high-speed global tag replacement).
   * @private
   * @param {Object} student Aggregated student data object.
   * @param {GoogleAppsScript.Drive.File} templateFile Source Google Doc template file.
   * @param {GoogleAppsScript.Drive.Folder} destinationFolder Target Drive folder.
   * @param {string} reportName Canonical report title.
   */
  _buildSingleDocument: function(student, templateFile, destinationFolder, reportName) {
    // Standardise admission numbers to 6 digits to guarantee consistent sorting and identification
    const safeAdNo = student.adNo ? String(student.adNo) : '000000';
    const paddedAdNo = safeAdNo.padStart(6, '0');
    
    let fileName = `${student.reg} ${student.name} ${paddedAdNo} ${student.shortName || ''}`.trim();
    if (reportName === CONFIG.REPORTS.NEXT_STEPS_SUMMARY.name) {
      fileName += " next-steps";
    }

    // 1. Create file copy in destination folder
    const newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    const docId = newDocFile.getId();
    
    // --- PHASE 1: Structural Table Building (DocumentApp) ---
    // DocumentApp is utilised specifically for table manipulation, which cannot be cleanly
    // performed via the REST API without complex structural index calculations.
    const newDoc = DocumentApp.openById(docId);
    const body = newDoc.getBody();
    
    this._populateSubjectTable(body, student.subjects);
    
    // Flush structural updates to Google Drive before handing off to the REST API in Phase 2
    newDoc.saveAndClose();

    // --- PHASE 2: High-Speed Global Text Replacement (Docs API) ---
    // The Advanced Google Docs API executes global replacements across body, headers,
    // and footers in a single atomic request, bypassing DocumentApp execution overhead.
    const requests = this._buildGlobalReplacementRequests(student, paddedAdNo);

    if (requests.length > 0) {
      Docs.Documents.batchUpdate({ requests: requests }, docId);
    }
  },

  /**
   * Constructs the payload required for the Google Docs API batchUpdate call.
   * @private
   * @param {Object} student Student data record.
   * @param {string} paddedAdNo Six-digit formatted admission number.
   * @return {Array<Object>} List of replaceAllText structural request objects.
   */
  _buildGlobalReplacementRequests: function(student, paddedAdNo) {
    const dateStr = Utilities.formatDate(new Date(), "Europe/London", "MMMM yyyy");
    
    // Map of template placeholder tokens to student metrics
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

    // Convert key-value pairs into the payload schema required by Docs.Documents.batchUpdate
    return Object.keys(replacements).map(tag => ({
      replaceAllText: {
        containsText: { text: tag, matchCase: true },
        replaceText: String(replacements[tag])
      }
    }));
  },

  /**
   * Locates the subject template row, clones it for each subject, and cleans up the placeholder.
   * @private
   * @param {GoogleAppsScript.Document.Body} body Document body element.
   * @param {Array<Object>} subjects Filtered list of subjects.
   */
  _populateSubjectTable: function(body, subjects) {
    const tables = body.getTables();
    if (tables.length === 0) return;

    // Search across document tables for the prototype row marker
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
          
          // Remove the prototype row to prevent blank placeholder remnants in the final document
          table.removeRow(r);
          break;
        }
      }
      if (targetTable) break;
    }

    if (!targetTable || !templateRow) return;

    // Inject a duplicated row populated with subject assessments
    if (subjects && subjects.length > 0) {
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
  }
};