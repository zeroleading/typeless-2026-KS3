/**
 * @fileoverview DataService.gs (KS3)
 * Handles data extraction, in-memory aggregation, dictionary lookups,
 * and translation of student assessment and attendance records across sheets.
 */

const DataService = {
  /**
   * Compiles the comprehensive student data payload required for Key Stage 3 reporting.
   * Pulls cohort parameters, constructs base student maps, and parses each subject sheet.
   * @param {Object} reportConfig Configuration definition for the active report profile.
   * @return {Array<Object>} Aggregated student record objects.
   */
  buildStudentDataPayload: function(reportConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Fetch Global Batch Values from the Control Panel named ranges
    const yearGroup = ss.getRangeByName(CONFIG.SCOPE.yearGroup)?.getValue() || '';
    const collection = ss.getRangeByName(CONFIG.SCOPE.collection)?.getValue() || '';
    const academicYear = ss.getRangeByName(CONFIG.SCOPE.academicYear)?.getValue() || '';
    const shortName = ss.getRangeByName(CONFIG.SCOPE.shortName)?.getValue() || '';
    const until = ss.getRangeByName(CONFIG.SCOPE.until)?.getValue() || '';

    // 2. Fetch Control Panel Maps & Dictionaries dynamically
    const fieldMap = this._getDynamicFieldMap(ss);
    const translations = this._getTranslationsDictionary(ss);

    // 3. Build base student directory and attach tutor attendance metrics
    const studentMap = this._getMasterStudentList(ss);
    this._attachTutorData(ss, studentMap, fieldMap, translations);

    // 4. Process all individual Subject Sheets matching standard naming patterns
    // Subject sheets follow 2-letter codes (e.g., 'Bi', 'Ch', 'Ma') or 'EnL'.
    const allSheets = ss.getSheets();
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;

    allSheets.forEach(sheet => {
      if (subjectRegex.test(sheet.getName())) {
        this._processSubjectSheet(ss, sheet, studentMap, fieldMap, translations, reportConfig);
      }
    });

    // 5. Convert map to array and inject global batch metadata into each student record
    return Object.values(studentMap).map(student => ({
      ...student,
      yearGroup: yearGroup,
      collection: collection,
      academicYear: academicYear,
      shortName: shortName,
      until: until
    }));
  },

  /**
   * Resolves dynamic header mappings from the Control Panel named range.
   * Falls back to FALLBACK_FIELD_MAP if the named range is unavailable.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet instance.
   * @return {Object} Dictionary mapping internal field keys to sheet column headers.
   */
  _getDynamicFieldMap: function(ss) {
    const map = { ...CONFIG.FALLBACK_FIELD_MAP }; 
    const range = ss.getRangeByName(CONFIG.SCOPE.fieldMap);
    if (!range) return map;

    const data = range.getValues();
    data.forEach(row => {
      const internalRef = String(row[0]).trim();
      const targetHeader = String(row[1]).trim();
      // Skip commented-out entries prefixed with asterisks
      if (internalRef && targetHeader && !internalRef.includes('**')) {
        map[internalRef] = targetHeader;
      }
    });
    return map;
  },

  /**
   * Resolves category translation dictionaries (e.g. Attitude to Learning codes).
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet instance.
   * @return {Object} Nested dictionary: { CATEGORY: { CODE: translation } }
   */
  _getTranslationsDictionary: function(ss) {
    const dict = {};
    const range = ss.getRangeByName(CONFIG.SCOPE.translations);
    if (!range) return dict;

    const data = range.getValues();
    data.forEach(row => {
      const category = String(row[0]).trim().toUpperCase();
      const code = String(row[1]).trim().toUpperCase();
      const translation = String(row[2]).trim();

      if (category && code && !category.includes('**')) {
        if (!dict[category]) dict[category] = {};
        dict[category][code] = translation;
      }
    });
    return dict;
  },

  /**
   * Translates raw MIS codes into user-facing descriptors using the translation tables.
   * @private
   * @param {*} rawValue Cell contents to be translated.
   * @param {string} category Dictionary category key (e.g., 'CRNT', 'CI', 'PSHE').
   * @param {Object} translationsDict Nested dictionary compiled from named ranges.
   * @return {string} Translated human-readable text.
   */
  _translate: function(rawValue, category, translationsDict) {
    if (rawValue === '' || rawValue === undefined) return '';
    const safeValue = String(rawValue).trim().toUpperCase();
    if (translationsDict[category] && translationsDict[category][safeValue]) {
      return translationsDict[category][safeValue];
    }
    return String(rawValue); 
  },

  /**
   * Extracts the master student cohort from the 'simpleStudentData' named range.
   * Initialises the audit tracking container for each individual record.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet instance.
   * @return {Object<string, Object>} Map of admission numbers to base student entities.
   */
  _getMasterStudentList: function(ss) {
    const studentMap = {};
    const range = ss.getRangeByName('simpleStudentData');
    if (!range) return studentMap; 
    
    const data = range.getValues();
    data.forEach(row => {
      const fullName = row[0];
      const rawAdNo = row[2]; 
      const reg = row[3];     
      const tutor = row[5];   
      
      if (rawAdNo && String(rawAdNo).toLowerCase() !== 'adno') { 
        const adNo = String(rawAdNo).trim();
        studentMap[adNo] = {
          adNo: adNo,
          name: fullName,
          reg: reg,
          tutor: tutor,
          tutorInfo: {}, 
          subjects: [],
          // Initialise an empty array to collect missing data warnings during sheet parsing
          auditIssues: []
        };
      }
    });
    return studentMap;
  },

  /**
   * Reads the 'tutorAssessment' named range to attach registration data,
   * attendance percentages, absence counts, and PSHE comments.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet instance.
   * @param {Object<string, Object>} studentMap Master student dictionary.
   * @param {Object} fieldMap Column header mapping dictionary.
   * @param {Object} translations Category translation dictionary.
   */
  _attachTutorData: function(ss, studentMap, fieldMap, translations) {
    const range = ss.getRangeByName('tutorAssessment');
    if (!range) return;

    const data = range.getValues();
    if (data.length < 3) return;

    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const adNoIdx = headers.indexOf((fieldMap['tut_adno'] || '').toLowerCase());
    const attIdx = headers.indexOf((fieldMap['tut_percent'] || '').toLowerCase());
    const possIdx = headers.indexOf((fieldMap['tut_poss'] || '').toLowerCase());
    const authIdx = headers.indexOf((fieldMap['tut_auth'] || '').toLowerCase());
    const unauthIdx = headers.indexOf((fieldMap['tut_unauth'] || '').toLowerCase());
    const latesIdx = headers.indexOf((fieldMap['tut_lates'] || '').toLowerCase());
    const psheIdx = headers.indexOf((fieldMap['tut_pshe'] || '').toLowerCase());

    if (adNoIdx === -1) return;

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rawAdNo = row[adNoIdx];
      if (!rawAdNo) continue;

      const adNo = String(rawAdNo).trim();
      if (studentMap[adNo]) {
        const rawPshe = psheIdx > -1 ? row[psheIdx] : '';
        const translatedPshe = this._translate(rawPshe, 'PSHE', translations);

        studentMap[adNo].tutorInfo = {
          percentAtt: attIdx > -1 ? row[attIdx] : '',
          possibleSessions: possIdx > -1 ? row[possIdx] : '',
          authAbsences: authIdx > -1 ? row[authIdx] : '',
          unauthAbsences: unauthIdx > -1 ? row[unauthIdx] : '',
          lates: latesIdx > -1 ? row[latesIdx] : '',
          pshe: translatedPshe
        };
      }
    }
  },

  /**
   * Reads an individual subject sheet, validates assessment completeness,
   * flags missing metrics, and appends the subject record to the student.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Active spreadsheet instance.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Individual subject sheet.
   * @param {Object<string, Object>} studentMap Master student dictionary.
   * @param {Object} fieldMap Column header mapping dictionary.
   * @param {Object} translations Category translation dictionary.
   * @param {Object} [reportConfig] Active report configuration.
   */
  _processSubjectSheet: function(ss, sheet, studentMap, fieldMap, translations, reportConfig) {
    const sheetName = sheet.getName();
    
    const nameRangeStr = `${sheetName}!${CONFIG.SCOPE.targetSubjectNameRange}`;
    const nameRange = ss.getRangeByName(nameRangeStr);
    const fullSubjectName = nameRange ? String(nameRange.getValue()).trim() : sheetName;

    const rangeName = `${sheetName}!thisSubjectAssessment`;
    const range = ss.getRangeByName(rangeName);
    if (!range) return; 

    const data = range.getValues();
    if (data.length < 3) return; 

    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const adNoColIdx = headers.indexOf((fieldMap['subj_adno'] || '').toLowerCase());
    const teacherIdx = headers.indexOf((fieldMap['subj_teacher'] || '').toLowerCase());
    const tgIdx = headers.indexOf((fieldMap['subj_tg'] || '').toLowerCase());
    const crntIdx = headers.indexOf((fieldMap['subj_crnt'] || '').toLowerCase());
    const ci1Idx = headers.indexOf((fieldMap['subj_ci1'] || '').toLowerCase());
    const ci2Idx = headers.indexOf((fieldMap['subj_ci2'] || '').toLowerCase());
    const ci3Idx = headers.indexOf((fieldMap['subj_ci3'] || '').toLowerCase());
    const ci4Idx = headers.indexOf((fieldMap['subj_ci4'] || '').toLowerCase());
    const ns1Idx = headers.indexOf((fieldMap['subj_ns1'] || '').toLowerCase());
    const ns2Idx = headers.indexOf((fieldMap['subj_ns2'] || '').toLowerCase());

    if (adNoColIdx === -1) return; 

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rawAdNo = row[adNoColIdx];
      
      if (!rawAdNo) continue; 
      const adNo = String(rawAdNo).trim();

      if (studentMap[adNo]) {
        const rawTg = tgIdx > -1 ? row[tgIdx] : '';
        const rawCrnt = crntIdx > -1 ? row[crntIdx] : '';
        const rawCi1 = ci1Idx > -1 ? row[ci1Idx] : '';
        const rawCi2 = ci2Idx > -1 ? row[ci2Idx] : '';
        const rawCi3 = ci3Idx > -1 ? row[ci3Idx] : '';
        const rawCi4 = ci4Idx > -1 ? row[ci4Idx] : '';
        const rawNs1 = ns1Idx > -1 ? String(row[ns1Idx]).trim() : '';
        const rawNs2 = ns2Idx > -1 ? String(row[ns2Idx]).trim() : '';

        // --- PRE-FLIGHT AUDIT VALIDATION ---
        // Flag empty values for core indicators so the UI modal can present actionable warnings.
        const missingElements = [];
        if (rawCrnt === '') missingElements.push('CRNT');
        if (rawCi1 === '') missingElements.push('CI1');
        if (rawCi2 === '') missingElements.push('CI2');
        if (rawCi3 === '') missingElements.push('CI3');
        if (rawCi4 === '') missingElements.push('CI4');
        
        // A next step entry is deemed missing only if both option columns are unselected
        if (rawNs1 === '' && rawNs2 === '') {
          missingElements.push('Next Steps');
        }

        if (missingElements.length > 0) {
          studentMap[adNo].auditIssues.push(`${fullSubjectName} (${missingElements.join(', ')})`);
        }
        // -----------------------------------

        const subjectData = {
          subjectName: fullSubjectName, 
          teacher: teacherIdx > -1 ? row[teacherIdx] : '',
          tg: this._translate(rawTg, 'CRNT', translations), 
          crnt: this._translate(rawCrnt, 'CRNT', translations), 
          ci1: this._translate(rawCi1, 'CI', translations),
          ci2: this._translate(rawCi2, 'CI', translations),
          ci3: this._translate(rawCi3, 'CI', translations),
          ci4: this._translate(rawCi4, 'CI', translations),
          nextSteps1: ns1Idx > -1 ? row[ns1Idx] : '', 
          nextSteps2: ns2Idx > -1 ? row[ns2Idx] : ''
        };

        studentMap[adNo].subjects.push(subjectData);
      }
    }
  }
};