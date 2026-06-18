/**
 * DataService.gs
 * Handles data extraction, in-memory aggregation, and translation of student records.
 */

const DataService = {

  buildStudentDataPayload: function(reportConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Fetch Global Batch Values
    const yearGroup = ss.getRangeByName(CONFIG.SCOPE.yearGroup)?.getValue() || '';
    const collection = ss.getRangeByName(CONFIG.SCOPE.collection)?.getValue() || '';
    const academicYear = ss.getRangeByName(CONFIG.SCOPE.academicYear)?.getValue() || '';
    const shortName = ss.getRangeByName(CONFIG.SCOPE.shortName)?.getValue() || '';

    // 2. Fetch Control Panel Maps & Dictionaries
    const fieldMap = this._getDynamicFieldMap(ss);
    const translations = this._getTranslationsDictionary(ss);

    // 3. Build base maps and attach data
    const studentMap = this._getMasterStudentList(ss);
    this._attachTutorData(ss, studentMap, fieldMap, translations);

    // 4. Process Subject Sheets
    const allSheets = ss.getSheets();
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;

    allSheets.forEach(sheet => {
      if (subjectRegex.test(sheet.getName())) {
        this._processSubjectSheet(ss, sheet, studentMap, fieldMap, translations);
      }
    });

    // 5. Convert map to array and inject globals
    return Object.values(studentMap).map(student => ({
      ...student,
      yearGroup: yearGroup,
      collection: collection,
      academicYear: academicYear,
      shortName: shortName
    }));
  },

  _getDynamicFieldMap: function(ss) {
    const map = { ...CONFIG.FALLBACK_FIELD_MAP }; 
    const range = ss.getRangeByName(CONFIG.SCOPE.fieldMap);
    if (!range) return map;

    const data = range.getValues();
    data.forEach(row => {
      const internalRef = String(row[0]).trim();
      const targetHeader = String(row[1]).trim();
      if (internalRef && targetHeader && !internalRef.includes('**')) {
        map[internalRef] = targetHeader;
      }
    });
    return map;
  },

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

  _translate: function(rawValue, category, translationsDict) {
    if (rawValue === '' || rawValue === undefined) return '';
    const safeValue = String(rawValue).trim().toUpperCase();
    if (translationsDict[category] && translationsDict[category][safeValue]) {
      return translationsDict[category][safeValue];
    }
    return String(rawValue); 
  },

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
          auditIssues: [] // <-- NEW: Array to hold our missing data warnings
        };
      }
    });
    return studentMap;
  },

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

  _processSubjectSheet: function(ss, sheet, studentMap, fieldMap, translations) {
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

        // --- NEW: AUDIT CHECK ---
        let missingElements = [];
        if (rawCrnt === '') missingElements.push('CRNT');
        if (rawCi1 === '') missingElements.push('CI1');
        if (rawCi2 === '') missingElements.push('CI2');
        if (rawCi3 === '') missingElements.push('CI3');
        if (rawCi4 === '') missingElements.push('CI4');

        if (missingElements.length > 0) {
          studentMap[adNo].auditIssues.push(`${fullSubjectName} (${missingElements.join(', ')})`);
        }
        // ------------------------

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