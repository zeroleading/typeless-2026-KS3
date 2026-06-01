/**
 * DataService.gs
 * Handles data extraction and in-memory aggregation of student records.
 */

const DataService = {

  /**
   * Main function to build the complete data payload for a specific report.
   * @param {Object} reportConfig The specific report configuration from CONFIG.REPORTS.
   * @returns {Array<Object>} An array of fully populated student objects.
   */
  buildStudentDataPayload: function(reportConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Build the base student map keyed by Admission Number (adno)
    const studentMap = this._getMasterStudentList(ss);

    // 2. Loop through all sheets and process the Subject Sheets
    const allSheets = ss.getSheets();
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      if (subjectRegex.test(sheetName)) {
        // The Why: If it matches our Regex (e.g., 'Ar', 'Bi', or 'EnL'), 
        // we process it and append the grades to our student map.
        this._processSubjectSheet(ss, sheet, studentMap, reportConfig);
      }
    });

    // 3. Convert our map back into a flat array of student objects ready for DocumentBuilder
    return Object.values(studentMap);
  },

  /**
   * Helper: Gets the master student list and builds the initial map.
   * @private
   */
  _getMasterStudentList: function(ss) {
    const studentMap = {};
    
    // Fetch the new simpler named range directly
    const range = ss.getRangeByName('simpleStudentData');
    if (!range) return studentMap; // Safeguard if range is missing
    
    const data = range.getValues();

    // Loop through the data. Assuming row[0] = fullName, row[1] = adno, row[2] = reg, row[3] = tutor
    data.forEach(row => {
      const fullName = row[0];
      const rawAdNo = row[1];
      const reg = row[2];
      const tutor = row[3];
      
      // Ensure we have an adno and skip the header row if it exists in the named range
      if (rawAdNo && String(rawAdNo).toLowerCase() !== 'adno') { 
        const adNo = String(rawAdNo).padStart(6, '0');
        studentMap[adNo] = {
          adNo: adNo,
          name: fullName,
          reg: reg,
          tutor: tutor,
          subjects: [] // Initialise an empty array to hold the dynamic subjects
        };
      }
    });

    return studentMap;
  },

  /**
   * Helper: Reads a subject sheet and pushes data to the matching students.
   * @private
   */
  _processSubjectSheet: function(ss, sheet, studentMap, reportConfig) {
    const sheetName = sheet.getName();
    
    // Construct the string for the sheet-specific named range
    const rangeName = `${sheetName}!thisSubjectAssessment`;
    const range = ss.getRangeByName(rangeName);
    
    // Safeguard: If a sheet exists but the named range hasn't been set up yet, skip it.
    if (!range) return; 

    const data = range.getValues();
    
    // Safeguard: Ensure the table has at least headers, spill row, and one data row
    if (data.length < 3) return; 

    // The Quirk: Row 0 is headers. We convert them to lowercase and trim spaces for robust matching.
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    // Dynamically find column indices based on exact new header names
    const adNoColIdx = headers.indexOf('adno');
    
    // The Why: We map these strictly against the lowercased, trimmed versions of your headers.
    const tgIdx = headers.indexOf('tg');
    const eoyIdx = headers.indexOf('eoy');
    const ucasIdx = headers.indexOf('ucas');
    const rankIdx = headers.indexOf('rank');
    const ucasRefIdx = headers.indexOf('✎ ucas ref.');
    const ci1Idx = headers.indexOf('ci1');
    const ci2Idx = headers.indexOf('ci2');
    const ci3Idx = headers.indexOf('ci3');
    const ci4Idx = headers.indexOf('ci4');
    const crntIdx = headers.indexOf('crnt');
    const nextSteps1Idx = headers.indexOf('≣ nextsteps1');
    const nextSteps2Idx = headers.indexOf('≣ nextsteps2');
    const attIdx = headers.indexOf('att');

    // If we cannot find the admission number column, we cannot map the data.
    if (adNoColIdx === -1) return; 

    // The Quirk: Row 1 (index 1) handles importing spill functions, so we skip it.
    // We start our data extraction loop at Row 2 (index 2).
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rawAdNo = row[adNoColIdx];
      
      if (!rawAdNo) continue; // Skip completely empty rows
      
      const adNo = String(rawAdNo).padStart(6, '0');

      // Check if this student exists in our master map
      if (studentMap[adNo]) {
        
        // Build the subject specific object.
        // The Why: We check if the index exists (> -1) to prevent errors if a teacher deletes a column.
        const subjectData = {
          subjectName: sheetName,
          tg: tgIdx > -1 ? row[tgIdx] : '',
          eoy: eoyIdx > -1 ? row[eoyIdx] : '',
          ucas: ucasIdx > -1 ? row[ucasIdx] : '',
          rank: rankIdx > -1 ? row[rankIdx] : '',
          ucasRef: ucasRefIdx > -1 ? row[ucasRefIdx] : '',
          crnt: crntIdx > -1 ? row[crntIdx] : '',
          nextSteps: nextStepsIdx > -1 ? row[nextStepsIdx] : '',
          att: attIdx > -1 ? row[attIdx] : ''
        };

        // Push this subject into the student's profile
        studentMap[adNo].subjects.push(subjectData);
      }
    }
  }

};