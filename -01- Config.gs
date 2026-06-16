/**
 * Config.gs
 * Global Configuration File
 * Acts as the single source of truth for the entire reporting system.
 */

const CONFIG = Object.freeze({
  
  // 1. Global Settings
  GLOBAL: {
    CONTROL_SHEET_NAME: 'control',
    OUTPUT_FOLDER_ID: '18JTL77flcaOV7Us93W_hJnf3_veATeFp', 
  },

  // 2. Authorisation Controls
  AUTH: {
    SUPER_USERS: [
      'jappleton@csg.school',
      'tnayagam@csg.school'
    ],
    REPORT_SPECIFIC: {}
  },

  // 3. Import Sheet Controls
  IMPORT: {
    targetSheetName: 'import',
    backupSheetName: 'import-backup',
    anchorRowStart: 6,
    anchorRowCount: 2, 
    statusCell: 'A1'   
  },

  // 4. Setup & Map Controls
  SCOPE: {
    subjectDetailsRange: 'scopeSubjectDetails',
    yearGroup: 'scopeYearGroup',
    keyStage: 'scopeKeyStage',
    academicYear: 'scopeAcademicYear',
    collection: 'scopeCollection',
    targetSubjectNameRange: 'thisSubjectName',
    
    // New dynamic tables on the Control Panel
    fieldMap: 'scopeFieldMap',
    translations: 'scopeTranslations'
  },

  // 5. Fallback Field Mapper 
  // Used only if the scopeFieldMap named range is missing or broken.
  FALLBACK_FIELD_MAP: {
    tut_adno: 'adno',
    tut_percent: 'att %',
    tut_poss: 'possible sessions',
    tut_auth: 'auth abs',
    tut_unauth: 'unauth abs',
    tut_lates: 'lates',
    tut_pshe: 'pshe engagement',
    subj_adno: 'adno',
    subj_teacher: 'teacher',
    subj_tg: 'tg',
    subj_crnt: 'crnt',
    subj_ci1: 'ci1',
    subj_ci2: 'ci2',
    subj_ci3: 'ci3',
    subj_ci4: 'ci4',
    subj_ns1: '≣ nextsteps1',
    subj_ns2: '≣ nextsteps2'
  },

  // 6. Report Profiles
  REPORTS: {
    PROGRESS_REVIEW: {
      name: 'Progress Review',
      templateId: '1G7XZ9CxNBsQ2JMx4xX9mAA4ap_ZdDLTGvnCLPDzXo-U'
    },
    NEXT_STEPS_SUMMARY: {
      name: 'Next Steps Summary',
      templateId: '1Z6O8k6C67vDBp3heHZ-9reT5Glfnc8lLZ-dqpUivf74'
    }
  }
});