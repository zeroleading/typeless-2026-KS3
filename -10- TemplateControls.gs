// ⚠️ Grants sheet and range editing access to configured users - verify list in CONFIG.PERMISSIONS.PROTECTION_EDITORS.
function createCopyWithPermissions() {
  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  
  // Make a copy of the file in the same folder
  const file = DriveApp.getFileById(sourceSs.getId());
  const destFolder = file.getParents().next();
  const newFile = file.makeCopy("Copy of " + sourceSs.getName(), destFolder);
  
  const newSs = SpreadsheetApp.openById(newFile.getId());
  
  // Pull protection editors directly from the immutable CONFIG object
  const superusers = CONFIG.PERMISSIONS.PROTECTION_EDITORS;
  
  // Apply permissions across all protected sheets and ranges
  const sheets = newSs.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    // Sheet-level protections
    let sheetProtections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (let j = 0; j < sheetProtections.length; j++) {
      if (!sheetProtections[j].isWarningOnly()) {
        sheetProtections[j].addEditors(superusers);
      }
    }
    
    // Range-level protections
    let rangeProtections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (let k = 0; k < rangeProtections.length; k++) {
      if (!rangeProtections[k].isWarningOnly()) {
        rangeProtections[k].addEditors(superusers);
      }
    }
  }
}
