/**
 * SheetServices.gs
 * Handles spreadsheet interaction events, such as onEdit triggers.
 */

const SheetServices = {
  
  /**
   * Main handler for edit events on the spreadsheet.
   * @param {Object} e The event object passed by the onEdit trigger.
   */
  handleOnEdit: function(e) {
    console.log('--- onEdit Trigger Fired ---');

    if (!e || !e.range) {
      console.log('Exit: No event object or range detected. (Was this run manually?)');
      return; 
    }
    
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    console.log(`Edit detected on sheet: ${sheetName}`);
    
    // 1. Gatekeeper: Exit immediately if not a subject sheet
    const subjectRegex = /^([A-Z][a-z]|EnL)$/;
    if (!subjectRegex.test(sheetName)) {
      console.log('Exit: Sheet name does not match subject regex.');
      return;
    }

    const ss = e.source;
    const rangeName = `${sheetName}!thisSubjectAssessment`;
    const tableRange = ss.getRangeByName(rangeName);
    
    if (!tableRange) {
      console.log(`Exit: Named range "${rangeName}" could not be found.`);
      return;
    }

    // 2. Map Reader: Calculate absolute physical boundaries of the named range on the grid
    const tableStartRow = tableRange.getRow();
    const tableStartCol = tableRange.getColumn();
    console.log(`Named range found starting at Row ${tableStartRow}, Column ${tableStartCol}`);
    
    // Extract headers and standardise them by removing ALL spaces to prevent accidental typos
    const headers = tableRange.getValues()[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
    
    // Find the relative indices (0-based) of our target columns
    const crntIdx = headers.indexOf('crnt');
    const ns1Idx = headers.indexOf('≣nextsteps1'); 
    const ns2Idx = headers.indexOf('≣nextsteps2');

    console.log(`Indices found -> CRNT: ${crntIdx}, NS1: ${ns1Idx}, NS2: ${ns2Idx}`);
    if (ns2Idx === -1) {
      console.log('NS2 not found. Exact headers seen by script:', headers);
    }

    // If there is no CRNT column found in the header, there is nothing to monitor
    if (crntIdx === -1) {
      console.log('Exit: "crnt" header not found in the named range.');
      return; 
    }
    
    // Convert the relative index to an absolute column number on the sheet
    const crntAbsCol = tableStartCol + crntIdx;
    console.log(`Calculated absolute CRNT column as: ${crntAbsCol}`);
    
    // 3. Intersection Check (Columns): Did the edit intersect the CRNT column?
    const editedColStart = e.range.getColumn();
    const editedColEnd = e.range.getLastColumn();
    
    if (crntAbsCol < editedColStart || crntAbsCol > editedColEnd) {
      console.log(`Exit: Edit occurred in col ${editedColStart}, which does not intersect CRNT col ${crntAbsCol}.`);
      return; 
    }

    // 4. Intersection Check (Rows): Did the edit happen on or below Row 3 of the named range?
    const dataStartRow = tableStartRow + 2; // Row 3 of the range (skipping headers and spill formulas)
    const editedRowStart = e.range.getRow();
    const editedRowEnd = e.range.getLastRow();

    if (editedRowEnd < dataStartRow) {
      return; 
    }

    // 5. Calculate exactly which rows to clear (handling multi-row selections gracefully)
    const startClearRow = Math.max(editedRowStart, dataStartRow);
    const numRowsToClear = editedRowEnd - startClearRow + 1;

    // 6. Action: Clear the dependent dropdowns silently
    if (ns1Idx > -1) {
      sheet.getRange(startClearRow, tableStartCol + ns1Idx, numRowsToClear, 1).clearContent();
    }
    
    if (ns2Idx > -1) {
      sheet.getRange(startClearRow, tableStartCol + ns2Idx, numRowsToClear, 1).clearContent();
    }
  }
};