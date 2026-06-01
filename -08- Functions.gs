/**
 * Returns a list of sheet names and link IDs that match a given Regular Expression.
 *
 * @param {"^202.*"} include_regex A regular expression string, or a range of regex strings.
 * @param {FALSE} trigger Optional: Checkbox reference to force a refresh (Boolean).
 * @return A 2-column array of matching Names and Link IDs.
 * @customfunction
 */
function GETSHEETNAMES_REGEX(include_regex, trigger) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let patterns = [];

  // Handle the regex input
  if (include_regex) {
    if (Array.isArray(include_regex)) {
      // If input is a range, flatten it and remove any empty cells
      patterns = include_regex.flat().filter(String).map(String);
    } else {
      // If input is a single cell, string, or number, treat it as one pattern
      patterns = [String(include_regex)];
    }
  }

  // If no pattern is provided, alert the user
  if (patterns.length === 0) {
    return [["Error", "Please provide a valid regex pattern."]];
  }

  // Convert string patterns into actual RegExp objects
  let regexObjects = [];
  try {
    // Note: We are using case-sensitive matching by default. 
    // To make it case-insensitive, change `new RegExp(p)` to `new RegExp(p, 'i')`
    regexObjects = patterns.map(p => new RegExp(p)); 
  } catch (e) {
    // Catch syntax errors in the user's regex
    return [["Regex Error", e.message]];
  }

  // Filter sheets: keep if the name matches ANY of the provided regex patterns
  const output = sheets
    .filter(sheet => {
      const name = sheet.getName();
      return regexObjects.some(regex => regex.test(name));
    })
    .map(sheet => [sheet.getName(), `#gid=${sheet.getSheetId()}`]);

  // Return the array, or a fallback message if nothing matches
  return output.length > 0 ? output : [["No sheets matched", ""]];
}