/**
 * Triggered automatically by Google Sheets when a user edits a cell.
 * @param {Object} e The edit event object.
 */
function onEdit(e) {
  // Delegate the logic to our dedicated SheetServices module
  SheetServices.handleOnEdit(e);
}