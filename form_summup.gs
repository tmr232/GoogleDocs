/* Usage:
 *   To copy to a new file, add an ID (like SUMMUP_ID), and add a
 *   new call to copySummup.
 */

var SUMMUP_ID = "document_key";

function onFormSubmit() {
  copySummup(SUMMUP_ID, 1, 5, [5,4,3,2]);
}

/* copySummup
*  Parameters:
*  docID - the document key copied from the URL
*  firstColumn - the column to start copying from
*  numberOfColumns - the number of columns to copy
*  sortObject - the sorting parameters: [1,3,5,7] (each number is a column to sort by.
*               the first one is the primary sort, the second is secondary, and so on.)
*/
function copySummup(docID, firstColumn, numberOfColumns, sortObject, targetColumn) {
   // Open the current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  // Open the target spreadsheet
  var summupSs = SpreadsheetApp.openById(docID);
  var summupSheet = summupSs.getSheets()[0];
  
  // Get the range to copy
  var sourceRange = sheet.getRange(1, firstColumn, sheet.getMaxRows(), numberOfColumns);
  
  // Get the range to copy to
  targetColumn = targetColumn || 1;
  var destRange = summupSheet.getRange(1, targetColumn, sheet.getMaxRows(), numberOfColumns);
  
  // Copy the source range to the destination range
  destRange.setValues(sourceRange.getValues());
  
  // Get the range to sort (without the headers row)
  var sortRange = summupSheet.getRange(2, 1, sheet.getMaxRows() - 1, numberOfColumns);
  // Sort the range
  try {
    sortRange.sort(sortObject);
  } catch(e) {
  }
}
