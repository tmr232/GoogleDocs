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


/*************************************************************************/
/* Functions for reading and copying filtered rows */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var rows2 = sheet.getRange(1,1,4,4);
  var rows3 = sheet.getRange(4,1,4,4);
  var numRows = rows2.getNumRows();
  var values = rows2.getValues();
  var testRows = [];
  
  testRows = getRows(sheet, 1, 1, 4, 4, makeFilterFirstCell("Test"));
  var tRange = setRows(sheet, testRows, 6, 1);
  tRange.sort([1,2,3,4]);
  
};

function makeFilterFirstCell(value) {
  filter = function(row) {
    return value == row[0];
  };
  
  return filter;
}

function getRows(sheet, sourceRow, sourceColumn, numRows, numColumns, filter) {
  var resultRows = [];
  
  var sourceRows = sheet.getRange(sourceRow, sourceColumn, numRows, numColumns).getValues();
  
  var filter = filter || function() {return true;};
  
  for (var i = 0; i < numRows; ++i) {
    var row = sourceRows[i];
    if (filter(row)) {
      resultRows.push(row);
    }
  }
  
  return resultRows;
}

function setRows(targetSheet, sourceRows, targetRow, targetColumn, rowLength) {
  var numColumns = rowLength || sourceRows[0].length;
  var numRows = sourceRows.length;
  
  var targetRange = targetSheet.getRange(targetRow, targetColumn, numRows, numColumns);
  
  targetRange.setValues(sourceRows);
  
  return targetRange;
}
