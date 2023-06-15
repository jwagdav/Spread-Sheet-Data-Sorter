function deleteAllEmptyRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var numRows = values.length;
  var numCols = values[0].length;
  
  var emptyRows = [];
  
  // Find all completely empty rows
  for (var i = 0; i < numRows; i++) {
    var empty = true;
    for (var j = 0; j < numCols; j++) {
      if (values[i][j] != "") {
        empty = false;
        break;
      }
    }
    
    // Mark the row as empty
    if (empty) {
      emptyRows.push(i + 1); // rows are 1-indexed
    }
  }
  
  var rowsToDelete = emptyRows.length;
  
  if (rowsToDelete > 0) {
    // Sort the array in descending order
    emptyRows.sort(function(a, b) { return b - a; });
    
    for (var i = 0; i < rowsToDelete; i++) {
      // Delete rows in batches
      var startRow = emptyRows[i];
      var endRow = startRow;
      while (i < rowsToDelete - 1 && emptyRows[i+1] == emptyRows[i] - 1) {
        endRow = emptyRows[i+1];
        i++;
      }
      sheet.deleteRows(endRow, startRow - endRow + 1);
    }
  }
}
