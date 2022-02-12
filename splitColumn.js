function splitColumn(range) {
  var output = [];

  for(var i in range) {
    var split = range[i][0].split(",");

    if(split.length == 1) {
      output.push([split[0]]);
    } else {
      for(var j in split) {
        output.push([split[j]]);
      }
    }
  }
  return output;  
}

function SplitColumnValues(sheet_name, column_index, delimiter) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheet_name);
    var lastRow = sheet.getLastRow();
  
    var range = sheet.getRange(1,column_index,lastRow-1, 1);
    range.splitTextToColumns(delimiter);
  }