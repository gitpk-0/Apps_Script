var activeSheet = SpreadsheetApp.getActiveSheet();
var ui = SpreadsheetApp.getUi();

function allFunctions() {

  function moveSheetToFirstPosition() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = activeSheet;

    if (sheet.getIndex() !== 1) {
      spreadsheet.setActiveSheet(sheet);
      spreadsheet.moveActiveSheet(1);
    }
  }
  moveSheetToFirstPosition();

  function deleteSheetsStartingFromPosition21() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
    var sheetsToDelete = [];

    if (sheets.length > 19) {
        // Gather a list of sheets to be deleted
      for (var i = 20; i < sheets.length; i++) {
        sheetsToDelete.push(sheets[i]);
      }
    }

    // Delete the gathered sheets
    sheetsToDelete.forEach(function(sheetToDelete) {
      spreadsheet.setActiveSheet(sheetToDelete);
      spreadsheet.deleteActiveSheet();
    });
  }
  deleteSheetsStartingFromPosition21();
  

  function hideLastVisibleTab() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
    
    // Show the first 5 tabs and hide the rest
    for (var i = 0; i < sheets.length; i++) {
      if (i < 5) {
        sheets[i].showSheet();
      } else {
        sheets[i].hideSheet();
      }
    }
  }
  hideLastVisibleTab();

  

  function freezeFirstRow() {
    activeSheet.setFrozenRows(1);
  }
  freezeFirstRow();

  function formatAndRemoveBackquoteFromUPC() {
    var upcColumn = activeSheet.getRange("B:B");
    upcColumn.setNumberFormat('@STRING@');
    upcColumn.createTextFinder("`").replaceAllWith("");
  }
  formatAndRemoveBackquoteFromUPC();

  function updateSheetNameWithSuffix() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
    let thisTabName = date;

    let suffix = 1;
    while (spreadsheet.getSheetByName(thisTabName + (suffix > 1 ? ` (${suffix})` : ''))) {
      suffix++;
    }

    if (suffix > 1) {
      thisTabName = `${date} (${suffix})`;
    }

    activeSheet.setName(thisTabName);
  }
  updateSheetNameWithSuffix();

  activeSheet.autoResizeColumns(1, 11);
  activeSheet.setColumnWidth(3, 170);
  activeSheet.setColumnWidth(4, 360);

  var ohCol = activeSheet.getRange("G:G");
  ohCol.setNumberFormat("0.00");
  var ooCol = activeSheet.getRange("I:I");
  ooCol.setNumberFormat("0.00");
}

function frankHole() {
  var alert = ui.alert('ALERT!', 'Why did you press the red button!?', ui.ButtonSet.OK);
}

function onOpen() {
  ui.createMenu('Scripts')
    .addItem('Clean Up Script', 'allFunctions')
    .addItem('Red Button', 'frankHole')
    .addToUi();
}



// function deleteSheetsNotStartingWith8() {
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var sheets = spreadsheet.getSheets();

//   for (var i = 0; i < sheets.length; i++) {
//     var sheet = sheets[i];
//     var sheetName = sheet.getName();
    
//     if (!sheetName.match(/^8/)) {
//       spreadsheet.deleteSheet(sheet);
//     }
//   }
// }