var activeSheet = SpreadsheetApp.getActiveSheet();
var ui = SpreadsheetApp.getUi();

function cleanUpSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

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

    // Get the current date and time
    let now = new Date();
    let hour = now.getHours();
    console.log("hour: " + hour);

    // If it's before 3 PM (15:00), use yesterday's date
    // if (hour < 15) {
    //     now.setDate(now.getDate() - 1);
    // }

    // Format the date
    let date = Utilities.formatDate(now, "GMT-05:00", "M.d");
    let thisTabName = date;

    // Check for existing sheet names and add suffix if necessary
    let suffix = 1;
    while (spreadsheet.getSheetByName(thisTabName + (suffix > 1 ? ` (${suffix})` : ''))) {
        suffix++;
    }

    // Add suffix to the tab name if needed
    if (suffix > 1) {
        thisTabName = `${date} (${suffix})`;
    }

    // Get the active sheet and set its name
    let activeSheet = spreadsheet.getActiveSheet();
    activeSheet.setName(thisTabName);
  }
  updateSheetNameWithSuffix();

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

  function deleteRowsUntilType() {
    var lastRow = sheet.getLastRow();
    var rangeA = sheet.getRange('A1:A' + lastRow);
    var valuesA = rangeA.getValues();
    for (var i = 0; i < valuesA.length; i++) {
      if (valuesA[i][0] === "Type") {
        if (i > 0) {
          sheet.deleteRows(1, i); // Delete rows before the row containing "Type"
        }
        break;
      }
    }
  }

  function formatCellsAsPlainText() {
    sheet.getDataRange().setNumberFormat('@STRING@');
  }

  function removeBackquoteFromColumnB() {
    var rangeB = sheet.getRange('B1:B' + sheet.getLastRow());
    var valuesB = rangeB.getValues();
    for (var i = 0; i < valuesB.length; i++) {
      valuesB[i][0] = valuesB[i][0].replace(/`/g, '');
    }
    rangeB.setValues(valuesB);
  }

  function insertDifferenceColumn() {
    sheet.insertColumnAfter(4);
    var lastRow = sheet.getLastRow();
    var rangeE = sheet.getRange('E1:E' + lastRow);
    rangeE.setFormula('=C1-D1');
    sheet.getRange('E1').setValue('Difference'); // Set header for the new column
  }

  function alignCells() {
    sheet.getDataRange().setVerticalAlignment('MIDDLE');
    sheet.getDataRange().setHorizontalAlignment('CENTER');
  }

  function autoResizeAndSetColumnWidths() {
    sheet.autoResizeColumns(1, 5);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 250);
    sheet.autoResizeColumns(8, 14);
  }

  deleteRowsUntilType();
  formatCellsAsPlainText();
  removeBackquoteFromColumnB();
  insertDifferenceColumn();
  alignCells();
  autoResizeAndSetColumnWidths();
}


function onOpen() {
  ui.createMenu('Scripts')
    .addItem('Clean Up Sheet', 'cleanUpSheet')
    .addToUi();
}
