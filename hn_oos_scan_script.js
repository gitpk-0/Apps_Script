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

    // Get the current date and time
    let now = new Date();
    let hour = now.getHours();
    console.log("hour: " + hour);

    // If it's before 3 PM (15:00), use yesterday's date
    if (hour < 15) {
        now.setDate(now.getDate() - 1);
    }

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

  function setColumnJHeader() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // Set the value of the first cell in column J to "Notes"
    sheet.getRange("J1").setValue("Notes");
  }
  setColumnJHeader();

  function sortDataByLocation() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // Assuming the data starts from row 2 (row 1 is headers)
    const range = sheet.getRange("A2:J" + sheet.getLastRow());

    // Sort the data by the 8th column (Location) in descending order (Z to A)
    range.sort({column: 8, ascending: false});
  }
  sortDataByLocation();

  function formatOnhandColToNum() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    // Assuming the data starts from row 2 (row 1 is headers)
    const range = sheet.getRange("G2:G" + sheet.getLastRow());

    // Set the number format for the range to display two decimal points
    range.setNumberFormat("0.00");
  }
  formatOnhandColToNum();


  activeSheet.autoResizeColumns(1, 11);
  activeSheet.setColumnWidth(3, 170);
  activeSheet.setColumnWidth(4, 360);
  activeSheet.setColumnWidth(10, 195);

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
