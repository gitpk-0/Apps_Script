function onEdit(e) {

  dateAdded(e);

  var range = e.range;
  var columnIndex = range.getColumn();
  var sheet = e.source.getActiveSheet();
  
  // Check if the edited sheet is not "Sale Sign Pull"
  if (sheet.getName() !== "Sale Sign Pull") {
    var range = e.range;
    var columnIndex = range.getColumn();
    
    // Check if columnIndex is less than or equal to 6 and not equal to 4
    if (columnIndex <= 6 && columnIndex !== 4) {
      // Apply your desired cell formatting here
      range.setFontWeight("normal"); // Set font weight to normal
      range.setFontFamily("Arial"); // Set font family to Arial
      range.setFontSize(11); // Set font size to 11
      range.setFontLine("none"); // Remove underline or strikethrough
      range.setHorizontalAlignment("center"); // Center the text horizontally
      range.setVerticalAlignment("middle"); // Center the text vertically
      range.setFontColor(null); // Set the font color to the default (automatic)
      range.setBackground(null); // Set the background color to the default (none)
      // Add more formatting options as needed
    }
  }

  function dateAdded(e) {
    let allOtherVendors = e.source.getSheetByName("||    ALL OTHER VENDORS    ||");
    let fourSeasons = e.source.getSheetByName("||    FOUR SEASONS    ||");
    
    // Process each sheet separately
    processSheet(allOtherVendors, e);
    processSheet(fourSeasons, e);

    function processSheet(sheet, event) {
      let watchedCols = [1,2,3];
      if (watchedCols.indexOf(event.range.columnStart) === -1) return;

      let row = event.range.getRow();
      let col1Value = sheet.getRange(row, 1).getValue();
      let col2Value = sheet.getRange(row, 2).getValue();
      let col3Value = sheet.getRange(row, 3).getValue();
      let dateCell = sheet.getRange(row, 12);
      let initialsCell = sheet.getRange(row, 11);

      if (col1Value === "" && col2Value === "" && col3Value === "") {
        dateCell.setValue(null);
        initialsCell.setValue(null);
      } else {
        if (!initialsCell.getValue()) {
          let email = Session.getActiveUser().getEmail();
          let initials = extractInitialsFromEmail(email);
          initialsCell.setValue(initials);
        }
        let date = Utilities.formatDate(new Date(), "GMT-05:00", "M/d/YYYY");
        if (!dateCell.getValue()) {
          dateCell.setValue(date);
        }
      }
    }

    function extractInitialsFromEmail(email) {
      let parts = email.split("@")[0].split(".");
      let initials = parts.map(part => part.charAt(0).toUpperCase()).join("");
      return initials;
    }
  }
}

function archiveCommittedItems() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = sheet.getSheetByName("Archive"); // Target tab name

  // Array of source sheets
  var sourceSheets = [
    sheet.getSheetByName("||    ALL OTHER VENDORS    ||"),
    sheet.getSheetByName("||    FOUR SEASONS    ||")
  ];

  // Process each source sheet
  sourceSheets.forEach(function(sourceSheet) {
    var data = sourceSheet.getDataRange().getValues(); // Get all data from the current source sheet

    // Iterate backwards to avoid index shifting when deleting rows
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][6] === true) { // Check if the cell in column G (index 6) is TRUE
        // Prepend a single quote to preserve leading zeros for columns A, B, and H
        data[i][0] = "'" + data[i][0]; // Column A
        data[i][1] = "'" + data[i][1]; // Column B
        data[i][7] = "'" + data[i][7]; // Column H

        targetSheet.appendRow(data[i]); // Append the entire row to the target sheet

        var appendedRowIndex = targetSheet.getLastRow(); // Get the index of the last row

        // Process the columns A, B, and H
        [1, 2, 8].forEach(function(columnIndex) {
          var cell = targetSheet.getRange(appendedRowIndex, columnIndex);
          cell.setNumberFormat("@"); // Set the number format to plain text

          // Remove the leading single quote from the displayed value
          var currentValue = cell.getDisplayValue();
          if (currentValue.startsWith("'")) {
            cell.setValue(currentValue.substring(1));
          }
        });

        sourceSheet.deleteRow(i + 1); // Delete the row from the source sheet
      }
    }
  });
}


// untested optimization of archiveCommittedItems function:
// function archiveCommittedItems() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet();
//   var targetSheet = sheet.getSheetByName("Archive"); // Target tab name

//   // Changed: Use sheet names directly in the array to simplify the logic
//   var sourceSheets = [
//     "||    ALL OTHER VENDORS    ||",
//     "||    FOUR SEASONS    ||"
//   ];

//   sourceSheets.forEach(function(sheetName) {
//     var sourceSheet = sheet.getSheetByName(sheetName);
//     var data = sourceSheet.getDataRange().getValues(); // Get all data from the current source sheet
//     var rowsToDelete = []; // Changed: Collect rows to delete in a batch
//     var dataToAppend = []; // Changed: Collect data to append in a batch

//     // Process data in memory to minimize Spreadsheet operations
//     for (var i = data.length - 1; i >= 0; i--) {
//       if (data[i][6] === true) { // Check if the cell in column G (index 6) is TRUE
//         // Adjust data for columns A, B, and H without prepending single quote
//         var rowData = data[i];

//         // Changed: Handle text conversion in memory to avoid manipulating cell formats individually later
//         rowData[7] = rowData[7].toString(); // Ensure column H is treated as text

//         dataToAppend.push(rowData); // Collect data to append
//         rowsToDelete.push(i + 1); // Collect row numbers to delete
//       }
//     }

//     // Append data in batches to reduce the number of appendRow calls
//     if (dataToAppend.length > 0) {
//       dataToAppend.reverse(); // Reverse to maintain original order when appending
//       dataToAppend.forEach(function(row) {
//         targetSheet.appendRow(row); // Append each row from the collected data
//       });
//     }

//     // Delete rows in batches to improve efficiency
//     // Note: Actual deletion is still one by one due to Google Apps Script limitations,
//     // but collecting and processing this way reduces the number of Spreadsheet operations.
//     rowsToDelete.reverse().forEach(function(rowNum) {
//       sourceSheet.deleteRow(rowNum); // Delete each row from the bottom up to avoid index issues
//     });
//   });
// }

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
      .addItem('Archive Committed Items', 'archiveCommittedItems')
      .addToUi();
}
