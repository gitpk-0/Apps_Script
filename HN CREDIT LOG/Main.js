function onEdit(e) {
  updateDateAndInitialsCellsOnEdit(e);
  applyCellFormatting(e);
}

function updateDateAndInitialsCellsOnEdit(e) {
  let allOtherVendors = e.source.getSheetByName("||    ALL OTHER VENDORS    ||");
  let fourSeasons = e.source.getSheetByName("||    FOUR SEASONS    ||");
  
  updateDateAndInitialsInSheet(allOtherVendors, e);
  updateDateAndInitialsInSheet(fourSeasons, e);
}

function updateDateAndInitialsInSheet(sheet, event) {
  let watchedColumns = [1, 2, 3];
  let editedRange = event.range;
  
  if (watchedColumns.includes(editedRange.columnStart)) {
    let startRow = editedRange.getRow();
    let numRows = editedRange.getNumRows();

    for (let i = startRow; i < startRow + numRows; i++) {
      let valueOfColumn1 = sheet.getRange(i, 1).getValue();
      let valueOfColumn2 = sheet.getRange(i, 2).getValue();
      let valueOfColumn3 = sheet.getRange(i, 3).getValue();
      let dateCell = sheet.getRange(i, 12);
      let initialsCell = sheet.getRange(i, 11);

      if (valueOfColumn1 === "" && valueOfColumn2 === "" && valueOfColumn3 === "") {
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
  }
}


function applyCellFormatting(e) {
  var range = e.range;
  var columnIndex = range.getColumn();
  var sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== "Sale Sign Pull" && columnIndex <= 6 && columnIndex !== 4) {
    range.setFontWeight("normal");
    range.setFontFamily("Arial");
    range.setFontSize(11);
    range.setFontLine("none");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    range.setFontColor(null);
    range.setBackground(null);
  }
}

function extractInitialsFromEmail(email) {
  let parts = email.split("@")[0].split(".");
  let initials = parts.map(part => part.charAt(0).toUpperCase()).join("");
  return initials;
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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
      .addItem('Archive Committed Items', 'archiveCommittedItems')
      .addToUi();
}
