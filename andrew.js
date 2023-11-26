// This function updates specific cells in a Google Sheet based on certain rules
function updateCells() {
  // Access the currently active sheet in the Google Sheets document
  // This is the sheet that is currently open or being viewed
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Retrieve all the data from the active sheet and store it as a two-dimensional array
  // This means each row and column in the sheet is represented in this array
  var data = activeSheet.getDataRange().getValues();
  
  // Create an empty array to store results that will be written back to the sheet
  var result = []; 

  // Loop through each row in the data, starting from the second row (skipping the header row)
  for (var i = 1; i < data.length; i++) {
    // Define the criteria for a 'parent with children' row:
    // A row is considered a parent with children if Column A is "Catalog, Search" and Column B is "Configurable"
    var isParentWithChildren = data[i][0] === "Catalog, Search" && data[i][1] === "Configurable";

    // Define the criteria for a 'child' row:
    // A row is considered a child if Column A is "Not Visible" and Column B is "Simple"
    var isChild = data[i][0] === "Not Visible" && data[i][1] === "Simple";

    // If the row meets the criteria for a 'parent with children'
    if (isParentWithChildren) {
      // Start with an empty string that will be used to concatenate child row data
      var concatenatedString = "";

      // Loop through subsequent rows to find and concatenate child rows
      for (var j = i + 1; j < data.length && (data[j][0] === "Not Visible" && data[j][1] === "Simple"); j++) {
        // Use a helper function to concatenate data from each child row and add it to the string
        concatenatedString += concatenateChildRow(data[j]);
      }

      // Add the final concatenated string for this parent row to the results array
      result.push([concatenatedString]);
    } else if (isChild) {
      // If the row is a child row, concatenate its own data using the helper function
      result.push([concatenateChildRow(data[i])]);
    } else {
      // If the row is neither a 'parent with children' nor a 'child', leave the corresponding result cell empty
      result.push([""]);
    }
  }

  // Write the results back to the sheet, starting from the second row, in the 10th column (assumed to be Column J)
  activeSheet.getRange(2, 10, result.length, 1).setValues(result);
}

// This helper function concatenates data from specific columns of a child row
function concatenateChildRow(row) {
  // Initialize an empty string for concatenation
  var concatenatedString = row[2] + row[3]; // Start by concatenating values from Columns C and D

  // If Column F is not blank, concatenate Columns E and F
  concatenatedString += row[5] ? row[4] + row[5] : ""; 

  // If Column H is not blank, concatenate Columns G and H
  concatenatedString += row[7] ? row[6] + row[7] : "";

  // Always add the value from Column I
  concatenatedString += row[8];

  // Return the final concatenated string
  return concatenatedString;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Scripts')
    .addItem('Run Script', 'updateCells')
    .addToUi();
}
