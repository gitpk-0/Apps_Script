function onOpen() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var s = ss.getActiveSheet();
}

function onEdit(e) {

  dateAdded(e);

  var range = e.range;
  var columnIndex = range.getColumn();
  
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

  function dateAdded(e) {
    let ss = e.source.getSheetByName("VENDORS");
    
    let watchedCols = [1,2,3];
    if (watchedCols.indexOf(e.range.columnStart) === -1) return;

    let row = e.range.getRow();
    let col1Value = ss.getRange(row, 1).getValue();
    let col2Value = ss.getRange(row, 2).getValue();
    let col3Value = ss.getRange(row, 3).getValue();
    let dateCell = ss.getRange(row, 11);
    let initialsCell = ss.getRange(row, 10);

    // Check if all the cells in columns 1, 2, and 3 are empty
    if (col1Value === "" && col2Value === "" && col3Value === "") {
      dateCell.setValue(null); // Clear the date value
      initialsCell.setValue(null); // Clear the initials
    } else {
      // Check if the initials cell is empty before setting the initials
      if (!initialsCell.getValue()) {
        // Extract initials from the Gmail address
        let email = Session.getActiveUser().getEmail();
        let initials = extractInitialsFromEmail(email);
        initialsCell.setValue(initials);
      }

      let date = Utilities.formatDate(new Date(), "GMT-05:00", "M/d/YYYY");
      // If the date cell is empty, set the date value
      if (!dateCell.getValue()) {
        console.log("setting date");
        dateCell.setValue(date);
      }
    }

    // Helper function to extract initials from an email
    function extractInitialsFromEmail(email) {
      let parts = email.split("@")[0].split(".");
      let initials = parts.map(part => part.charAt(0).toUpperCase()).join("");
      return initials;
    }
  }
}