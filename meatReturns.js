// ALL HEADERS [global]
var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");


function meatReturnsFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  // header
  s.insertRowBefore(1); // inserts new empty first row

  var meat = headers.getRange('A1:J1');
  var destinationS = s.getRange("A1:J1");
  meat.copyTo(destinationS); // pastes headers to first row

  // horizontal align -- center (columns H thru AK)
  var colRange = s.getRange("D:J");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // resize columns
  // setColumnWidth(column number, pixel width)  -- single column
  // setColumnWidths(column number, number of columns, pixel width) -- multiple columns
  s.setColumnWidth(1, 105);
  s.setColumnWidth(2, 265);
  s.setColumnWidth(3, 185);
  s.setColumnWidth(4, 50);
  s.setColumnWidths(5, 6, 96);

}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('Meat Returns Formatting', 'meatReturnsFormat')   
    .addSeparator()

    .addToUi(); 
}


// to do
// update alteryx workflow for currency and percentage columns