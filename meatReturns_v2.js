// ALL HEADERS [global]
var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Headers");


function meatReturnsFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  // header
  s.insertRowBefore(1); // inserts new empty first row
  // s.insertRowBefore(1); // inserts another new empty first row

  var meat = headers.getRange('A1:L2');
  var destinationS = s.getRange("A1:L2");  // current tab destination
  meat.copyTo(destinationS); // pastes headers to first row

  // horizontal align -- center (columns H thru AK)
  var colRange = s.getRange("C:L");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // resize columns
  // setColumnWidth(column number, pixel width)  -- single column
  // setColumnWidths(column number, number of columns, pixel width) -- multiple columns
  s.setColumnWidth(1, 105);
  s.setColumnWidth(2, 265);
  s.setColumnWidth(3, 185);
  s.setColumnWidth(4, 50);
  s.setColumnWidths(3, 8, 80);
  s.setColumnWidths(11, 2, 100);

  // freeze rows and columns
  s.setFrozenRows(2);
  s.setFrozenColumns(2);

  // background colors
  var days7 = s.getRange("C:F");
  var days30 = s.getRange("G:J");
  var rate = s.getRange("K:L");
  days7.setBackground("#fff2cc");
  days30.setBackground("#f4cccc");
  rate.setBackground("#c9daf8")
}



function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('Meat Returns Formatting', 'meatReturnsFormat')   
    .addSeparator()

    .addToUi(); 
}