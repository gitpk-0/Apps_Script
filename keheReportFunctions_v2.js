// ALL HEADERS [global]
var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");

/*  ********************  */
/*  OH OO TAB FORMATTING  */
/*  ********************  */

function ohOoFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  // date formatting
  var dateRange = s.getRange("L:N");
  dateRange.setNumberFormat("m/dd/yyyy");


  // horizontal align -- center (columns H thru AJ)
  var colRange = s.getRange("H:AJ");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


  // header
  var ohOo = headers.getRange('A2:AJ2');
  s.insertRowBefore(1); // inserts new empty first row
  var destinationS = s.getRange("A1:AJ1");
  ohOo.copyTo(destinationS); // pastes headers to first row


  // freeze rows and columns
  s.setFrozenRows(1);
  s.setFrozenColumns(7);


  // duplicate UPC conditional formatting

  // invalid APT conditional formatting
  var aptRange = s.getRange("N:N");
  var aptRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(new Date("01/01/2001"))
    .setBackground("#f4cccc")
    .setRanges([aptRange])
    .build();
  var aptRules = s.getConditionalFormatRules();
  aptRules.push(aptRule);
  s.setConditionalFormatRules(aptRules);
  

}



/*  ********************  */
/*  PROMO TAB FORMATTING  */
/*  ********************  */

function promoFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  // date formatting
  var dateRange = s.getRange("L:N");
  dateRange.setNumberFormat("m/dd/yyyy");


  // horizontal align -- center, (columns H thru AA)
  var colRange = s.getRange("H:AA");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


  // header vertical align -- middle
  var headRange = s.getRange("A2:AA2");
  headRange.setVerticalAlignment("middle");


  // header
  var promo = headers.getRange('A4:AA5');
  s.insertRowBefore(1); // inserts new empty first row
  s.insertRowBefore(1); // inserts another new empty first row
  var destinationS = s.getRange("A1:AA2");
  promo.copyTo(destinationS); // pastes headers to first row


  // freeze rows and columns
  s.setFrozenRows(2);
  s.setFrozenColumns(7);


  // invalid APT conditional formatting
  var aptRange = s.getRange("N:N");
  var aptRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(new Date("01/01/2001"))
    .setBackground("#f4cccc")
    .setRanges([aptRange])
    .build();
  var aptRules = s.getConditionalFormatRules();
  aptRules.push(aptRule);
  s.setConditionalFormatRules(aptRules);

}



/*  *********************  */
/*  ENDCAP TAB FORMATTING  */
/*  *********************  */

function endcapFormat() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  // date formatting
  var dateRange = s.getRange("I:J");
  dateRange.setNumberFormat("m/dd/yyyy");


  // horizontal align -- center, vertical align -- middle (all cells)
  var colRange = s.getRange("A:L");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle").setVerticalAlignment("middle");


  // header
  var endcap = headers.getRange('A7:L7');
  s.insertRowBefore(1); // inserts new empty first row
  var destinationS = s.getRange("A1:L1");
  endcap.copyTo(destinationS); // pastes headers to first row


  // freeze header row
  s.setFrozenRows(1);


  // invalid APT conditional formatting
  var aptRange = s.getRange("N:N");
  var aptRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(new Date("01/01/2001"))
    .setBackground("#f4cccc")
    .setRanges([aptRange])
    .build();
  var aptRules = s.getConditionalFormatRules();
  aptRules.push(aptRule);
  s.setConditionalFormatRules(aptRules);

}


function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('OH OO Tab Formatting', 'ohOoFormat')   
    .addSeparator()
    .addItem('Promo Tab Formatting', 'promoFormat')
    .addSeparator()
    .addItem('Endcap Tab Formatting', 'endcapFormat')

    .addToUi(); 
}








// // need to implement:
// convert decimal to percentage 
// convert to currency
// autoresize columns (autoResizeColumns())  **** header rows prevent this from working properly
// duplicate UPC conditional formatting

