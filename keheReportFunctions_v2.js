// ALL HEADERS [global]
var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");

/*  ********************  */
/*  OH OO TAB FORMATTING  */
/*  ********************  */

function ohOoFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating doc and sheet names with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let docName = "KeHE Report (updated " + date + ")";  
  ss.setName(docName);

  let mainTabName = "OH OO " + date;
  s.setName(mainTabName);


  // date formatting
  var dateRange = s.getRange("L:N");
  dateRange.setNumberFormat("m/dd/yyyy");


  // horizontal align -- center (columns H thru AH)
  var colRange = s.getRange("H:AH");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


  // header
  var ohOo = headers.getRange('A2:AH2');
  s.insertRowBefore(1); // inserts new empty first row
  var destinationS = s.getRange("A1:AH1");
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


  // ETA passed conditional formatting
  var currentDate = new Date();
  var etaRange = s.getRange("L:L");
  var etaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(currentDate)
    .setBackground("#f4cccc")
    .setRanges([etaRange])
    .build();
  var etaRules = s.getConditionalFormatRules();
  etaRules.push(etaRule);
  s.setConditionalFormatRules(etaRules);


  // No ETA conditional formatting
  // var currentDate = new Date();
  // var etaRange = s.getRange("L:L");
  // var etaRule = SpreadsheetApp.newConditionalFormatRule()
  //   .whenDateEqualTo(undefined)
  //   .setBackground("#111111")
  //   .setRanges([etaRange])
  //   .build();
  // var etaRules = s.getConditionalFormatRules();
  // etaRules.push(etaRule);
  // s.setConditionalFormatRules(etaRules);


  // ETA too far out conditional formatting
  var currentDate = new Date();
  var tooFarDate = new Date(currentDate.setDate(currentDate.getDate()+10));
  var etaRange = s.getRange("L:L");
  var etaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateAfter(tooFarDate)
    .setBackground("#f4cccc")
    .setRanges([etaRange])
    .build();
  var etaRules = s.getConditionalFormatRules();
  etaRules.push(etaRule);
  s.setConditionalFormatRules(etaRules);

  // set tab color to green
  s.setTabColor("#00ff00");
  

}



/*  ********************  */
/*  PROMO TAB FORMATTING  */
/*  ********************  */

function promoFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let mainTabName = "Flagged Promo Items " + date;
  s.setName(mainTabName);


  // date formatting
  var dateRange = s.getRange("L:N");
  dateRange.setNumberFormat("m/dd/yyyy");


  // horizontal align -- center, (columns H thru AA)
  var colRange = s.getRange("H:AA");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


  // headers already formatted in Template tab
  // // header vertical align -- middle
  // var headRange = s.getRange("A2:AA2");
  // headRange.setVerticalAlignment("middle");


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


  // ETA passed conditional formatting
  var currentDate = new Date();
  var etaRange = s.getRange("L:L");
  var etaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(currentDate)
    .setBackground("#f4cccc")
    .setRanges([etaRange])
    .build();
  var etaRules = s.getConditionalFormatRules();
  etaRules.push(etaRule);
  s.setConditionalFormatRules(etaRules);
  

  // ETA too far out conditional formatting (10 days or more)
  var currentDate = new Date();
  var tooFarDate = new Date(currentDate.setDate(currentDate.getDate()+10));
  var etaRange = s.getRange("L:L");
  var etaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateAfter(tooFarDate)
    .setBackground("#f4cccc")
    .setRanges([etaRange])
    .build();
  var etaRules = s.getConditionalFormatRules();
  etaRules.push(etaRule);
  s.setConditionalFormatRules(etaRules);


  // set tab color to green
  s.setTabColor("#00ff00");

}



/*  *********************  */
/*  ENDCAP TAB FORMATTING  */
/*  *********************  */

function endcapFormat() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let mainTabName = "ENDCAP " + date;
  s.setName(mainTabName);


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
  var aptRange = s.getRange("J:J");
  var aptRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(new Date("01/01/2001"))
    .setBackground("#f4cccc")
    .setRanges([aptRange])
    .build();
  var aptRules = s.getConditionalFormatRules();
  aptRules.push(aptRule);
  s.setConditionalFormatRules(aptRules);


  // ETA passed conditional formatting
  var currentDate = new Date();
  var etaRange = s.getRange("I:I");
  var etaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(currentDate)
    .setBackground("#f4cccc")
    .setRanges([etaRange])
    .build();
  var etaRules = s.getConditionalFormatRules();
  etaRules.push(etaRule);
  s.setConditionalFormatRules(etaRules);
  

  // ETA too far out conditional formatting
  var currentDate = new Date();
  var tooFarDate = new Date(currentDate.setDate(currentDate.getDate()+10));
  var etaRange = s.getRange("I:I");
  var etaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateAfter(tooFarDate)
    .setBackground("#f4cccc")
    .setRanges([etaRange])
    .build();
  var etaRules = s.getConditionalFormatRules();
  etaRules.push(etaRule);
  s.setConditionalFormatRules(etaRules);

  // set tab color to green
  s.setTabColor("#00ff00");

}


function mondayArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  s.setTabColor("#ff0000") // change sheet color to bright red
  s.hideSheet();
}

function wednesdayArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  s.setTabColor("#ffff00") // change sheet color to yellow
  s.hideSheet();
}

function fridayArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  s.setTabColor("#4a86e8") // change sheet color to blue
  s.hideSheet();
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
    .addSeparator()
    .addSeparator()
    .addItem('Archive - Monday', 'mondayArchive')
    .addItem('Archive - Wednesday', 'wednesdayArchive')
    .addItem('Archive - Friday', 'fridayArchive')

    .addToUi(); 
}


/* tests */
// var now = new Date();
// console.log("date:" + now);

// var currentDate = new Date();
// var tooFarDate = new Date(currentDate.setDate(currentDate.getDate()+10));
// // currentDate.setDate(currentDate.getDate()+1);
// console.log(tooFarDate);




// // need to implement:
// autoresize columns (autoResizeColumns())  **** header rows prevent this from working properly
// duplicate UPC conditional formatting


// batch tab hiding
