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
  
  // resize columns
  s.setColumnWidth(4, 70);
  s.setColumnWidth(7, 180);
  s.setColumnWidths(8, 4, 39);
  s.setColumnWidths(12, 3, 71);
  s.setColumnWidths(15, 3, 60);
  s.setColumnWidth(18, 155);
  s.setColumnWidths(19, 3, 65);
  s.setColumnWidths(23, 3, 51);
  s.setColumnWidths(26, 4, 85);
  s.setColumnWidths(30, 4, 76);
  s.setColumnWidth(34, 500);

}



/*  ********************  */
/*  PROMO TAB FORMATTING  */
/*  ********************  */

function promoFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let promoTabName = "Flagged Promo Items " + date;
  s.setName(promoTabName);


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

  // resize columns
  s.setColumnWidth(4, 70);
  s.setColumnWidth(7, 180);
  s.setColumnWidths(8, 4, 39);
  s.setColumnWidths(12, 3, 71);
  s.setColumnWidths(15, 3, 60);
  s.setColumnWidth(18, 155);
  s.setColumnWidths(19, 3, 65);
  s.setColumnWidths(23, 3, 51);
  s.setColumnWidths(26, 2, 85);

}



/*  *********************  */
/*  ENDCAP TAB FORMATTING  */
/*  *********************  */

function endcapFormat() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let endcapTabName = "ENDCAP " + date;
  s.setName(endcapTabName);


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

  // resize columns
  s.autoResizeColumns(1, 11);
  s.setColumnWidth(12, 72);
}

// Hide previous/old report tabs
function hideTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var visibleSheets = SpreadsheetApp.getActive().getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName())
  var lenVis = visibleSheets.length

  if (lenVis == 6){  
    for (let i = 0; i < 3; i++) {
      var sheet = ss.getSheetByName(visibleSheets[i])
      sheet.setTabColor("#ff0000") // change sheet color to bright red
      sheet.hideSheet();
    }
  }
  return;
}

// Delete check rows
function deleteCheckRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var visibleSheets = SpreadsheetApp.getActive().getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName())
  var lenVis = visibleSheets.length

  if (lenVis == 3){  
    for (let i = 0; i < lenVis; i++) {
      var sheet = ss.getSheetByName(visibleSheets[i])
      if (i == 1) {
        sheet.deleteRow(3);
      } else {
        sheet.deleteRow(2);
      }      
    }
  }
  return;
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('***Run First*** Hide Old Tabs', 'hideTabs')
    .addSeparator()
    .addSeparator()
    .addSeparator()
    .addItem('OH OO Tab Formatting', 'ohOoFormat')   
    .addSeparator()
    .addItem('Promo Tab Formatting', 'promoFormat')
    .addSeparator()
    .addItem('Endcap Tab Formatting', 'endcapFormat')
    .addSeparator()
    .addSeparator()
    .addSeparator()
    .addItem('***Check headers before running*** Delete Check Rows All Tabs', 'deleteCheckRows')
    

    .addToUi(); 
}




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
  
  // resize columns
  s.setColumnWidth(4, 70);
  s.setColumnWidth(7, 180);
  s.setColumnWidths(8, 4, 39);
  s.setColumnWidths(12, 3, 71);
  s.setColumnWidths(15, 3, 60);
  s.setColumnWidth(18, 155);
  s.setColumnWidths(19, 3, 65);
  s.setColumnWidths(23, 3, 51);
  s.setColumnWidths(26, 4, 85);
  s.setColumnWidths(30, 4, 76);
  s.setColumnWidth(34, 500);

}



/*  ********************  */
/*  PROMO TAB FORMATTING  */
/*  ********************  */

function promoFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let promoTabName = "Flagged Promo Items " + date;
  s.setName(promoTabName);


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

  // resize columns
  s.setColumnWidth(4, 70);
  s.setColumnWidth(7, 180);
  s.setColumnWidths(8, 4, 39);
  s.setColumnWidths(12, 3, 71);
  s.setColumnWidths(15, 3, 60);
  s.setColumnWidth(18, 155);
  s.setColumnWidths(19, 3, 65);
  s.setColumnWidths(23, 3, 51);
  s.setColumnWidths(26, 2, 85);

}



/*  *********************  */
/*  ENDCAP TAB FORMATTING  */
/*  *********************  */

function endcapFormat() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let endcapTabName = "ENDCAP " + date;
  s.setName(endcapTabName);


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

  // resize columns
  s.autoResizeColumns(1, 11);
  s.setColumnWidth(12, 72);
}

// Hide previous/old report tabs
function hideTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var visibleSheets = SpreadsheetApp.getActive().getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName())
  var lenVis = visibleSheets.length

  if (lenVis == 6){  
    for (let i = 0; i < 3; i++) {
      var sheet = ss.getSheetByName(visibleSheets[i])
      sheet.setTabColor("#ff0000") // change sheet color to bright red
      sheet.hideSheet();
    }
  }
  return;
}

// Delete check rows
function deleteCheckRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var visibleSheets = SpreadsheetApp.getActive().getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName())
  var lenVis = visibleSheets.length

  if (lenVis == 3){  
    for (let i = 0; i < lenVis; i++) {
      var sheet = ss.getSheetByName(visibleSheets[i])
      if (i == 1) {
        sheet.deleteRow(3);
      } else {
        sheet.deleteRow(2);
      }      
    }
  }
  return;
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('***Run First*** Hide Old Tabs', 'hideTabs')
    .addSeparator()
    .addSeparator()
    .addSeparator()
    .addItem('OH OO Tab Formatting', 'ohOoFormat')   
    .addSeparator()
    .addItem('Promo Tab Formatting', 'promoFormat')
    .addSeparator()
    .addItem('Endcap Tab Formatting', 'endcapFormat')
    .addSeparator()
    .addSeparator()
    .addSeparator()
    .addItem('***Check headers before running*** Delete Check Rows All Tabs', 'deleteCheckRows')
    

    .addToUi(); 
}




