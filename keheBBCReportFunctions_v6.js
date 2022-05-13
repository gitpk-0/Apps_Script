// ALL HEADERS [global]
var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");


/*  ********************  */
/*  BBC TAB FORMATTING  */
/*  ********************  */

function bbcFormat() {

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

      var s = ss.getActiveSheet();

      // updating doc and sheet names with current date
      let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
      
      let docName = "BBC/NL KeHE Report (updated " + date + ")";  
      ss.setName(docName);

      let mainTabName = "OH OO " + date;
      s.setName(mainTabName);


      // date formatting
      var dateRange = s.getRange("M:O")
      dateRange.setNumberFormat("m/dd/yyyy");


      // header
      var bbc = headers.getRange('A2:AL2');
      s.insertRowBefore(1); // inserts new empty first row
      var destinationS = s.getRange("A1:AL1");
      bbc.copyTo(destinationS); // pastes headers to first row


      // horizontal & vertical align -- center, middle (columns I thru AL)
      var colRange = s.getRange("I:AL");
      colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


      // freeze rows and columns
      s.setFrozenRows(1);
      s.setFrozenColumns(8);



      // invalid APT conditional formatting
      var aptRange = s.getRange("O:O");
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
      var etaRange = s.getRange("M:M");
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
      var etaRange = s.getRange("M:M");
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
        // setColumnWidth(column number, pixel width)  -- single column
        // setColumnWidths(column number, number of columns, pixel width) -- multiple columns
      s.setColumnWidth(4, 58);
      s.setColumnWidth(5, 70);
      s.setColumnWidth(8, 180);
      s.setColumnWidths(9, 4, 39);
      s.setColumnWidths(13, 3, 71);
      s.setColumnWidths(16, 3, 60);
      s.setColumnWidths(19, 3, 72);
      s.setColumnWidth(22, 155);
      s.setColumnWidths(23, 3, 65);
      s.setColumnWidths(27, 3, 51);
      s.setColumnWidth(30, 125);
      s.setColumnWidths(31, 3, 85);
      s.setColumnWidths(34, 4, 76);
      s.setColumnWidth(38, 310);

    } else {
      var ui = SpreadsheetApp.getUi();
      var Alert = ui.alert("The are more (or less) than 6 visible tabs/sheets. This script will not run unless there are exactly 6 unhidden tabs/sheets. Please make the necessary adjustments and rerun the script.");
    }
    return;
  }
  hideTabs();
}

  



/*  ********************  */
/*  PROMO TAB FORMATTING  */
/*  ********************  */

function promoFormat() {

  // local variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();


  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let promoTabName = "Flagged Promo Items " + date;
  s.setName(promoTabName);


  // date formatting
  var dateRange = s.getRange("L:N");
  dateRange.setNumberFormat("m/dd/yyyy");


  // header
  var promo = headers.getRange('A4:AD5');
  s.insertRowBefore(1); // inserts new empty first row
  s.insertRowBefore(1); // inserts another new empty first row
  var destinationS = s.getRange("A1:AD2");
  promo.copyTo(destinationS); // pastes headers to first row


  // horizontal align -- center, (columns H thru AD)
  var colRange = s.getRange("H:AD");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


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
  s.setColumnWidth(5, 70);
  s.setColumnWidth(7, 180);
  s.setColumnWidths(8, 4, 39);
  s.setColumnWidths(12, 3, 71);
  s.setColumnWidths(15, 3, 60);
  s.setColumnWidths(18, 3, 72);
  s.setColumnWidth(21, 155);
  s.setColumnWidths(22, 3, 65);
  s.setColumnWidths(26, 3, 51);
  s.setColumnWidths(29, 2, 125);
}




/*  ********************  */
/*  NL/BS TAB FORMATTING  */
/*  ********************  */


function nlFormat() {

  // local variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();

  // updating sheet name with current date
  let date = Utilities.formatDate(new Date(), "GMT-05:00", "M.d");
  let nlTabName = "NL/BS " + date;
  s.setName(nlTabName);


  // date formatting
  var dateRange = s.getRange("L:N");
  dateRange.setNumberFormat("m/dd/yyyy");


  // horizontal align -- center (columns H thru AJ)
  var colRange = s.getRange("H:AD");
  colRange.setHorizontalAlignment("center").setVerticalAlignment("middle");


  // header
  var nlbs = headers.getRange('A7:AD7');
  s.insertRowBefore(1); // inserts new empty first row
  var destinationS = s.getRange("A1:AD1");
  nlbs.copyTo(destinationS); // pastes headers to first row


  // freeze rows and columns
  s.setFrozenRows(1);
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
  s.setColumnWidths(18, 3, 72);
  s.setColumnWidth(21, 155);
  s.setColumnWidths(22, 3, 65);
  s.setColumnWidths(26, 3, 51);
  s.setColumnWidths(29, 2, 125);
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
  } else {
    var ui = SpreadsheetApp.getUi();
    var Alert = ui.alert("There are more (or less) than 3 visible tabs/sheets. This script will not run unless there are exactly 3 unhidden tabs/sheets. Please make the necessary adjustments and rerun the script.");
  }
  return;
}




function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('BBC Tab Formatting', 'bbcFormat')
    .addSeparator()
    .addItem('Promo Tab Formatting', 'promoFormat')
    .addSeparator()
    .addItem('NL/BS Formatting', 'nlFormat')
    .addSeparator()
    .addSeparator()
    .addSeparator()
    .addItem('*** Run Last and Check Headers Before Running *** Delete Check Rows All Tabs', 'deleteCheckRows')
    

    .addToUi(); 
}