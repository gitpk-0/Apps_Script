function hideColumns() { 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[1];
    var sheet1 = ss.getSheets()[2];
    var sheet2 = ss.getSheets()[3];
    // Hides the first column
    sheet.hideColumns(1,2);
    //sheet.hideColumns(3);
    sheet1.hideColumns(1);
    sheet2.hideColumns(1);
  }
  
  
  function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Hide Columns').addItem('Hide Columns', 'hideColumns').addToUi();
  }