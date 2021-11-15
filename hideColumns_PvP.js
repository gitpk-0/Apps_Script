function hideColumns() { 
    ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getSheets()[12];
    // Hides the first column
    sheet.hideColumns(3,3);
    sheet.hideColumns(11,6);
    sheet.hideColumns(18,3);
    sheet.hideColumns(25,3);
    sheet.hideColumns(29,4);
    sheet.hideColumns(34,2);
    sheet.hideColumns(35,1);
    sheet.hideColumns(37,1);
}
  
  
function onOpen() {
    ui = SpreadsheetApp.getUi();
    ui.createMenu('Hide Columns').addItem('Hide Columns', 'hideColumns').addToUi();
}