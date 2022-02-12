function freezeRowsAndColumns() {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    
    activeSheet.setFrozenRows(1);
    activeSheet.setFrozenColumns(7);
  }
  
  function freezeRowsAndColumns2() {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    
    activeSheet.setFrozenRows(1);
  }
  
  function centerColumns() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    var range = s.getRange("A:L")
  
    range.setHorizontalAlignment("center");
  }
  
  function centerColumns2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    var range = s.getRange("H:AL")
  
    range.setHorizontalAlignment("center");
  }
  
  function dateFormat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    var range = s.getRange("L:N")
  
    range.setNumberFormat("m/dd/yyyy");
  }
  
  function autoResizeCol() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    s.autoResizeColumns(1,5);
  }
  
  
  
  function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // menu button name, function button name, function for button to perform
    ui.createMenu('Scripts')
      .addItem('Freeze Row 1', 'freezeRowsAndColumns2')
      .addSeparator()
      .addItem('Freeze Row 1 & Up to Col H', 'freezeRowsAndColumns')
      .addSeparator()
      .addItem('Center Col A - L', 'centerColumns')
      .addSeparator()
      .addItem('Center Col H - AL', 'centerColumns2')
      .addSeparator()
      .addItem('Date Format Col L - N', 'dateFormat')
      // .addSeparator()
      // .addItem('Resize Columns', 'autoResizeCol')
      .addToUi(); 
  }
  
  
    
  /* NEED TO ADD */
  // format percentages
  // format currencies
  // NL formatting
  // promo formatting
  