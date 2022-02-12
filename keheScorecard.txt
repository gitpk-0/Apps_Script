function freezeRowsAndColumns() {
    var longTermOOS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Long_Term_OOS");
    var pendingDiscos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending_Discos");
    var invalidsSummary = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invalids_Summary");
    var constrainedItemsInv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Constrained_Items_Invalids");
    var topBrandOuts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Top_Brand_outs");
  
  
    longTermOOS.setFrozenRows(1);
    longTermOOS.setFrozenColumns(1);
    pendingDiscos.setFrozenRows(1);
    pendingDiscos.setFrozenColumns(1);
    invalidsSummary.setFrozenRows(1); 
    invalidsSummary.setFrozenColumns(1); 
    constrainedItemsInv.setFrozenRows(1);
    constrainedItemsInv.setFrozenColumns(1);
    topBrandOuts.setFrozenRows(3);
    topBrandOuts.setFrozenColumns(2);
  }
  
  function prependHeaderRows() {  
      // headers
      var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Headers");
    
      // header ranges
      var pendingDiscos = headers.getRange('A2:R2'); // pending discos headers
      var invalidsSummary = headers.getRange('A4:R4'); // invalidsSummary headers
      var constrainedItems = headers.getRange('A6:R6'); // constrainedItems headers
      var longTermOOS = headers.getRange('A8:R8'); // longTermOOS headers
      var topBrandOuts = headers.getRange('A10:012'); // topBrandOuts headers
    
      // pending discos sheet
      var pd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending_Discos");
      pd.insertRowBefore(1);
      var destinationPd = pd.getRange("A1:R1");
      pendingDiscos.copyTo(destinationPd);
  
  
      // invalids summary sheet
      var is = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invalids_Summary");
      is.insertRowBefore(1);
      var destinationIs = is.getRange("A1:R1");
      invalidsSummary.copyTo(destinationIs);
  
      // constrained items invalids sheet
      var cii = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Constrained_Items_Invalids");
      cii.insertRowBefore(1);
      var destinationCii = cii.getRange("A1:R1");
      constrainedItems.copyTo(destinationCii);
  
      // long term oos sheet
      var  ltoos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Long_Term_OOS");
      ltoos.insertRowBefore(1);
      var destinationLt = ltoos.getRange("A1:R1");
      longTermOOS.copyTo(destinationLt);
  
  
      // top brand outs sheet
      var tbo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Top_Brand_Outs");
      tbo.insertRowBefore(1);
      tbo.insertRowBefore(1);
      tbo.insertRowBefore(1);
      var destinationTbo = tbo.getRange("A1:03");
      topBrandOuts.copyTo(destinationTbo);
    }
    
    
    
    function onOpen() {
      let ui = SpreadsheetApp.getUi();
      // menu button name, function button name, function for button to perform
      ui.createMenu('Scripts')
        .addItem('Prepend Header Rows', 'prependHeaderRows')
        .addSeparator()
        .addItem('Freeze Rows & Columns', 'freezeRowsAndColumns')   
        .addSeparator()
    
        .addToUi(); 
    }