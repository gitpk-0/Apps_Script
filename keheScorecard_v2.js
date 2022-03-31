function script1() {  
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

    /*   */

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    function reorderSheets() {
  
      // make a list of the sheets in your deferred sort order
      // it is NOT necessary to list every every sheet in the Spreadsheet, 
      // just the ones that are import to list from left to right
      var mysheetOrder = ["Headers","Top_Brand_outs","Pending_Discos","Invalids_Summary","Constrained_Items_Invalids","Long_Term_OOS"];
    
      // get the actual sheets
      var sheets = ss.getSheets()
    
      // Reorder the sheets.
      for( var j = 0; j < mysheetOrder.length; j++ ) {
        ss.setActiveSheet(ss.getSheetByName(mysheetOrder[j]));
        ss.moveActiveSheet(j + 1);
      }
    }
  reorderSheets();

    function hideSheets() {
      var longTermOOSKehe = ss.getSheetByName("Long Term OOS");
      var pendingDiscosKehe = ss.getSheetByName("Pending Discos");
      var invalidsSummaryKehe = ss.getSheetByName("Invalids Summary");
      var constrainedItemsInvKehe = ss.getSheetByName("Constrained Items Invalids");
      var topBrandOutsKehe = ss.getSheetByName("Top Brand outs");
      var headers = ss.getSheetByName("Headers");
      
      
      // hide the following sheets
      headers.hideSheet();
      longTermOOSKehe.hideSheet();
      pendingDiscosKehe.hideSheet();
      invalidsSummaryKehe.hideSheet();
      constrainedItemsInvKehe.hideSheet();
      topBrandOutsKehe.hideSheet();
    }
  hideSheets();

    
} 


function script2() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();

  var longTermOOS = ss.getSheetByName("Long_Term_OOS");
  var pendingDiscos = ss.getSheetByName("Pending_Discos");
  var invalidsSummary = ss.getSheetByName("Invalids_Summary");
  var constrainedItemsInv = ss.getSheetByName("Constrained_Items_Invalids");
  var topBrandOuts = ss.getSheetByName("Top_Brand_Outs");


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

  // set tab color to green
  longTermOOS.setTabColor("#00ff00");
  pendingDiscos.setTabColor("#00ff00");
  invalidsSummary.setTabColor("#00ff00");
  constrainedItemsInv.setTabColor("#00ff00");
  topBrandOuts.setTabColor("#00ff00");


  // // delete check row (all tabs)
  longTermOOS.deleteRow(2)
  pendingDiscos.deleteRow(2)
  invalidsSummary.deleteRow(2)
  constrainedItemsInv.deleteRow(2)
  topBrandOuts.deleteRow(4)

  // invalids tab formatting
  // invalidsSummary

  // copy/paste date on Top Brand Outs sheet
    var topBrandOutsKehe = ss.getSheetByName("Top Brand Outs");
    var date = topBrandOutsKehe.getRange("E3:I3");
    var dest = topBrandOuts.getRange("E2:I2")
    date.copyTo(dest);

}



function script3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var longTermOOS = ss.getSheetByName("Long_Term_OOS");
  var pendingDiscos = ss.getSheetByName("Pending_Discos");
  var invalidsSummary = ss.getSheetByName("Invalids_Summary");
  var constrainedItemsInv = ss.getSheetByName("Constrained_Items_Invalids");
  var topBrandOuts = ss.getSheetByName("Top_Brand_Outs");

  /* ------------------------------- */
  /* Top Brands Out Sheet Formatting */
  /* ------------------------------- */

  // horizontal & vertical alignment
  var tboColRange = topBrandOuts.getRange("A4:040");
  tboColRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // column widths
  topBrandOuts.setColumnWidths(1, 2, 150);
  topBrandOuts.setColumnWidths(3, 7, 75);
  topBrandOuts.setColumnWidths(10, 4, 95);
  topBrandOuts.setColumnWidth(14, 65);
  topBrandOuts.setColumnWidth(15, 300);
  
  // text wrap 
  var wrapRange1 = topBrandOuts.getRange("B:B");
  wrapRange1.setWrap(true);
  var wrapRange2 = topBrandOuts.getRange("O:O");
  wrapRange2.setWrap(true);



  /* ------------------------------- */
  /* Pending Discos Sheet Formatting */
  /* ------------------------------- */

  // horizontal & vertical alignment
  var penDiscRange = pendingDiscos.getRange("A:R")
  penDiscRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // column widths
  pendingDiscos.autoResizeColumns(6,2);
  pendingDiscos.setColumnWidths(9, 4, 85);
  pendingDiscos.setColumnWidths(13, 2, 65);


  /* ------------------------------- */
  /* Invalids Summary Sheet Formatting */
  /* ------------------------------- */

  // horizontal & vertical alignment
  var invSumRange = invalidsSummary.getRange("A:L")
  invSumRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // column widths
  invalidsSummary.autoResizeColumns(5,4);
  invalidsSummary.autoResizeColumns(10,3);



  /* ------------------------------- */
  /* Constrained Item Invalids Sheet Formatting */
  /* ------------------------------- */

  // horizontal & vertical alignment
  var consItemRange = constrainedItemsInv.getRange("A:L")
  consItemRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // column widths
  constrainedItemsInv.autoResizeColumns(5,4);
  constrainedItemsInv.setColumnWidths(9, 2, 88);

  // date format
  var conDateRange = constrainedItemsInv.getRange("K:L")
  conDateRange.setNumberFormat("m/dd/yyyy");


  /* ------------------------------- */
  /* LTOOS Sheet Formatting */
  /* ------------------------------- */

  // horizontal & vertical alignment
  var ltoosRange = longTermOOS.getRange("A:Q")
  ltoosRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

  // column widths
  longTermOOS.autoResizeColumns(5,3);
  longTermOOS.setColumnWidth(8, 215);
  longTermOOS.setColumnWidths(9, 5, 64);
  longTermOOS.setColumnWidth(17, 60);
  
  // date format
  var conDateRange = longTermOOS.getRange("P:P")
  conDateRange.setNumberFormat("m/dd/yyyy");

  // text wrap
  var wrapRange3 = longTermOOS.getRange("H:H");
  wrapRange3.setWrap(true);


}

  
  
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // menu button name, function button name, function for button to perform
  ui.createMenu('Scripts')
    .addItem('Run First', 'script1')
    .addSeparator()
    .addItem('Wait Until First Script Finishes -- Run Second', 'script2')   
    .addSeparator()
    .addItem('Wait Until Second Script Finishes -- Run Third', 'script3')   
    .addSeparator()

    .addToUi(); 
}



// top brands outs
// conditional formatting (less than 0%)
