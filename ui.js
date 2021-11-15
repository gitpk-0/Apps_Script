
function FileMenu()
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sales Tool')
      .addItem('Generate Sales', 'RunGetSales')  
      .addItem('Clear Log', 'clearLogSheet')
      .addToUi();
  
}