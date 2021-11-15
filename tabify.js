function Tabify() 
{
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    /*
      ss: Spreadsheet
      deptcol: Integer
      catcol: Integer
      row: Integer
    */
    DataUtilities.Tabify(ss, 3, null, 2);
    DataUtilities.FormatSheets(ss, 2,'PVP');
    DataUtilities.AutoResizeAll(ss);
    ss.toast('Done');
    SpreadsheetApp.flush();
}



function DeleteTabs()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var SH_EXCLUSIONS = ['master','master_list','log','Notes'];
  var sh;
  for(var s in sheets)
  {
    sh = sheets[s];
    
    if(SH_EXCLUSIONS.indexOf(sh.getName())<0)
    {
      console.log('removing: '+sh.getName());
      ss.deleteSheet(sh);
    }
  }
}