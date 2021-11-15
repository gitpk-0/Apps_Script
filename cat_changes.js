var APP_SCRIPTNAME='ss-HandleCCUpdates';

function UpdatesHandler(){
  var s = SpreadsheetApp.getActiveSheet();
  var sh = s.getSheetName();
  if(sh === 'SpecialOrders-Mispicks' || sh === 'SpecialOrders-UNFI'){
    CatapultChanges.UpdateSheetOnChange();
  }
  else if(sh === 'Products to be Updated' || sh === 'KeHe Item Issues' ){
    CatapultChanges.UpdateProductUpdateRequest();
  }else{}
}

function cleanUp(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sosh = s.getSheetByName('SpecialOrders-Mispicks');
  var psh = s.getSheetByName('Products to be Updated');
  var ksosh = s.getSheetByName('SpecialOrders-UNFI');
  var kpsh = s.getSheetByName('KeHe Item Issues');
  var delete_keywords = [];
  delete_keywords.push('No Information');
  delete_keywords.push('no information');
  delete_keywords.push('Delete Row');
  delete_keywords.push('delete row');
  var keyword;
  for(var a=0;a<delete_keywords.length;a++){
    keyword = delete_keywords[a];
    CatapultChanges.RemoveBadRows(sosh,keyword);
    CatapultChanges.RemoveBadRows(psh,keyword);
    
    CatapultChanges.RemoveBadRows(ksosh,keyword);
    CatapultChanges.RemoveBadRows(kpsh,keyword);
  }
  checkBlankRows();
}

function ListUpdate(){
  CatapultChanges.GetFileData();
}

function Archive(){
  CatapultChanges.archiveCompletedCC('SpecialOrders-Mispicks');
  CatapultChanges.archiveCompletedCC('Products to be Updated');
  CatapultChanges.archiveCompletedCC('SpecialOrders-UNFI');
  CatapultChanges.archiveCompletedCC('KeHe Item Issues');
}

function test(){
  CatapultChanges.archiveCompletedCC('KeHe Item Issues');
}

function PopUp(){
  var s = SpreadsheetApp.getActiveSpreadsheet();  
  CatapultChanges.displayText(s);
}

function checkBlankRows()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var shname;
  var loaded, empty, rows;
  for(var s=0;s<sheets.length;s++)
  {
    shname = sheets[s].getName();
    if(shname.indexOf('Completed')===-1 && shname.indexOf('Guidelines')===-1 && shname!=='data')
    {
      loaded = sheets[s].getLastRow();
      rows = sheets[s].getMaxRows();
      empty = rows-loaded;
      if(empty<100)
      {
        sheets[s].insertRowsAfter(loaded, 100);
        Logger.log('Insufficient rows, adding 100 to '+shname);
      }
      else
      {
        Logger.log('Sufficient rows');
      }
    }
    
  }
}