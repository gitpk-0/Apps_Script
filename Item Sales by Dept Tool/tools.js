/**
 @param sh Sheet being trimmed 
**/
function removeEmptyRows(sh)
{
  var lastdatarow = sh.getLastRow(); // rows populated with data
  var lastsheetrow = sh.getMaxRows(); // total rows in sheet
  // Logger.log(lastdatarow+' - '+lastsheetrow+' >> '+lastsheetrow-(lastdatarow));
  if(lastsheetrow > lastdatarow)
  {sh.deleteRows(lastdatarow+1, lastsheetrow-(lastdatarow));}
  else{/* no extra rows */}
}

/**
 @param sh Sheet being trimmed 
**/
function clearSheet(sh,start_row)
{
  var lastdatarow = start_row; // rows populated with data
  var lastsheetrow = sh.getMaxRows(); // total rows in sheet
  Logger.log('clearing rows: '+start_row+' to '+lastsheetrow);
  
  sh.getRange(start_row,1,sh.getLastRow(),sh.getLastColumn()).clearContent();
  
  removeEmptyRows(sh);
  /*
  if(lastsheetrow > lastdatarow)
  {sh.deleteRows(lastdatarow, lastsheetrow-(lastdatarow));}
  else{Logger.log('no extra rows');}
  Logger.log('clearing row '+start_row);
  sh.getRange(start_row,1,1,sh.getLastColumn()).clearContent();
  */
  
}


/**
 @param sh Sheet being trimmed 
**/
function removeEmptyColumns(sh)
{
  var lastdatacol = sh.getLastColumn(); // rows populated with data
  var lastsheetcol = sh.getMaxColumns(); // total rows in sheet
  // Logger.log(lastdatarow+' - '+lastsheetrow+' >> '+lastsheetrow-(lastdatarow));
  if(lastsheetcol > lastdatacol)
  {sh.deleteColumns(lastdatacol+1, lastsheetcol-(lastdatacol));}
  else{/* no extra rows */}
}


function CheckIfSheetExists(ss,sheetname)
{
  var sh = ss.getSheetByName(sheetname);
  if(sh === 'undefined' || sh === undefined || sh === null)
  {
    sh = ss.insertSheet();
    sh.setName(sheetname);
  }
  else
  {
    Logger.log('Sheet exists: '+sheetname);
  }
  return sh;
}

/*
  source: https://stackoverflow.com/questions/20059111/trying-to-subtract-5-days-from-a-defined-date-google-app-script
*/
function subDaysFromDate(date,d){
  // d = number of day ro substract and date = start date
  var result = new Date(date.getTime()-d*(24*3600*1000));
  return result;
}

function addDaysFromDate(date,d){
  // d = number of day ro substract and date = start date
  var result = new Date(date.getTime()+d*(24*3600*1000));
  return result;
}

function formatDateYMDGMT4(date)
{
  return Utilities.formatDate(date,'GMT-4:00','yyyy-MM-dd')
}

function formatDateMDYGMT4(date)
{
  return Utilities.formatDate(date,'GMT-4:00','MM/dd/yyyy')
}

function autoResizeSheetColumns(sh)
{
  var cols = sh.getLastColumn();
  sh.autoResizeColumns(1, cols);
}

/* Based on https://gist.github.com/erickoledadevrel/91d3795949e158ab9830 */

function isTimeUp_(start) {
  var now = new Date();
  return now.getTime() - start.getTime() > 300000; // 5 minutes
}