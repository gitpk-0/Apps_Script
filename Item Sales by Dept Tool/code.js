function GetSales() 
{
  var start = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh_config = ss.getSheetByName('config');
  var sh_sales = ss.getSheetByName('sales');
  var sh_log = ss.getSheetByName('log');
  
  // get parameters
  var start_date = sh_config.getRange(REF_START_DATE).getValue();
  var days = sh_config.getRange(REF_NUM_OF_DAYS).getValue();
  var periods = sh_config.getRange(REF_PERIODS).getValue();
  var departments = sh_config.getRange(REF_DEPARTMENT).getValue();
  var categories = sh_config.getRange(REF_CATEGORY).getValue();
  var rep_status_range = sh_config.getRange(REF_REPORT_STATUS);
  var run_status_range = sh_config.getRange(REF_RUN_STATUS);
  /* --------------------------------------------------------------------------------------------------- */
  /* --------------------------------------------------------------------------------------------------- */
  // Log what we have
  console.log(start_date+', '+days+', '+periods+', '+departments+', '+categories);
  
  var startdate,enddate,department,category,query,start_row,data;
  
  query = SQL_ITEM_SALES;
  /* --------------------------------------------------------------------------------------------------- */
  // get department numbers
  var d = departments.split(',');
  if(d.length>1)
  {
    // loop through department names and get numbers
    department = BQGetDepartment(d[0].trim());
    for(var i=1;i<d.length;i++)
    {
      department += ','+BQGetDepartment(d[i].trim());
    }
  }
  else if(departments !== '')
  {
    // get single department
    department = BQGetDepartment(departments);
  }
  else
  {
    // * all departments
    department = '*';
  }
  
  console.log(department);
  /* --------------------------------------------------------------------------------------------------- */
  // get category
  var c = categories.split(',');
  if(c.length>1)
  {
    // loop through department names and get numbers
    category = '\''+c[0].trim()+'\'';
    for(var i=1;i<d.length;i++)
    {
      category += ',\''+c[i].trim()+'\'';
    }
    query += '  where trim(p1.PI1_Description) IN ('+category+')';
  }
  else if(categories !== '')
  {
    category = '\''+categories+'\'';
    query += '  where trim(p1.PI1_Description) IN ('+category+')';
  }
  else
  {
    category = '*';
  }
  console.log(category);  
  /* --------------------------------------------------------------------------------------------------- */
  /* --------------------------------------------------------------------------------------------------- */
  var log_p,p;
  var log_lastrow = sh_log.getLastRow();
  
  if(log_lastrow>1)
  {
    log_p = sh_log.getRange(log_lastrow, 1).getValue();
  }
  else
  {
    log_p = 0;
  }
  
  if(log_p === periods)
  {    
    p=log_p;
    console.log('Log P: '+log_p+', p: '+p);
  }
  else
  {
    for(p=log_p;p<=periods;p++)
    {
      if (isTimeUp_(start)) 
      {
        console.log("Times up");
        break;
      }
      else
      {
        startdate = (addDaysFromDate(start_date,days*(p-1)));
        enddate = (addDaysFromDate(startdate,days-1));
        
        // Log date range
        console.log((startdate)+', '+(enddate));
        
        ss.toast('Generating sales for range '+(p)+' of '+(periods)+': '+formatDateMDYGMT4(startdate)+' to '+formatDateMDYGMT4(enddate), 'Progress', -1);
        
        data = GetBQData(sh_sales,query,formatDateYMDGMT4(startdate),formatDateYMDGMT4(enddate),department);
        
        if(data !== undefined && data !== 'undefined' && data !== null)
        {
          if(p === 1)
          // clear sheet 
          {
            sh_sales.getRange(2,1,sh_sales.getLastRow(), sh_sales.getLastColumn()).clearContent();
            sh_sales.getRange(2,1,data.length,data[0].length)
            .setNumberFormat("@")
            .setValues(data);
          }
          else
          {
            start_row = sh_sales.getLastRow()+1;
            sh_sales.getRange(start_row,1,data.length,data[0].length)
            .setNumberFormat("@")
            .setValues(data);
            
          }
        }
        // Log the range
        Log(sh_log,p,startdate,enddate);
      }
    }
  }
  /* --------------------------------------------------------------------------------------------------- */
  // format columns
  // if nothing added 
  if(p !== undefined && p !== 'undefined' && p >= periods)
  {
    sh_sales.getRange(2,11,sh_sales.getLastRow(),4).setNumberFormat("#,##0.00");
    sh_sales.getRange(2,12,sh_sales.getLastRow(),3).setNumberFormat("$#,##0.00");
    removeEmptyRows(sh_sales);
    removeEmptyColumns(sh_sales);
    autoResizeSheetColumns(sh_sales);
    // dismiss toast
    ss.toast('Sales generated :)','Progress',-1);
    // Stop the report from generating
    rep_status_range.setValue('Stop');  // ss.toast('Setting status to Stop','Progress',-1);
    // clear log
    clearSheet(sh_log,2);
  }
  else
  {
    // dismiss toast
    ss.toast('Nothing added :)','Progress',-1);
  }  
}

function RunGetSales()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh_config = ss.getSheetByName('config');
  var sh_sales = ss.getSheetByName('sales');
  var run_status_range = sh_config.getRange(REF_RUN_STATUS);
  var rep_status_range = sh_config.getRange(REF_REPORT_STATUS);
  
  // indicate script is running 
  if(run_status_range.getValue()==='Running')
  {
    console.log('Already running ...');
  }
  else
  {
    try
    {
      run_status_range.setValue('Running');
      
      if(rep_status_range.getValue() === 'Start')
      {
        GetSales();
        sh_sales.getDataRange().removeDuplicates();
        run_status_range.setValue('Finished');
      }
      else
      {
        run_status_range.setValue('Finished');
      }
    }
    catch(e)
    {
      run_status_range.setValue('Finished');
    }
  }  
}

function Log(sh_log,p,startdate,enddate)
{
  var lastrow = sh_log.getLastRow()+1;
  sh_log.getRange(lastrow, 1).setValue(p);
  sh_log.getRange(lastrow, 2).setValue(startdate);
  sh_log.getRange(lastrow, 3).setValue(enddate);
}

function clearLogSheet()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh_log = ss.getSheetByName('log');
  var start_row = 2;
  clearSheet(sh_log,start_row);
  ss.toast('Log sheet cleared', 'Clear Logs', 5);
}