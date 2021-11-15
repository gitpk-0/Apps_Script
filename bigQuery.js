

function GetBQData(sheet,query,start_date,end_date,dpt_numbers) 
{
  var projectId = BQ_PROJECTID;
  var requestQuery;
  
  if(start_date === null || start_date === 'undefined' || start_date === undefined)
  {
    return null; 
  }
  else
  {
    requestQuery = "#standardSQL\n"+ query
    .replace(/START_DATE/g,start_date)
    .replace(/END_DATE/g,end_date)
    .replace(/DEPARTMENT_NUMBERS/g,dpt_numbers); 
  }
  
  Logger.log(requestQuery);
  
  var request = {query: requestQuery};
  
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;
  
  // check status of the query job
  var sleepTimeMs = 500;
  while(!queryResults.jobComplete)
  {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *=2; // increment sleep times 2
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
    
  }
  
  // get all rows of result
  var rows = queryResults.rows;
  var cols;
  
  while(queryResults.pageToken)
  {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken:queryResults.pageToken
    });
    rows = rows.concat(queryResults.rows);
  }
  
  // if there's data
  if(rows)
  {
    var data = new Array(rows.length);
    for(var i=0;i<rows.length;i++)
    {
      cols = rows[i].f;
      data[i] = new Array(cols.length);
      
      for(var j=0;j<cols.length;j++)
      {
        data[i][j] = cols[j].v;
      }
    }
    
    return data;
  } 
  else 
  {
    Logger.log("No rows returned");
  }
}

/*  
  Get Department Numbers
*/
function BQGetDepartment(dept)
{
  var projectId = 'momsdatawarehouse';
  var requestQuery,res;
  
  requestQuery = "#standardSQL\n"+ "select DPT_Number from catapult.Departments where DPT_Name like \'"+dept+"\';"; 
  
  Logger.log(requestQuery);
  
  var request = {query: requestQuery};
  
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;
  
  // check status of the query job
  var sleepTimeMs = 500;
  while(!queryResults.jobComplete)
  {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *=2; // increment sleep times 2
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }
  
  // get all rows of result
  var rows = queryResults.rows;
  var cols;
  
  while(queryResults.pageToken)
  {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken:queryResults.pageToken
    });
  }
  
  // if there's data
  if(rows)
  {
    var data = new Array(rows.length);
    for(var i=0;i<rows.length;i++)
    {
      cols = rows[i].f;
      data[i] = new Array(cols.length);
      
      for(var j=0;j<cols.length;j++)
      {
        data[i][j] = cols[j].v;
      }
    }
  } 
  else 
  {
    Logger.log("No rows returned");
  }
  if(data.length>0)
  {
    res = data[0][0];
    for(var r=1;r<data.length;r++)
    {
      for(var c=0;c<data[0].length;c++)
      {
        res += ','+data[r][c];
      }
    }
  }
  return res;
}