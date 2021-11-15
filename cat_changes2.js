/**
* @author
* @date
* @name {String} notifyApprovedUpdates
*/

function notifyApprovedUpdates()
{
  var fnName = APP_SCRIPTNAME+'notifyApprovedUpdates';
  var TABS = ['SpecialOrders-Mispicks','Products to be Updated','SpecialOrders-UNFI','Naked Lunch, Bake Shop & Coffee Bar Menu Items'];
  var HEADER_ROWS = [3,3,3,1];
  var DATA_ROWS = [4,4,4,2]
  var COORD_COL = 3;
  var COORD_IDX = COORD_COL-1;
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  
  // rows of approved requests
  var header = [];
  var sh;
  var data,approved_data,colspan;
  var message = 
  "<html>\
  <body style='font:\'Trebuchet MS Helvetica\';'>\
  <p>Hello,<br><br>\
  <a href='"+SS.getUrl()+"'>Catapult Changes</a> has approved items: \
  </p>\
  <table cellspacing='0' style='border-collapse:collapse;'>";
  
  // loop through each tab and get rows that are approved
  for(var t=0;t<TABS.length;t++)
  {
    sh = SS.getSheetByName(TABS[t]);
    header = [];
    
    if(sh)
    {
      
      data = sh.getRange(DATA_ROWS[t], 1, sh.getLastRow()-(DATA_ROWS[t]-1), sh.getLastColumn()).getValues();
      
      if(TABS[t].indexOf('Naked Lunch, Bake Shop & Coffee Bar Menu Items')>-1)
      {
        approved_data = data.filter(function(r){
          return r[0] === ''
        });
        console.log('%s - approved_data size: %s',fnName,approved_data.length);
      }
      else
      {
        approved_data = data.filter(function(r){
          return r[COORD_IDX] !== '' && r[COORD_IDX].toString().toUpperCase().indexOf('APPROVAL')===-1 !== '' && r[COORD_IDX-1].toString().toUpperCase() !== 'X' && r[COORD_IDX-2].toString().toUpperCase().indexOf('WAIT')==-1 && r[COORD_IDX-2].toString().toUpperCase().indexOf('EXAMPLE')==-1
        });
      }
      
      if(approved_data && approved_data.length>0)
      {
        header = sh.getRange(HEADER_ROWS[t], 1, 1, sh.getLastColumn()).getValues();;
        
        console.log('%s - header width: %s',fnName,header[0].length-1);
        
        // build email
        message += "<tr><th colspan=2 style='padding:8px;background-color:rgb(71, 107, 107);color:white;border-bottom:1px solid black;'>"+TABS[t]+"</th><th style='padding:8px;background-color:rgb(71, 107, 107);color:white;border-bottom:1px solid black;'>"+approved_data.length+"</th></tr>";
        
        // TODONE: separate function to write HTML message  
        message += writeHTMLMessage(header,'header');
        
        message += writeHTMLMessage(approved_data,'body');                
        
        message+= "<tr><td><hr></td></tr>"
      }
      
    }
    else
    {
    }
    
  } // # End TABS for loop
  
  message += 
  "</table>\
  </body>\
  </html>";
  
  var to = 'teamdata@momsorganicmarket.com';
  var subject = 'Catapult Changes Approvals';
  
  try
  {
    MailApp.sendEmail({to:to, subject:subject, htmlBody:message, noReply: true});
  }
  catch(e)
  {console.error('%s - %s',fnName,e);}
  
}

/**
* @author {Person} everard.selkridge
* @param {Array} array - 2 dimensional
* @param {String} type - header or body
* @return {String} HTML
*/
function writeHTMLMessage(array,type)
{
  var coltag, padding, cellstyle, message_part = "";
  var rows, cols;
  if(type === 'header')
  {
    coltag = 'th';
    padding = '8px';
    cellstyle = "'padding:"+padding+";background-color:rgb(194, 214, 214);border-bottom:1px solid grey;border-right:1px solid grey;border-left:1px solid grey;'";
  }
  else
  {
    coltag = 'td';
    padding = '5px';
    cellstyle = "'padding:"+padding+";border:1px solid grey;'";
  }
  
  rows = array.length;
  cols = array[0].length;
  
  for(var r=0;r<rows;r++)
  {
    message_part += "<tr>";
    
    for(var c=0;c<cols;c++)
    {
      if(array[r])
      {message_part += "<"+coltag+" style="+cellstyle+">"+array[r][c]+"</"+coltag+">";}
        
    } // End column loop
    
    message_part += "</tr>";
    
  } // End row loop
  
  return message_part;
}