var ticketSubjects = {};
/**
 * Retrieves all the Time Entries from Redmine for current week
 */
function getCurrentWeekEntries() {
  var offset=0;
  var limit=100;
  
  var ui = SpreadsheetApp.getUi();
  var username = ui.prompt('Enter Redmine User Name:');
  var password = ui.prompt('Enter Redmine Password:'); 
  var recordsCount = 0;
  
  do {
    var result = getDataFromRedmine(username.getResponseText(),password.getResponseText(),offset,limit); 
    var data = JSON.parse(result);
    var userdata = fetchUserData(data.time_entries, username.getResponseText(), password.getResponseText());   
    offset=offset + limit;
    var sheetName='TimeEntry_'+formatDate(new Date());
    createSheet(userdata,sheetName,recordsCount);   
    recordsCount = recordsCount + userdata.length;
  }
  while (data.time_entries[0].spent_on >= getMonday(new Date()));  
  //SpreadsheetApp.getUi().alert("This week entries imported from Redmine");
};

/**
 * Retrieves all the Time Entries from Redmine for user provided dates
 */
function getCustomEntries() {  
  var limit=100;
  var ui = SpreadsheetApp.getUi();
  var username = ui.prompt('Enter Redmine User Name:');
  var password = ui.prompt('Enter Redmine Password:'); 
  var recordsCount = 0;  
  
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("DataInput");
  //Read Offset from sheet
  var offset=sheet.getRange(2,1).getValue();
  if(offset!=0){
    recordsCount = offset;
  }
  
  //var sd=ui.prompt('Enter Start Date in YYYY-MM-DD format:');
  //var startDate=sd.getResponseText();  
  //var ed=ui.prompt('Enter End Date in YYYY-MM-DD format:');
  //var endDate=ed.getResponseText();
  
  
  var startDate = formatDate(new Date(sheet.getRange(2,2).getValue()));
  var endDate = formatDate(new Date(sheet.getRange(2,3).getValue()));
  
  //SpreadsheetApp.getUi().alert(offset);
  //SpreadsheetApp.getUi().alert(startDate);
  //SpreadsheetApp.getUi().alert(endDate);
  //return;
  
  if(startDate>endDate){
    SpreadsheetApp.getUi().alert("End Date should be greater than Start Date");
    return;
  }
  
  do {
    var result = getDataFromRedmine(username.getResponseText(),password.getResponseText(),offset,limit); 
    var data = JSON.parse(result);
    var userdata = fetchCustomUserData(data.time_entries, startDate, endDate, username.getResponseText(), password.getResponseText());
    offset=offset + limit;
    var sheetName='TimeEntry_'+startDate+':'+endDate;
    createSheet(userdata,sheetName,recordsCount);
    recordsCount = recordsCount + userdata.length;
    sheet.getRange(2,1).setValue(offset);    
  }
  while (data.time_entries[data.time_entries.length-1].spent_on >= startDate); 
  //SpreadsheetApp.getUi().alert("Time entries imported from Redmine");
};



/**
 * Retrieves all the Time Entries from Redmine for last week
 */
function getLastWeekEntries() {
  var offset=0;
  var limit=100;
  var ui = SpreadsheetApp.getUi();
  var username = ui.prompt('Enter Redmine User Name:');
  var password = ui.prompt('Enter Redmine Password:'); 
  var recordsCount = 0;
  
  var startDate=getLastWeekStartDate();
  var endDate=getLastWeekEndDate();
  do {
    var result = getDataFromRedmine(username.getResponseText(),password.getResponseText(),offset,limit); 
    var data = JSON.parse(result);
    var userdata = fetchCustomUserData(data.time_entries, startDate, endDate, username.getResponseText(), password.getResponseText());   
    offset=offset + limit;
    var sheetName='TimeEntry_'+startDate+':'+endDate;
    createSheet(userdata,sheetName,recordsCount);
    recordsCount = recordsCount + userdata.length;
  }
  while (data.time_entries[data.time_entries.length-1].spent_on >= startDate);
  //SpreadsheetApp.getUi().alert("Last week entries imported from Redmine");
};


/**
 * Fetch data from the redmine API Data
 */
function fetchCustomUserData(data, startDate, endDate,  username, password){    
  var timeEntry = [];
  
  for(var i in data) {    
    var item = data[i];
    if (item.spent_on >= startDate && item.spent_on <= endDate) {
      var subject = getSubjectFromRedmine(username, password, item.issue.id);
      timeEntry.push({ 
        "userName" : item.user.name,
        "projectType" : getProjectType(item.project.name),
        "projectName"  : item.project.name,
        "ticketNumber" : item.issue.id,
        "ticketDescription" : subject, 
        "date" : item.spent_on,
        "spentHours" : item.hours,
        "billingCategory" : item.custom_fields[0].value,
        "activity" : item.activity.name,
        "comments" : item.comments      
      });
    }
  }
  return timeEntry;
}


function getProjectType(projectName){
  if(projectName.indexOf("QA") > -1){
    return "QA"
  }
  if(projectName.indexOf("DevBI") > -1){
    return "DevBI";
  }
  if(projectName.indexOf("Mobile") > -1){
    return "Mobile";
  }
  if(projectName.indexOf("Induction") > -1){
    return "Induction";
  }
  if(projectName.indexOf("API") > -1){
    return "API";
  }
  if(projectName.indexOf("Leaves") > -1){
    return "Leave";
  }
  if(projectName.indexOf("Dev/BI") > -1){
    return "DevBI";
  }
  if(projectName.indexOf("BI") > -1){
    return "BI";
  }  
  else {
    return "Organization";
  }
}


/**
 * Fetch data from the redmine API Data
 */
function fetchUserData(data, username, password){    
  var timeEntry = [];
  for(var i in data) {    
    var item = data[i];
    if (item.spent_on >= getMonday(new Date())) { 
      var subject = getSubjectFromRedmine(username, password, item.issue.id);
      timeEntry.push({ 
        "userName" : item.user.name,
        "projectType" : getProjectType(item.project.name),
        "projectName"  : item.project.name,
        "ticketNumber" : item.issue.id,
        "ticketDescription" : subject, 
        "date" : item.spent_on,
        "spentHours" : item.hours,
        "billingCategory" : item.custom_fields[0].value,
        "activity" : item.activity.name,
        "comments" : item.comments     
      });
    }
  }
  return timeEntry;
}

function getLastWeekStartDate(){
  var curr = new Date; // get current date
  var first = curr.getDate() - curr.getDay() - 6; 
  var last = first + 6; // last day is the first day + 6
  var startDate = formatDate(new Date(curr.setDate(first)));
  return startDate;
}

function getLastWeekEndDate(){
  var curr = new Date; // get current date
  var first = curr.getDate() - curr.getDay() - 6;
  var last = first + 6; // last day is the first day + 6
  var endDate = formatDate(new Date(curr.setDate(last))); 
  return endDate;
}



function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(),
      diff = d.getDate() - day + (day == 0 ? -6:1); // adjust when day is sunday
  return formatDate(new Date(d.setDate(diff)));
}


function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}



/**
 * Get Data from redmine API
 */
function getDataFromRedmine(username,password,offset,limit){

  var url = 'https://portal.optimusinfo.com/redmine/time_entries.json?offset='+offset+'&limit='+limit; 
  var blob = Utilities.newBlob(username+":"+password);
  var encoded = Utilities.base64Encode(blob.getBytes());  
  var headers = {
    'Authorization': "Basic " + encoded    
  };  
  var options = {
    'method': 'get',
    'contentType': 'text/json',
    'headers': headers
  };
  
  var response = UrlFetchApp.fetch(url,options);  
  return response.getContentText();  
}

/**
 * Get Issue Subject from redmine API
 */
function getSubjectFromRedmine(username,password,issueId){ 
  
  if(ticketSubjects[issueId] != null && ticketSubjects[issueId] != " "){
    return ticketSubjects[issueId];
  }
  var url = 'https://portal.optimusinfo.com/redmine/issues/'+issueId+'.json'; 
  var blob = Utilities.newBlob(username+":"+password);
  var encoded = Utilities.base64Encode(blob.getBytes());  
  var headers = {
    'Authorization': "Basic " + encoded    
  };  
  var options = {
    'method': 'get',
    'contentType': 'text/json',
    'headers': headers
  };
  
  try {
    var response = UrlFetchApp.fetch(url,options);
    var result = JSON.parse(response.getContentText());
    ticketSubjects[issueId] = result.issue.subject;
    return result.issue.subject;
  } catch (e) {    
    ticketSubjects[issueId] = " ";
    return " ";  
  }   
}



/**
 * Creates sheet and sets header
 */
function createSheet(data,sheetname, rows){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var temp = doc.getSheetByName("TMP");
  if (!doc.getSheetByName(sheetname)){
    var sheet = doc.insertSheet(sheetname, {template:temp});
    var headerNames = ["User Name", "Project Type", "Project Name", "Ticket Number", "Ticket Description",  "Date", "Spent Hours", "Billing Category", "Activity", "Comments"];
    sheet.getRange(1, 1, 1, headerNames.length).setValues([headerNames]); 
  } else {
    var sheet = doc.getSheetByName(sheetname);
    //sheet.getRange(2, 1, sheet.getLastRow(), sheet.getMaxColumns()).clear({contentsOnly:true});    
    var headerNames = ["User Name", "Project Type", "Project Name", "Ticket Number", "Ticket Description",  "Date", "Spent Hours", "Billing Category", "Activity", "Comments"];
    sheet.getRange(1, 1, 1, headerNames.length).setValues([headerNames]); 
  } 
  insertData(sheet,data, rows);
}


/**
 * Fetch data in the sheet
 */
function insertData(sheet, data, rows){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (data.length>0){
    ss.toast("Inserting "+data.length+" rows");
    sheet.insertRowsAfter(rows + 1, data.length);
    setRowsData(sheet, data, rows + 2);
  }
}

/**
 * Inserts data in the row
 */
function setRowsData(sheet, objects, optFirstDataRowIndex, optHeadersRange) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders(headersRange.getValues()[0]);
  
  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
    }
    data.push(values);
  }
  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
                                        objects.length, headers.length);
  destinationRange.setValues(data);
} 



// Returns an Array of normalized Strings.
// Arguments:
// - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}
 
// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
// - header: string to normalize
// Examples:
// "First Name" -> "firstName"
// "Market Cap (millions) -> "marketCapMillions
// "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    //if (!isAlnum(letter)) {
    // continue;
    //}
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}
 
// Returns true if the cell where cellData was read from is empty.
// Arguments:
// - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}
 
// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
      isDigit(char);
}
 
// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}
// http://jsfromhell.com/array/chunk
function chunk(a, s){
  for(var x, i = 0, c = -1, l = a.length, n = []; i < l; i++)
    (x = i % s) ? n[c][x] = a[i] : n[++c] = [a[i]];
  return n;
} 

/**
 * Adds a custom menu to the active spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [{
      name : "This week entries",
      functionName : "getCurrentWeekEntries"},
    { 
      name : "Last week entries",
      functionName : "getLastWeekEntries"},
    { 
      name : "Custom entries",
      functionName : "getCustomEntries"}];
  
  spreadsheet.addMenu("Redmine", menuItems);
};





