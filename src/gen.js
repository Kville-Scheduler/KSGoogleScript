var spreadsheet;
var file;
var captainName;

var lastInsertedDate;

var DEV = true;

var PUBLIC_TEMPLATE_ID = "1BnG71G8WHl3XLwKxZtjpOWtQhECC_xJ8_ALk7ZyQkTg";
var DEV_ID = "16a4ukXaFV00BQJNlesRWDwdBDoqrq03MWn96haAePHo";

function hook() {

  startDate = "2015-01-15T23:00:00.000Z"
  blackStartDate = "2015-01-08T23:00:00.000Z"
  blueStartDate = "2015-01-18T23:00:00.000Z"
  whiteStartDate = "2015-02-01T23:00:00.000Z"
  whiteEndDate = "2015-02-11T23:00:00.000Z"
  username = "Davis Gossage"
  generateSchedule(DEV_ID,startDate,blackStartDate,blueStartDate,whiteStartDate,whiteEndDate,username);
}

function createSheet(){
  if (DEV){
    //access testbed spreadsheet
    spreadsheet = SpreadsheetApp.openById(DEV_ID);
  }
  else{
    //access template spreadsheet and make a copy for the user
    templateSS = SpreadsheetApp.openById(PUBLIC_TEMPLATE_ID);
    spreadsheet = templateSS.copy("Tenting Schedule");
  }

  return spreadsheet.getId();
}

function generateSchedule(spreadsheetId, startDate,blackStartDate,blueStartDate,whiteStartDate,whiteEndDate,username) {
  //initialize dates string->moment
  startDate = new moment(startDate);
  blackStartDate = new moment(blackStartDate);
  blueStartDate = new moment(blueStartDate);
  whiteStartDate = new moment(whiteStartDate);
  whiteEndDate = new moment(whiteEndDate);

  captainName = username;

  spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  //access drive file for sharing options that SpreadsheetApp doesn't offer
  file = DriveApp.getFileById(spreadsheet.getId());
  file.setSharing(DriveApp.Access.ANYONE,DriveApp.Permission.EDIT);

  crawlerDate = moment(startDate);
  trackingRange = [];

  var trackingRange = generateSheet("Black",startDate,blueStartDate,2,10);
  
  notifyParseOfCompletion(trackingRange);
}

function generateSheet(name,startDate,endDate,numDay,numNight) {
  sheet = spreadsheet.getSheetByName(name);
  var dataStartRow = 5
  var dataStartColumn = 1
  var offset = 0

  crawlerDate = moment(startDate);
  var wasNight = true;
  var trackingRange = [];


  while (startDate < endDate){
    range = sheet.getRange(dataStartRow + offset, dataStartColumn,1,13);

    //skip over the night time, will display it as a single row ex: '9PM-8AM'
    while(isNight(crawlerDate)){
      crawlerDate.add(1,'h')  
    }
    offset = buildSlotContent(sheet, range, startDate, crawlerDate, wasNight, offset)
    trackingRange.push(buildSlotGridAndValidate(sheet,range,startDate,crawlerDate,numDay,numNight))

    //add a single hour for the next row, unless time was added for night
    if(!isNight(startDate)){
      crawlerDate.add(1,'h')
    }
    //only want to display date in the spreadsheet following a night
    wasNight = isNight(startDate);
    //reset the startDate to match the date we crawled to
    startDate = moment(crawlerDate);
  }

  return trackingRange
}

/**
Build a slot object for storing in db

@param range
  Range of slot to build
@param dateStart
  Starting time for night slot (moment.js)
@param dateStop
  Stopping time for night slot (moment.js)
*/
function buildSlotObject(range,dateStart,dateStop) {
  //no end date implies day slot
  return {startRow:range.getRow(),endRow:range.getLastRow(),startColumn:range.getColumn(),
    endColumn:range.getLastColumn(),startDate:moment(dateStart),endDate:moment(dateStop)};
}

/**
Generic slot builder.

@param sheet
  Google sheet
@param range
  Range of slot to build
@param values
  Empty or partially completed array representing range values
*/
function buildSlotContent(sheet, range, date, endDate, wasNight, offset) {
  sheet.insertRowAfter(range.getLastRow());
  offset++;

  values = [];
  values[0] = [];
  formattedDate = date.format("MMMM Do")
  values[0][0] = wasNight ? formattedDate : ""
  if (isNight(date)){
    values[0][1] = date.format("h a")+" - "+endDate.format("h a")
    sheet.insertRowAfter(range.getLastRow());
    offset++;
  }
  else{
    values[0][1] = date.format("h a")
  }
  for (i = values[0].length; i<range.getWidth(); i++){
    values[0][i] = "";
  }

  range.setValues(values);

  return offset;
}

function buildSlotGridAndValidate(sheet, range, date, endDate, numDay, numNight) {
  //this accounts for Day and Time columns
  TENTER_OFFSET = 2;

  var numberTenting = isNight(date) ? numNight : numDay;

  buildGrid(sheet, range.getRow(), range.getColumn() + TENTER_OFFSET, 1, numberTenting);

  var validatedRange = sheet.getRange(range.getRow(), range.getColumn() + TENTER_OFFSET, 1, numberTenting);
  appendNamesToCellValidation([captainName],validatedRange)

  return buildSlotObject(validatedRange,date,endDate)
}

/**
Generic grid builder.  Puts a border around every cell in a given range.

@param sheet
  Google sheet
@param row
  Initial row
@param col
  Initial col
@param numRows
  Number of rows
@param numCols
  Number of columns
*/
function buildGrid(sheet, row, col, numRows, numCols) {
  var griddedDayRange = sheet.getRange(row,col,numRows,numCols);
  griddedDayRange.setBorder(true,true,true,true,true,true);

  //hack for setting border color
  //feature request at https://code.google.com/p/google-apps-script-issues/issues/detail?id=2002
  /*
  broken pending copyFormatToRange fix
  https://code.google.com/p/google-apps-script-issues/issues/detail?id=5304
  var firstCell = sheet.getRange(1,1);
  firstCell.copyFormatToRange(sheet,row,col,numRows,numCols);
  */
}

function appendNamesToCellValidation(names, range) {
  var rule = range.getDataValidation();
  if (rule == null){
    rule = SpreadsheetApp.newDataValidation().requireValueInList(names).setAllowInvalid(false);
  }
  else{
    ruleBuilder = rule.copy()
    var allNames = rule.getCriteriaValues()
    allNames.push(names)
    ruleBuilder.requireValueInList(allNames)
    rule = ruleBuilder.build()
  }
  range.setDataValidation(rule);
}

function notifyParseOfCompletion(trackingRange) {
  
  var headers = {
    "X-Parse-Application-Id" : PropertiesService.getScriptProperties().getProperty('PARSE_APPLICATION_ID'),
    "X-Parse-REST-API-Key" : PropertiesService.getScriptProperties().getProperty('PARSE_REST_API_KEY'),
    "Content-Type" : "application/json"
  };
  
  var payload = {
    "jsonSlots" : JSON.stringify(trackingRange)
  }

  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : payload
  }

  UrlFetchApp.fetch("https://api.parse.com/1/functions/recordSlots",options)
}

/**
Convenience method for determining if a moment.js object is a night according to k-ville policy

@param date
  Moment.js date object
*/
function isNight(date) {
  if (date.day() == 0){
    return date.hour() >= 2 && date.hour() < 10 || date.hour() == 23;
  }
  else if (date.day() == 1 || date.day() == 2){
    return date.hour() < 7 || date.hour() == 23;
  }
  else if (date.day() == 3){
    return date.hour() < 7;
  }
  else if (date.day() == 4 || date.day() == 5){
    return date.hour() >= 2 && date.hour() < 7;
  }
  else{
    return date.hour() >= 2 && date.hour() < 10;
  }  
}
