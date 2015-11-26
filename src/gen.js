var spreadsheet;
var dataStartRow = 5;
var dataStartColumn = 1;
var hoursSheet;

var lastInsertedDate;

var DEV = true;

var PUBLIC_TEMPLATE_ID = "1BnG71G8WHl3XLwKxZtjpOWtQhECC_xJ8_ALk7ZyQkTg";
var DEV_ID = "16a4ukXaFV00BQJNlesRWDwdBDoqrq03MWn96haAePHo";

function hook() {
  startDate = new moment("01-09-15 20", "MM-DD-YY HH");
  blackStartDate = new moment("01-08-15 23", "MM-DD-YY HH");
  blueStartDate = new moment("01-18-15 23", "MM-DD-YY HH");
  whiteStartDate = new moment("02-01-15 23", "MM-DD-YY HH");
  whiteEndDate = new moment("02-11-15 23", "MM-DD-YY HH");
  generateSchedule(startDate,blackStartDate,blueStartDate,whiteStartDate,whiteEndDate);
}

function generateSchedule(startDate,blackStartDate,blueStartDate,whiteStartDate,whiteEndDate) {
  if (DEV){
    //access testbed spreadsheet
    spreadsheet = SpreadsheetApp.openById(DEV_ID);
  }
  else{
    //access template spreadsheet and make a copy for the user
    templateSS = SpreadsheetApp.openById(PUBLIC_TEMPLATE_ID);
    spreadsheet = templateSS.copy("Tenting Schedule");
  }
  crawlerDate = moment(startDate);
  trackingRange = [];

  offset = 0;
  while (startDate < blueStartDate){ 
    sheet = spreadsheet.getSheetByName("Black");
    //select 1x12 range representing each black tenting row
    //Day**Time**[begin gridded]Slot 1**Slot 2[end gridded]**Empty**[begin gridded]Slot 3**....**Slot 10[end gridded]
    var BLACK_WIDTH = 13;
    var BLACK_DAY_WIDTH = 3;
    var BLACK_NIGHT_WIDTH = 8;

    range = sheet.getRange(dataStartRow + offset, dataStartColumn,1,BLACK_WIDTH);
    if (isNight(crawlerDate)){
      while (isNight(crawlerDate)){
        crawlerDate.add(1,'h');
      }
      trackingRange[startDate] = buildNightSlot(sheet,range,startDate,crawlerDate);
      sheet.insertRowAfter(dataStartRow + offset);
      offset++;
    }
    else{
      trackingRange[startDate] = buildDaySlot(sheet,range,startDate,wasNight);
      crawlerDate.add(1,'h');
    }
    buildGrid(sheet,range.getRow(),2,1,BLACK_DAY_WIDTH);
    buildGrid(sheet,range.getRow(),6,1,BLACK_NIGHT_WIDTH);
    offset++;
    //only want to display date in the spreadsheet following a night
    var wasNight = isNight(startDate);
    //reset the startDate to match the date we crawled to
    startDate = moment(crawlerDate);
  }
  
}

/**
Build a slot for nights covering a time interval instead of a specific hour.

@param sheet
  Google sheet
@param range
  Range of slot to build
@param dateStart
  Starting time for night slot (moment.js)
@param dateStop
  Stopping time for night slot (moment.js)
*/
function buildNightSlot(sheet, range, dateStart, dateStop) {
  values = [];
  values[0] = [];
  values[0][0] = "";
  values[0][1] = dateStart.format("h a")+" - "+dateStop.format("h a");
  buildSlot(sheet,range,values);
}

/**
Build a slot for days covering a specific hour.

@param sheet
  Google sheet
@param range
  Range of slot to build
@param date
  Time for slot (moment.js)
*/
function buildDaySlot(sheet, range, date, isNewDay) {
  values = [];
  values[0] = [];
  formattedDate = date.format("MMMM Do");
  values[0][0] = isNewDay ? formattedDate : "";
  values[0][1] = date.format("h a");
  buildSlot(sheet,range,values);
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
function buildSlot(sheet, range, values) {
  sheet.insertRowAfter(range.getLastRow());

  trackingRange = [range.getRow(),range.getLastRow(),range.getColumn(),range.getLastColumn()];
  for (i = values[0].length; i<range.getWidth(); i++){
    values[0][i] = "";
  }
  range.setValues(values);
  return trackingRange;
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
