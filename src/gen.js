var spreadsheet;
var dataStartRow = 5;
var dataStartColumn = 1;
var hoursSheet;

var lastInsertedDate;

function hook() {
  startDate = new moment("01-09-15 20", "MM-DD-YY HH");
  blackStartDate = new moment("01-08-15 23", "MM-DD-YY HH");
  blueStartDate = new moment("01-18-15 23", "MM-DD-YY HH");
  whiteStartDate = new moment("02-01-15 23", "MM-DD-YY HH");
  whiteEndDate = new moment("02-11-15 23", "MM-DD-YY HH");
  generateSchedule(startDate,blackStartDate,blueStartDate,whiteStartDate,whiteEndDate);
}

function generateSchedule(startDate,blackStartDate,blueStartDate,whiteStartDate,whiteEndDate) {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  crawlerDate = startDate;
  trackingRange = [];
  offset = 0;
  while (crawlerDate < blueStartDate){ 
    sheet = spreadsheet.getSheetByName("Black");
    //select 1x12 range representing each black tenting row
    //Day**Time**Slot 1**Slot 2**Empty**Slot 3**....**Slot 10
    range = sheet.getRange(dataStartRow + offset, dataStartColumn,1,13);
    trackingRange[crawlerDate] = buildSlot(sheet,range,crawlerDate);
    if (isNight(crawlerDate)){
      while (isNight(crawlerDate)){
        crawlerDate.add(1,'h');
      }
      sheet.insertRowAfter(dataStartRow + offset);
      offset++;
    }
    else{
      crawlerDate.add(1,'h');
    }
    offset++;
  }
  
}

function buildSlot(sheet, range, date) {
  sheet.insertRowAfter(range.getLastRow());
  values = [];
  values[0] = [];
  
  if (isNight(date)){
    values[0][0] = "";
    values[0][1] = date.format("h a")+" - "+"EOT";
  }
  else{
    formattedDate = date.format("MMMM Do");
    values[0][0] = lastInsertedDate != formattedDate ? formattedDate : "";
    lastInsertedDate = formattedDate;
    values[0][1] = date.format("h a");
  }
  trackingRange = [range.getRow(),range.getLastRow(),range.getColumn(),range.getLastColumn()];
  for (i = 2; i<range.getWidth(); i++){
    values[0][i] = "";
  }
  range.setValues(values);
  return trackingRange;
}

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

function generateScheduleOld() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ui = SpreadsheetApp.getUi();
  if (spreadsheet.getSheetByName("Black") != null || spreadsheet.getSheetByName("Blue") != null || spreadsheet.getSheetByName("White") != null){
    ui.alert("You'll need to delete the 'Black' 'Blue' and 'White' sheets before generating a schedule.");
    return;
  }
  //do some cleanup
  hoursSheet = spreadsheet.getSheetByName("Hours")
  hoursSheet.getRange(1, 7, 13, 6).clear();
  
  var userStartDate = new Date(spreadsheet.getActiveSheet().getRange(8,2).getValue());
  var userStartTime = new Date(spreadsheet.getActiveSheet().getRange(9,2).getValue());
  userStartDate.setHours(userStartTime.getHours());
  //black sheet creation
  var blackStartDate = new Date(2015,0,8,23);
  var blackEndDate = new Date(2015,0,18,23);
  if (blackEndDate > userStartDate){
    var blackSheet = createOrEmptySheet("Black");
    formatSheet(blackSheet,Math.max(blackStartDate,userStartDate),blackEndDate,2,10,"Black");
  }
  //blue sheet creation
  var blueStartDate = new Date(blackEndDate);
  var blueEndDate = new Date(2015,1,1,23);
  if (blueEndDate > userStartDate){
    var blueSheet = createOrEmptySheet("Blue");
    formatSheet(blueSheet,Math.max(blueStartDate,userStartDate),blueEndDate,1,6,"Blue");
  }
  //white sheet creation
  var whiteStartDate = new Date(blueEndDate);
  var whiteEndDate = new Date(2015,1,11,23);
  var whiteSheet = createOrEmptySheet("White");
  formatSheet(whiteSheet,Math.max(whiteStartDate,userStartDate),whiteEndDate,1,2,"White");
}

function createOrEmptySheet(name){
  if (spreadsheet.getSheetByName(name) != null){
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(name));
  }
  //100 makes sure the sheet inserts after other sheets
  var createdSheet = spreadsheet.insertSheet(name,100);
  return createdSheet;
}

function formatSheet(sheet, startDate, endDate, dayMembers, nightMembers, tName){
  var dayRanges = [];
  var nightRanges = [];
  row = 1;
  column = 2;
  for (var i=0; i<Math.max(dayMembers,nightMembers); i++){
    adjustedIndex = i+1;
    sheet.getRange(row,column).setValue("Person "+adjustedIndex);
    sheet.getRange(row, column).setFontWeight("bold");
    column++;
  }
  sheet.setFrozenRows(1);
  row++;
  dateIter = new Date(startDate);
  while(dateIter < endDate){
    sheet.getRange(row,1).setValue(dateIter.format("ddd. (mmm d)"));
    sheet.getRange(row,1).setFontWeight("bold");
    row++;
    nextDate = dateIter.addDays(1);
    newDayFlag = false;
    while(!newDayFlag){
      if (!dateIter.isNight()){
        if (dateIter.getHours() == 1){
          //slight hack for the 2:30am night start nonsense
          sheet.getRange(row,1).setValue("1 AM - 2:30 AM");
        }
        else{
          sheet.getRange(row,1).setValue(dateIter.format("h tt"));
          sheet.getRange(row,1).setNumberFormat("h am/pm");
        }
        var dayHourRange = sheet.getRange(row,2,1,dayMembers);
        dayHourRange.setBackground("yellow");
        dayRanges.push(dayHourRange);
        dateIter = dateIter.addHours(1);
        row++;
      }
      else{
        var nightStartText;
        if (dateIter.getHours() == 2){
          //slight hack for the 2:30am night start nonsense
          nightStartText = "2:30 AM";
        }
        else{
          nightStartText = dateIter.format("h TT");
        }
        while(dateIter.isNight()){
          dateIter = dateIter.addHours(1);
        }
        newDayFlag = true;
        nightEndText = dateIter.format("h TT");
        sheet.getRange(row,1).setValue(nightStartText+" - "+nightEndText);
        var nightHourRange = sheet.getRange(row,2,1,nightMembers);
        nightHourRange.setBackground("yellow");
        nightRanges.push(nightHourRange);
        row++;
      }
    }
  }
  assignFormula(dayRanges,nightRanges,tName);
}

function assignFormula(dayRanges,nightRanges,tName){
  assignFormulaForRange(dayRanges,"Day Hours",tName);
  assignFormulaForRange(nightRanges,"Nights",tName);
}

function assignFormulaForRange(rangeArray,type,tName){
  var hourHeader = hoursSheet.getRange(1,hourDataStartIndex);
  hourHeader.setValue(tName+" "+type);
  hourHeader.setFontWeight("bold");
  for (var i=0; i<12; i++){
    var adjustedIndex = i+2;
    var hourSlot = hoursSheet.getRange(adjustedIndex,hourDataStartIndex);
    var nameToCheck = hoursSheet.getRange(adjustedIndex, 2);
    var formulaChain = "";
    for (var j=0; j<rangeArray.length; j++){
      if (j != 0){
        formulaChain += "+";
      }
      formulaChain += "COUNTIF("+tName+"!"+rangeArray[j].getA1Notation()+","+nameToCheck.getA1Notation()+")";
    }
    hourSlot.setFormula(formulaChain);
  }
  hourDataStartIndex++;
}

//helper method for adding days to a date
Date.prototype.addDays = function(days)
{
    var dat = new Date(this.valueOf());
    dat.setDate(dat.getDate() + days);
    return dat;
}

//helper method for adding hours to a date
Date.prototype.addHours= function(h){
    var dat = new Date(this.getTime());
    dat.setHours(dat.getHours()+h);
    return dat;
}
