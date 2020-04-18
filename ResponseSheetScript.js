// For help on sheet class, https://developers.google.com/apps-script/reference/spreadsheet/sheet
// For help on form class, https://developers.google.com/apps-script/reference/forms/form
// For help on SpreadsheetApp class, https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app

// function onFromSbmit(e) will be executated each time a form response is submitted.
// For each farm, its own spreadsheet will be updated with list of orders. (It's a brute force 
// implementation by overwriting those sheets rather than incrementally updating them)

// TODO 
// Run it once to generate sheet for last
// Change this line to the line above:     if (dateInMilliSec < lastFriday.getTime() - milliSecInOneDay * 7) {
// Change farmNames
// Set correct farmspreadsheet IDs.

var refFriday = 'Fri Apr 10 2020 00:00:01 GMT-0700 (Pacific Daylight Time)'
var milliSecInOneDay = 24 * 60 * 60 * 1000;
var milliSecInOneWeek = 24 * 60 * 60 * 1000 * 7;
var MAX_ORDERS = 200;
var email_address = 'foo@gmail.com'

var farmNames = ["farm 1", "farm 2", "farm 3"];
var userInfoStartColIndex = 0;
var userInfoEndColIndex   = 3;
var farmStartColIndex     = [4, 8, 10];
var farmEndColIndex       = [7, 9, 24];
var farmSheetIDs = ["1zcdlcn3P",
                    "1G-Q3BCSu", 
                    "19MGX5yYf"];

function getFormId() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var formURL = sheet.getFormUrl();
  var form = FormApp.openByUrl(formURL);
  var formId = form.getId();
  return formId;
}

// Figure out the date of next Friday, and return it as a string
function getNextFriday() {
  let dateOfRefFriday = Date.parse(refFriday);  // In milliseconds
  let dateOfToday = new Date();                 // Date object 
  let intervalInWeeks = (dateOfToday.getTime() - dateOfRefFriday) / milliSecInOneWeek;
  let numWeeksRoundedUp = Math.ceil(intervalInWeeks);
  let targetDate = new Date(dateOfRefFriday + numWeeksRoundedUp * milliSecInOneWeek);
  return targetDate;
}

function sendDebugMsgEmail(msg) {
  var htmlBody = '<ol>';
  htmlBody += '<li>' + msg + '</li>';
  htmlBody += '</ol>';
  GmailApp.sendEmail(email_address, msg, '', {htmlBody:htmlBody});
}

function stopAcceptingFormResponses(){
  var msg = 'Maximum number of orders has been reached for this week. Thank you and please try it next week.';
  var formId = getFormId();
  var form = FormApp.openById(formId);
  //sendDebugMsgEmail(msg);
  form.setAcceptingResponses(false).setCustomClosedFormMessage(msg);
}

// This function is called whenever someone submit the form (one row is added to the sheet)
function onFormSubmit(e){
  // Get spreadsheet data in variable 'values'
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  Logger.log('Number of rows : ' + numRows);
  Logger.log('Number of cols : ' + numCols);
  //sendDebugMsgEmail('Number of rows : ' + numRows);

  // Use date of next Friday as the name of the new sheet in Farmer's spreadsheet.
  let nextFriday = getNextFriday();
  let sheetName = nextFriday.toDateString();
  let lastFriday = new Date(nextFriday.getTime() - milliSecInOneWeek);
  
  // "sheets" is an array of sheets to be worked on, each from a spreadsheet of a farm
  // If the sheet of name sheetName doesn't exist, this is the first order for 
  // the week. Insert a new sheet to the left in the spreadsheet with sheetName.  
  // Otherwise use the existing sheet found
  var sheets = new Array(farmNames.length);
  for (let n = 0; n < farmNames.length; n++) {
    let spreadSheet = SpreadsheetApp.openById(farmSheetIDs[n]);
    let newsheet = spreadSheet.getSheetByName(sheetName);
    if (newsheet == null) {
      sheets[n] = spreadSheet.insertSheet(sheetName, 0);
    } else {
      sheets[n] = newsheet;
    }
  }

  // In the form response sheet, the groups of columns are for each farmer
  // lined from left to right.  When these columns are taken to each farmer's 
  // sheet, the columns needs to shift left, except for the first farmer.
  // Compute the left shift offset and store them in offsets array
  var offsets = new Array(farmNames.length);
  offsets[0] = 0;
  for (let n = 1; n < farmNames.length; n++) {
    offsets[n] = farmEndColIndex[n-1] - farmStartColIndex[n-1] + 1 + offsets[n-1];
  }  
  
  // Write Farmer's sheets.  
  // Source data is from 2D array "values", which starts from [0][0]
  // Destination uses sheet[n].getRange(x,y), which starts from (1, 1)
  
  // Write headers for each farm's spreadsheet
  // Array response is for debug purpose, it's printed out by logger in View->logs
  for (let n = 0; n < farmNames.length; n++) {
    for (let j = userInfoStartColIndex; j <= userInfoEndColIndex; j++) {
      let tmpRange = sheets[n].getRange(0+1, j+1);  // + 1 is for array to range offset
      tmpRange.setValue(values[0][j]);
    }
    for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) {
      let tmpRange = sheets[n].getRange(0+1, j-offsets[n]+1); 
      tmpRange.setValue(values[0][j]);
    }
  }
  
  // Go through all rows in the sheet and update sheet for each farm
  // Ignore rows that's earlier than lastFriday as those data must have
  // been recorded in the previous sheets in farmer's spreadsheets.
  // rowIndices[n] stores which row to be written to the farmer's sheet 
  // as loop variable i can't be used because of skipped rows. It starts
  // from 2, since the first row is the header.
  rowIndices = new Array(farmNames.length).fill(2);
  for (let i = 0; i < values.length; i++) {
    // If the leftmost cell of a row isn't a date, it's not a valid order, ignore it.
    let cell = values[i][0];
    let dateInMilliSec = Date.parse(cell);
    let rowIdx = i + 1;
    if (isNaN(dateInMilliSec)) {
      continue;
    }
    
    // If the date of the order is earlier than lastFriday, then it has been recorded
    // in previous sheet in the farmer's spreadsheets.  Ignore it.
//    if (dateInMilliSec < lastFriday.getTime()) {
    if (dateInMilliSec < lastFriday.getTime() - milliSecInOneDay * 7) {
      continue;
    }
    
    // For each farm, if any produce of it is ordered, 
    // copy the user info and the ordered pruduces to the farm's sheet
    for (let n = 0; n < farmNames.length; n++) {
      // Set hasOrder to true if there is any cell that isn't empty
      let hasOrder = false;
      for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) {
        cell = values[i][j].length;
        if (values[i][j] !== "") { 
          hasOrder = true;  
        }
      }
      if (hasOrder) {
        for (let j = userInfoStartColIndex; j <= userInfoEndColIndex; j++) {
          let tmpRange = sheets[n].getRange(rowIndices[n], j+1);  // + 1 is for array to range offset
          tmpRange.setValue(values[i][j]);
        }
        
        // Set hasOrder to true if there is any cell that isn't empty
        for (let j = farmStartColIndex[n]; j <= farmEndColIndex[n]; j++) {
          let tmpRange = sheets[n].getRange(rowIndices[n], j-offsets[n]+1); 
          tmpRange.setValue(values[i][j]);
        }
        
        rowIndices[n]++;
      }
    }
  }
  
  var numOrderThisWeek = 0;
  // Stop accepting form response if the number of orders within the week is more than a threshold
  if (numOrderThisWeek > MAX_ORDERS) {
    stopAcceptingFormResponses();
  }
}
