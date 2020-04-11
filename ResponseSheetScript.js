var MAX_ORDERS = 4;

function getFormId() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var formURL = sheet.getFormUrl();
  var form = FormApp.openByUrl(formURL);
  var formId = form.getId();
  return formId;
}

function sendDebugMsgEmail(msg) {
  var htmlBody = '<ol>';
  htmlBody += '<li>' + msg + '</li>';
  htmlBody += '</ol>';
  GmailApp.sendEmail('btbingtian@gmail.com', msg, '', {htmlBody:htmlBody});
}

function stopAcceptingFormResponses(e){
  // Count how many rows are in the sheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  //Logger.log('Number of rows : ' + numRows);
  //sendDebugMsgEmail('Number of rows : ' + numRows);
  
  if (numRows > MAX_ORDERS) {
    var msg = 'Maximum number of orders has been reached for this week. Thank you and please try it next week.';
    var formId = getFormId();
    var form = FormApp.openById(formId);
    //sendDebugMsgEmail(msg);
    form.setAcceptingResponses(false).setCustomClosedFormMessage(msg);
  }
}

