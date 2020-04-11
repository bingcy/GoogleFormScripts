// For help on sheet class, https://developers.google.com/apps-script/reference/spreadsheet/sheet
// For help on form class, https://developers.google.com/apps-script/reference/forms/form

// The format of the spreadsheet looks like the following.  Each farm has 4 rows.
// FARM 1 NAME, 
// PRODUCT 1           , PRODUCT 2           , ..., 
// PRODUCT 1 unit      , PRODUCT 2 unit      , ..., 
// PRODUCT 1 unit price, PRODUCT 2 unit price, ..., 
// FARM 2 NAME, 
// ...

let NUM_ROWS_PER_PRODUCT = 3;

function ReadSheetAndCreateForm() {
  
   // Create & name Form  
   var item = "Creekside Farmers' Market, Cupertino";  
   var form = FormApp.create(item)  
       .setTitle(item)
       .setDescription('Be Safe.  Stay Healthy!')
       .setCollectEmail(true)
       .setConfirmationMessage('An order summary has been sent to your email address for your record.  Thank you for shopping at your local farmers market.')
       .setProgressBar(true)
      
   var img = DriveApp.getFileById('13o2pAUJogA9NBItao9oTPYhRfMpA806l');
   form.addImageItem()
       .setImage(img);
  
   item = "Customer's First Name";  
   form.addTextItem()  
       .setTitle(item)  
       .setRequired(true);  
   
   item = "Customer's Contact Phone # (Tell the seller your phone # when pick up)";  
   form.addTextItem()  
       .setTitle(item)  
       .setRequired(true);  

  // Read the spreadsheet and create one form section per farm
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Get a Range bounded by A1 and (Range.getLastColumn(), Range.getLastRow())
  var range = sheet.getDataRange();
  var data = range.getValues();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var rowIndex = 0;
  // data is a group of a two-dimentional array of cells in the spreadsheet 
  while (rowIndex < numRows) {

    // Create a new section in the form for each farm
    form.addPageBreakItem()
        .setTitle(data[rowIndex][0]);
    //Logger.log('Farm Name: ' + data[rowIndex][0]);
    img = DriveApp.getFileById(data[rowIndex][1]);
    form.addImageItem()
        .setImage(img);
    rowIndex++;
  
    for (let i = 0; i < numCols; i++) {
      if (data[rowIndex][i].length > 0) {
        item = data[rowIndex][i] + ' ($' + data[rowIndex+2][i] + ' per ' + data[rowIndex+1][i] + ')';
        //Logger.log('Product name: ' + data[rowIndex][i]);
        //Logger.log('Unit: ' + data[rowIndex+1][i]);
        //Logger.log('Unit price: ' + data[rowIndex+2][i]);
        form.addScaleItem()
            .setTitle(item)
            .setLabels(data[rowIndex+1][i], '')
            .setBounds(0, 10);
      }
    }
    rowIndex = rowIndex + NUM_ROWS_PER_PRODUCT;
  }
}
