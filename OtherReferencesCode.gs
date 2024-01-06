
// Add Spaces to Links

function addSpaceToLinks() {
    // Set the sheet name and range
    const sheetName = "Expenses";
    const range = "K19";
    
    // Get the sheet and range
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const cell = sheet.getRange(range);
    
    // Get the cell value and split the links using comma as delimiter
    const links = cell.getValue().split(",");
     Logger.log(links);
  
    // Loop through the links and append a space after each link
    const updatedLinks = links.map(link => link.trim() + "  ");
    Logger.log(updatedLinks);
  
    // Set the updated links back to the cell
    cell.setValue(updatedLinks.join(" , "));
  }
  
  function getFileByUrl(url) {
    var fileId = url.match(/[-\w]{25,}/); // Extracts the file ID from the URL
    Logger.log(fileId[0]);
    if (fileId) {
      var file = DriveApp.getFileById(fileId[0]);
      return file;
    } else {
      return null;
    }
  }


///////////////////////////////////////////////   Menuopts.gs
/////////
/////
////
///
//
//////
/////
//


  function createDocument() {
  
    var ui = SpreadsheetApp.getUi();
    var fileName = ui.prompt("File Name");
    var reportHeader = ui.prompt("Report Header");
    
    var ss= SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName('Expenses'); 
    var dataRange = dataSheet.getDataRange();
    var dataValues = dataRange.getValues();
    
    var destination_id = '1mJu-DWuVgHjdF2iFLGWEK_cxvV9iLaMn';  // ID OF GOOGLE DRIVE DIRECTORY;
    var destination = DriveApp.getFolderById(destination_id);
    
    var doc = DocumentApp.create(fileName.getResponseText());
    var docID = doc.getId();
    var file = DriveApp.getFileById(docID);
    file.moveTo(destination);
    
    var body = doc.getBody();
    body.insertParagraph(0, reportHeader.getResponseText())
       .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    table = body.appendTable(dataValues);
    
    
    var tableHeader = table.getRow(0);
    tableHeader.editAsText().setBold(true).setForegroundColor('#ffffff');
    var getCells = tableHeader.getNumCells();
    
    for(var i = 0; i < getCells; i++)
    {
      tableHeader.getCell(i).setBackgroundColor('#BBB9B9');
    }
  }
  
  function addMenu()
  {
    var menu = SpreadsheetApp.getUi().createMenu('Custom');
    menu.addItem('Create Doc', 'createDocument');
    menu.addToUi(); 
  }
  
  function onOpen(e)
  {
    addMenu(); 
  }


///////////////////////////////////////////////Macros.gs
/////////
/////
////
///
//
//////
/////
//

function copyTemplateDoc() {
  var templateDocUrl = "https://docs.google.com/document/d/1zeZP2FepGn7ZDVW0pTX9BQEHgeSrANyLCP-fe1RpkSU/edit";
  var dstFolderUrl = "https://drive.google.com/drive/folders/1HTY82ojyno5V6e4KYPTA7K_uO_7btYzi?usp=share_link";
  var copyFileName = "Gani01";
  var templateFile = DriveApp.getFileById(getID(templateDocUrl)); // replace TEMPLATE_DOC_ID with the ID of your template document
  var folder = DriveApp.getFolderById(getID(dstFolderUrl)); // replace FOLDER_ID with the ID of your destination folder
  var newFile = templateFile.makeCopy(copyFileName, folder); // replace COPY_NAME with the desired name for the copied document
  newFile.moveTo(folder);
  Logger.log(" Doc:"+ copyFileName+ " ->" + newFile.getUrl());
}

function replaceTextInDoc() {
  var doc = DocumentApp.openById('1uawSL-isYDFAWXaRUh5bxM0lWsekVNNDOJfMr2phvNE'); // Replace 'DOCUMENT_ID' with the ID of your document
  var searchText = 'Category Expense'; // Replace 'TEXT_TO_BE_REPLACED' with the text you want to replace
  var replacementText = 'JPNCE Lab Expense'; // Replace 'REPLACEMENT_TEXT' with the text you want to replace the old text with
  var body = doc.getBody();
  body.replaceText(searchText, replacementText);
}

function getID(url){
  var fileId = url.match(/[-\w]{25,}/); // Extracts the file ID from the URL
  Logger.log( fileId[0] );
  if( fileId[0]){
    return fileId[0] ;
  }
  else{
    return null ;
  }
}

//===============================================


function createTableWithPhoto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Expenses");
  var data = sheet.getDataRange().getValues();

 var body = DocumentApp.openById("11HI_c6HhibEDtNTZqb0wgW4qAoCyksKt4hqBeFeoaQ8").getBody();

  // Create table
  var table = "<table><tr><th>S. No</th><th>Date</th><th>Expense Description</th><th>Amount</th><th>Remarks</th></tr>";
  for (var i = 1; i < data.length; i++) {
    table += "<tr><td>" + i + "</td><td>" + data[i][1] + "</td><td>" + data[i][2] + "</td><td>" + data[i][3] + "</td><td>" + data[i][4] + "</td></tr>";
  }
  table += "</table>";

  // Add table to document
  body.appendParagraph(table);

  // Add photo to document
  var photoUrl = data[2][10];
  if (photoUrl) {
    var response = UrlFetchApp.fetch(photoUrl);
    var contentType = response.getHeaders()['Content-Type'];
    var image = response.getBlob().setContentType(contentType);
    var blob = image.setName("Expense Receipt");
    body.appendImage(blob);
  }
}

function getData00(){
  //getting bills from sheets
  var personCategory ="Ganesh+Texoham" ;
  var name = "Ganesh";
  var date ="28/05/2023";
  var ss = SpreadsheetApp.openById("1IFi0nkBTzG98Op_DI9BJuMUk3BIlPDkWw6d3SDgnJ7k").getSheetByName("Expenses");
  var dataValues = ss.getDataRange().getValues();
  var currentBills = [] ;
  var docBills =[];
  var billCount = 0 ;
  var totalPrice = 0;
  Logger.log(dataValues);

  for(let count = 0 ; count < dataValues.length ; count++){
     Logger.log(dataValues[count][12]);
     if( dataValues[count][12] == personCategory ){ //column of the "Paid"
         currentBills.push(dataValues[count]); 
         billCount += 1;
         totalPrice += dataValues[count][8] ; //column of the  "debit"
         //Logger.log(Utilities.formatDate(dataValues[count][1], Session.getScriptTimeZone(), "dd-MMM-yyyy"));  formating date 
         //s.no , date , reason , cost ,remarks
         docBills.push([billCount,Utilities.formatDate(dataValues[count][1], Session.getScriptTimeZone(), "dd-MMM-yyyy"),dataValues[count][5],dataValues[count][8],""])   
      }
  }
  Logger.log(totalPrice);
  Logger.log(currentBills);
  Logger.log(docBills);
  if(billCount < 20){
  editDoc(name,date,docBills,billCount,totalPrice,currentBills);
  }
  else {
    return "failed";
  }
}

function editDoc(name , date , bills,billCount,totPrice,currentBills){
  let id = "1tyfTani7aZC4ryr63czIFaHu8nIXUOj8mvMXhgalDYE";
  let imageUrl ="https://drive.google.com/open?id=1ri1T4JV7QmEgg3JWZ8Rj3GgO1sas8HoR";
  var doc = DocumentApp.openById(id);
  var table = doc.getBody().getTables()[0];
  Logger.log(bills);
  for(let i = 0 ; i < bills.length ; i++ ){
    let currentRow = table.getRow(i+1);
    for(let j= 0 ; j < 5  ; j++ ){  //num of columns
      Logger.log(bills[i][j]);
      currentRow.getCell(j).setText(bills[i][j]);
    }
  }
  let numRows = table.getNumRows();  //num of rows
  Logger.log(numRows);
  var row = table.removeRow(numRows-1); //row last index = num of rows - 1
  for(let i = (numRows-2) ; i > bills.length ;i-- ){  //one row is removed ,current last rows index = now of rows - 2  
   table.removeRow(i);
  }
  table.appendTableRow(row);
  table.getRow(table.getNumRows()-1).getCell(3).setText(totPrice); // last row index = num of rows - 1

   //adding image to doc
  //  const image = UrlFetchApp.fetch(imageUrl).getBlob();
  //  //doc.getCursor.insertInlineImage(image);
  //  const inlineImage = doc.getBody().appendImage(image);
  
   var fileId = '1ri1T4JV7QmEgg3JWZ8Rj3GgO1sas8HoR';
 var img = DriveApp.getFileById(fileId).getBlob();
  doc.getBody().insertImage(0, img); 
 //doc.getChild(0).asParagraph().appendInlineImage(resp.getBlob());

}



function test1(){
let id1 ="11HI_c6HhibEDtNTZqb0wgW4qAoCyksKt4hqBeFeoaQ8"; 
let id2 ="1p_DlV24tR899gpIFAkoKXDe3t7flm_CmcB2zIx-i2Zg";
let id3 ="117ReMBhPmqpHTjF1kL-o9a8ZJo3YiAQphr23d_5y6bM";
var doc = DocumentApp.openById(id3);
var headers =["5","16-Feb-2023","10 Cab Charges", "gani","4"];
var table = doc.getBody().getTables()[0];
var headerRow = table.getRow(1);

for (var i = 0; i < headers.length; i++) {
    headerRow.getCell(i).setText(headers[i]);
}
  table.appendTableRow();
  for(let i= 5 ;i > 1  ; i--){
  table.removeRow(i);}
}


///////////////////////////////////////////////     Other Codes
/////////
/////
////
///
//
//////
/////
//

function getSchool1Msg() {
let schoolId = activeSheet.getRange("B3").getValue();
var schoolName = activeSheet.getRange("B5").getValue();
let date = activeSheet.getRange("B7").getValue();
let startTime = getCellTime12HourFormat("B9");
let closeTime = getCellTime12HourFormat("B11");
let reportingTime = getCellTime12HourFormat("B13");
let bringLunch =activeSheet.getRange("B15").getValue();
let locationLink = activeSheet.getRange("B17").getValue();
let folderName = activeSheet.getRange("B19").getValue();
let subFolderName = activeSheet.getRange("B21").getValue();
let trackerSheetName = activeSheet.getRange("B23").getValue();
let sourceSheetId= activeSheet.getRange("B25").getValue();
let destFolderId=activeSheet.getRange("B27").getValue();
let teamLead = activeSheet.getRange("E4").getValue(); 
let trainers = getTrainers("E5:E21");
let msgBox = "D26";
let msg = getMessage(schoolName,date,startTime,closeTime,reportingTime,bringLunch,locationLink,teamLead,trainers,"Don't Bring Lunch Box as school is providing lunch for us","School");
let schoolDetails = createSchools(folderName,subFolderName,trackerSheetName,sourceSheetId,destFolderId);  //has return [newFolder.getUrl(),newSubFolder.getUrl(),newSpreadsheet.getUrl(),msg];
msg += schoolDetails[3];
Logger.log(msg);
activeSheet.getRange(msgBox).setValue(msg);
setTrainersinSheet(schoolDetails[2],teamLead,trainers); //spreadsheetUrl
var trainersList ="";
for(let count = 0 ; count < trainers.length ; count++){
  if(count < (trainers.length -1)){
  trainersList+= trainers[count]+", ";}
  else{
    trainersList += trainers[count];
  }
}
Logger.log(trainersList);
let schoolData = [schoolId,schoolName,date,startTime,closeTime,locationLink,schoolDetails[0],schoolDetails[2],teamLead,trainersList];
Logger.log(schoolData);
addData(schoolData);
}



function test() {
  var doc = DocumentApp.openById("1p_DlV24tR899gpIFAkoKXDe3t7flm_CmcB2zIx-i2Zg");
  var body = doc.getBody();
  var sourceTable = DocumentApp.openById('11HI_c6HhibEDtNTZqb0wgW4qAoCyksKt4hqBeFeoaQ8').getBody().getTables()[0];
  var table = body.insertTable(0, sourceTable.copy());
  var row = table.removeRow(3);
   table.appendTableRow(row);
}


function createTable() {
  // Get the active Google Doc
  var doc = DocumentApp.openById("11HI_c6HhibEDtNTZqb0wgW4qAoCyksKt4hqBeFeoaQ8").getBody();

  // Define the table headings
  var headers = ['S. No',
'Date',
'Expense Description',
'Amount',
'Remarks'
];

  // Define the table rows
  var rows = [
    ['1', 'Cab Charges for Team Ria Morning Office to Bandimet School', 'Row 1, Cell 3','',''],
    ['2', 'Cab Charges for Team Ria Evening Bandimet School to office', 'Row 2, Cell 3','',''],
    ['3', 'Row 3, Cell 2', 'Row 3, Cell 3','','']
  ];

var tables = doc.getTables()[0];
var tRow = tables.getRow(0);
 tables.insertTableRow(2,tables);

Logger.log(tRow.getText());
Logger.log(tables);
Logger.log(tables.getRow(0));

  // Create the table
  // var table = doc.appendTable(rows);

  // // Set the table headings
  // var headerRow = table.getRow(0);
  // for (var i = 0; i < headers.length; i++) {
  //   headerRow.getCell(i).setText(headers[i]);
  // }
  // table.setColumnWidth(0,0.45);
  // //table.autoResize(AutoResizeMethod.CONTENTS);
}

