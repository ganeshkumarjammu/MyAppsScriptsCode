

function copyTemplateDoc() {
  //setting up reimbus template docs and creating new doc and moving it to destination;
  let getDocSheet = SpreadsheetApp.openById("1IFi0nkBTzG98Op_DI9BJuMUk3BIlPDkWw6d3SDgnJ7k").getSheetByName("Get Doc");
  var reimbusDocUrl = getDocSheet.getRange("B12").getValue();
  //var reimbusDocUrl = "https://docs.google.com/document/d/1zeZP2FepGn7ZDVW0pTX9BQEHgeSrANyLCP-fe1RpkSU/edit";
  //var dstFolderUrl = "https://drive.google.com/drive/folders/1HTY82ojyno5V6e4KYPTA7K_uO_7btYzi?usp=share_link"; //RiA Expenses folder
  var dstFolderUrl = getDocSheet.getRange("B14").getValue(); 
  //var copyFileName = getDocSheet.getRange("B10").getValue();  ///new doc name
  let copyFileName =  SpreadsheetApp.openById("1IFi0nkBTzG98Op_DI9BJuMUk3BIlPDkWw6d3SDgnJ7k").getSheetByName("Get Doc").getRange("B10").getValue();  ///new doc name
  Logger.log(copyFileName);
  var templateFile = DriveApp.getFileById(getID(reimbusDocUrl)); // replace TEMPLATE_DOC_ID with the ID of your template document
  var folder = DriveApp.getFolderById(getID(dstFolderUrl)); // replace FOLDER_ID with the ID of your destination folder
  var newFile = templateFile.makeCopy(copyFileName, folder); // replace COPY_NAME with the desired name for the copied document
  newFile.moveTo(folder);
  Logger.log(" Doc:"+ copyFileName+ " \n and URL:" + newFile.getUrl());
  getDocSheet.getRange("B20").setValue(newFile.getUrl());
  return newFile.getUrl();
  
}

function getData(){
  //getting bills from sheets
  var getDocSheet = SpreadsheetApp.openById("1IFi0nkBTzG98Op_DI9BJuMUk3BIlPDkWw6d3SDgnJ7k").getSheetByName("Get Doc");
  var name =  getDocSheet.getRange("B4").getValue();
  var date = getDocSheet.getRange("B6").getValue();
   var personCategory = getDocSheet.getRange("B8").getValue();
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
  Logger.log("Total Price"+totalPrice);
  Logger.log(currentBills);
  Logger.log(docBills);
  Logger.log("editDoc\n");

  if(billCount < 20){
  Logger.log("Bills Count:"+billCount);
  editDoc(name,date,docBills,billCount,totalPrice,currentBills);
  }
  else {
     Logger.log("Error greater than 20");
    return "failed";
  }

}


function editDoc(name , date , bills,billCount,totPrice,currentBills){
  let newDocUrl = copyTemplateDoc() ;
  Logger.log("newDocUrl :\n"+newDocUrl);
  let id = getID(newDocUrl);
  //let imageUrl ="https://drive.google.com/open?id=1ri1T4JV7QmEgg3JWZ8Rj3GgO1sas8HoR";
  
  replaceTextInDoc(id,name,date);
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
  Logger.log("NewDocUrl :\n"+newDocUrl);
}

function replaceTextInDoc(docId,name,date) {
  var doc = DocumentApp.openById(docId); // Replace 'DOCUMENT_ID' with the ID of your document
  var searchText1 = 'Category Expense'; // Replace 'TEXT_TO_BE_REPLACED' with the text you want to replace
  var replacementText = 'JPNCE Lab Expense'; // Replace 'REPLACEMENT_TEXT' with the text you want to replace the old text with
  var searchText2 ="YourName "; //Name in temp Doc
  var searchText3 ="  -  -2023"; //date in temp Doc
  date = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy");
  var body = doc.getBody();
  body.replaceText(searchText1, replacementText);
  body.replaceText(searchText2, name);
  body.replaceText(searchText3, date);
  Logger.log("Temp Text Replaced with current name , date");
}


// function appendImagesToDoc() {
//   // ID of the Google Doc
//   const docId = '1GUIvS7Ydyc4qKGpgMEqr61x7KiCP8JjQtAqJu6akJMk';
  
//   // Array of file IDs of images in Google Drive
//   const imageIds = ["https://drive.google.com/open?id=1sX4t0b5snNpliBaa4rw3XZqZrmsottQG","https://drive.google.com/open?id=1XbFeSzsGx3Q8eSVzYHn_w_vfrPb-uFrD ","https://drive.google.com/open?id=1ri1T4JV7QmEgg3JWZ8Rj3GgO1sas8HoR "];
  
//   // Open the Google Doc
//   const doc = DocumentApp.openById(docId);
  
//   // Get the body of the Google Doc
//   const body = doc.getBody();
  
//   // Loop through the image IDs and add each image to the Google Doc
//   imageIds.forEach(imageId => {
//     // Get the file object for the image
//     const imageFile = DriveApp.getFileById(getID(imageId));
    
//     // Get the URL of the image in Google Drive
//     const imageUrl = imageFile.getThumbnailLink(0, 1, { 'height': 400, 'width': 400 });
    
//     // Create a new paragraph in the Google Doc
//     const imageParagraph = body.appendParagraph('');
    
//     // Add the image to the new paragraph
//     const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
//     const image = imageParagraph.addPositionedImage(imageBlob);
    
//     // Set the image size to medium
//     image.setWidth(400);
//     image.setHeight(400);
//   });
// }

// function getThumbnailLink(fileId) {
//   var imageFile = DriveApp.getFileById(fileId);
//   var mimeType = imageFile.getMimeType();
  
//   if (mimeType.indexOf('image/') === 0) {
//     var thumbnailUrl = imageFile.getThumbnailLink(200, 200);
//     return thumbnailUrl;
//   } else {
//     return 'File is not an image.';
//   }
// }

function insertImages(docId,) {
  // Set the ID of the Google Drive folder containing the images
  var folderId = '1WxFpb3LPpv5kGcDXfIZd151mbBlMd0m81RTZnMm63TdLEhavIIO1z_bREYIeOMT3X2ktt2uM';
  var docId = '1GUIvS7Ydyc4qKGpgMEqr61x7KiCP8JjQtAqJu6akJMk';
  // Get the Google Drive folder
  var folder = DriveApp.getFolderById(folderId);
  
  // Get all images in the folder
  var images = folder.getFilesByType('image/jpeg');
  
  // Get the Google Doc where the images will be inserted
  var doc =DocumentApp.openById(docId);
  var lastParagraph = doc.getBody().getChild(doc.getBody().getNumChildren() - 1).asParagraph();
  // Get the width and height of the page in pixels
  var pageWidth = doc.getPageWidth();
  var pageHeight = doc.getPageHeight();
  
  // Loop through the images and append them to the document
  while (images.hasNext()) {
    var imageFile = images.next();
    
    // Get the thumbnail link for the image
    //var thumbnailLink = imageFile.getThumbnailLink(200, 200);
    
    // Create an inline image from the thumbnail link
    var inlineImage = doc.getBody().insertImage(lastParagraph.getNumChildren(), imageFile);
    
    // Get the width and height of the inline image in pixels
    var imageWidth = inlineImage.getWidth();
    var imageHeight = inlineImage.getHeight();
    
    // Calculate the aspect ratio of the image
    var aspectRatio = imageWidth / imageHeight;
    
    // Calculate the maximum width and height of the image that will fit on the page
    var maxWidth = pageWidth - 100;
    var maxHeight = pageHeight - 100;
    
    // Adjust the width and height of the image to fit within the page margins
    if (imageWidth > maxWidth) {
      inlineImage.setWidth(maxWidth);
      inlineImage.setHeight(maxWidth / aspectRatio);
    }
    
    if (inlineImage.getHeight() > maxHeight) {
      inlineImage.setHeight(maxHeight);
      inlineImage.setWidth(maxHeight * aspectRatio);
    }
  }
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

function getFileLink(fileId) {
  var fileId ="18hUjE5kWYaOqXDwWq884LkU3BWY_KxgtvicAsQ3K_e8";
  var file = DriveApp.getFileById(fileId);
  var link = file.getUrl();
  Logger.log(link);
}
