var sourceSheetID = '1c4kM3FwAuAo3jZOUvWojZ84-wtUIDHr_KAllQxv-rOg';
var sourceSheetName = 'Sheet11';
var destDocID = "113b3xJ9DzxOrC80V5iQl1QY5i-mmkvXjXxKbFgTfEtA";
//Dont forget set the fontsize of the text in the material required table -> fontSize variable is in createInvisibleTableFromText()
function appendDataToDoc() {
    var sheet = SpreadsheetApp.openById(sourceSheetID).getSheetByName(sourceSheetName); 
    // Replace 'Sheet1' with your sheet name
    var data = sheet.getDataRange().getValues();
  
    var doc = DocumentApp.openById(destDocID);
    var body = doc.getBody();
  
    for (var i = 2; i < data.length; i++) {
      var project = data[i][2]; // Project Name
      var aim = data[i][3]; // Aim
      var materials = data[i][4]; // Materials Required
      var circuitDiagram = data[i][5]; // Circuit Diagram Link
      var outputPhoto = data[i][6]; // Output Photo Link
      var resultText = data[i][7]; // Result Text
      
      var newPage = body.appendPageBreak();
      var projectHeading = body.appendParagraph(project);
      projectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      projectHeading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph(''); // Add an empty line
      Logger.log("project:"+project);
      body.appendParagraph('Aim:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph(aim);
      Logger.log("Aim:"+aim);
      body.appendParagraph('');
  
      body.appendParagraph('Materials Required:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      //body.appendParagraph(materials);
      createInvisibleTableFromText(materials);
      body.appendParagraph('');
      //Logger.log("materials:"+materials);
      
      body.appendParagraph('Circuit Diagram:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph('');
      var image1Pos = readParagraphsAndCount();
      body.appendParagraph('');
      body.appendParagraph('Output Photo:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      //body.appendParagraph('');
      var image2Pos = readParagraphsAndCount();
       body.appendParagraph('');
      body.appendParagraph('Result Text:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph(resultText);
      body.appendParagraph('');
      //append Image error occur please uncomment this lines
      // doc.saveAndClose();
      // doc = DocumentApp.openById(destDocID);
      // body = doc.getBody();
      // body.appendParagraph('');
      insertImageSV(circuitDiagram,image1Pos);
      insertImageSV(outputPhoto,image2Pos);
     // body.appendPageBreak();
      doc.saveAndClose();
      Utilities.sleep(1000);
    }
  }

//Adds elements in the column wise and also adds serial number
function createInvisibleTableFromText(text) {
  var fontSize = 11;
  var doc = DocumentApp.openById(destDocID)
  var body = doc.getBody();
  // var text = "Arduino UNO R3 - 1, Ultrasonic Sensor - 1,Breadboard - 1, Buzzer - 1,LEDs - 6,Several Jumper Wires,Resistors - 6"; 
  // even if it total no of elements are odd then also it works.
  var items = text.split(",");
  var numRows = Math.ceil(items.length / 2);
  var columnWidth = (body.getPageWidth() ) / 2;
  var tableData = [];
  for (var i = 0; i < numRows; i++) {
    var rowData = [];
    if (i < items.length) {
      rowData.push((i + 1) + ". " + items[i]);
    }
    if (i + numRows < items.length) {
      rowData.push((i + numRows + 1) + ". " + items[i + numRows]);
    }
    tableData.push(rowData);
  }
  
  Logger.log("Here is the tableData:");
  Logger.log(tableData);
  var table = body.appendTable(tableData);
  table.setColumnWidth(0, columnWidth);
  table.setColumnWidth(1, columnWidth);
  table.setBorderWidth(0);

  for (var i = 0; i < table.getNumRows(); i++) {
    for (var j = 0; j < table.getRow(i).getNumCells(); j++) {
      table.getRow(i).getCell(j).setBackgroundColor(null);
      table.getRow(i).getCell(j).getChild(0).asParagraph().setFontSize(fontSize);
    }
  }
}



function insertImageSV(image,targetParagraphIndex) {
  var doc = DocumentApp.openById(destDocID)
  var body = doc.getBody();

  // Find the target paragraph index after which you want to insert the image 
  //var targetParagraphIndex = readParagraphsAndCount()-3; // Change this index to the desired paragraph location

  // Get the target paragraph in the body
  var targetParagraph = body.getParagraphs()[targetParagraphIndex];

  // Replace with your image file ID or URL
 // var image = 'https://drive.google.com/open?id=1wDZpqeh7eKxlGrZfVevh_BCPI-Cg7Fw4'; 
  var fileID = image.match(/[\w\_\-]{25,}/).toString();
  var blob = DriveApp.getFileById(fileID).getBlob();

  var inlineImage = targetParagraph.appendInlineImage(blob);
  //inlineImage.setLayout(DocumentApp.WrapMode);
  var width = inlineImage.getWidth();

  //ADJUST THE IMAGE SIZE HERE YOU CAN USE '1' ALSO INSTEAD OF 0.8
  var scaledWidth = body.getPageWidth() * 0.8; // Adjust the percentage as needed

  // Scale the image width to fit 60% or less of the page width
  if (width > scaledWidth) {
    var height = inlineImage.getHeight() * (scaledWidth / width);
    inlineImage.setWidth(scaledWidth).setHeight(height);
  }
  //var leftMargin = ((body.getPageWidth() - inlineImage.getWidth()) / 2); // if incase image went out of the page uncomment this.
}


function readParagraphsAndCount() {
  // Open the active document
  var doc = DocumentApp.openById(destDocID)
  // Get the body of the document
  var body = doc.getBody();
  // Get all paragraphs
  var paragraphs = body.getParagraphs();
  // Get the number of paragraphs
  var paragraphCount = paragraphs.length;
  // Read each paragraph and print its text
  for (var i = 0; i < paragraphCount; i++) {
    var paragraphText = paragraphs[i].getText();
   // Do something with the paragraph text (e.g., print it to the console)
   // console.log(paragraphText);
  }
  console.log("Number of paragraphs:",paragraphCount );
  // Return the last paragraph number
  return paragraphCount ;
}

//////
//////
//////
///////
///////
////////////////Below Code are for Learning and reusage top code are successful previous version while botom code have some errors
////////
///////////
///////
////////////
///////
//////////
//////
//////
//////
///////
///////
////////////////  you can delete the below codes to be free of same functions name error 
////////
///////////
///////
////////////
///////
//////////

 //Basic Code for doc creating
 function createAndAppendToDoc() {
  //Create a new Google Doc
 var doc = DocumentApp.create('test appscript1');

  //Get the document body
 var body = doc.getBody();

  //Append information to the document
  body.appendParagraph('This is a new paragraph.');
  body.appendParagraph('This is another paragraph.');
  body.appendPageBreak(); // Add a page break
  // Save the document
}


//Successfull appends the elements in the column wise but not row wise
function createInvisibleTableFromTextSV4() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var text = "Arduino UNO R3 - 1,4. Ultrasonic Sensor - 1,Breadboard - 1,5. Buzzer - 1,LEDs - 6,6. Several Jumper Wires,Ganesh";
  var items = text.split(",");

  var numRows = Math.ceil(items.length / 2);
  var columnWidth = (body.getPageWidth() * 0.8) / 2;

  var tableData = [];
  for (var i = 0; i < numRows; i++) {
    var rowData = [];
    if (i < items.length) {
      rowData.push(items[i]);
    }
    if (i + numRows < items.length) {
      rowData.push(items[i + numRows]);
    }
    tableData.push(rowData);
  }

  var table = body.appendTable(tableData);

  table.setColumnWidth(0, columnWidth);
  table.setColumnWidth(1, columnWidth);

  table.setBorderWidth(0);

  for (var i = 0; i < table.getNumRows(); i++) {
    for (var j = 0; j < table.getRow(i).getNumCells(); j++) {
      table.getRow(i).getCell(j).setBackgroundColor(null);
      table.getRow(i).getCell(j).getChild(0).asParagraph().setFontSize(12);
    }
  }
}

//Successfully running adds items row wise not column wise
function createInvisibleTableFromTextSV3(text) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

 // var text = "Arduino UNO R3 - 1,4. Ultrasonic Sensor - 1,Breadboard - 1,5. Buzzer - 1,LEDs - 6,6. Several Jumper Wires";

  var items = text.split(",");

  var numRows = items.length % 2 === 0 ? items.length / 2 : Math.floor(items.length / 2) + 1;

  var columnWidth = (body.getPageWidth() * 0.8) / 2;

  var tableData = [];
  var index = 0;
  for (var i = 0; i < numRows; i++) {
    var rowData = [];
    for (var j = 0; j < 2; j++) {
      if (index < items.length) {
        var serialNumber = index + 1; // Generating serial numbers starting from 1
        var itemWithSerial = serialNumber + ". " + items[index];
        rowData.push(itemWithSerial);
        index++;
      } else {
        rowData.push(""); 
      }
    }
    tableData.push(rowData);
  }

  var table = body.appendTable(tableData);

  table.setColumnWidth(0, columnWidth);
  table.setColumnWidth(0, columnWidth);
  table.setBorderWidth(0);     

  for (var i = 0; i < table.getNumRows(); i++) {
    for (var j = 0; j < table.getRow(i).getNumCells(); j++) {
      table.getRow(i).getCell(j).setBackgroundColor(null);
      table.getRow(i).getCell(j).getChild(0).asParagraph().setFontSize(12); 
    }
  }
}


//Successful this version does have adding components number in the table
function createInvisibleTableFromTextSV2(text) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var fontsize = 11;
  var text = "Arduino UNO R3 - 1,4. Ultrasonic Sensor - 1,Breadboard - 1,5. Buzzer - 1,LEDs - 6,6. Several Jumper Wires";
  var items = text.split(",");
  var numRows = items.length % 2 === 0 ? items.length / 2 : Math.floor(items.length / 2) + 1;
  var columnWidth = (body.getPageWidth() * 0.8) / 2;
  var tableData = [];
  var index = 0;
  for (var i = 0; i < numRows; i++) {
    var rowData = [];
    for (var j = 0; j < 2; j++) {
      if (index < items.length) {
        rowData.push(items[index]);
        index++;
      } else {
        rowData.push(""); 
      }
    }
    tableData.push(rowData);
  }
  var table = body.appendTable(tableData);
  table.setColumnWidth(0, columnWidth);
  table.setColumnWidth(0, columnWidth);
  table.setBorderWidth(0); 
  for (var i = 0; i < table.getNumRows(); i++) {
    for (var j = 0; j < table.getRow(i).getNumCells(); j++) {
      table.getRow(i).getCell(j).setBackgroundColor(null);
      table.getRow(i).getCell(j).getChild(0).asParagraph().setFontSize(fontsize); 
    }
  }
}

//table creating basic code
function createInvisibleTableFromTextSV1(text) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Create a 2D array to represent the table data (optional, but recommended for clarity)
  var tableData = [
    ["Row 1, Column 1", "Row 1, Column 2"],
    ["Row 2, Column 1", "Row 2, Column 2"],
    ["Row 3, Column 1", "Row 3, Column 2"]
  ];

  // Create the table with 2 columns
  var table = body.appendTable(tableData);

  // Make the table invisible
  table.setBorderWidth(0).setBackgroundColor("#ffffff");  // Remove all borders
  //table.setWidth(0);       // Set width to 0 to prevent any visual spacing

  // Set font size to 0 for complete invisibility (optional)
  for (var i = 0; i < table.getNumRows(); i++) {
    for (var j = 0; j < table.getRow(i).getNumCells(); j++) {
      table.getRow(i).getCell(j).getChild(0).asParagraph().setFontSize(0);
    }
  }
}


function insertImageSV1(image) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Find the target paragraph index after which you want to insert the image
  var targetParagraphIndex = 15; // Change this index to the desired paragraph location

  // Get the target paragraph in the body
  var targetParagraph = body.getParagraphs()[targetParagraphIndex];
  
  // Replace with your image file ID or URL
  var image = 'https://drive.google.com/open?id=1wDZpqeh7eKxlGrZfVevh_BCPI-Cg7Fw4'; 
  var fileID = image.match(/[\w\_\-]{25,}/).toString();
  var blob = DriveApp.getFileById(fileID).getBlob();

  var inlineImage = targetParagraph.appendInlineImage(blob);
  var width = inlineImage.getWidth();
  var scaledWidth = body.getPageWidth() * 0.8; // Adjust the percentage as needed

  // Scale the image width to fit 60% or less of the page width
  if (width > scaledWidth) {
    var height = inlineImage.getHeight() * (scaledWidth / width);
    inlineImage.setWidth(scaledWidth).setHeight(height);
  }

  var leftMargin = (body.getPageWidth() - inlineImage.getWidth()) / 2;

  // Set the left margin of the inline image to center it horizontally
  //inlineImage.getParent().asText().setLeftIndent(leftMargin);
}









function insertImage1(targetParagraph) {
  var body = targetParagraph.getParent();
  
  // Replace with your image file ID or URL
  var image = 'https://drive.google.com/open?id=1wDZpqeh7eKxlGrZfVevh_BCPI-Cg7Fw4'; 
  var fileID = image.match(/[\w\_\-]{25,}/).toString();
  var blob = DriveApp.getFileById(fileID).getBlob();

  var inlineImage = body.appendImage(blob);
  var width = inlineImage.getWidth();
  var scaledWidth = body.getPageWidth() * 0.6; // Adjust the percentage as needed

  // Scale the image width to fit 60% or less of the page width
  if (width > scaledWidth) {
    var height = inlineImage.getHeight() * (scaledWidth / width);
    inlineImage.setWidth(scaledWidth).setHeight(height);
  }

  var leftMargin = (body.getPageWidth() - inlineImage.getWidth()) / 2;

  // Set the left margin of the inline image to center it horizontally
  inlineImage.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  targetParagraph = body.appendParagraph(''); // Insert an empty paragraph for spacing
}
function insertImage2() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  let imagelink =  'https://drive.google.com/open?id=1wDZpqeh7eKxlGrZfVevh_BCPI-Cg7Fw4'; 
  var image = body.appendImage(imagelink);
  var width = 0.6 * doc.getPageWidth();  // Set width to 60% of page width
  var height = image.getHeight() * width / image.getWidth();
  image.setWidth(width);
  image.setHeight(height);
}


function appendDataToDoc2() {
  var sheet = SpreadsheetApp.openById('1c4kM3FwAuAo3jZOUvWojZ84-wtUIDHr_KAllQxv-rOg').getSheetByName('Sheet11');
  var data = sheet.getDataRange().getValues();

  var docId = '113b3xJ9DzxOrC80V5iQl1QY5i-mmkvXjXxKbFgTfEtA';
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  for (var i = 2; i < data.length; i++) {
    var project = data[i][2];
    var aim = data[i][3];
    var materials = data[i][4];
    var circuitDiagram = data[i][5];
    var outputPhoto = data[i][6];
    var resultText = data[i][7];

    var newPage = body.appendPageBreak();
    var projectHeading = body.appendParagraph(project);
    projectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    projectHeading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');

    body.appendParagraph('Aim:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(aim);
    body.appendParagraph('');

    body.appendParagraph('Materials Required:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(materials);
    body.appendParagraph('');

    body.appendParagraph('Circuit Diagram:').setHeading(DocumentApp.ParagraphHeading.HEADING3);

    try {
      var imageBlob1 = UrlFetchApp.fetch(circuitDiagram).getBlob();
      var image1 = body.appendImage(imageBlob1);
      var width1 = image1.getWidth();
      var height1 = image1.getHeight();

      // Calculate new dimensions for the image
      var newWidth1 = width1 * 0.5; // Adjust the percentage as needed
      var newHeight1 = (newWidth1 / width1) * height1;

      // Set the new size of the image
      image1.setWidth(newWidth1).setHeight(newHeight1);
    } catch (error) {
      Logger.log('Error fetching Circuit Diagram for row ' + (i + 1) + ':', error);
      body.appendParagraph('Error fetching Circuit Diagram');
    }

    body.appendParagraph('');

    // Add output photo with resizing
    body.appendParagraph('Output Photo:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    try {
      var imageBlob = UrlFetchApp.fetch(outputPhoto).getBlob();
      var image = body.appendImage(imageBlob);
      var width = image.getWidth();
      var height = image.getHeight();

      // Calculate new dimensions for the image
      var newWidth = width * 0.5; // Adjust the percentage as needed
      var newHeight = (newWidth / width) * height;

      // Set the new size of the image
      image.setWidth(newWidth).setHeight(newHeight);
    } catch (error) {
      Logger.log('Error fetching Output Photo for row ' + (i + 1) + ':', error);
      body.appendParagraph('Error fetching Output Photo');
    }

    body.appendParagraph('');
    body.appendParagraph('Result Text:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(resultText);
    body.appendParagraph('');

    body.appendPageBreak();
    Utilities.sleep(1000);
  }
}

function appendDataToDoc3() {
  var sheet = SpreadsheetApp.openById('1c4kM3FwAuAo3jZOUvWojZ84-wtUIDHr_KAllQxv-rOg').getSheetByName('Sheet11');
  var data = sheet.getDataRange().getValues();

  var docId = '113b3xJ9DzxOrC80V5iQl1QY5i-mmkvXjXxKbFgTfEtA';
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  for (var i = 2; i < data.length; i++) {
    var project = data[i][2];
    var aim = data[i][3];
    var materials = data[i][4];
    var circuitDiagram = data[i][5];
    var outputPhoto = data[i][6];
    var resultText = data[i][7];

    var newPage = body.appendPageBreak();
    var projectHeading = body.appendParagraph(project);
    projectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    projectHeading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');

    body.appendParagraph('Aim:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(aim);
    body.appendParagraph('');

    body.appendParagraph('Materials Required:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(materials);
    body.appendParagraph('');

    body.appendParagraph('Circuit Diagram:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    try {
      var imageBlob1 = UrlFetchApp.fetch(circuitDiagram).getBlob();
      var image1 = body.appendImage(imageBlob1);
      var width1 = image1.getWidth();
      var height1 = image1.getHeight();

      // Calculate new dimensions for the image
      var newWidth1 = width1 * 0.5; // Adjust the percentage as needed
      var newHeight1 = (newWidth1 / width1) * height1;

      // Set the new size of the image
      image1.setWidth(newWidth1).setHeight(newHeight1);
    } catch (error) {
      Logger.log('Error fetching Circuit Diagram:', error);
      body.appendParagraph('Error fetching Circuit Diagram');
    }

    body.appendParagraph('');

    // Add output photo with resizing
    body.appendParagraph('Output Photo:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    try {
      var imageBlob = UrlFetchApp.fetch(outputPhoto).getBlob();
      var image = body.appendImage(imageBlob);
      var width = image.getWidth();
      var height = image.getHeight();

      // Calculate new dimensions for the image
      var newWidth = width * 0.5; // Adjust the percentage as needed
      var newHeight = (newWidth / width) * height;

      // Set the new size of the image
      image.setWidth(newWidth).setHeight(newHeight);
    } catch (error) {
      Logger.log('Error fetching Output Photo:', error);
      body.appendParagraph('Error fetching Output Photo');
    }

    body.appendParagraph('');
    body.appendParagraph('Result Text:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(resultText);
    body.appendParagraph('');

    body.appendPageBreak();
    Utilities.sleep(1000);
  }
}


function appendDataToDoc1() {
   var sheet = SpreadsheetApp.openById('1c4kM3FwAuAo3jZOUvWojZ84-wtUIDHr_KAllQxv-rOg').getSheetByName('Sheet11'); // Replace 'Sheet1' with your sheet name
  var data = sheet.getDataRange().getValues();

var docId = '113b3xJ9DzxOrC80V5iQl1QY5i-mmkvXjXxKbFgTfEtA'; // Replace 'YOUR_DOCUMENT_ID' with your Google Doc ID
 var doc = DocumentApp.openById(docId);

  var body = doc.getBody();

  for (var i = 2; i < data.length; i++) {
    var project = data[i][2]; // Project Name
    var aim = data[i][3]; // Aim
    var materials = data[i][4]; // Materials Required
    var circuitDiagram = data[i][5]; // Circuit Diagram Link
    var outputPhoto = data[i][6]; // Output Photo Link
    var resultText = data[i][7]; // Result Text
    var newPage = body.appendPageBreak();
    var projectHeading = body.appendParagraph(project);
    projectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    projectHeading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(''); // Add an empty line
    Logger.log("project:"+project);
    body.appendParagraph('Aim:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(aim);
    Logger.log("Aim:"+aim);
    body.appendParagraph('');

    body.appendParagraph('Materials Required:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(materials);
    body.appendParagraph('');
    Logger.log("materials:"+materials);
    
    body.appendParagraph('Circuit Diagram:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    //insertImage();
     var imageBlob1 = UrlFetchApp.fetch(circuitDiagram).getBlob();
    let image1 = body.appendImage(imageBlob1);
    var width1 = image1.getWidth();
    var height1 = image1.getHeight();

    // Calculate new dimensions for the image
    var newWidth1 = width1 * 0.5; // Adjust the percentage as needed
    var newHeight1 = (newWidth1 / width1) * height1;

    // Set the new size of the image
    image1.setWidth(newWidth1).setHeight(newHeight1);
    body.appendParagraph('');

    // Add output photo with resizing
    body.appendParagraph('Output Photo:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    var imageBlob = UrlFetchApp.fetch(outputPhoto).getBlob();
    var image = body.appendImage(imageBlob);
    var width = image.getWidth();
    var height = image.getHeight();

    // Calculate new dimensions for the image
    var newWidth = width * 0.5; // Adjust the percentage as needed
    var newHeight = (newWidth / width) * height;

    // Set the new size of the image
    image.setWidth(newWidth).setHeight(newHeight);
    body.appendParagraph('');
    Logger.log(result);
    body.appendParagraph('Result Text:').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(resultText);
    body.appendParagraph('');

    body.appendPageBreak();
   doc.saveAndClose();
    Utilities.sleep(1000);
  }
}


function getPresentParagraphIndexAfterAppend() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Append a new paragraph
  var newParagraph = body.appendParagraph('New content');
  
  // Get the index of the new paragraph after appending


  Logger.log('Present Paragraph Index: ' + presentParagraphIndex);
  return presentParagraphIndex;
}


