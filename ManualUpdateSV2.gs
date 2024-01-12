

//Adds elements in the column wise and also adds serial number
function createInvisibleTableFromText(text) {
  var fontSize = 11;
  var doc = DocumentApp.openById(destDocID)
  var body = doc.getBody();
  //var text = "Arduino UNO R3 - 1, Ultrasonic Sensor - 1,Breadboard - 1, Buzzer - 1,LEDs - 6,Several Jumper Wires,Resistors - 6";
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
      table.getRow(i).getCell(j).getChild(0).asParagraph().setFontFamily("Calibri").setFontSize(12).setItalic(false);
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
  var scaledWidth = body.getPageWidth() * 1; // Adjust the percentage as needed

  // Scale the image width to fit 60% or less of the page width
  if (width > scaledWidth) {
    var height = inlineImage.getHeight() * (scaledWidth / width);
    inlineImage.setWidth(scaledWidth).setHeight(height);
  }
  //var leftMargin = ((body.getPageWidth() - inlineImage.getWidth()) / 2); // if incase image went out of the page uncomment this.
}

function appendDataToDoc() {
  var sourceSheetID = '1OLaYvIZct11myC4X2NxbwTdXGRGVO9hIWUS9riW4mlo';
var sourceSheetName = 'Sheet1';
var destDocID = "166uMELxwBMQGx66vsGGuqQ23HCeEB95iFyDPaAdvzbE";
var normalFont ="Calibri";
var normalFontSize ="12";
  var sheet = SpreadsheetApp.openById(sourceSheetID).getSheetByName(sourceSheetName); 
  // Replace 'Sheet1' with your sheet name
  var data = sheet.getDataRange().getValues();
  
 // var data = sheet.getRange("C4:I4").getValues();
  Logger.log(data);
  for (var i = 2; i < 3; i++) {
    var project = data[i][2]; // Project Name
    var aim = data[i][3]; // Aim
    var materials = data[i][4]; // Materials Required
    var circuitDiagram = data[i][5]; // Circuit Diagram Link
    var outputPhoto = data[i][7]; // Output Photo Link
    var ouputText = data[i][6]; // Result Text
    var doc = DocumentApp.openById(destDocID);
    var body = doc.getBody();
    var newPage = body.appendPageBreak();
    var projectHeading = body.appendParagraph(project);
    projectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    projectHeading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
   // body.appendParagraph(''); // Add an empty line
    Logger.log("project:"+project);
    body.appendParagraph('Goal:').setHeading(DocumentApp.ParagraphHeading.HEADING3).setItalic(true);
    body.appendParagraph(aim).setFontFamily("Calibri").setFontSize(12).setItalic(false);
    Logger.log("Goal:"+aim);
    //body.appendParagraph('');

    body.appendParagraph('Material Required:').setHeading(DocumentApp.ParagraphHeading.HEADING3).setItalic(true);
    //body.appendParagraph(materials);
    if (materials != "" ) 
    createInvisibleTableFromText(materials);
    else {
    body.appendParagraph("No Materials");
    //body.appendParagraph('');
    }
    //body.appendParagraph('');
    Logger.log("materials:"+materials);
    
    body.appendParagraph('Circuit Diagram:').setHeading(DocumentApp.ParagraphHeading.HEADING3).setItalic(true);
    //body.appendParagraph('');
    var image1Pos = readParagraphsAndCount();
     body.appendParagraph('');

    body.appendParagraph('Output Photo:').setHeading(DocumentApp.ParagraphHeading.HEADING3).setItalic(true);

    //body.appendParagraph('');
    var image2Pos = readParagraphsAndCount();
    // body.appendParagraph('');

    var output = body.appendParagraph('Output:').setHeading(DocumentApp.ParagraphHeading.HEADING3).setItalic(true);
    body.appendParagraph(ouputText).setFontFamily("Calibri").setFontSize(12);
    //body.appendParagraph('');
    //append Image error occur please uncomment this lines
    // doc.saveAndClose();
    // doc = DocumentApp.openById(destDocID);
    // body = doc.getBody();
    // body.appendParagraph('');
    Logger.log("CircuitDiagram"+circuitDiagram);
    Logger.log("outputPhoto:"+outputPhoto);

  if(circuitDiagram ==""){
    body.appendParagraph("No Circuit Diagram");
    //body.appendParagraph('');
  }
  else if (circuitDiagram !=""){
    insertImageSV(circuitDiagram,image1Pos);
  }
  else{
   body.appendParagraph("ERROR");
  }

 if(outputPhoto == "" ){
    body.appendParagraph("No Output Photo");
   // body.appendParagraph('');
  }
  else if (outputPhoto != ""){
      // insertImageSV(outputPhoto,image2Pos);
  }
     doc.saveAndClose();
    Utilities.sleep(1000);
}
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
