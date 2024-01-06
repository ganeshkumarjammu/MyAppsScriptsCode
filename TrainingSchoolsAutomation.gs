var activeSheet = SpreadsheetApp.getActiveSheet();
var bringLunchBox = "Bring your Lunch Box";
// var dontbringLunchBox ="Don't Bring Lunch Box as school is providing lunch for us";
// var colegeDontBringLunchBox ="Don't Bring Lunch Box as Institute is providing lunch for us" ;

//get school 1 message
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

//get school 2 message
function getSchool2Msg() {
let schoolId = activeSheet.getRange("I3").getValue();
var schoolName = activeSheet.getRange("I5").getValue();
let date = activeSheet.getRange("I7").getValue();
let startTime = getCellTime12HourFormat("I9");
let closeTime = getCellTime12HourFormat("I11");
let reportingTime = getCellTime12HourFormat("I13");
let bringLunch =activeSheet.getRange("I15").getValue();
let locationLink = activeSheet.getRange("I17").getValue();
let folderName = activeSheet.getRange("I19").getValue();
let subFolderName = activeSheet.getRange("I21").getValue();
let trackerSheetName = activeSheet.getRange("I23").getValue();
let sourceSheetId= activeSheet.getRange("I25").getValue();
let destFolderId=activeSheet.getRange("I27").getValue();
let teamLead = activeSheet.getRange("L4").getValue(); 
let trainers = getTrainers("L5:L21");
let msgBox = "K26";
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


function getMessage(schoolName,date,startTime,closeTime,reportingTime,bringLunch,locationLink,teamLead,trainers,dontbringLunchBox,educationLevel){
var msg = "Hello Everyone,\n";
msg += "*Confirmed List for ";
msg +=  schoolName+"*\n\n";
msg += "*Team Lead: "+teamLead+"*\n\n";
for(let count= 0 ; count < trainers.length ; count++){
msg += (count+1)+"."+trainers[count] +"\n";
} 
msg += "\n"+educationLevel+": *"+ schoolName + "*\n";
msg += "Date: "+ date+"\n";
msg += "Start Time: "+startTime+"\n";
msg += "Close Time: "+closeTime+"\n";
msg += "Reporting Time: "+reportingTime+"\n\n";
msg += "Note: ";
if( bringLunch == "Yes" || bringLunch == "yes"){
  msg+= bringLunchBox+"\n\n"; 
}
else{
  msg+= dontbringLunchBox+"\n\n";
}
msg += "Location Link: " 
msg +=  locationLink+"\n\n" ;
msg += "Thank You"+"\n"+"Regards"+"\nSoham Academy\n\n";
Logger.log(msg);
//storing message in message box
return msg;
}


function createSchools(newFolderName,subFolderName,newSheetName,srcSheetId,dstFolder) {
  //getting a src folder id
  var parentFolder = DriveApp.getFolderById(dstFolder);
  
  //creating a new folder in main folder
  var newFolder = parentFolder.createFolder(newFolderName);

  //getting src sheet details
  var originalSpreadsheet = SpreadsheetApp.openById(srcSheetId);
  var file = DriveApp.getFileById(originalSpreadsheet.getId());

  //copy sheet and moving it to school folder
  var newSpreadsheet = file.makeCopy(newSheetName,newFolder);

  //creating photos folder inside the school folder
  var newFolderId = DriveApp.getFolderById(newFolder.getId())
  var newSubFolder= newFolderId.createFolder(subFolderName);
  //var newSubFolderId = newSubFolder.getId();

  //setting access to everyone
  //var folder = DriveApp.getFolderById("Folder ID");
  //var sheet = SpreadsheetApp.openById("Sheet ID");
  newSubFolder.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  newSpreadsheet.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  let msg ="School ID : "+newFolderName+"\n\nPhotos and Videos Folder: "+ newSubFolder.getUrl()+"\n\nTracker Sheet Link: "+ newSpreadsheet.getUrl()+"\n";
  Logger.log(msg);
  return [newFolder.getUrl(),newSubFolder.getUrl(),newSpreadsheet.getUrl(),msg];
}

function setTrainersinSheet(sheetUrl,teamLead,trainers){
 let ss  = SpreadsheetApp.openByUrl(sheetUrl);
 //let teamLead = "Ghero";
 //let trainers = ["ganesh","kumar","how","hello","are","you"];
 let sheet = ss.getSheetByName("Trainers & Club Members");
 sheet.getRange(4, 1).setValue(1);
 sheet.getRange(4,2).setValue(teamLead);
 for(let count = 3 ; count <= (trainers.length+1) ; count++ )// Add the serial number and student name to the sheet
  {
    sheet.getRange(count+2, 1).setValue(count-1);
    sheet.getRange(count+2, 2).setValue(trainers[count-3]);
  }
}

function addData(data) {
  let sheet = SpreadsheetApp.openById("1wpMeV_gBYR5XvFVF3XcFP2Ool_Q2tOj0_Ze80Phj3as").getSheetByName("Schools Data");
  // Get the last row in the sheet
  let lastRow = sheet.getLastRow();
  Logger.log(lastRow);
  sheet.getRange(lastRow+1,1).setValue(lastRow-1);
  // Get the range to write the data to 
  sheet.getRange(lastRow+1, 2, 1, data.length).setValues([data]);
  // Write the data to the range
}


function getCellTime12HourFormat(range) { 
  let cell = activeSheet.getRange(range);
  // Get the value of the cell
  let value = cell.getValue();
  // Get the time of the cell value
  let time = value.getTime();
  // Get the date of the cell value
  let date = new Date(time);
  // Get the hours of the time
  let hours = date.getHours();
  // Get the minutes of the time
  let minutes = date.getMinutes();
  // Get the am/pm designation
  let ampm = hours >= 12 ? 'pm' : 'am';
  // Convert the hours to 12-hour format
  hours = hours % 12;
  hours = hours ? hours : 12;
  // Pad the minutes with a leading zero if necessary
  minutes = minutes < 10 ? '0' + minutes : minutes;
  // Concatenate the hours, minutes, and am/pm designation
  let timeString = hours + ':' + minutes + ' ' + ampm;
  // Log the time in 12-hour format
  Logger.log(timeString);
  return timeString ;
}

function getTrainers(givenRange) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getRange(givenRange);
  // Get the values in the specified range
  let values = range.getValues();
  let trainers =[];
  // Initialize the count of non-empty cells
  let count = 0;
  // Loop through the values in the range
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      // If the value is not empty, increment the count
      if (values[i][j]) {
        count++;
        trainers.push(values[i][j]);
      }
    }
  } 
  // Log the count of non-empty cells
  Logger.log("Number of non-empty cells: " + count);
  return trainers ;
}



//let rDate ="B7";
// function getDayName(rDate) {
//   let sheet = SpreadsheetApp.getActiveSheet();
//   let cell = sheet.getRange("B7");
  
//   // Get the value in the cell as a date
//   let date = new Date(cell.getValue());
  
//   // Get the day of the week as a number (0 = Sunday, 1 = Monday, etc.)
//   let dayOfWeek = date.getDay();
  
//   // Create an array of day names
//   let dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  
//   // Get the name of the day
//   let dayName = dayNames[dayOfWeek];
  
//   // Log the name of the day
//   Logger.log("Day: " + dayName);
// }

// function onEdit(e) {
//   // Set a comment on the edited cell to indicate when it was changed.
//   const range = e.range;
//   var value = range.value();
//   Logger.log(value);
//   if(value == "Yes" || value == "yes"){
//     getSchool1Msg();
//   }
// }


// //get college 1 message
// function getCollege1Msg() {
//  let schoolname ="B3";
// let folderName ="B5";
// let date ="B7";
// let startTime = "B9";
// let closeTime ="B11";
// let reportingTime = "B13";
// let bringLunch ="B15";
// let locationLink ="B17";
// let teamLead = "E4";
// let trainers ="E5:E21";
// let msgBox = "D24";
//  getMessage(schoolname,date,startTime,closeTime,reportingTime,bringLunch,locationLink,teamLead,trainers,msgBox,"Don't Bring Lunch Box as Institute is providing lunch for us","Institute");
// }

// //get College 2 message
// function getCollege2Msg(){
// let schoolname ="I3";
// let folderName ="I5";
// let date ="I7";
// let startTime = "I9";
// let closeTime ="I11";
// let reportingTime = "I13";
// let bringLunch ="I15";
// let locationLink ="I17";
// let teamLead = "I4";
// let trainers ="L5:L21";
// let msgBox = "K24";
//   getMessage(schoolname,date,startTime,closeTime,reportingTime,bringLunch,locationLink,teamLead,trainers,msgBox,"Don't Bring Lunch Box as Institute is providing lunch for us","Institute");
// }




///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



// function onEdit(e) {
//   // Set a comment on the edited cell to indicate when it was changed.
//   const range = e.range;
//   var value = range.value();
//   if(value == "Yes" || value == "yes"){

//   }
//   range.setNote('Last modified: ' + new Date());
// }

function getfullnames() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("set School Trainers"); // change to the name of your sheet
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings"); // change to the name of your data sheet
  var names = sheet.getRange("E5:E19"  ).getValues().flat(); // get the list of names in column A, excluding the header row  //sheet.getLastRow()
  var data = dataSheet.getRange("B3:C19"  ).getValues(); // get the name and fullname data from the data sheet  //dataSheet.getLastRow()
  var fullnameMap = new Map(data.map(row => [row[0], row[1]])); // create a map of name -> fullname from the data
  
  // create a 2D array of names and fullnames to write back to the sheet
  //var output = [["Name", "FullName"]]; // add headers to the output
  var output = []; // add headers to the output
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var fullname = fullnameMap.get(name);
    if(name !="") output.push([name, fullname || ""]); // add the name and fullname (or a blank cell if the fullname is not found) to the output
  }
  Logger.log(output);
  for(let count = 0 ; count < output.length;count++){
    Logger.log(output[count][1]);
  }
  //sheet.getRange(1, 3, output.length, output[0].length).setValues(output); // write the output back to the sheet starting at cell C1
}