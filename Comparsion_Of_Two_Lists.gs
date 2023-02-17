function COMPARE(range1,range2) {
  // Get the data range for the two columns
  // var range1 = SpreadsheetApp.getActiveSheet().getRange("A1:A10"); // First column
  // var range2 = SpreadsheetApp.getActiveSheet().getRange("B1:B10"); // Second column
  // Logger.log(range1.getValues());
  
  // // Get the values from the two columns as arrays
  // var column1Values = range1.getValues().flat(); 
  // var column2Values = range2.getValues().flat();
  //   // Get the values from the two columns as arrays

  var column1Values = range1.flat(); 
  var column2Values = range2.flat();
  // Filter the values that are not in the second column
  // var notInColumn2 = column1Values.filter(function(value) {
  //   return !column2Values.includes(value);
  // });
  
  // // Display the values not in the second column
  // Logger.log("Components not in column B:");
  // Logger.log(notInColumn2);
  // return notInColumn2 ;

  list1Values = range1.flat().map(function (x) { return x.toString().toLowerCase(); });
  list2Values = range2.flat().map(function (x) { return x.toString().toLowerCase(); });
  
  // Get the missing components in list1Values that are not in list2Values
  var missingComponents = list1Values.filter(function(x) { return !list2Values.includes(x); });
  
  // Return the missing components
  return missingComponents;
}


function listComponents() {
  // Get the data range for the two columns
  var range1 = SpreadsheetApp.getActiveSheet().getRange("A1:A10"); // First column
  var range2 = SpreadsheetApp.getActiveSheet().getRange("B1:B10"); // Second column
  
  // Get the values from the two columns as arrays
  var column1Values = range1.getValues().flat(); 
  var column2Values = range2.getValues().flat();
  
  // Filter the values that are not in the second column
  var notInColumn2 = column1Values.filter(function(value) {
    return !column2Values.includes(value);
  });
  
  // Display the values not in the second column
  Logger.log("Components not in column B:");
  Logger.log(notInColumn2);
}
