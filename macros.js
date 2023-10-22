/** @OnlyCurrentDoc */

  var spreadsheet;
  var parentFolder;
  var contactData;
  var surveyData;
  var surveyFolder;
  var contactFolder;



function processSurvey(){
  setup();
  setupSheet("surveyDataTemp");
  setupFiles();
  processXLSFilesInZip(surveyFolder);
  surveyData.getRange(2,1,surveyData.getLastRow()-1,surveyData.getLastColumn()).sort(3);
  surveyData.getRange(2,1,surveyData.getLastRow()-1,surveyData.getLastColumn()).removeDuplicates([1,2,3,4,5,6,7,8]);

  

};

function processContact(){
  setup();
  setupSheet("contactDataTemp");
  setupFiles();
  processXLSFilesInZip(contactFolder);
  contactData.getRange(2,1,contactData.getLastRow()-1,contactData.getLastColumn()).sort(3);
  contactData.getRange(2,1,contactData.getLastRow()-1,contactData.getLastColumn()).removeDuplicates([1,2,3,4,5,6,7,8,9,10]);
}




function setup() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  parentFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();


  surveyHeader = [["Voter File VANID","Contact Name","Date Canvassed","Survey Response","Survey Question Name","Survey Question Type","Canvassed By","Contact Type","Instance of VANID","Filename","Date/Time Added"]];

  contactHeader = [["Voter File VANID","Contact Name","Date Canvassed","Result","Contact Type","Canvassed By","Created By","Date Created","Input Type Name","MiniVAN Campaign","Instance of VANID","Filename","Date/Time Added"]]
}

function setupSheet(sheetName){
    if(sheetName == 'surveyDataTemp'){
      surveyData = spreadsheet.getSheetByName(sheetName);
      if (surveyData != null) {
        surveyData.getDataRange().clear();
      }else{
        surveyData = spreadsheet.insertSheet(sheetName);
      }
      surveyData.getRange(1,1,1,surveyHeader[0].length).setValues(surveyHeader)      
  } else{
    contactData = spreadsheet.getSheetByName(sheetName);
    if (contactData != null) {
      contactData.getDataRange().clear();
    }else{
      contactData = spreadsheet.insertSheet(sheetName);
    }
    contactData.getRange(1,1,1,contactHeader[0].length).setValues(contactHeader)       
  }
}



function setupFiles() {
  // Get the active spreadsheet and its parent folder

  // Define folder names
  var folderNames = ["Survey Response Reports", "Contact History Reports"];

  // Create the folders that don't exist
  folderNames.forEach(function (folderName) {
    var folder = parentFolder.getFoldersByName(folderName);
    if (!folder.hasNext()) {
      parentFolder.createFolder(folderName);
      Logger.log(folderName + " folder created.");
    } else {
      Logger.log(folderName + " folder already exists.");
    }
  });
  surveyFolder = parentFolder.getFoldersByName("Survey Response Reports").next();
  contactFolder = parentFolder.getFoldersByName("Contact History Reports").next();

};


function processXLSFilesInZip(folder) {
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (fileName.endsWith(".zip")) {
      var zipBlob = file.getBlob();
    // Unzip the ZIP file
      var zip = Utilities.unzip(zipBlob);

      // Check for a single XLS file in the ZIP
      var xlsFiles = zip.filter(function (blob) {
        return blob.getName().endsWith(".xls");
      });

      if (xlsFiles.length == 1) {
        Logger.log("Processing: "+ fileName)
        processXLSData(folder,xlsFiles[0])
      } else if (xlsFiles.length == 0) {
        Logger.log("No XLS file found in the ZIP file: " + fileName);
        // Handle the case where no XLS file is found
      } else {
        Logger.log("Multiple XLS files found in the ZIP file: " + fileName);
        // Handle the case where multiple XLS files are found
      }
    }
  }
};


function processXLSData(folder,xlsBlob) {
  var newFile = {
    title : xlsBlob.getName(),
    parents: [{id: folder.getId()}] //  Added
  };
  var file = Drive.Files.insert(newFile, xlsBlob, {
    convert: true
  });
  var otherSpreadsheet = SpreadsheetApp.openById(file.id);
  var otherSheet = otherSpreadsheet.getSheets()[0]; // Change to the desired sheet name
  var importedData = otherSheet.getRange(2,1,otherSheet.getLastRow()-1,otherSheet.getLastColumn()).getValues();


  Drive.Files.remove(file.id); // Added // If this line is run, the original XLSX file is removed. So please be careful this.

  // Append the imported data to the existing or new sheet



  if(folder.getName()=='Survey Response Reports'){
    var startRow = surveyData.getLastRow() + 1
    var filename = Array(importedData.length).fill([xlsBlob.getName()]);
    var dateTimeCol = Array(importedData.length).fill([Utilities.formatDate(new Date(), "America/New_York", "MM/dd/yyyy h:mm a")]);

    surveyData.getRange(startRow, 1, importedData.length, importedData[0].length).setValues(importedData);


    surveyData.getRange(2,surveyData.getLastColumn()-2,importedData.length,1).setFormula("=COUNTIF(A$1:A2,A2)");;
    surveyData.getRange(startRow,surveyData.getLastColumn()-1,importedData.length,1).setValues(filename);
    surveyData.getRange(startRow,surveyData.getLastColumn(),importedData.length,1).setValues(dateTimeCol);
  }else{
    var startRow = contactData.getLastRow() + 1
    var filename = Array(importedData.length).fill([xlsBlob.getName()]);
    var dateTimeCol = Array(importedData.length).fill([Utilities.formatDate(new Date(), "America/New_York", "MM/dd/yyyy h:mm a")]);
    contactData.getRange(startRow, 1, importedData.length, importedData[0].length).setValues(importedData);


    contactData.getRange(2,contactData.getLastColumn()-2,importedData.length,1).setFormula("=COUNTIF(A$1:A2,A2)");;
    contactData.getRange(startRow,contactData.getLastColumn()-1,importedData.length,1).setValues(filename);
    contactData.getRange(startRow,contactData.getLastColumn(),importedData.length,1).setValues(dateTimeCol);

  }

};