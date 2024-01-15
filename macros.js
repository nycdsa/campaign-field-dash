/** @OnlyCurrentDoc */

var spreadsheet;
var parentFolder;
var contactData;
var surveyData;
var surveyFolder;
var contactFolder_;
var tempFolder;

function compareRangeAndArray(range, stringArray) {
  // Convert the range to a string array
  var rangeValues = range.getValues();
  var rangeArray = rangeValues.map(row => row.map(cell => String(cell)));

  // Compare the converted range array with the provided string array
  var compareArrays = function (array1, array2) {
    if (array1.length !== array2.length) {
      return false;
    }

    return array1.every((row, index) => row.toString() === array2[index].toString());
  };

  return compareArrays(rangeArray, stringArray);
}

function processSurvey() {
  setup();
  setupSheet("surveyDataTemp");
  setupFiles();
  processXLSFilesInZip(surveyData, surveyFolder, "Add Survey Data");
  cleanupSheet(surveyData);
  showDoneMessage();
}

function processContact() {
  setup();
  setupSheet("contactDataTemp");
  setupFiles();
  processXLSFilesInZip(contactData, contactFolder_, "Add Contact Data");
  cleanupSheet(contactData);
  showDoneMessage();
}

function cleanupSheet(sheet) {
  var Array;
  if (sheet.getName() == "surveyDataTemp") {
    Array = [1, 2, 3, 4, 5, 6, 7, 8];
  } else {
    Array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  }
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort(3);
  sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .removeDuplicates(Array);
}

function createTempFolder() {
  var folderName = 'temp'; // Change 'TemporaryFolder' to your desired folder name

  // Check if the folder already exists
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    tempFolder = folders.next();
    Logger.log('Temporary folder already exists with ID: ' + tempFolder.getId());
    return; // Exit the function if the folder exists
  }

  // If the folder doesn't exist, create it
  tempFolder = parentFolder.createFolder(folderName);
  Logger.log('Temporary folder created with ID: ' + tempFolder.getId());
}

function showPleaseWait() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange('A1'); // Adjust the cell as per your preference
  cell.setValue('Please wait...');
  
  // Enhance text visibility and disable text wrapping
  cell.setFontSize(14).setFontWeight("bold").setFontColor("red");
  cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  cell.setWrap(false);
}

function updateProgress(processedXLSFiles, totalXLSFiles) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange('B1'); // Adjust the cell as per your preference

  var progressPercentage = Math.round((processedXLSFiles / totalXLSFiles) * 100);
  cell.setValue('Progress: ' + progressPercentage + '%'); // Displaying progress percentage in cell B1

  // Enhance text visibility and disable text wrapping
  cell.setFontSize(12).setFontWeight("bold").setFontColor("green");
  cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  cell.setWrap(false);
}
function showDoneMessage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange('A1'); // Same cell as the "Please wait..." message
  cell.setValue('Task Completed!');
  
  // Enhance text visibility and disable text wrapping
  cell.setFontSize(14).setFontWeight("bold").setFontColor("blue");
  cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  cell.setWrap(false);
}

function clearProgress() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var progressCell = sheet.getRange('B1'); // Assuming progress is displayed in cell B1
  var doneCell = sheet.getRange('A1'); // Assuming "done" message is displayed in cell C1

  progressCell.clearContent(); // Clear progress
  doneCell.clearContent(); // Clear "done" message
}

function setup() {
  clearProgress();
  showPleaseWait();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  parentFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();

  surveyHeader = [
    [
      "Voter File VANID",
      "Contact Name",
      "Date Canvassed",
      "Survey Response",
      "Survey Question Name",
      "Survey Question Type",
      "Canvassed By",
      "Contact Type",
      "Instance of VANID",
      "Filename",
      "Date/Time Added",
    ],
  ];

  contactHeader = [
    [
      "Voter File VANID",
      "Contact Name",
      "Date Canvassed",
      "Result",
      "Contact Type",
      "Canvassed By",
      "Created By",
      "Date Created",
      "Input Type Name",
      "MiniVAN Campaign",
      "Instance of VANID",
      "Filename",
      "Date/Time Added",
    ],
  ];
  createTempFolder();
}

function setupSheet(sheetName) {
  if (sheetName == "surveyDataTemp") {
    surveyData = spreadsheet.getSheetByName(sheetName);
    if (surveyData != null) {
      surveyData.getDataRange().clear();
    } else {
      surveyData = spreadsheet.insertSheet(sheetName);
    }
    surveyData
      .getRange(1, 1, 1, surveyHeader[0].length)
      .setValues(surveyHeader);
  } else {
    contactData = spreadsheet.getSheetByName(sheetName);
    if (contactData != null) {
      contactData.getDataRange().clear();
    } else {
      contactData = spreadsheet.insertSheet(sheetName);
    }
    contactData
      .getRange(1, 1, 1, contactHeader[0].length)
      .setValues(contactHeader);
  }
}

function setupFiles() {
  // Get the active spreadsheet and its parent folder

  // Define folder names
  var folderNames = ["temp","Survey Response Reports", "Contact History Reports"];

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
  surveyFolder = parentFolder
    .getFoldersByName("Survey Response Reports")
    .next();
  contactFolder_ = parentFolder
    .getFoldersByName("Contact History Reports")
    .next();
}

function processXLSFilesInZip(sheet, folder, output) {
  var filesIterator = folder.getFiles();
  var files = [];
  while (filesIterator.hasNext()) {
    files.push(filesIterator.next());
  }
  isFirstFile = true;
  for (var j = 0; j < files.length; j++) {
    // Update progress after each ZIP file processed
    updateProgress(j, files.length-1);
    var file = files[j];
    var fileName = file.getName();

    if (fileName.endsWith(".zip")) {
      var zipBlob = file.getBlob();
      var zip = Utilities.unzip(zipBlob);

      // Process the ZIP file contents here
      Logger.log("Processing ZIP file: " + fileName);
      try {
        for (var i = 0; i < zip.length; i++) {
          var innerFile = zip[i];
          var innerFileName = innerFile.getName();

          // Handle specific file types within the ZIP (e.g., XLSX)
          if (innerFileName.endsWith(".xls")) {
            Logger.log("Processing file within ZIP: " + innerFileName);

            // Perform operations on the XLS file
            try {
              processXLSData(sheet, tempFolder, innerFile, output, j === 0);
            } catch (e) {
              addRowsToTable(output, innerFileName, null, null, null, e, isFirstFile);
            }
          } else {
            Logger.log("Skipping non-XLS file within ZIP: " + innerFileName);
          }
        }
      } catch (e) {
        addRowsToTable(output, fileName, null, null, null, e, isFirstFile);
      }

    }
    isFirstFile = false;
  }
}



function processXLSData(sheet, folder, xlsBlob, outputSheet, reset) {
  var newFile = {
    title: xlsBlob.getName(),
    parents: [{ id: folder.getId() }], //  Added
  };
  var file = Drive.Files.insert(newFile, xlsBlob, {
    convert: true,
  });

  var otherSpreadsheet = SpreadsheetApp.openById(file.id);
  var otherSheet = otherSpreadsheet.getSheets()[0]; // Change to the desired sheet name
  var dateRange = otherSheet.getRange(2, 3, otherSheet.getLastRow() - 1, 3);
  var dateValues = dateRange.getValues().flat().filter(Number);
  var hdr = otherSheet
    .getRange(1, 1, 1, otherSheet.getLastColumn());

  var contactHdr = Array([
    "Voter File VANID",
    "Contact Name",
    "Date Canvassed",
    "Result",
    "Contact Type",
    "Canvassed By",
    "Created By",
    "Date Created",
    "Input Type Name",
    "MiniVAN Campaign"
  ]);

  var surveyHdr = Array([
    "Voter File VANID",
    "Contact Name",
    "Date Canvassed",
    "Survey Response",
    "Survey Question Name",
    "Survey Question Type",
    "Canvassed By",
    "Contact Type"
  ]);

  if (
    ((outputSheet == "Add Survey Data") & !compareRangeAndArray(hdr,surveyHdr)) |
    ((outputSheet == "Add Contact Data") & !compareRangeAndArray(hdr,contactHdr))
  ) {
    Drive.Files.remove(file.id); // Added // If this line is run, the original XLSX file is removed. So please be careful this.

    throw new Error(
      "Header does not align with expected output, please check file"
    );
  }

  var maxRange = Utilities.formatDate(
    new Date(Math.max.apply(null, dateValues)),
    "America/New_York",
    "MM/dd/yyyy"
  );
  var minRange = Utilities.formatDate(
    new Date(Math.min.apply(null, dateValues)),
    "America/New_York",
    "MM/dd/yyyy"
  );

  addRowsToTable(
    outputSheet,
    xlsBlob.getName(),
    minRange,
    maxRange,
    dateRange.getNumRows(),
    "OK",
    reset
  );

  Logger.log(
    "File " +
      xlsBlob.getName() +
      ", starts on " +
      minRange +
      " and ends on " +
      maxRange
  );

  var importedData = otherSheet
    .getRange(2, 1, otherSheet.getLastRow() - 1, otherSheet.getLastColumn())
    .getValues();

  Drive.Files.remove(file.id); // Added // If this line is run, the original XLSX file is removed. So please be careful this.

  // Append the imported data to the existing or new sheet

  var startRow = sheet.getLastRow() + 1;
  var filename = Array(importedData.length).fill([xlsBlob.getName()]);
  var dateTimeCol = Array(importedData.length).fill([
    Utilities.formatDate(new Date(), "America/New_York", "MM/dd/yyyy h:mm a"),
  ]);

  sheet
    .getRange(startRow, 1, importedData.length, importedData[0].length)
    .setValues(importedData);

  sheet
    .getRange(2, sheet.getLastColumn() - 2, importedData.length, 1)
    .setFormula("=COUNTIF(A$1:A2,A2)");
  sheet
    .getRange(startRow, sheet.getLastColumn() - 1, importedData.length, 1)
    .setValues(filename);
  sheet
    .getRange(startRow, sheet.getLastColumn(), importedData.length, 1)
    .setValues(dateTimeCol);
}

function addRowsToTable(
  sheetName,
  filename,
  min,
  max,
  numRowsProcessed,
  statusColumn,
  isFirstValue
) {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  var sheet = spreadsheet.getSheetByName(sheetName);

  // Define the specific range you want to work with
  var range = sheet.getRange("C3:H" + sheet.getLastRow());

  // If isFirstValue is true, clear the table within the range
  if (isFirstValue) {
    range.clearContent();
    sheet
      .getRange(3, 3, 1, 6)
      .setValues([
        [
          "File Name",
          "First Date",
          "Last Date",
          "# Records",
          "Error Notes",
          "Timestamp",
        ],
      ]);
  }

  // Get the last row in the specified range
  var lastRow = getLastRowInRange(spreadsheet, "C", "H");

  // Calculate the next row number
  var nextRow = lastRow + 1;

  // Get the current timestamp
  var timestamp = Utilities.formatDate(
    new Date(),
    "America/New_York",
    "MM/dd/yyyy hh:mm:ss aa"
  ); //Utilities.formatDate(new Date(),"America/New_York", "MM/dd/yyyy")

  // Create an array with the data to add to the table
  var rowData = [filename, min, max, numRowsProcessed, statusColumn, timestamp];

  // Set the values in the next row of the specified range
  sheet.getRange(nextRow, 3, 1, 6).setValues([rowData]);
}

function getLastRowInRange(sheet, startColumn, endColumn) {
  var data = sheet
    .getRange(startColumn + "2:" + endColumn + sheet.getLastRow())
    .getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    if (
      row.some(function (cell) {
        return cell !== "";
      })
    ) {
      return i + 2; // Adding 2 to account for 1-based indexing and header row
    }
  }
  return 1; // If the range is completely empty, return 1 as the first row.
}
