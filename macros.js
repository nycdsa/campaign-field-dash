/* eslint-disable no-undef */
/** @OnlyCurrentDoc */
/** @type {*} */
/** @param {GoogleAppsScript.Spreadsheet.Sheet} activeSheet*/
var spreadsheet;
var parentFolder;
var contactData;
var surveyData;
var surveyFolder;
var contactFolder_;
var activeSheet;
function compareArrays(rangeArray, stringArray) {
  if (rangeArray.length !== stringArray.length) {
    return false;
  }

  return rangeArray.every((row, index) => {
    var isEqual = row.toString() === stringArray[index].toString();
    return isEqual;
  });
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function processSurvey() {
  activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Add Survey Data")
  setup();
  setupSheet("surveyDataTemp");
  setupFiles();
  processCSVFilesInZip(surveyData, surveyFolder, "Add Survey Data");
  var numberOfDuplicates = cleanupSheet(surveyData);
  sortSummaryandUpdateCounts(numberOfDuplicates);
  showDoneMessage();
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function processContact() {
  activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Add Contact Data")
  setup();
  setupSheet("contactDataTemp");
  setupFiles();
  processCSVFilesInZip(contactData, contactFolder_, "Add Contact Data");
  var numberOfDuplicates = cleanupSheet(contactData);
  sortSummaryandUpdateCounts(numberOfDuplicates);
  showDoneMessage();
}

function cleanupSheet(sheet) {
  var Array;
  if (sheet.getName() == "surveyDataTemp") {
    Array = [1, 2, 3, 4, 5, 6, 7, 8];
  } else {
    Array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  }

  var lastRow = sheet.getLastRow();

  // Check if there is at least one row of data
  if (lastRow <= 1) {
    // No data, you can handle this case as needed
    Logger.log("No data in the range. Skipping cleanup.");
    return;
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());

  // Sort the data
  dataRange.sort(3);
  // Remove duplicates
  dataRange.removeDuplicates(Array);

  return numberOfDuplicates = lastRow - sheet.getLastRow();


}

/**
 *
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Number} numberOfDuplicates
 * @paramm {Number} total
 */
function sortSummaryandUpdateCounts(numberOfDuplicates) {
  var dataRange = activeSheet.getRange(4, 3, activeSheet.getLastRow() - 3, 6);
  dataRange.sort(4);
  activeSheet.getRange(5, 11).setValue(numberOfDuplicates);
  total = activeSheet.getRange(4, 11).getValue();
  activeSheet.getRange(6, 11).setValue(numberOfDuplicates / total);


}

function showPleaseWait() {
  var cell = activeSheet.getRange("A1"); // Adjust the cell as per your preference
  cell.setValue("Please wait...");

  // Enhance text visibility and disable text wrapping
  cell.setFontSize(14).setFontWeight("bold").setFontColor("red");
  cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  cell.setWrap(false);
}

function updateProgress(processedXLSFiles, totalXLSFiles) {
  var cell = activeSheet.getRange("B1"); // Adjust the cell as per your preference

  var progressPercentage = Math.round(
    (processedXLSFiles / totalXLSFiles) * 100
  );
  cell.setValue("Progress: " + progressPercentage + "%"); // Displaying progress percentage in cell B1

  // Enhance text visibility and disable text wrapping
  cell.setFontSize(12).setFontWeight("bold").setFontColor("green");
  cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  cell.setWrap(false);
}
function showDoneMessage() {
  var cell = activeSheet.getRange("A1"); // Same cell as the "Please wait..." message
  cell.setValue("Task Completed!");

  // Enhance text visibility and disable text wrapping
  cell.setFontSize(14).setFontWeight("bold").setFontColor("blue");
  cell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  cell.setWrap(false);
}

function clearProgress() {
  var progressCell = activeSheet.getRange("B1"); // Assuming progress is displayed in cell B1
  var doneCell = activeSheet.getRange("A1"); // Assuming "done" message is displayed in cell C1

  progressCell.clearContent(); // Clear progress
  doneCell.clearContent(); // Clear "done" message
}

function setup() {
  clearTableIfFirstValue();
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

/**
 * Sets up the necessary files and folders for the import process.
 */
function setupFiles() {
  // Get the active spreadsheet and its parent folder

  // Define folder names
  var folderNames = ["temp", "Survey Response Reports", "Contact History Reports"];

  // Create the folders that don't exist
  folderNames.forEach(function (folderName) {
    var folder = parentFolder.getFoldersByName(folderName);
    if (!folder.hasNext()) {
      parentFolder.createFolder(folderName);
      // eslint-disable-next-line no-undef
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

/**
 * Processes CSV files within a ZIP folder.
 * 
 * @param {Sheet} sheet - The sheet to write the output to.
 * @param {Folder} folder - The folder containing the ZIP files.
 * @param {string} output - The output destination.
 */
function processCSVFilesInZip(sheet, folder, output) {
  var filesIterator = folder.getFiles();
  var files = [];
  while (filesIterator.hasNext()) {
    files.push(filesIterator.next());
  }
  for (var j = 0; j < files.length; j++) {
    // Update progress after each ZIP file processed
    updateProgress(j, files.length - 1);
    var file = files[j];
    var fileName = file.getName();

    if (fileName.endsWith(".zip")) {
      var zipBlob = file.getBlob();
      var zip = Utilities.unzip(zipBlob);

      // Process the ZIP file contents here
      Logger.log("Processing ZIP file: " + fileName);
      try {
        var innerFile = zip[0];
        var innerFileName = innerFile.getName();

        // Handle specific file types within the ZIP (e.g., XLSX)
        if (innerFileName.endsWith(".xls")) {
          Logger.log("Processing file within ZIP: " + innerFileName);

          // Perform operations on the XLS file
          try {
            processCSVData(sheet, fileName, innerFile, output);
          } catch (e) {
            addRowsToTable(
              fileName,
              null,
              null,
              null,
              e);
          }
        } else {
          throw new Error("Invalid file type: " + innerFileName + " in " + fileName);
        }
      } catch (e) {
        addRowsToTable(fileName, null, null, null, e);
      }
    }
  }
}
/**
 *
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Base.Blob} csvBlob
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet
 * @param {string} filename
 * @param {boolean} reset
 */
function processCSVData(sheet,filename_, csvBlob, outputSheet) {

  var csvData = Utilities.parseCsv(csvBlob.getDataAsString("utf-16"), "\t",);

  var contactHdr = [
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
  ];

  var surveyHdr = [
    "Voter File VANID",
    "Contact Name",
    "Date Canvassed",
    "Survey Response",
    "Survey Question Name",
    "Survey Question Type",
    "Canvassed By",
    "Contact Type",
  ];

  // Get the header values
  var hdr = csvData[0];


  if (
    ((outputSheet == "Add Survey Data") & !compareArrays(hdr, surveyHdr)) |
    ((outputSheet == "Add Contact Data") & !compareArrays(hdr, contactHdr))
  ) {
    throw new Error(
      "Header does not align with expected output, please check file"
    );
  }

  // Get the data values excluding headers
  csvData = csvData.slice(1).map(row => {
    // Iterate over each cell in the row
    return row.map((cell, i) => {
      // If the header of this column contains the word "Date", reformat the cell
      if (hdr[i].includes("Date")) {
        var parts = cell.split('-');
        if (parts.length === 3) {
          var date = new Date(parts[0], parts[1] - 1, parts[2]);
          return Utilities.formatDate(date, "America/New_York", "MM/dd/yyyy");
        } else {
          return NaN;
        }
      } else {
        // If the header of this column does not contain the word "Date", return the cell as is
        return cell;
      }
    });
  });
  var dateValues = csvData.map(row => new Date(row[2]))

  var maxRange = Utilities.formatDate(new Date(Math.max.apply(null, dateValues)), "America/New_York", "MM/dd/yyyy");
  var minRange = Utilities.formatDate(new Date(Math.min.apply(null, dateValues)), "America/New_York", "MM/dd/yyyy");

  addRowsToTable(
    filename_,
    minRange,
    maxRange,
    csvData.length,
    "OK"
  );

  Logger.log(
    "File " +
    filename_ +
    ", starts on " +
    minRange +
    " and ends on " +
    maxRange
  );
  // Append the imported data to the existing or new sheet

  var startRow = sheet.getLastRow() + 1;
  var filenameArray = Array(csvData.length).fill([filename_]);
  var dateTimeCol = Array(csvData.length).fill([
    Utilities.formatDate(new Date(), "America/New_York", "MM/dd/yyyy h:mm a"),
  ]);

  sheet
    .getRange(startRow, 1, csvData.length, csvData[0].length)
    .setValues(csvData);

  sheet
    .getRange(2, sheet.getLastColumn() - 2, csvData.length, 1)
    .setFormula("=COUNTIF(A$1:A2,A2)");
  sheet
    .getRange(startRow, sheet.getLastColumn() - 1, csvData.length, 1)
    .setValues(filenameArray);
  sheet
    .getRange(startRow, sheet.getLastColumn(), csvData.length, 1)
    .setValues(dateTimeCol);


}
/**
 *
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startColumn
 * @param {number} endColumn
 * @return {number} 
 */
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

/**
 *
 *
 * @param {String} filename
 * @param {Number} min
 * @param {Number} max
 * @param {Number} numRowsProcessed
 * @param {String} statusColumn
 * @param {boolean} isFirstValue
 */
function addRowsToTable(
  filename,
  min,
  max,
  numRowsProcessed,
  statusColumn
) {

  // Get the last row in the specified range
  var lastRow = getLastRowInRange(activeSheet, "C", "H");

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
  activeSheet.getRange(nextRow, 3, 1, 6).setValues([rowData]);
}

function clearTableIfFirstValue() {
  var range = activeSheet.getRange("C3:H" + activeSheet.getLastRow());

  // If isFirstValue is true, clear the table within the range
  range.clearContent();
  activeSheet
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
