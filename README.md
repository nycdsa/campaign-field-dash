# Campaign in a Box, VAN Report Importer

This is a Google Apps Script project that processes CSV files within ZIP files in a specified Google Drive folder.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

You need to have a Google account to run this script in the Google Apps Script environment. 

### Installing

1. Open Google Apps Script in your browser.
2. Click on `New Project`.
3. Copy and paste the code from `macros.js` into the script editor.
4. Save the project with your desired name.

## Running the script

To run the script, follow these steps:

1. Select the function you want to run from the select function dropdown.
2. Click on the play icon to run the function.

## Built With

* [Google Apps Script](https://developers.google.com/apps-script) - The scripting language used

## Authors

* **Sebastian Caceres** - *Campaign in a Box, VAN Report Importer* - [SebastianCaceres](https://github.com/SebastianCaceres/)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Using the Prebuilt Spreadsheet

If you want to get started quickly, you can use our prebuilt Google Spreadsheet. Here's how:

1. Go to [Prebuilt Spreadsheet](https://docs.google.com/spreadsheets/d/1ZO4C2JnS4Z-QF5EEXGjZcA6inW1ZOpqtoqtByZeuCc4/edit?usp=drive_link).
2. Click on `File` > `Make a copy...` to create a copy of the spreadsheet in your Google Drive.
3. Open the Script Editor from `Extensions` > `Apps Script`.
4. Paste the script code into the Script Editor and save the project.
5. Now, you can run the script from the select function dropdown in the Script Editor.

Please note that you need to have a Google account to use this feature.

If you need access to the template, please contact Nick or Dave Mahler.

## Script Functionality

The main function of the script is `processSheet`. This function processes a Google Spreadsheet sheet and performs several operations:

1. Retrieves and preprocesses header data from the 11th column of the first row in the active sheet.
2. Concatenates additional column headers to the retrieved header data.
3. Calls `setupSheet` function with the value from the 8th column of the first row in the active sheet. This function is expected to perform some setup tasks on the sheet.
4. Calls `setupFiles` function with the value from the 5th column of the first row in the active sheet. This function is expected to perform some setup tasks on the files.
5. Calls `processCSVFilesInZip` function to process CSV files within ZIP files in a specified folder.
6. Calls `cleanupSheet` function to remove duplicates from the sheet.
7. Calls `sortSummaryandUpdateCounts` function to sort a summary and update counts based on the number of duplicates.
8. Calls `showDoneMessage` function to show a message indicating that the processing is done.

Please note that the actual functionality of `setupSheet`, `setupFiles`, `processCSVFilesInZip`, `cleanupSheet`, `sortSummaryandUpdateCounts`, and `showDoneMessage` functions depends on their implementation in the script.
