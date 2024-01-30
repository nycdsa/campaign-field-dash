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
5. Configure the buttons in the Google Sheets interface. To do this, go back to your Google Sheet, right-click on the button (drawing), and select `Assign Script`. In the text box that appears, type `processSheet` and click `OK`.
6. Now, you can run the script either from the select function dropdown in the Script Editor or by clicking the button with the assigned script in the Google Sheets interface.

Please note that you need to have a Google account to use this feature.

If you need access to the template, please contact Nick or Dave Mahler.

## Configuration Sheet

The Configuration sheet is a crucial part of the script. It is used to store and retrieve various configuration parameters that the script uses to process the data.

The Configuration sheet is expected to have the following structure:

- The third column contains the names of the folders where the script will look for files to process.
- The fourth column contains the names of the data sheets where the script will write the processed data.
- The fifth column contains the header names that the script will use when processing the data.

The script reads these values from either the third or fourth row of the Configuration sheet, depending on the name of the active sheet. If the active sheet's name is "Add Survey Data", the script reads the values from the third row. Otherwise, it reads the values from the fourth row.

Please ensure that the Configuration sheet is properly set up with the correct values in these cells before running the script.

## File Setup

When you run the script, it checks if a specific folder exists in the same location as your Google Spreadsheet. If the folder doesn't exist, the script will create it for you. This folder is where you should place any ZIP files that you want the script to process. Once the ZIP files are in this folder, the script can read and process them accordingly.

## Script Functionality

When you run this script on a Google Spreadsheet, it performs several operations:

1. It reads certain information from the first row of the active sheet.
2. It prepares the active sheet and a specific folder for data processing.
3. It processes CSV files within ZIP files located in the specified folder.
4. It removes any duplicate data from the sheet.
5. It sorts a summary and updates counts based on the number of duplicates.
6. Finally, it shows a message indicating that the processing is done.

