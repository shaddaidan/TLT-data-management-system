function onOpen() {
    const ui = SpreadsheetApp.getUi();
    
    // Create a menu option to rename sheets
    ui.createMenu('New Month Sheets')
      .addItem('Rename and Create New Sheet', 'renameActiveSheet')
      .addToUi();
}

function renameActiveSheet() {
    const ui = SpreadsheetApp.getUi();
    
    // Prompt user for input (new name for old "Active Sheet")
    const response = ui.prompt('Rename Sheet', 'Enter a new name for the current active sheets:', ui.ButtonSet.OK_CANCEL);
    const newSheetName = response.getResponseText().trim();
    
    if (response.getSelectedButton() == ui.Button.OK && newSheetName) {
      // Dictionary of tutor names and sheet IDs
      const tutorSheets = {
        "ndw week sheet": "1Jf4C0NedWBR5NlwyfqMeHdZSRVLFxqQ2ix-ERdtgfWo",
        "001 - oladele": "1CM9nRhLQTjD0nA5q8jH57mqr4lgBnV73sTQDvmaLmbs",
        "Mr Shaddai": "1l_DXrdHK1c79U4oX5Xm8MliOtwn8Z41pcUv_aDgevuo",
        "Aiden Josiah": "1w6YEmAqzMyF98jjNT8wRfivqc2_X00JvHznO0PCKmbs",
      };
      
      // Loop through each tutor's spreadsheet
      for (const tutor in tutorSheets) {
        const spreadsheetId = tutorSheets[tutor];
        const ss = SpreadsheetApp.openById(spreadsheetId);
        
        // Get the old "Active Sheet" (the sheet we are renaming and calculating total duration for)
        const oldActiveSheet = ss.getSheetByName("Active Sheet");
        
        if (oldActiveSheet) {
          // Step 1: Calculate the total duration on the old "Active Sheet"
          calculateTotalDuration(oldActiveSheet);
          
          // Step 2: Sum the green time cells
          sumGreenTimeCells(oldActiveSheet);
          
          // Step 3: Rename the old "Active Sheet"
          oldActiveSheet.setName(newSheetName);
          
          // Step 4: Create a new "Active Sheet"
          const newActiveSheet = ss.insertSheet("Active Sheet");
          
          // Apply the header and formatting to the new active sheet
          applyHeaderToSheet(newActiveSheet);
        }
      }
      
      // Notify the user the process is complete
      ui.alert("The sheets have been successfully renamed and the new 'Active Sheet' created.");
    } else {
      ui.alert('Operation cancelled or no new name was entered.');
    }
}

// Function to apply the header and logic to any provided sheet
function applyHeaderToSheet(sheet) {
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Set up column headers
    const headers = ["Date", "Start Time", "End Time", "Duration", "Waiting Time", "Student", "Year", "Subject", "Topic", "Tutor/Parent", "Status", "Timestamp"];

    // Find the last row with data
    const lastRow = sheet.getLastRow();

    // Start the header on row 2 if no data, or after the last row + 3 rows
    const nextRow = lastRow > 0 ? lastRow + 3 : 2;

    // Add headers at the next available row
    sheet.getRange(nextRow, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    // Set column widths
    sheet.setColumnWidths(1, headers.length, 150);

    // Format the header
    const headerRange = sheet.getRange(nextRow, 1, 1, headers.length);
    headerRange.setBackground("#FFA500");  // Orange background
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");

    // Add thick border around the entire header
    headerRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Store the current date in the last header column
    sheet.getRange(nextRow, headers.length).setValue(formattedDate);
}

// Function to calculate the total duration from the previous header
function calculateTotalDuration(sheet) {
    // Function to find the first empty row above the current header
    function findFirstEmptyRowAbove(sheet, startRow, column) {
        for (let row = startRow; row > 0; row--) {
            let cellValue = sheet.getRange(row, column).getValue();
            if (cellValue === "") {
                return row;  // Return the row number of the first empty row
            }
        }
        return null;  // Return null if no empty row is found
    }

    const lastRow = sheet.getLastRow();
    
    // Start row of the previous header and calculate the duration sum
    const prevHeaderRow = findFirstEmptyRowAbove(sheet, lastRow, 1) + 2; // Two rows after the previous header
    const endRow = lastRow; // Last row before the new header
    
    if (prevHeaderRow > 1 && prevHeaderRow < endRow) {
        const durationRange = sheet.getRange(prevHeaderRow, 4, endRow - prevHeaderRow + 1, 1); // Duration column (D)
        
        // Calculate the total duration between the last header and the new header
        const totalDuration = durationRange.getValues().reduce((sum, row) => {
            const time = row[0];
            if (time instanceof Date) {
                return sum + time.getHours() * 60 + time.getMinutes(); // Convert time to total minutes
            }
            return sum;
        }, 0);

        // Convert total minutes back to a time format (hours and minutes)
        const totalHours = Math.floor(totalDuration / 60);
        const totalMinutes = totalDuration % 60;

        // Write the total duration below the last row of the previous data
        const totalRow = lastRow; // Total goes to the last row of data
        sheet.getRange(totalRow + 1, 4).setValue(Utilities.formatString("%02d:%02d", totalHours, totalMinutes));

        // Format the cell for total duration
        const totalCell = sheet.getRange(totalRow + 1, 4);
        totalCell.setBackground("#90EE90"); // Light green color
        totalCell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
}

// Function to sum green time cells on the old active sheet
function sumGreenTimeCells(sheet) {
    var lastRow = sheet.getLastRow();  // Get the last row with content in the sheet
    var totalMinutes = 0;  // Total duration in minutes
    var greenCellTimes = [];  // Array to store times from green cells
    var greenShades = ['#00ff00', '#b7e1cd', '#00cc00', '#ccffcc', '#90ed91', '#90ee90'];  // Array of possible green shades

    // Loop through column D (4th column)
    for (var row = 1; row <= lastRow; row++) {
        var cell = sheet.getRange(row, 4);  // Get the cell in column D of the current row
        var bgColor = cell.getBackground();  // Get the background color of the cell

        // Check if the background color matches any green shade
        if (greenShades.includes(bgColor)) {
            var timeValue = cell.getDisplayValue();  // Get the displayed value (time) from the cell
            greenCellTimes.push(timeValue);  // Store the time from green cells in the array

            // Convert time string (HH:MM) to minutes and add to totalMinutes
            var timeParts = timeValue.split(":"); // Split the time string into hours and minutes
            if (timeParts.length === 2) {
                var hours = parseInt(timeParts[0], 10);
                var minutes = parseInt(timeParts[1], 10);
                totalMinutes += (hours * 60) + minutes; // Add to total minutes
            }
        }
    }

    // If green cells were found, log the total time
    if (greenCellTimes.length > 0) {
        // Convert total minutes back to 'HH:MM' format
        var totalHours = Math.floor(totalMinutes / 60);
        var remainingMinutes = totalMinutes % 60;
        var totalTimeFormatted = totalHours.toString().padStart(2, '0') + ":" + remainingMinutes.toString().padStart(2, '0');
        
// Output the total time in a blue cell below the last entry in column D
var outputRow = lastRow + 2;  // Two rows below the last entry
var labelCell = sheet.getRange(outputRow, 3);  // Cell in column 3 to display the label
var outputCell = sheet.getRange(outputRow, 4);  // Cell in column 4 to display the total time

// Set the label for the total time (no background color)
labelCell.setValue("Total Time:").setFontWeight("bold");

// Set the total time value (blue background and bold font)
outputCell.setValue(totalTimeFormatted)
    .setFontWeight("bold")
    .setBackground("#add8e6")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);  // Add thick border around the total time
    }
}