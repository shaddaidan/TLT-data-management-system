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
        
        // Rename the current "Active Sheet" if it exists
        const oldActiveSheet = ss.getSheetByName("Active Sheet");
        if (oldActiveSheet) {
          oldActiveSheet.setName(newSheetName); // Rename the old "Active Sheet"
        }

        // Create a new "Active Sheet"
        const newActiveSheet = ss.insertSheet("Active Sheet");
        
        // Apply the header and total duration logic to the new active sheet
        applyHeaderToSheet(newActiveSheet);
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

    // Call the function to calculate and display the total duration from the previous header
    calculateTotalDuration(sheet, lastRow, nextRow);
}

// Function to calculate the total duration from the previous header
function calculateTotalDuration(sheet, lastRow, nextRow) {
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

    // Start row of the previous header and calculate the duration sum
    const prevHeaderRow = findFirstEmptyRowAbove(sheet, lastRow, 1) + 2; // Two rows after the previous header
    const endRow = nextRow - 1; // Last row before the new header
    
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
        const totalRow = nextRow - 3; // Total goes to the last row before the new header
        sheet.getRange(totalRow + 1, 4).setValue(Utilities.formatString("%02d:%02d", totalHours, totalMinutes));

        // Format the cell for total duration
        const totalCell = sheet.getRange(totalRow + 1, 4);
        totalCell.setBackground("#90EE90"); // Light green color
        totalCell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
}