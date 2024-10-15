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
          // Calculate and insert total duration before renaming the old sheet
          calculateTotalDuration(oldActiveSheet);
          
          oldActiveSheet.setName(newSheetName); // Rename the old "Active Sheet"
        }
  
        // Create a new "Active Sheet"
        const newActiveSheet = ss.insertSheet("Active Sheet");
        
        // Apply the header to the new active sheet
        applyHeaderToSheet(newActiveSheet);
      }
      
      // Notify the user the process is complete
      ui.alert("The sheets have been successfully renamed and the new 'Active Sheet' created.");
      
    } else {
      ui.alert('Operation cancelled or no new name was entered.');
    }
  }
  
  // Function to calculate the total duration and add it to the last row before the new sheet
  function calculateTotalDuration(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // No data to calculate
  
    // Assume duration is in column 4 (D), starting from row 2
    const durationRange = sheet.getRange(2, 4, lastRow - 1, 1);
  
    // Calculate the total duration
    const totalDuration = durationRange.getValues().reduce((sum, row) => {
      const time = row[0];
      if (time instanceof Date) {
        return sum + time.getHours() * 60 + time.getMinutes(); // Convert time to total minutes
      }
      return sum;
    }, 0);
  
    // Convert total minutes back to hours and minutes
    const totalHours = Math.floor(totalDuration / 60);
    const totalMinutes = totalDuration % 60;
  
    // Write the total duration at the last row of the sheet
    const totalRow = lastRow + 1;
    sheet.getRange(totalRow, 4).setValue(Utilities.formatString("%02d:%02d", totalHours, totalMinutes));
  
    // Format the total duration cell
    const totalCell = sheet.getRange(totalRow, 4);
    totalCell.setBackground("#90EE90");  // Light green for total duration
    totalCell.setFontWeight("bold");
    totalCell.setHorizontalAlignment("center");
    totalCell.setVerticalAlignment("middle");
    totalCell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
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