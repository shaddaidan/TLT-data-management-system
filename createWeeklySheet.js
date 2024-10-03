function createWeeklyHeader() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Active Sheet";  // Set the sheet name where the header will be created
    let newSheet = ss.getSheetByName(sheetName);
  
    // If the sheet doesn't exist, create a new one
    if (!newSheet) {
      newSheet = ss.insertSheet(sheetName);
    }
  
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
    // Set up column headers
    const headers = ["Date", "Start Time", "End Time", "Duration", "Waiting Time", "Student", "Year", "Subject", "Topic", "Tutor", "Status", "Timestamp"];
  
    // Find the last row with data
    const lastRow = newSheet.getLastRow();
  
    // Start the header on row 2 if no data, or after the last row + 3 rows
    const nextRow = lastRow > 0 ? lastRow + 3 : 2; 
  
    // Add headers at the next available row
    newSheet.getRange(nextRow, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
    // Set column widths
    newSheet.setColumnWidths(1, headers.length, 150);
  
    // Format the header
    const headerRange = newSheet.getRange(nextRow, 1, 1, headers.length);
    headerRange.setBackground("#FFA500");  // Orange background
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
  
    // Add thick border around the entire header
    headerRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
    // Store the current date in the last header column
    newSheet.getRange(nextRow, headers.length).setValue(formattedDate);
  
    // Set the format of the entire Duration column (D) to time format
    newSheet.getRange("D:D").setNumberFormat("[h]:mm"); 
  
    // Find the first empty row above the current header
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
    const prevHeaderRow = findFirstEmptyRowAbove(newSheet, lastRow, 1) + 2; // Two rows after the previous header
    const endRow = nextRow - 1; // Last row before the new header
    
    if (prevHeaderRow > 1 && prevHeaderRow < endRow) {
      const durationRange = newSheet.getRange(prevHeaderRow, 4, endRow - prevHeaderRow + 1, 1); // Duration column (D)
      
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
      newSheet.getRange(totalRow + 1, 4).setValue(Utilities.formatString("%02d:%02d", totalHours, totalMinutes));
  
      // Format the cell for total duration
      const totalCell = newSheet.getRange(totalRow + 1, 4);
      totalCell.setBackground("#90EE90"); // Light green color
      totalCell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }



//   ----------------------second version with the dictionaries-----------
function createWeeklyHeaderForTutorsAndParents() {
    // Define two dictionaries: one for tutors and one for parents
    const tutorSheets = {
      "Miss Blossom": "1Xoyz_PzsVHHdKePwnSpXl8UpAe1tZToGC7MgR6jCFHY",
      "Miss Dolapo": "1rFAp41FP2MN6d2Fx58g4JnS4R4S7QENIHzj3M-2wwdw",
    };
  
    const parentSheets = {
      "Doudo - Jesse, Joel": "1pIgB1DoNnJ4lYYwWp2GvwFiKkGQbRmgSoBybT6za1V4",
      "Eniola Shemilore": "14i6PwlW3VTS0QeH7a1Pla91rne1wbeXYtj7iS3oJvOk",
    };
  
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
  
      // Set the format of the entire Duration column (D) to time format
      sheet.getRange("D:D").setNumberFormat("[h]:mm");
  
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
  
    // Loop through the tutor sheets and apply the header
    for (let tutor in tutorSheets) {
      const sheetId = tutorSheets[tutor];
      const ss = SpreadsheetApp.openById(sheetId);
      let sheet = ss.getActiveSheet();  // Get the active sheet for the current tutor
      applyHeaderToSheet(sheet);  // Apply header logic to the tutor's sheet
    }
  
    // Loop through the parent sheets and apply the same header logic
    for (let parent in parentSheets) {
      const sheetId = parentSheets[parent];
      const ss = SpreadsheetApp.openById(sheetId);
      let sheet = ss.getActiveSheet();  // Get the active sheet for the current parent
      applyHeaderToSheet(sheet);  // Apply header logic to the parent's sheet
    }
  }