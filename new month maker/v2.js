// Function to apply the header and logic to any provided sheet
function applyHeaderToSheet(sheet) {
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
    // Set up column headers
    const headers = ["Date", "Start Time", "End Time", "Duration", "Waiting Time", "Student", "Year", "Subject", "Topic", "Tutor/Parent", "Status", "Timestamp"];
  
    // Find the last row with data
    const lastRow = sheet.getLastRow();
  
    // Locate the last data section's duration column (Assuming duration is in column 4)
    const durationRange = sheet.getRange(2, 4, lastRow - 1, 1); // Column D has the duration data starting from row 2
  
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
    const totalRow = lastRow + 1; // Total goes to the last row after the previous data
    sheet.getRange(totalRow, 4).setValue(Utilities.formatString("%02d:%02d", totalHours, totalMinutes));
  
    // Format the cell for total duration
    const totalCell = sheet.getRange(totalRow, 4);
    totalCell.setBackground("#90EE90"); // Light green color for total duration
    totalCell.setFontWeight("bold");
    totalCell.setHorizontalAlignment("center");
    totalCell.setVerticalAlignment("middle");
    totalCell.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
    // Now proceed to create the new header for the new section
    const nextRow = totalRow + 3; // Start the new header 3 rows after the total duration row
    sheet.getRange(nextRow, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
    // Set column widths
    sheet.setColumnWidths(1, headers.length, 150);
  
    // Format the new header
    const headerRange = sheet.getRange(nextRow, 1, 1, headers.length);
    headerRange.setBackground("#FFA500");  // Orange background for the new header
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
  
    // Add a thick border around the entire header
    headerRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
    // Store the current date in the last header column
    sheet.getRange(nextRow, headers.length).setValue(formattedDate);
  }
  
  // Main function to go through all tutor sheets, create the new active sheet, and apply the header
  function updateTutorSheets() {
    const tutorSheets = {
      "Mr Shaddai": "11YJdkV2lCRv_FjVNqyVhOtnDMFLRQ5_kXPZxusmE3UI",
      "Miss Blossom": "1Xoyz_PzsVHHdKePwnSpXl8UpAe1tZToGC7MgR6jCFHY",
      "Miss Dolapo": "1rFAp41FP2MN6d2Fx58g4JnS4R4S7QENIHzj3M-2wwdw",
      "Miss Erioluwa": "12m6PDGFK27lj3GAncdmAoxZG39k3or6-dq8ifB00YQE",
      "Miss Priscilla": "1KGzdYGCFWHUYIbPXZFZNlP3tCrEfnB146rmU1OsmsMo",
      "Mr Gideon": "1KaiE12sQBgZLdDHqHABcm1a8WF_g2AqWJU5O2IraWhg"
    };
  
    // Go through each tutor's spreadsheet
    for (const [tutorName, sheetId] of Object.entries(tutorSheets)) {
      const ss = SpreadsheetApp.openById(sheetId);
      
      // Get the active sheet (this will be renamed later)
      const oldActiveSheet = ss.getSheetByName("Active Sheet");
      
      // If old active sheet exists, rename it to tutor's name with timestamp
      if (oldActiveSheet) {
        const newSheetName = tutorName + " - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        oldActiveSheet.setName(newSheetName);
  
        // Apply the total duration logic and header formatting before creating the new active sheet
        applyHeaderToSheet(oldActiveSheet);
      }
  
      // Create a new sheet named "Active Sheet"
      const newActiveSheet = ss.insertSheet("Active Sheet");
  
      // Apply header to the new active sheet
      applyHeaderToSheet(newActiveSheet);
    }
  }