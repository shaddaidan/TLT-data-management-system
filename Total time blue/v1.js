function sumGreenTimeCells() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Active Sheet');  // Access the sheet named 'Active Sheet'
    var lastRow = sheet.getLastRow();  // Get the last row with content in the sheet
    var totalMinutes = 0;
    var greenCellTimes = [];  // Array to store times from green cells
    var greenShades = ['#00ff00', '#b7e1cd', '#00cc00', '#ccffcc', '#90ed91', '#90ee90'];  // Array of possible green shades
    
    // Loop through column D (4th column)
    for (var row = 1; row <= lastRow; row++) {
      var cell = sheet.getRange(row, 4);  // Get the cell in column D of the current row
      var bgColor = cell.getBackground();  // Get the background color of the cell
      
      // Log the background color for troubleshooting
      console.log("Row " + row + " - Cell D" + row + " background color: " + bgColor);
      
      // Check if the background color matches any green shade
      if (greenShades.includes(bgColor)) {
        console.log("Green cell found at row " + row + " with background color: " + bgColor);
        var timeValue = cell.getValue();  // Get the value (time) from the cell
        
        // Append the time value to the greenCellTimes array
        greenCellTimes.push(timeValue);  // Store the time from green cells in the array
        
        // Log the time found
        console.log("Time found in green cell: " + timeValue);
        
        // If timeValue is a Date object (as seen in the log), extract the hours and minutes
        if (timeValue instanceof Date) {
          var hours = timeValue.getHours();
          var minutes = timeValue.getMinutes();
          
          // Add the total minutes for this time
          totalMinutes += (hours * 60) + minutes;
        } else if (typeof timeValue === 'string' && timeValue.indexOf(":") !== -1) {
          // Handle the case where the value is in 'HH:MM' string format (fallback)
          var parts = timeValue.split(":");
          var hours = parseInt(parts[0], 10);  // Get the hours part
          var minutes = parseInt(parts[1], 10);  // Get the minutes part
          
          // Add the total minutes for this time
          totalMinutes += (hours * 60) + minutes;
        }
      }
    }
    
    // If green cells were found, log the total time
    if (greenCellTimes.length > 0) {
      // Convert total minutes back to 'HH:MM' format
      var totalHours = Math.floor(totalMinutes / 60);
      var remainingMinutes = totalMinutes % 60;
      var totalTimeFormatted = totalHours.toString().padStart(2, '0') + ":" + remainingMinutes.toString().padStart(2, '0');
      
      // Log the times found and the sum of times from green cells
      console.log("Times found in green cells: " + greenCellTimes.join(", "));
      console.log("Total time sum: " + totalTimeFormatted);
    } else {
      // If no green cells were found
      console.log("No green cells found.");
    }
  }