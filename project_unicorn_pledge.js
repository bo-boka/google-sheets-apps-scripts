
/**
* Puts tools in a map, iterates through Pledge list vendors, to find companies based on either product or company name
* @returns {array} Pastes company name & id in end columns of pledge list
*/
function match_companies_and_products() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Access the sheets
  var toolsSheet = ss.getSheetByName("2024_04_16_tools");
  var pledgeSheet = ss.getSheetByName("Copy of pledge_list");

  // Get the data from both sheets
  var toolsData = toolsSheet.getDataRange().getValues();
  var pledgeData = pledgeSheet.getDataRange().getValues();

  // Define regex pattern to match various forms of business designations
  var businessPattern = /,?\s*(LLC|INC|LTD)\.?/gi;

  // Create a map for faster searching
  var toolsMap = new Map();
  toolsData.forEach(function(row, rowIndex) {

    var keyA = row[0].toString().replace(/\s+/g, '').replace(businessPattern, '').toLowerCase();
    var keyC = row[2].toString().replace(/\s+/g, '').replace(businessPattern, '').toLowerCase();

    if (keyA || keyC) { // Add to map if either keyA or keyC is non-empty
      if (!toolsMap.has(keyA)) {
        toolsMap.set(keyA, []);
      }
      if (!toolsMap.has(keyC)) {
        toolsMap.set(keyC, []);
      }
      toolsMap.get(keyA).push(rowIndex);
      toolsMap.get(keyC).push(rowIndex);
    }

  });

  // Iterate over pledgeData to find matches and update values
  pledgeData.forEach(function(row, rowIndex) {
    if (row[0]) {
      // Remove specific text patterns from column A values
      var pledgeKey = row[0].toString().replace(/\s+/g, '').replace(businessPattern, '').toLowerCase();

      if (toolsMap.has(pledgeKey)) {
        var matchingRows = toolsMap.get(pledgeKey);
        var uniqueMatchesI = new Set();
        var uniqueMatchesJ = new Set();

        matchingRows.forEach(function(matchingRowIndex) {
          var toolRow = toolsData[matchingRowIndex];
          // Only aggregate non-duplicate matches
          if (!uniqueMatchesI.has(toolRow[0])) {
            row[8] = row[8] ? row[8] + ", " + toolRow[0] : toolRow[0]; // Column K (index 10)
            uniqueMatchesI.add(toolRow[0]);
          }
          if (!uniqueMatchesJ.has(toolRow[1])) {
            row[9] = row[9] ? row[9] + ", " + toolRow[1] : toolRow[1]; // Column L (index 11)
            uniqueMatchesJ.add(toolRow[1]);
          }
        });
      }
    }
  });

  // Write updated data back to the pledge sheet for columns K and L
  var rangeToUpdate = pledgeSheet.getRange(1, 9, pledgeData.length, 2); // Starting from column K (index 11), 2 columns wide
  var updatedValues = pledgeData.map(function(row) {
    return [row[8], row[9]];
  });
  rangeToUpdate.setValues(updatedValues);
}

/**
* Puts tools in a map, iterates through Pledge list vendors urls, strips for domain, sees if urls contain domain
* @returns {array} Pastes company name & id in end columns of pledge list
*/
function match_url_domains() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Access the sheets
  var toolsSheet = ss.getSheetByName("2024_04_16_tools");
  var pledgeSheet = ss.getSheetByName("Copy of pledge_list");

  /// Get the data from both sheets, skipping the first row
  var toolsData = toolsSheet.getRange(2, 1, toolsSheet.getLastRow() - 1, toolsSheet.getLastColumn()).getValues();
  var pledgeData = pledgeSheet.getRange(2, 1, pledgeSheet.getLastRow() - 1, pledgeSheet.getLastColumn()).getValues();

  // Define the pattern to remove URL prefixes
  var urlPattern = /^(https?:\/\/|www\.)+/i;

  // Iterate over pledgeData to find matches and update values
  pledgeData.forEach(function(row, rowIndex) {
    if (row[1]) { // Check if there is a value in Column B
      // Clean the URL from specific patterns
      var pledgeKey = row[1].toString().replace(urlPattern, '').toLowerCase();

      // Ensure that there is a value in the expected column and it's converted to a string
      var columnKContent = (row[12] || "").toString(); // Handling undefined or other types
      var columnLContent = (row[13] || "").toString(); // Handling undefined or other types

      // Prepare sets to keep track of unique matches
      var uniqueMatchesK = new Set(columnKContent ? columnKContent.split(", ") : []);
      var uniqueMatchesL = new Set(columnLContent ? columnLContent.split(", ") : []);

      // Iterate through toolsData to check for substring matches in columns G and H
      toolsData.forEach(function(toolRow) {
        var columnG = toolRow[6] ? toolRow[6].toString().toLowerCase() : "";
        var columnH = toolRow[7] ? toolRow[7].toString().toLowerCase() : "";

        if (columnG.includes(pledgeKey) || columnH.includes(pledgeKey)) {
          // If there's a match, add to set if not already present
          if (!uniqueMatchesK.has(toolRow[0])) {
            uniqueMatchesK.add(toolRow[0]);
          }
          if (!uniqueMatchesL.has(toolRow[1])) {
            uniqueMatchesL.add(toolRow[1]);
          }
        }
      });

      // Join sets into strings for column K and L
      row[12] = Array.from(uniqueMatchesK).join(", ");
      row[13] = Array.from(uniqueMatchesL).join(", ");
    }
  });

  // Write updated data back to the pledge sheet for columns K and L
  var rangeToUpdate = pledgeSheet.getRange(2, 13, pledgeData.length, 2); // Starting from column K (index 11), 2 columns wide
  var updatedValues = pledgeData.map(function(row) {
    return [row[12], row[13]];
  });
  rangeToUpdate.setValues(updatedValues);
}


/**
* Highlights rows that have unmatching values between 2 specified columns
* @returns sheet highlight red
*/
function highlightDifferences() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Copy of pledge_list");

  // Get the range that excludes the first row, considering values in columns L and N
  var range = sheet.getRange(2, 12, sheet.getLastRow() - 1, 3); // Columns L (12) and N (14), with a span of 3 columns including an extra column between them
  var values = range.getValues();

  // Loop through each row in the retrieved range
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var columnL = row[0]; // First element in the array (column L)
    var columnN = row[2]; // Second element in the array (column N)

    // Convert both values to string to avoid issues with different data types
    // and trim to avoid issues with extra spaces
    columnL = (columnL === null || columnL === undefined) ? "" : String(columnL).trim();
    columnN = (columnN === null || columnN === undefined) ? "" : String(columnN).trim();

    // Check if the values in columns L and N are different
    if (columnL !== columnN) {
      // Highlight the entire row in red, adjusting for the offset and skipping the header row
      var rowToHighlight = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn());
      rowToHighlight.setBackground("red");
    } else {
      // Optionally clear any previous highlighting if values match
      var rowToClear = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn());
      rowToClear.setBackground(null); // Set to default background
    }
  }
}

/**
* Puts all comma separated values within a column into their own rows
* @returns {array} ID values
*/
function transferUniqueValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Accessing the source sheet and the target sheet
  var sourceSheet = ss.getSheetByName("Copy of pledge_list");
  var targetSheet = ss.getSheetByName("Scratch");

  // Getting the values from column J, skipping the first row
  var sourceRange = sourceSheet.getRange(2, 10, sourceSheet.getLastRow() - 1);
  var sourceValues = sourceRange.getValues();

  // Initialize a set to keep track of unique values
  var uniqueValues = new Set();

  // Process each value in the source column
  sourceValues.forEach(function(row) {
    var cellValue = row[0];
    if (cellValue) {
      // Split by commas to handle multiple entries in one cell
      var items = cellValue.toString().split(",");
      items.forEach(function(item) {
        // Trim spaces and add to set if not empty
        var trimmedItem = item.trim();
        if (trimmedItem) {
          uniqueValues.add(trimmedItem);
        }
      });
    }
  });

  // Convert set back to array for output
  var outputArray = Array.from(uniqueValues).map(function(value) {
    return [value];  // Each value needs to be an array for setValues to work
  });

  // Clear existing data in target column to prevent old data overlap
  targetSheet.getRange(2, 1, targetSheet.getMaxRows() - 1, 1).clearContent();

  // Write unique values to the target sheet starting from the second row
  if (outputArray.length > 0) {
    targetSheet.getRange(2, 1, outputArray.length, 1).setValues(outputArray);
  }
}


/**
* Finds all products that match a company ID to list in separate sheet (multiple products to a company)
* @returns {array} ID values and names for products and company
*/
function copyMatchesToNewSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Access the "Scratch" and "2024_04_16_tools" sheets
  var scratchSheet = ss.getSheetByName("Scratch");
  var toolsSheet = ss.getSheetByName("2024_04_16_tools");

  // Prepare the "Scratch2" sheet, creating it if it does not exist
  var scratch2Sheet = ss.getSheetByName("Scratch2");
  if (!scratch2Sheet) {
    scratch2Sheet = ss.insertSheet("Scratch2");
  }

  // Get values from Column A in "Scratch" skipping the first row
  var scratchData = scratchSheet.getRange(2, 1, scratchSheet.getLastRow() - 1).getValues();

  // Get values from "2024_04_16_tools" for matching
  var toolsData = toolsSheet.getRange(1, 1, toolsSheet.getLastRow(), 4).getValues(); // Fetching A, B, C, D columns

  // Clear old data from "Scratch2"
  // scratch2Sheet.clear();

  // Set headers in "Scratch2" if needed (optional)
  // scratch2Sheet.getRange(1, 1, 1, 4).setValues([["Column B", "Column A", "Column D", "Column C"]]);

  // To store matching rows to be written to Scratch2
  var matchingRows = [];

  // Iterate over each value in "Scratch" column A
  scratchData.forEach(function(scratchRow) {
    var scratchValue = scratchRow[0].toString().trim();
    if (scratchValue) { // Ensure non-empty value
      toolsData.forEach(function(toolsRow, index) {
        if (index > 0) { // Skip header row
          var toolsValue = toolsRow[1].toString().trim(); // Column B in "2024_04_16_tools"
          if (toolsValue === scratchValue) {
            // Add a row to matchingRows in the order B, A, D, C
            matchingRows.push([toolsRow[1], toolsRow[0], toolsRow[3], toolsRow[2]]);
          }
        }
      });
    }
  });

  // Write the collected rows to "Scratch2" if any matches were found
  if (matchingRows.length > 0) {
    scratch2Sheet.getRange(2, 1, matchingRows.length, 4).setValues(matchingRows);
  }
}
