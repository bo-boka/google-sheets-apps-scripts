/**
* Simple matching for products names and company names.
* Doesn't account for duplicates, which will be found when manually checking
*/
function updateIKSFull() {
  var iksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("iks_full");
  var lpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("lp_full_april_2024");

  // Column A iks tool names
  var iksValues = iksSheet.getRange("A:A").getValues().map(function(row) {
        return row[0] ? row[0].toString().replace(/\s+/g, '').toLowerCase() : "";
  });

  // Column C LP tool names
  var lpValuesC = lpSheet.getRange("C:C").getValues().map(function(row) {
    return row[0].toString().replace(/\s+/g, '').toLowerCase();
  });

  // Column A LP company names
  var lpValuesA = lpSheet.getRange("A:A").getValues().map(function(row) {
    return row[0].toString().replace(/\s+/g, '').toLowerCase();
  });

  for (var i = 0; i < iksValues.length; i++) {
    var rowIndexC = lpValuesC.indexOf(iksValues[i]);
    var rowIndexA = lpValuesA.indexOf(iksValues[i]);

    if (rowIndexC !== -1) {
      var rowData = lpSheet.getRange(rowIndexC + 1, 1, 1, 4).getValues()[0];
      iksSheet.getRange(i + 1, 2, 1, 4).setValues([rowData]);

    } else if (rowIndexA !== -1) {
      var rowData = lpSheet.getRange(rowIndexA + 1, 1, 1, 4).getValues()[0];
      iksSheet.getRange(i + 1, 2, 1, 4).setValues([rowData]);

    }

  }
}
