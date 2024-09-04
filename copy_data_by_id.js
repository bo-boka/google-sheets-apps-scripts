function updateIKSFull() {


  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Access the sheets
  var toolsSheet = ss.getSheetByName("sql_pull");
  var destSheet = ss.getSheetByName("dest_sheet");



  var pullSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sql_pull");
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dest_sheet");

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
