
/**
* Fuzzy match for companies that looks at both company and tool columns
* Removes all spacing, business suffixes, and makes lowercase for matching
* @returns {array} updates 2 columns, 1st with comma separated list of tool names, 2nd with comma separated list of ids
*/
function findMatches() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dpSheet = ss.getSheetByName("DP_New");
  var lpSheet = ss.getSheetByName("LP_Tools");

  var dpData = dpSheet.getRange("A2:A" + dpSheet.getLastRow()).getValues();
  var lpDataA = lpSheet.getRange("A2:A" + lpSheet.getLastRow()).getValues();
  var lpDataC = lpSheet.getRange("C2:C" + lpSheet.getLastRow()).getValues();
  var lpDataD = lpSheet.getRange("D2:D" + lpSheet.getLastRow()).getValues();

  var businessPattern = /\b(?:inc\.?|llc\.?|ltd\.?|limited|corp\.?)\b/gi;

  Logger.log(dpData.length);

  // if there are a bunch of blank rows, update length to just the num of populated rows.
  for (var i = 0; i < dpData.length; i++) {
    var dpValue = String(dpData[i][0]).trim();

    // Only process if there's a value in column A
    if (dpValue !== "") {
      dpValue = dpValue.toLowerCase().replace(businessPattern, "").replace(/\s+/g, '');
      var matchesName = new Set();
      var matchesID = new Set();

      for (var j = 0; j < lpDataA.length; j++) {
        var lpValueA = String(lpDataA[j][0]).toLowerCase().replace(businessPattern, "").replace(/\s+/g, '');
        var lpValueC = String(lpDataC[j][0]).toLowerCase().replace(businessPattern, "").replace(/\s+/g, '');

        if (dpValue === lpValueA || dpValue === lpValueC) {
          matchesName.add(lpDataC[j][0]);
          matchesID.add(lpDataD[j][0]);
        }
      }

      // updates current row, starting at col 10, on 1 row, for 2 col cells
      dpSheet.getRange(i + 2, 10, 1, 2).setValues([[Array.from(matchesName).join(", "), Array.from(matchesID).join(", ")]]);
    }
  }
}
