
/**
* Fuzzy match for companies that looks at both company and tool columns and domains to find list of matches that can be
* manually narrowed down to single id that can then be matched to get the rest of the tool data in another function.
* Removes all spacing, business suffixes, and makes lowercase for matching
* @returns {array} updates 2 columns, 1st with comma separated list of tool names, 2nd with comma separated list of ids
*/
function findMatches() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dpSheet = ss.getSheetByName("DP_New");
  var lpSheet = ss.getSheetByName("LP_Tools");

  // TODO: below I began to work on domain matching, but wondered if that should be its own matching since it holds a priority

  // sheet 1
  var dpDataA = dpSheet.getRange("A2:A" + dpSheet.getLastRow()).getValues();  // Tool name
  var dpDataB = dpSheet.getRange("B2:B" + dpSheet.getLastRow()).getValues();  // Tool domain
  // sheet 2
  var lpDataA = lpSheet.getRange("A2:A" + lpSheet.getLastRow()).getValues();  // Company name
  var lpDataC = lpSheet.getRange("C2:C" + lpSheet.getLastRow()).getValues();  // Tool name
  var lpDataD = lpSheet.getRange("D2:D" + lpSheet.getLastRow()).getValues();  // Tool ID
  var lpDataG = lpSheet.getRange("D2:D" + lpSheet.getLastRow()).getValues();  // Tool domains
  var lpDataH = lpSheet.getRange("D2:D" + lpSheet.getLastRow()).getValues();  // Tool url

  var businessPattern = /\b(?:inc\.?|llc\.?|ltd\.?|limited|corp\.?)\b/gi;

  // if there are a bunch of blank rows, update length to just the num of populated rows.
  Logger.log(dpDataA.length);

  for (var i = 0; i < dpDataA.length; i++) {
    var dpValueA = String(dpDataA[i][0]).trim();

    // Only process if there's a value in column A
    if (dpValueA !== "") {
      dpValueA = dpValueA.toLowerCase().replace(businessPattern, "").replace(/\s+/g, '');
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
