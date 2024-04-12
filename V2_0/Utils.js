/**
 * Fichiers contenant une multitude de fonctions utilitaires pour le reste des fonctions.
 */

function test() {
  console.log("fonctionne")
}

function getRangeExercises(sheet) {

}

function getRangeLoadData(sheet) {

}

function getCurrentMacro(date) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");

  var cellDate = findCellMonth(date, sheet);
  try {
    var macro = sheet.getRange(21, cellDate.getColumn()).getMergedRanges()[0].getValue()
    return macro;
  } catch (e) {throw new Error("Pas de macro pour la date.")}
}

function hexToRgb(hex) {
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}

function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1);
}

function findCellMonth(date, sheet) {
  const startDate = sheet.getRange("AK9:AO9").getMergedRanges()[0].getValues()[0][0];

  var dayBetween = Math.floor((date - startDate) / (24 * 60 * 60 * 1000));

  var col = 9 + Math.floor(dayBetween / 7);

  var row = findCellDate(date, sheet, col)

  if (row != null) {
    return sheet.getRange(row, col)
  }
}




function findCellDate(date, sheet, colonne) {
  const day = date.getDate();

  const range = sheet.getRange(13, colonne, 19 - 13 + 1, 1)
  const values = range.getValues();

  for (let row = 0; row < values.length; row++) {
    const cellValues = values[row][0]

    if (cellValues.getDate() == day) {
      return row + 13;
    }
  }

  return null;
}

function getCurrentWeekNo(date) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");
  var cell = findCellMonth(date, sheet);
  var column = cell.getColumn();

  var macroRange = sheet.getRange(21, cell.getColumn());
  var startColumn = macroRange.getMergedRanges()[0].getColumn();

  return column - startColumn + 1;
}