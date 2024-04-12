function getCurrentMacro(date) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");

  var cellDate = findCellMonth(date, sheet);
  try {
  var macro = sheet.getRange(21, cellDate.getColumn()).getMergedRanges()[0].getValue()
  return macro;
  } catch (e) {throw new Error("Pas de macro pour la date.")}
}

