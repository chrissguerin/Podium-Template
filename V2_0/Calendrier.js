/**
 * Contiens toutes les fonctions reliés au calendrier.
 */

function updateCalendar(dateParam) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");

  const currentDate = dateParam;
  const yesterday = dateParam;

  //const currentDate = new Date();
  //const yesterday = new Date();
  yesterday.setDate(currentDate.getDate() - 1)

  var newDate = findCellMonth(currentDate, sheet);
  var oldDate = findCellMonth(yesterday, sheet)

  setStyleNewDate(newDate)
  if (oldDate != null) {
    if (oldDate.getRow() % 2 == 0) {
      setStyleOldDate(oldDate)
    } else {
      setStyleOldDateOdd(oldDate)
    }
  }
  resetMeetDay();
  if (oldDate != null) {
    insertCouleurDansDash(yesterday);
  } else {
    console.log(false);
  }
  insertMeetDay();
}

function insertMeetDay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");
  var range = sheet.getRange("AW38:BB40");

  for (var i = 0; i < range.getNumRows() - 1; i++) {
    var date = range.getMergedRanges()[i].getValue();
    var cell = findCellMonth(date, sheet);

    setStyleMeetDay(cell);
  }
}

function resetMeetDay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");
  var range = sheet.getRange("I13:BH19");
  var backgrounds = range.getBackgrounds();

  for (var row in backgrounds) {
    if (backgrounds[row].indexOf("#a61c00") != -1) {
      for (var column in backgrounds[row]) {
        if (backgrounds[row][column] == "#a61c00") {
          var cell = sheet.getRange(parseInt(row) + 13, parseInt(column) + 8 + 1);
          if (row % 2 == 1) {
            setStyleOldDate(cell);
          } else {
            setStyleOldDateOdd(cell);
          }
        }
      }
    }
  }
}

function setStyleMeetDay(cell) {
  cell.setBackground("#a61c00");
  cell.setFontColor("white")
  cell.setFontWeight("bold")
  cell.setFontStyle("normal")
  cell.setFontLine("normal")
}

function setStyleNewDate(cell) {
  cell.setBackground("#000000");
  cell.setFontColor("white")
  cell.setFontWeight("bold")
  cell.setFontStyle("italic")
  cell.setFontLine("underline")
}

function setStyleOldDate(cell) {
  cell.setBackground("#9fc5e8");
  cell.setFontColor("black")
  cell.setFontWeight("normal")
  cell.setFontStyle("normal")
  cell.setFontLine("none")
}

function setStyleOldDateOdd(cell) {
  cell.setBackground("#cfe2f3");
  cell.setFontColor("black")
  cell.setFontWeight("normal")
  cell.setFontStyle("normal")
  cell.setFontLine("none")
}
