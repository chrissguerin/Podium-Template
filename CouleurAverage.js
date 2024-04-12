function insertCouleurDansDash(yesterday) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getCurrentMacro(yesterday));
  var macroStartingDate = sheet.getRange("AJ23").getValue();

  var spacing_x = 19;
  var spacing_y = 23;
  var spacing_meso = 138;

  var weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  var dayOfTheWeek = weekdays[yesterday.getDay()];
  var dayBetween = Math.floor((yesterday - macroStartingDate) / (24 * 60 * 60 * 1000));
  var weeks = Math.floor(dayBetween / 7);
  var meso = getMeso(sheet, getCurrentWeekNo(yesterday));

  for (var i = 1; i < meso; i++) {
    weeks = weeks - sheet.getRange("AB" + (23 + spacing_meso * (i - 1))).getValue();
  }

  var mesoRange = getMesoRange(meso, sheet);
  var completedOnRange = mesoRange.getCell(22, 4);
  var fatigueRange = sheet.getRange(22 + mesoRange.getRow() - 1, 5 + mesoRange.getColumn() - 1, 1, 10);

  for (var i = 0; i < 6; i++) {
    var completedOnRange_offset = completedOnRange.offset(spacing_y * i, spacing_x * weeks);
    if (completedOnRange_offset.getValue() == dayOfTheWeek) {
      var fatigueRange_offset = fatigueRange.offset(spacing_y * i, spacing_x * weeks);

      var averageColor = getAverageColor(fatigueRange_offset);
      var cellDate = findCellMonth(yesterday, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD"));
      cellDate.setBackground(averageColor);
    }
  }
}

function getAverageColor(range) {
  var totalR = 0;
  var totalG = 0;
  var totalB = 0;
  var count = 0;
  for (var col = 1; col <= range.getNumColumns(); col++) {
    var cell = range.offset(0, col - 1).getMergedRanges()[0].getValues()[0][0];
    var backgroundColor = getColorForCellValue(cell);

    if (backgroundColor !== "#0a53a8" && backgroundColor !== "") {
      var color = hexToRgb(backgroundColor);
      totalR += color.r;
      totalG += color.g;
      totalB += color.b;
      count++;
    }
  }

  var avgR = Math.round(totalR / count);
  var avgG = Math.round(totalG / count);
  var avgB = Math.round(totalB / count);

  return rgbToHex(avgR, avgG, avgB);
}

function getColorForCellValue(cellValue) {
  cellValue = cellValue.toLowerCase().replace(/ - /g, '-').replace("!", "").split("-")[0];

  if (cellValue == "fill this") {
    return "#0a53a8";
  } else if (cellValue == "fairly sore" || cellValue == "tired") {
    return "#ffe5a0";
  } else if (cellValue == "very sore" || cellValue == "exhausted") {
    return "#b10202";
  } else if (cellValue == "ready" || cellValue == "not sore") {
    return "#d4edbc";
  }
}

