function insertCouleurDansDash(yesterday) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getCurrentMacro(yesterday));
  var macroStartingDate = sheet.getRange("AL21").getValue();

  var spacing_x = 14;

  var weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  var dayOfTheWeek = weekdays[yesterday.getDay()];
  var dayBetween = Math.floor((yesterday - macroStartingDate) / (24 * 60 * 60 * 1000));
  var weeks = Math.floor(dayBetween / 7);

  //for (var i = 1; i < meso; i++) {
    //weeks = weeks - sheet.getRange("AB" + (23 + spacing_meso * (i - 1))).getValue();
  //}

  //var mesoRange = getMesoRange(meso, sheet);

  var rangeCompletedOn = sheet.getRange("AP27:AP200"); //range de la premiere semaine
  var rangeCompletedOn_Offset_Weeks = rangeCompletedOn.offset(0, spacing_x * weeks); //change le range pour la bonne semaine

  console.log(rangeCompletedOn_Offset_Weeks.getA1Notation());

  var results = rangeCompletedOn_Offset_Weeks.createTextFinder("COMPLETED ON").findAll(); //trouve toutes les cellules avec "completed on" (a la bonne semaine)

  for (var i = 0; i < results; i++) { 
    var resultRow = results[i].getRow();
    var rangeCompletedOn_Offset_Days = sheet.getRange(resultRow + 1, rangeCompletedOn_Offset_Weeks.getColumn());


    if (rangeCompletedOn_Offset_Days.getValue() == dayOfTheWeek) {
      var fatigueRange = sheet.getRange(resultRow + 1, rangeCompletedOn_Offset_Weeks.getColumn() + 2, 1, 9)

      var averageColor = getAverageColor(fatigueRange);
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

