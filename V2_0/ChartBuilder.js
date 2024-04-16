function updateCharts(dateParam) {
  var date = dateParam;
  //var date = new Date();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getCurrentMacro(date));
  var liftGroupArrays = ["sq", "bn", "dl"];

  var sqRange = sheet.getRange(28, 8, 2, 18);
  var bnRange = sheet.getRange(32, 8, 2, 18);
  var dlRange = sheet.getRange(36, 8, 2, 18);
  var volRange = sheet.getRange(44, 8, 2, 18);

  var firstColumn = "AX";
  var spacing = 14;

  var firstColumnIndex = sheet.getRange(firstColumn + "1").getColumn();

  var weeksInMeso = sheet.getRange("AL21").getValue();

  for (var columnIndex = firstColumnIndex, j = 0; j < weeksInMeso; columnIndex += spacing, j++) { //loop pour parcourir toutes les semaines
    var intentionRange = sheet.getRange("AJ27:AJ200");
    var range = sheet.getRange(columnIndex, 27, columnIndex, 200);

    var weeklyVolume = getVolumeTotal(sheet, range);

    var results = intentionRange.createTextFinder("main").findAll();

    for (var i = 0; i < results.length; i++) { //loop pour parcourir tous les "main" trouve dans AJ
      var result = results[i];
      var rowIndex = result.getRow();

      var range = sheet.getRange(rowIndex, columnIndex)
      var value = range.getValue();

      var liftGroupRange = sheet.getRange(rowIndex, columnIndex - 15);
      var liftGroupValue = liftGroupRange.getValue();

      var week = sheet.getRange(21, columnIndex - 11).getValue();
      var weekNo = parseInt(week.toString().charAt(week.toString().length - 1));

      week = "WEEK " + weekNo;

      if (value != null) {
        if (!isNaN(value)) {
          if (liftGroupArrays.includes(liftGroupValue)) {
            if (liftGroupValue == "sq") {
              var weekRange = findWeekRange(sqRange, week);
              sheet.getRange(weekRange.getRow(), weekRange.getColumn()).setValue(value);
            } else if (liftGroupValue == "bn") {
              var weekRange = findWeekRange(bnRange, week);
              sheet.getRange(weekRange.getRow(), weekRange.getColumn()).setValue(value);
            } else if (liftGroupValue == "dl") {
              var weekRange = findWeekRange(dlRange, week);
              sheet.getRange(weekRange.getRow(), weekRange.getColumn()).setValue(value);
            }
          }
        }
      }
    }
    var weekRange = findWeekRange(volRange, week);
    weekRange.setValue(weeklyVolume);
  }
}

function findWeekRange(liftGroupRange, week) {
  var values = liftGroupRange.getValues();
  for (var i = 0; i < values[0].length; i++) {
    if (values[0][i] === week) {
      return liftGroupRange.offset(1, i, 1, 1);
    }
  }
  return null;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function getVolumeTotal(sheet, range) {
  var total = 0;

  var values = sheet.getRange(range).getValues();
  
  for (var i = 0; i < values.length; i++) {
    var cellValue = values[i][0];
    if (!isNaN(cellValue)) { 
      sum += Number(cellValue);
    }
  }

  return total;
}