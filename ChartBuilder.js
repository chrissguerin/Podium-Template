function copyData() {
  var date = new Date();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getCurrentMacro(date));
  var liftGroupArrays = ["sq", "bn", "dl"];
  var liftArrays = ["Comp Squat", "Comp Bench", "Comp Deadlift", "Comp Dead", "Comp bench", "Comp squat", "Comp deadlift", "Comp dead"];

  var sqRange = sheet.getRange(28, 8, 2, 18);
  var bnRange = sheet.getRange(32, 8, 2, 18);
  var dlRange = sheet.getRange(36, 8, 2, 18);
  var volRange = sheet.getRange(44, 8, 2, 18);

  var firstColumn = "AV";
  var spacing = 19;
  var spacing_meso = 138; 

  var firstColumnIndex = sheet.getRange(firstColumn + "1").getColumn();

  var meso = getMeso(sheet, getCurrentWeekNo(date));
  var mesoRange = getMesoRange(meso, sheet);
  var weeksInMeso = sheet.getRange("AB" + (23 + (spacing_meso * (meso - 1)))).getValue();

  for (var columnIndex = firstColumnIndex, j = 0; j < weeksInMeso; columnIndex += spacing, j++) {
    var intentionRange = sheet.getRange(mesoRange.getRow(), columnIndex - 14, 138, 1);

    var weeklyVolume = getVolumeTotal(sheet, mesoRange, columnIndex);

    var results = intentionRange.createTextFinder("main").findAll();

    for (var i = 0; i < results.length; i++) {
      var result = results[i];
      var rowIndex = result.getRow();

      var range = sheet.getRange(rowIndex, columnIndex)
      var value = range.getValue();

      var liftGroupRange = sheet.getRange(rowIndex, columnIndex - 15);
      var liftGroupValue = liftGroupRange.getValue();

      var liftRange = sheet.getRange(rowIndex, columnIndex - 12);
      var liftValue = liftRange.getValue();

      var week = sheet.getRange(21, columnIndex - 11).getValue();
      var weekNo = parseInt(week.toString().charAt(week.toString().length - 1));

      var counter = 1;
      while (counter < meso){
        var weeksInMeso_previous = sheet.getRange("AB" + (23 + (spacing_meso * (counter - 1)))).getValue();
        weekNo += parseInt(weeksInMeso_previous);
        counter++;
      }

      week = "WEEK " + weekNo;

      if (value != null) {
        if (!isNaN(value)) {
          if (liftGroupArrays.includes(liftGroupValue)) {
            if (liftArrays.includes(liftValue)) {
              if (liftGroupValue == "sq") {
                var weekRange = findWeekRange(sqRange, week);
                sheet.getRange(weekRange.getRow() ,weekRange.getColumn()).setValue(value);
              } else if (liftGroupValue == "bn") {
                var weekRange = findWeekRange(bnRange, week);
                sheet.getRange(weekRange.getRow() ,weekRange.getColumn()).setValue(value);
              } else if (liftGroupValue == "dl") {
                var weekRange = findWeekRange(dlRange, week);
                sheet.getRange(weekRange.getRow() ,weekRange.getColumn()).setValue(value);
              }
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

function getVolumeTotal(sheet, mesoRange, columnIndex) {
  var total = 0;
  for (var i = 0; i < 6; i++){
      var range = sheet.getRange(mesoRange.getRow() + 4 + (23 * i), columnIndex + 1, 9, 1);
      var values = range.getValues();

      for (var j = 0; j < values.length; j++) {
        for (var k = 0; k < values[j].length; k++) {
          total += parseFloat(values[j][k]);
    }
  }
  }
  return total;
}