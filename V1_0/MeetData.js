function getCsvData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");
  var lifterName = removeAccents(sheet.getRange("Q9").getMergedRanges()[0].getValue().toString().toLowerCase().replace(" ", "").replace("-", ""));
  var url = "https://www.openipf.org/api/liftercsv/" + lifterName;
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() == 200){
    return csvData = response.getContentText();
  }
}

function removeAccents(inputString) {
  var accentMap = {
    'À': 'A',
    'Á': 'A',
    'Â': 'A',
    'Ã': 'A',
    'Ä': 'A',
    'Å': 'A',
    'à': 'a',
    'á': 'a',
    'â': 'a',
    'ã': 'a',
    'ä': 'a',
    'å': 'a',
    'È': 'E',
    'É': 'E',
    'Ê': 'E',
    'Ë': 'E',
    'è': 'e',
    'é': 'e',
    'ê': 'e',
    'ë': 'e',
    'Ì': 'I',
    'Í': 'I',
    'Î': 'I',
    'Ï': 'I',
    'ì': 'i',
    'í': 'i',
    'î': 'i',
    'ï': 'i',
    'Ò': 'O',
    'Ó': 'O',
    'Ô': 'O',
    'Õ': 'O',
    'Ö': 'O',
    'Ø': 'O',
    'ò': 'o',
    'ó': 'o',
    'ô': 'o',
    'õ': 'o',
    'ö': 'o',
    'ø': 'o',
    'Ù': 'U',
    'Ú': 'U',
    'Û': 'U',
    'Ü': 'U',
    'ù': 'u',
    'ú': 'u',
    'û': 'u',
    'ü': 'u',
    'Ç': 'C',
    'ç': 'c',
    'Ñ': 'N',
    'ñ': 'n',
    'ß': 'ss'
  };

  return inputString.replace(/[À-ÖØ-öø-ÿ]/g, function(match) {
    return accentMap[match] || match;
  });
}

function filterCSVData(csvData) {
  var filteredData = [];
  
  // Define the indices of the columns you want to keep
  var divisionIndex = 7;
  var bodyweightIndex = 8;
  var bestSquatIndex = 14;
  var bestBenchIndex = 19;
  var bestDeadliftIndex = 24;
  var totalIndex = 25;
  var meetNameIndex = 40;
  var dateIndex = 36;
  var goodliftIndex = 30;
  
  for (var i = 1; i < csvData.length; i++) {
    var row = csvData[i];
    var filteredRow = [
      row[dateIndex],
      row[meetNameIndex],
      row[divisionIndex],
      row[bodyweightIndex],
      row[bestSquatIndex],
      row[bestBenchIndex],
      row[bestDeadliftIndex],
      row[totalIndex],
      row[goodliftIndex]
    ];
    filteredData.push(filteredRow);
  }
  
  return filteredData;
}

function insertData(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COACH DASHBOARD");
  var data = filterCSVData(Utilities.parseCsv(getCsvData()));

  var colorOne = "#cfe2f3";
  var colorTwo = "#9fc5e8";
  var white = "#ffffff";
  var black = "#000000";

  for (var i = 0; i < data.length; i++){
    var row = 33 + i;
    sheet.getRange("H" + row).setValue(data[i][0]);
    sheet.getRange("I" + row + ":M" + row).merge().setValue(data[i][1]);
    sheet.getRange("N" + row + ":P" + row).merge().setValue(data[i][2]);
    sheet.getRange("Q" + row + ":S" + row).merge().setValue(data[i][3]);
    sheet.getRange("T" + row + ":W" + row).merge().setValue(data[i][4]);
    sheet.getRange("X" + row + ":AA" + row).merge().setValue(data[i][5]);
    sheet.getRange("AB" + row + ":AE" + row).merge().setValue(data[i][6]);
    sheet.getRange("AF" + row + ":AI" + row).merge().setValue(data[i][7]);
    sheet.getRange("AJ" + row + ":AK" + row).merge().setValue(data[i][8]);
    
    if (i % 2 == 0){
      sheet.getRange("H" + row + ":" + "AK" + row).setBackground(colorOne);
    } else {
      sheet.getRange("H" + row + ":" + "AK" + row).setBackground(colorTwo);
    }

    var range = sheet.getRange("H33:AK" + row);
    range.setBorder(true, true, true, true, true, true, white, SpreadsheetApp.BorderStyle.SOLID).setBorder(true, true, true, true, null, null, black, SpreadsheetApp.BorderStyle.SOLID_THICK).setFontFamily("Rajdhani").setHorizontalAlignment("center").setFontSize(10);
  }
}
