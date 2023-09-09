function populateWordData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var dataValues = dataRange.getValues();
  
  // Add a header to column A if it doesn't exist
  if (dataValues[0][0] !== "Serial No") {
    sheet.getRange(1, 1).setValue("Serial No");
  }
  
  for (var i = 1; i < dataValues.length; i++) {
    var word = dataValues[i][1]; // Assuming word is in Column B (index 1)
    
    if (word) {
      // Fetch word data from the Dictionary API
      var wordData = getWordDataFromDictionaryAPI(word);
      
      // Update the corresponding columns
      sheet.getRange(i + 1, 3).setValue(wordData.meaning); // Column C
      sheet.getRange(i + 1, 4).setValue(wordData.example);  // Column D
      sheet.getRange(i + 1, 5).setValue(wordData.synonyms.join(', ')); // Column E
      sheet.getRange(i + 1, 6).setValue(wordData.antonyms.join(', ')); // Column F

      // Enable text wrapping for columns C, D, E, and F
      enableTextWrapping(sheet, i + 1, 3); // Column C
      enableTextWrapping(sheet, i + 1, 4); // Column D
      enableTextWrapping(sheet, i + 1, 5); // Column E
      enableTextWrapping(sheet, i + 1, 6); // Column F
      
      // Add serial numbers to column A
      sheet.getRange(i + 1, 1).setValue(i);
    } else {
      // Clear the columns if the word is empty
      sheet.getRange(i + 1, 1, 1, 6).clear(); // Columns A, B, C, D, E, F
    }
  }
}

function enableTextWrapping(sheet, row, col) {
  var cell = sheet.getRange(row, col);
  cell.setWrap(true);
}

function enableTextWrappingForColumnB() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numRows = sheet.getMaxRows();
  
  // Enable text wrapping for Column B (assuming it's Column 2)
  for (var row = 1; row <= numRows; row++) {
    var cell = sheet.getRange(row, 2); // Column B
    cell.setWrap(true);
  }
}

function enableTextWrappingForAllCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numRows = sheet.getMaxRows();
  var numCols = sheet.getMaxColumns();
  
  // Loop through all rows and columns to enable text wrapping
  for (var row = 1; row <= numRows; row++) {
    for (var col = 1; col <= numCols; col++) {
      var cell = sheet.getRange(row, col);
      cell.setWrap(true);
    }
  }
}

function getWordDataFromDictionaryAPI(word) {
  // Construct the API URL for the Dictionary API
  var apiUrl = 'https://api.dictionaryapi.dev/api/v2/entries/en/' + encodeURIComponent(word.toLowerCase());

  try {
    var response = UrlFetchApp.fetch(apiUrl);
    var data = response.getContentText();
    var wordData = parseDictionaryAPIResponse(data);

    return wordData;
  } catch (e) {
    throw new Error("Failed to fetch word data from the Dictionary API: " + e.message);
  }
}

function parseDictionaryAPIResponse(data) {
  try {
    var jsonData = JSON.parse(data);

    // Extract meaning, example usage, synonyms, and antonyms from the response
    var meanings = jsonData[0]?.meanings;
    if (meanings && meanings.length > 0) {
      var meaning = meanings[0]?.definitions[0]?.definition || '';
      var example = meanings[0]?.definitions[0]?.example || '';
      var synonyms = meanings[0]?.definitions[0]?.synonyms || [];
      var antonyms = meanings[0]?.definitions[0]?.antonyms || [];

      return {
        meaning: meaning,
        example: example,
        synonyms: synonyms,
        antonyms: antonyms
      };
    } else {
      throw new Error("Word data not found in the Dictionary API response.");
    }
  } catch (e) {
    throw new Error("Failed to parse Dictionary API response: " + e.message);

  }
}
