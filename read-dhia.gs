function onOpen() {
  initMenu();
}


function initMenu(){

  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('My Macros');
  menu.addItem('Read DHIA Files', 'readDHIA');
  
  menu.addToUi();
}


function readDHIA() {
  DHIA_FOLDER_ID = '1AdV9v9aNSmmSEpKivd87wGguSUnGZxWF';
  START_LINE_INDEX = 6;
  
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1,1).setValue('Two');
  
  var dhiaFolder = DriveApp.getFolderById(DHIA_FOLDER_ID);
//  Logger.log(dhiaFolder.getName());
  var folders = dhiaFolder.getFolders();
  while (folders.hasNext()) {
    folder = folders.next();
    var files = folder.getFiles();
    while (files.hasNext()) {
      file = files.next();
//      Logger.log(file);
      
      str = file.getBlob().getDataAsString();
      parseText(str);
//      Logger.log(blob.getDataAsString())
    }
  }

}

function parseText(str) {
  var lines = str.split('\n');
  var headerLine = lines[START_LINE_INDEX]
  
  var widths = lines[START_LINE_INDEX + 1].split(' ').map(function(a) {return a.length});
  Logger.log(widths);
  
  var start = startIndex(widths, cumsum(widths));
  
  headers = parseLine(headerLine, start, widths);
  
  
  
  
  
  Logger.log(headers);
  Logger.log(lines);
  lines.forEach(findLength);
  
  function findLength(str) {Logger.log(str.trim().length)}
  
  
  
}

function parseLine(line, arrStart, arrWidths) {
  var arr = [];
  for (i=0; i < arrStart.length; i++) {
    arr[i] = line.substr(arrStart[i], arrWidths[i]).trim();
  }
  return arr;
}


function cumsum(arr) {
  var newArray = [];
  arr.reduce(function(a,b,i) {return newArray[i] = a + b;},0);
  return newArray;
}

function startIndex(arrLengths, arrCumsum) {
  var arr = [];
  for (i=0; i < arrLengths.length; i++) {
    arr[i] = arrCumsum[i] - arrLengths[i] + i;
  }
  return arr;
}



