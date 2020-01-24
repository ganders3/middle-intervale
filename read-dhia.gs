// Run this function when the spreadsheet is opened
function onOpen() {
  initMenu();
}


// Initialize the menu - add a menu item to read DIHA files
function initMenu(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Middle Intervale Macros');
  menu.addItem('Read DHIA Files', 'readDHIA');
  
  menu.addToUi();
}


function deleteSheets(sheetNameToKeep) {
//  sheetNameToKeep = 'MasterSheet'
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (i=0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != sheetNameToKeep) {
      ss.deleteSheet(sheets[i])
    }
  }
}


// A function to read in DHIA files
function readDHIA() {
  SHEET_NAME_TO_KEEP = 'MasterSheet';
  DHIA_FOLDER_ID = '1AdV9v9aNSmmSEpKivd87wGguSUnGZxWF';
  DATE_LINE_INDEX = 4;
  HEADER_LINE_INDEX = 6;
  
  masterArray = [];
  
  deleteSheets(SHEET_NAME_TO_KEEP);
  
  // Get the folder on Google Drive containing the DHIA folder
  var dhiaFolder = DriveApp.getFolderById(DHIA_FOLDER_ID);
  // Get each subfolder
  var folders = dhiaFolder.getFolders();
  // Search through each folder and the files within the folder
  while (folders.hasNext()) {
    folder = folders.next();
    var files = folder.getFiles();
    while (files.hasNext()) {
      file = files.next();
      str = file.getBlob().getDataAsString();
      newData = parseText(str);
      newArray = arrayToObjectArray(newData, true);
      Logger.log(newArray);
//      masterArray = appendArray(masterArray, newData);
    }
  }
}


function parseText(str) {  
  const LINES_TO_REMOVE = [
    'MIDDLE INTERVALE FARM',
    'Command',
    'Expanded',
    'Dairy One',
    'in PEN',
    'Avg',
    '[tT]otal',
    '^[\\s\\=\\_\\-\\n]+$',
    '^$',
    
    'Cows To Breed',
    'Cows To Calve',
    'To Dry Off',
    
    'Barn Name, 7 Characters',
    'Date of Last Breeding', 
    'Date to breed \\(60 days\\)',
    'Days in Milk',          
    'Dry date',
    'Dry Date for 60 Day Dry Period',
    'Due date',
    'Due Date if PG to Last Breeding',
    'Fresh \\(calving\\) date',
    'Lact\\. to Date Milk \\(Internal\\)',
    'Lactation number',
    'Last Calf Info',
    'Last sire used',
    'Last test day milk weight',
    'Last test raw somatic cell coun[t]*',
    'Milk \\@ Next2Last TestDate',
    'Pen or String number',
    'Projected 305 milk production',
    'Raw SCC \\(x1000\\) \\@Next2LastTest',
    'Repro code \\(FRESH,BRED,DRY etc\\)',
    'Sire ID',
    'Times bred'
  ];
  
  var re = new RegExp(LINES_TO_REMOVE.join('|'));
  var lines = str.split('\n');
             
  var dateLine = lines[DATE_LINE_INDEX]
  var date = dateLine.match(/[0-9]{1,2}\/\s*[0-9]{1,2}\/\s*[0-9]{1,2}/)
  var headerLine = lines[HEADER_LINE_INDEX]
  
  var widthsLine = lines[HEADER_LINE_INDEX + 1]  
  var widths = widthsLine.split(' ').map(function(a) {return a.length});
  var widthsLength = widthsLine.trim().length;
  
  var start = startIndices(widths, cumsum(widths));
  
  var header = parseLine(headerLine, start, widths);
  
  
  // Search lines for all patterns to remove
  var searchForLinesToRemove = lines.map(function(a) {return a.search(re)});
  // Keep all rows that do not match to a pattern to remove
  var keepRows = indicesOf(searchForLinesToRemove, -1);

  var dataArray = createDataArray(lines, header, start, widths, keepRows);
  
//  var masterArray = [];
//  var dataArray = combineArrays(dataArray)
//  Logger.log(dataArray);
  
  var ss = SpreadsheetApp;
  var sht = ss.getActiveSpreadsheet().insertSheet();
  sht.getRange(1, 1, dataArray.length, dataArray[1].length).setValues(dataArray);
  
  return dataArray;

  function createDataArray(lines, header, start, widths, keepRows) {
    var arr = [];
    var len = keepRows.length;
    var keep = keepRows.slice().reverse();
    for (var i=len-1; i>=0; i--) {
      var ind = keep[i];
      var data = parseLine(lines[ind], start, widths);
      arr[i] = data;
    }
    return arr.reverse();
  }
}


function arrayToObjectArray(array, containsHeader) {
	var header = [];
	if (containsHeader) {
		header = array[0];
		array.splice(0,1);
	} else {
		for (j=0; j < array[0].length; j++) {
			header.push('x' + j);
		}
	}
  
  var arrObj = [];
  for (i=0; i < array.length; i++) {
    line = array[i]
//  }
//	array.forEach((line) => {
		arrObj.push({});
		for (j=0; j < line.length; j++) {
			arrObj[arrObj.length-1][header[j]] = line[j];
		}
	};
	return arrObj;
}

//function appendArray(oldArray, newArray) {
//  for (i=0; i < newArray.length; i++) {
//    oldArray.push({
//      lact: newArray[i]['LACT'],
//      pen: newArray[i]['PEN']
//    
//    });
//  }
//
//
//}

// Parses a line into an array, using the start position and column widths
function parseLine(line, arrStart, arrWidths) {
  var arr = [];
  var len = arrStart.length;
  var start = arrStart.slice().reverse();
  var widths =  arrWidths.slice().reverse();
  for (var i=len-1; i>=0; i--) {
    arr[i] = line.substr(start[i], widths[i]).trim();
  }
  return arr.reverse();
}


// Find all indices containing a value
function indicesOf(array, find) {
  var indices = [];
  var len = array.length;
  for (var i=0; i < len; i++) {
    if (array[i] == find) {indices.push(i)}
  }
  return indices;
}


// Calculate cumulative sum of an array and return an array
function cumsum(arr) {
  var newArray = [];
  arr.reduce(function(a,b,i) {return newArray[i] = a + b;},0);
  return newArray;
}

// Find the starting indices of a fixed width file, using the column widths (arrLengths) and cumulative widths (arrCumsum)
function startIndices(arrLengths, arrCumsum) {
  var arr = [];
  var len = arrLengths.length;
  for (i=0; i < len; i++) {
    arr[i] = arrCumsum[i] - arrLengths[i] + i;
  }
  return arr;
}
