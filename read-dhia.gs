// Run this function when the spreadsheet is opened
function onOpen() {
  initMenu();
}


// Initialize the menu - add a menu item to read DIHA files
function initMenu(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Middle Intervale Macros');
  menu.addItem('Step 1: Read DHIA Files', 'readDHIA');
  
  menu.addToUi();
}


function deleteSheets(sheetNameToKeep) {
//  sheetNameToKeep = 'MasterSheet'
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var nsheets = sheets.length;
  for (i=0; i < nsheets; i++) {
    if (sheets[i].getSheetName() != sheetNameToKeep) {ss.deleteSheet(sheets[i])}
  }
}


// A function to read in DHIA files
function readDHIA() {
  SHEET_NAME_TO_KEEP = 'PreviousData';
  DHIA_FOLDER_ID = '1AdV9v9aNSmmSEpKivd87wGguSUnGZxWF';
  DATE_LINE_INDEX = 4;
  HEADER_LINE_INDEX = 6;
  
  FINAL_HEADER = [
    {varname: 'bname', dispname: 'BNAME', order: 1},
    {varname: 'lact', dispname: 'LACT', order: 2},
    {varname: 'pen', dispname: 'PEN', order: 3},
    {varname: 'dim', dispname: 'DIM', order: 4},
    {varname: 'milk', dispname: 'MILK', order: 5},
    {varname: 'pmilk', dispname: 'PMILK', order: 6},
    {varname: 'scc', dispname: 'SCC', order: 7},
    {varname: 'rpro', dispname: 'RPRO', order: 8},
    {varname: 'ltdm', dispname: 'LTDM', order: 9},
    {varname: 'x05me', dispname: '305ME', order: 10},
    {varname: 'brdat', dispname: 'BRDAT', order: 11},
    {varname: 'sid', dispname: 'SID', order: 12},
    {varname: 'tbrd', dispname: 'TBRD', order: 13},
    {varname: 'lbdat', dispname: 'LBDAT', order: 14},
    {varname: 'lsir', dispname: 'LSIR', order: 15},
    {varname: 'pscc', dispname: 'PSCC', order: 16},
    {varname: 'dry60', dispname: 'DRY60', order: 17},
    {varname: 'ddat', dispname: 'DDAT', order: 18},
    {varname: 'dueif', dispname: 'DUEIF', order: 19},
    {varname: 'fdat', dispname: 'FDAT', order: 20},
    {varname: 'calf', dispname: 'CALF', order: 21}
  ];
  //-----
  cumArray = [];
  
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
      cumArray = appendArray(cumArray, newArray);
    }
  }

  processedArray = makeObjectArray(cumArray);
  processedArray = fixSpecificData(processedArray);
  consolidatedArray = consolidateObjects(processedArray, 'bname');
  spreadsheetArray = makeSpreadsheetArray(consolidatedArray, FINAL_HEADER);
  
  var ss = SpreadsheetApp;
  var sht = ss.getActiveSpreadsheet().insertSheet('NewData');
  //Print the completed array to the spreadsheet
  sht.getRange(1, 1, spreadsheetArray.length, spreadsheetArray[1].length).setValues(spreadsheetArray);
}

function makeObjectArray(array) {
  var arr = [];
  var len = array.length;
  for (i=0; i<len; i++) {
    var row = array[i];

    arr.push({
      lact: row['LACT'] | row['L'],
      pen: row['PEN'],
      bname: row['BNAME'],
      dim: row['DIM'],
      milk: row['MILK'],
      pmilk: row['PMILK'],
      scc: row['SCC'],
      rpro: row['RPRO'],
      ltdm: row['LTDM'],
      x05me: row['305ME'],
      
      brdat: row['BRDAT'],
      sid: row['SID'],
      tbrd: row['TBRD'],
      lbdat: row['LBDAT'],
      lsir: row['LSIR'],
      
      pscc: row['PSCC'],
      dry60: row['DRY60'],
      ddat: row['DDAT'],
      
      dueif: row['DUEIF'],
      fdat: row['FDAT'],
      calf: row['CALF']
    });
  }
  return arr;
}

function makeSpreadsheetArray(objArray, headerConfig) {  
  // Determine the number of rows of data and the number of variables in the final header
  var nrows = objArray.length;
  var nvars = headerConfig.length;
  
  var output = [];
  // for each row + 1 (for the header) make an empty value in the output array
  for (i=0; i < nrows+1; i++) {output.push('')};
  
  //sort the header array by the 'order' property
  headerConfig.sort(function(a, b) {return a.order - b.order});
  
  // The first row of the output array will be the header to be displayed in the spreadsheet
  var header = headerConfig.map(a => a.dispname);
  output[0] = header;

  // Iterate through each row in the data
  for (i=0; i < nrows; i++) {
    var rowOutput = [];
    var currentRow = objArray[i];
    // Then iterate through each variable to be extracted
    for (j=0; j < nvars; j++) {
      var currentVar = headerConfig[j]['varname'];
      // Add each extracted variable to the row output array
      rowOutput.push(currentRow[currentVar]);
    }
    //Offset the row by 1 since the first row is the header
    output[i+1] = rowOutput;
  }
  return output;
}

function fixSpecificData(objArray) {
  //==================Fix lact values - replace 0 with the correct value=======================
  //Find all bnames
  var bnames = objArray.map(a => a.bname);
  //Returns an array of the unique values from the array of bnames above
  bnames = [...new Set(bnames)];
  //Iterate through each bname
  var nB = bnames.length;
  for (i=0; i<nB; i++) {
    var bn = bnames[i];
    var objFiltered = objArray.filter(a => {return a.bname == bn});
    //Find all lact values within the current bname
    var lacts = objFiltered.map(a => a.lact);
    var correctLact = Math.max(...lacts);
    objArray = updateData(objArray, bn, correctLact);
  }
  
  //Update each row in the object array that matches the bname, to give it the correct lact value
  function updateData(arr, bn, l) {
    var len = arr.length;
    for (j=0; j<len; j++) {
      if(arr[j].bname == bn) {arr[j].lact = l}
    }
    return arr;
  }
  //===============================================================================
  return objArray;
}

// Join 2 objects on a specified variable
function joinOn(obj1, obj2, idVar) {
  var output = {};
  //Check whether arrays each contain the id variable specified, and whether it is the same - if not, return an empty array
  if (obj1[idVar] == undefined || obj2[idVar] == undefined || obj1[idVar] != obj2[idVar]) {
    return output;
  // If the first object exists but not the 2nd, just use the 1st
  } else if (obj1.length != 0 && obj2.length == 0) {
    return obj1;
  // If the 2nd object exists but not the first, just use the 2nd
  } else if (obj2.length == 0 && obj2.length != 0) {
    return obj2;
  // If all required components exist (both objects, with matching id variables), proceed to join them
  } else {
    // Iterate through each key in the first object
    for (key in obj1) {
      if (isEmpty(obj1[key])) {
        if (!isEmpty(obj2[key])) {
          output[key] = obj2[key];
        }
      } else {
        output[key] = obj1[key];
      }
    }

    // Iterate through each key in the second object    
    for (key in obj2) {
      if (isEmpty(obj2[key])) {
        if (!isEmpty(obj1[key])) {
          output[key] = obj1[key];
        }
      } else {
        output[key] = obj2[key];
      }
    }
  }
  return output;
}


// Consolidates objects within an object array, so that there is only one object for each unique ID
// Allows the user to determine the object key that is the ID
function consolidateObjects(objArray, idVar) {
  var output = [];
  //Returns an array of all IDs in the object array
  var allIds = objArray.map(a => a[idVar]);
  //Returns an array of the unique values from the array of IDs above
  var allIds = [...new Set(allIds)];
  
  // Iteratre through each id and filter to find only those rows with matching ids
  var nIds = allIds.length;
  for (i=0; i<nIds; i++) { 
    var id = allIds[i];
    // Filter to find objects with matching id
    var objFiltered = objArray.filter(a => {return a[idVar] == id});
    // Iterate through each matched row (object)
    var nFiltered = objFiltered.length;
    var blob = objFiltered[0];
    var j = 0;
    while (j < nFiltered-1) {
      var next = objFiltered[j+1];
      blob = joinOn(blob, next, idVar);
      j++;
    }
    output.push(blob);
  }
  return output;
}


function appendArray(oldArray, newArray) {
  var nRowNew = newArray.length;
  
  for (i=0; i<nRowNew; i++) {
    oldArray.push(newArray[i]);
  }  
  return oldArray;
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
  var len = array.length;
  for (i=0; i < len; i++) {
    line = array[i]
		arrObj.push({});
		for (j=0; j < line.length; j++) {
			arrObj[arrObj.length-1][header[j]] = line[j];
		}
	};
	return arrObj;
}

function appendArray(oldArray, newArray) {
  var nRowNew = newArray.length;
  
  for (i=0; i < nRowNew; i++) {
    oldArray.push(newArray[i]);
  }
  return oldArray;
}


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

function isEmpty(x) {
    return (!x || x.length === 0);
}
