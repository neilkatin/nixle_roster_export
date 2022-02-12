'use strict';

// Generate a sheet that can be uploaded to Nixle with people's contact info
//

function generateNixleCERT(e) {
  var ui = SpreadsheetApp.getUi();
  generateNixleCERTUi(e, ui);

}

function generateNixleCERTUi(e, ui) {

  var error;

  var nameFunc = function(row, titleMap) {
    // should probably do error checking here
    let lastCol = titleMap.get('Last');
    let firstCol = titleMap.get('First');

    if (lastCol !== undefined && firstCol !== undefined) {
      if (lastCol < row.length && firstCol < row.length) {
        return row[firstCol] + " " + row[lastCol];
      }
    }
    return null;
  };



  try {
    generateNixle(
      ui,
      'Nixle-CERT',
      'Active Roster',
      1,
      [ nameFunc, 'Home Ph', 'Mobile', 'Email', 'Zip'  ],
      function(row, titleMap) { return certFilter(row, titleMap, 'Cert') } );

    generateNixle(
      ui,
      'Nixle-Sups',
      'Active Roster',
      1,
      [ nameFunc, 'Home Ph', 'Mobile', 'Email', 'Zip'  ],
      function(row, titleMap) { return certFilter(row, titleMap, 'Sup') } );

    generateNixle(
      ui,
      'Nixle-Recon',
      'Active Roster',
      1,
      [ nameFunc, 'Home Ph', 'Mobile', 'Email', 'Zip'  ],
      function(row, titleMap) { return certFilter(row, titleMap, 'Recon') } );

  } catch (error) {
    ui.alert(error.message);
  }
}

function generateNixleECC(e) {

  var ui = SpreadsheetApp.getUi();
  var error;
  try {
    generateNixle(
      ui,
      'Nixle-ECC',
      'Current Roster',
      3,
      [ 'Name', 'Home Phone', 'Cell Phone', 'Email', null  ],
      eccFilter);

  } catch (error) {
    ui.alert(error.message);
  }
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  onOpenUi(ui);
}



function onOpenUi(ui) {

  var name = SpreadsheetApp.getActiveSpreadsheet().getName();

  Logger.log("Spreadsheet name '%s'", name);

  if (name.startsWith("ECC")) {
    ui.createMenu('Nixle').addItem("Generate Nixle Sheet", "generateNixleECC").addToUi();
  } else if (name.startsWith("CERT")) {
    ui.createMenu('Nixle').addItem("Generate Nixle Sheets", "generateNixleCERT").addToUi();
  } else {
    // do nothing
    Logger.log("No matches for file named '%s'", name);
  }

}

function generateNixle(ui, nixle_sheet_name, data_sheet_name, data_title_row, sheet_col_list, rowFilter) {

  Logger.log("generateNixle: called");

  // get the current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // see if there is an existing Nixle sheet
  var sheets = ss.getSheets();
  Logger.log("sheets length is " + sheets.length)
  var sheet = null;
  for(let i = 0; i < sheets.length; i++) {
    var trySheet = sheets[i];

    if (trySheet.getName() ===  nixle_sheet_name) {
      sheet = trySheet;
      break;
    }
  }

  if (sheet === null) {
    // no existing sheet: create it
    Logger.log("no existing sheet found; creating one");
    sheet = ss.insertSheet(nixle_sheet_name, sheets.length);
  } else {
    sheet.clear();
  }

  // Add the title row
  let titleData = [ "Full Name", "Home Phone Number (Voice)", "Mobile Phone Number (SMS, Voice)", "Email", "Zip" ];
  sheet.appendRow(titleData);


  var dataSheet = ss.getSheetByName(data_sheet_name);
  if (dataSheet == null) {
    throw new Error("Could not find data sheet (" + data_sheet_name + ")");
  }

  var titleMap = getTitleRow(dataSheet, data_title_row);

  //Logger.log("sheet getLastColumn %s getMaxColumns %s", dataSheet.getLastColumn(), dataSheet.getMaxColumns());
  
  var dataRange = dataSheet.getRange(data_title_row + 1, 1, dataSheet.getLastRow() - data_title_row, dataSheet.getLastColumn());

  // combine the SHEET_COL_MAP with the actual column indexes in the sheet
  var colList = [];
  for (let i = 0; i < sheet_col_list.length; i++) {
    let key = sheet_col_list[i];
    if (key === null) {
      colList[i] = null
    } else if (typeof key === 'function') {
      colList[i] = key;
    } else {
      if (titleMap.has(key) ) {
        colList[i] = titleMap.get(key);
      } else {
        throw new Error('could not find expected column "' + key + ' in list source sheet title row');
      }
    }
  }

  copyData(dataRange, sheet, colList, titleMap, rowFilter);
  //
  // for some reason the CERT sheet has rotated text; try and clear it
  let lastRow = sheet.getLastRow();
  let a1notation = Utilities.formatString("A1:E%d", lastRow);
  let allRangeList = sheet.getRangeList([a1notation]);
  allRangeList.setTextRotation(0);
}

function getTitleRow(sheet, titleRow) {
  // read the specified row from the spreadsheet; return a Map of column names to column numbers

  var titleRange = sheet.getRange(titleRow, 1, 1, sheet.getLastColumn());
  var values = titleRange.getValues();

  if (values.length < 1) {
    throw new Error('Could not parse roster sheet title row')
  }

  // iterate through the first (and only) row
  const row = values[0];
  const map = new Map() 
  for (var i = 0; i < row.length; i++) {
    let v = row[i]
    if (v !== "") {
      let col = i + 1   // change to origin 1 index
      //Logger.log("setting key %s to %d", v, col);
      map.set(v, i);
    }
  }
  return map;
}


/**
 * copy the data from the source data to the destination sheet.  titleMap shows the column labels; colList shows desired columns to source column indexes
 */
function copyData(dataRange, sheet, colList, titleMap, filter) {
  // generate a new map that goes directly to source column numbers
  var values = dataRange.getValues()
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    let row = values[rowIndex];
    let newRow = []

    // filter out unwanted rows
    if (! filter(row, titleMap)) continue;

    for (let i = 0; i < colList.length; i++) {
      let col = colList[i];

      if (col == null) continue;
      if (typeof col === 'function') {
        newRow[i] = col(row, titleMap);
      } else {
        if (col < row.length) {
          newRow[i] = row[col];
        }
      }

    }
    Logger.log("appending newRow %s", newRow);
    sheet.appendRow(newRow);
  }
}


// filter out un-wanted rows
function eccFilter(row, titleMap) {
  let val = commonFilter(row, titleMap, "ECC");
  return val === 1 || val === 2;
}

function certFilter(row, titleMap, colName) {
  let val = commonFilter(row, titleMap, colName);

  return val != null && val != "";
}

// filter out un-wanted rows
function commonFilter(row, titleMap, colName) {
  if ( ! titleMap.has(colName)) {
    throw new Error("Cannot find 'ECC' column in source sheet");
  }

  let colIndex = titleMap.get(colName)

  // row isn't long enough
  if (row.length < colIndex) {
    return null;
  }

  let colVal = row[colIndex];

  return colVal;
}
