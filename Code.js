/** 
 * @OnlyCurrentDoc Make macros available across G Suite domain
 */

/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Creates a new menu entry in the Sheet Add-On menu when Sheet is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the tigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Convert formulas to values (active tab only)','formulasToValuesActiveSheet')
      .addItem('Convert formulas to values (globally)','formulasToValuesGlobal')
      .addItem('Sort tabs','sortSheets')
      .addItem('Unhide rows and columns (active tab only)','unhideRowsColumnsActiveSheet')
      .addItem('Unhide rows and columns (globally)','unhideRowsColumnsGlobal')
      .addItem('Set all tab colors to red','setTabColor')
      .addItem('Reset all tab colors','resetTabColor')
      .addItem('Hide all tabs except active one','hideAllSheetsExceptActive')
      .addItem('Unhide all tabs','unhideAllSheets')
      .addItem('Reset filters','resetFilter')
      .addToUi();
}

/** 
 * Convert all formulas to values in the active sheet
 */ 
function formulasToValuesActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  range.copyValuesToRange(sheet, 1, range.getLastColumn(), 1, range.getLastRow());
};

/**
 * Convert all formulas to values in every sheet of the Google Sheet
 */
function formulasToValuesGlobal() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    var range = sheet.getDataRange();
    range.copyValuesToRange(sheet, 1, range.getLastColumn(), 1, range.getLastRow());
  });
};

/** 
 * Sort sheets alphabetically
 */
function sortSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var sheetNames = [];
  sheets.forEach(function(sheet,i) {
    sheetNames.push(sheet.getName());
  });
  sheetNames.sort().forEach(function(sheet,i) {
    spreadsheet.getSheetByName(sheet).activate();
    spreadsheet.moveActiveSheet(i + 1);
  });
};
    
/** 
 * Unhide all rows and columns in current Sheet data range
 */
function unhideRowsColumnsActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  sheet.unhideRow(range);
  sheet.unhideColumn(range);
}

/** 
 * Unhide all rows and columns in data ranges of entire Google Sheet
 */
function unhideRowsColumnsGlobal() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    var range = sheet.getDataRange();
    sheet.unhideRow(range);
    sheet.unhideColumn(range);
  });
};

/** 
 * Set all Sheets tabs to red
 */
function setTabColor() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    sheet.setTabColor("ff0000");
  });
};
  
/** 
 * Remove all Sheets tabs color
 */
function resetTabColor() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    sheet.setTabColor(null);
  });
};

/** 
 * Hide all sheets except the active one
 */
function hideAllSheetsExceptActive() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    if (sheet.getName() != SpreadsheetApp.getActiveSheet().getName()) 
      sheet.hideSheet();
  });
};

/** 
 * Unhide all sheets in Google Sheet
 */
function unhideAllSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    sheet.showSheet();
  });
};


/** 
 * Reset all filters for a data range on current Sheet
 */
function resetFilter() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  range.getFilter().remove();
  range.createFilter();
}
