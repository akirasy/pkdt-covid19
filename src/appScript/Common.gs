/**
 * Initialize set variable in `appscript.gs` sheet and return as `key:value` object/dictionary.
 */
function getProjectVariables() {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheetAppscript = activeSpreadsheet.getSheetByName('appScript.gs');
  let output = {
    // sheetName
    sheetAppscript         : sheetAppscript,
    sheetKesPositif        : activeSpreadsheet.getSheetByName(sheetAppscript.getRange('B4').getValue()),
    sheetKesPositifArchive : activeSpreadsheet.getSheetByName(sheetAppscript.getRange('B5').getValue()),
    sheetLaporanEpid       : activeSpreadsheet.getSheetByName(sheetAppscript.getRange('B6').getValue()),
    // googleDoc output
    valueClerkingTemplateId      : sheetAppscript.getRange('B10').getValue(),
    valueGeneratedFolderMain     : sheetAppscript.getRange('B11').getValue(),
    rangeGeneratedFolderToday    : sheetAppscript.getRange('B12'),
    rangeTodayDate               : sheetAppscript.getRange('B13'),
    // case registration
    valuePatientIdPrefix  : sheetAppscript.getRange('B17').getValue(),
    rangePatientIdCurrent : sheetAppscript.getRange('B18'),
    // permission
    valueReqAccessFormId      : sheetAppscript.getRange('B22').getValue(),
    valueThisSpreadsheetId    : sheetAppscript.getRange('B23').getValue(),
    valueArchiveSpreadsheetId : sheetAppscript.getRange('B24').getValue(),
  };
  return output
}

/**
 * Get header key index value and return as `name:index` object/dictionary.
 * @param {Object} sheetObj The spreadsheet object to read from.
 */
function getHeaderKey(sheetObj) {
  let output = new Object();
  let headerKey = sheetObj.getRange(1, 1, 1, sheetObj.getMaxColumns()).getValues();
  headerKey[0].forEach((item, index) => { output[item] = index });
  return output
}

/**
 * Check if still enough time to run another process. Appscript only allow 6 minutes of execution time.
 * @param {Date} initialTime Instance of `new Date()` from the initial execution time.
 * @param {Number} processDuration Estimated time (in seconds) for the process to complete
 */
function isEnoughTime(initialTime, processDuration) {
  let currentTime = new Date();
  let milisecondsDifference = currentTime.getTime() - initialTime.getTime();
  let secondsLeft = 360 - (milisecondsDifference / 1000);
  return secondsLeft > processDuration ? true : false;
}

/**
 * Check if user has allow permission to run this app.
 */
function aquireGooglePermission() {
  SpreadsheetApp.getUi().alert(
    'Success',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
