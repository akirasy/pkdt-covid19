/**
 * Move completed/done case to archive Sheet.
 * @param {Range} selectedRange Range to evaluate and move to archive.
 */
function moveToArchive(selectedRange) {
  let initialTime = new Date();
  let projectVar             = getProjectVariables();
  let headerKey              = getHeaderKey(projectVar.sheetKesPositif);
  let sheetKesPositif        = projectVar.sheetKesPositif;
  let sheetKesPositifArchive = projectVar.sheetKesPositifArchive;
  let selectedRowIndex       = selectedRange.getRowIndex();

  let conditionalArray = sheetKesPositif.getRange(selectedRowIndex, headerKey.gen_url+1, selectedRange.getNumRows(), 3).getValues();
  let selectedRowid = conditionalArray.map((item, index) => {
    if (item[0] != '' && item[1] == 'DONE' && item[2] == 'DONE') {
      let rowid = index + selectedRowIndex;
      return rowid
    };
  });
  selectedRowid.filter(item => item).forEach(rowid => {
    if (isEnoughTime(initialTime, 10)) {
      Logger.log('-- Move case to archive for rowid: ' + rowid)
      let doneCase = sheetKesPositif.getRange(rowid, 1, 1, sheetKesPositif.getLastColumn());
      let targetSheetArchive = sheetKesPositifArchive.getRange(sheetKesPositifArchive.getLastRow()+1, 1);
      doneCase.copyTo(targetSheetArchive);
      doneCase.clear();
      doneCase.setBackground('#cccccc');
    };
  });
}

/**
 * Grant permission for edit and view from requesting user at GoogleForm.
 */
function grantPermission() {
  let projectVar         = getProjectVariables();
  let activeSpreadsheet  = SpreadsheetApp.getActiveSpreadsheet();
  let archiveSpreadsheet = DriveApp.getFileById(projectVar.valueArchiveSpreadsheetId);
  let driveFolder        = DriveApp.getFolderById(projectVar.valueGeneratedFolderMain);

  // STEP 1 : Get current user
  let activeSpreadsheetUser  = activeSpreadsheet.getEditors().map(user => { return user.getEmail().toLowerCase() });
  let archiveSpreadsheetUser = archiveSpreadsheet.getViewers().map(user => { return user.getEmail().toLowerCase() });
  let driveFolderUser        = driveFolder.getEditors().map(user => { return user.getEmail().toLowerCase() });

  // STEP 2 : Get requesting user from GoogleForm
  let requestAccessForm = FormApp.openById(projectVar.valueReqAccessFormId);
  let formResponses = requestAccessForm.getResponses();
  let requestingUser = formResponses.map(item => {
    return item.getRespondentEmail().toLowerCase();
  });

  // STEP 3 : Grant permission
  requestingUser.forEach(user => {
    try {
      if (!activeSpreadsheetUser.includes(user)) {
        activeSpreadsheet.addEditor(user);
        Logger.log('Grant Active Spreadsheet edit access for: ' + user);
      };
      if (!archiveSpreadsheetUser.includes(user)) {
        archiveSpreadsheet.addViewer(user);
        Logger.log('Grant Archive Spreadsheet edit access for: ' + user);
      };
      if (!driveFolderUser.includes(user)) {
        driveFolder.addEditor(user);
        Logger.log('Grant Folder edit access for: ' + user);
      };
    } catch {
      Logger.log('Invalid email: ' + user);
    };
  });
}

/**
 * Action to move completed/done case to archive Sheet.
 */
function actionMoveToArchive() {
  let selectedRange = SpreadsheetApp.getActiveRange();
  moveToArchive(selectedRange);
}

/**
 * Prompt user for password before further execution. Returns `boolean`.
 */
function promptPassword() {
  Logger.log('Waiting for user input: Yes/No');
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt('Password protected command', 'Please enter password.', ui.ButtonSet.YES_NO);
  let password = '123qwe';

  // Process the user's response.
  let allowUsage;
  let correctPassword = response.getResponseText() == password;
  if (response.getSelectedButton() == ui.Button.YES && correctPassword) {
    allowUsage = true;
  } else if (response.getSelectedButton() == ui.Button.NO) {
    allowUsage = false;
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  } else {
    allowUsage = false;
    ui.alert('Access denied', 'Wrong password!', ui.ButtonSet.OK);
  }
  return allowUsage
}

function deleteGreyedRow() {
  let initialTime = new Date();
  let projectVar = getProjectVariables();
  let sheetKesPositif = projectVar.sheetKesPositif;
  let allRow = sheetKesPositif.getRange(1,1,sheetKesPositif.getLastRow());
  let greyedRowA1Notation = allRow.getBackgrounds().map((item, index) => {
    if (item[0] == '#cccccc') {
      let a1Notation = (index+1).toString() + ':' + (index+1).toString();
      return a1Notation
    };
  });
  greyedRowA1Notation.filter(item => item).reverse().forEach(item => {
    if (isEnoughTime(initialTime, 1)) {
      Logger.log('Delete rowid: ' + item);
      sheetKesPositif.getRange(item).deleteCells(SpreadsheetApp.Dimension.ROWS);
    };
  });
}
