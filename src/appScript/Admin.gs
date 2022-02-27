function addListedUser() {
  let var_source = getVarSource();
  Logger.log('Get actual user in spreadsheet.');
  let current_editor = SpreadsheetApp.getActiveSpreadsheet().getEditors();
  let current_editor_list = current_editor.map(user => { return user.getEmail() });
  Logger.log('--- Total user in spreadsheet: ' + (current_editor_list.length - 1)); // minus one because owner is included as actual user

  Logger.log('Get listed user in `User Access List` sheet');
  let sheet_uac = SpreadsheetApp.openById(var_source.spreadsheet_uac_id).getSheetByName('User Access List');
  let listed_editor_range = sheet_uac.getRange(2, 2, sheet_uac.getLastRow()).getValues();
  let listed_editor = listed_editor_range.map(user => {
    return user[0].toLowerCase().replace(/\s+/g, '');
  })
  listed_editor.pop(); // remove empty element at end of array (which always appear for unknown reason)
  Logger.log('--- Total user in list: ' + listed_editor.length);

  Logger.log('Check if user is listed and then add to awaiting list');
  let list_awaiting = new Array();
  listed_editor.forEach(user => {
    if (!(current_editor_list.includes(user))) {
      list_awaiting.push(user);
    }
  })
  Logger.log('--- Number of user in awaiting list: ' + list_awaiting.length);

  if (list_awaiting.length != 0) {
    Logger.log('Processing user in awaiting list to grant permissions');
    Logger.log('--- Adding listed user to spreadsheet. (Quick mode)');
    SpreadsheetApp.getActiveSpreadsheet().addEditors(list_awaiting);
    Logger.log('--- Adding listed user to Drive folder. (Quick mode)');
    DriveApp.getFolderById(var_source.path_tlh_folder).addEditors(list_awaiting);
  } else { Logger.log('--- No new user found') }
}

function removeOneUser() {
  let gmail_address = 'kerdautemerloh2011@gmail.com';
  SpreadsheetApp.getActiveSpreadsheet().removeEditor(gmail_address);
  SpreadsheetApp.getActiveSpreadsheet().removeViewer(gmail_address);
  DriveApp.getFolderById(getVarSource().path_tlh_folder).removeViewer(gmail_address);
  DriveApp.getFolderById(getVarSource().path_tlh_folder).removeEditor(gmail_address);
}

function removeAllUser() {
  let var_source = getVarSource();
  let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let drive_folder = DriveApp.getFolderById(var_source.path_tlh_folder);

  let current_editor = active_spreadsheet.getEditors();
  let current_editor_list = current_editor.map(user => { return user.getEmail() });
  current_editor_list.forEach(user => {
    active_spreadsheet.removeEditor(user);
    active_spreadsheet.removeViewer(user);
    drive_folder.removeEditor(user);
    Logger.log('User removed: ' + user);
  })
}

function removeUnlistedUser() {
  let var_source = getVarSource();
  Logger.log('Get actual user in spreadsheet.');
  let current_editor = SpreadsheetApp.getActiveSpreadsheet().getEditors();
  let current_editor_list = current_editor.map(user => { return user.getEmail() });
  Logger.log('--- Total user in spreadsheet: ' + (current_editor_list.length - 1)); // minus one because owner is included as actual user

  Logger.log('Get listed user in `User Access List` sheet');
  let sheet_uac = SpreadsheetApp.openById(var_source.spreadsheet_uac_id).getSheetByName('User Access List');
  let listed_editor_range = sheet_uac.getRange(2, 2, sheet_uac.getLastRow()).getValues();
  let listed_editor = listed_editor_range.map(user => {
    return user[0].toLowerCase().replace(/\s+/g, '');
  })
  listed_editor.pop(); // remove empty element at end of array (which always appear for unknown reason)
  Logger.log('--- Total user in list: ' + listed_editor.length);

  Logger.log('Check if user is listed and then add to awaiting list');
  let list_removal = new Array();
  current_editor.forEach(user => {
    if (listed_editor.includes(user)) {
      list_removal.push(user);
    }
  })
  Logger.log('--- Number of user in removal list: ' + list_removal.length);

  if (list_removal.length != 0) {
    Logger.log('Removing unlisted user from spreadsheet.');
    list_removal.forEach(item => {
      SpreadsheetApp.getActiveSpreadsheet().removeEditor(item);
      SpreadsheetApp.getActiveSpreadsheet().removeViewer(item);
      DriveApp.getFolderById(var_source.path_tlh_folder).removeViewer(item);
      DriveApp.getFolderById(var_source.path_tlh_folder).removeEditor(item);
      Logger.log('--- User removed: ' + item);    
    })
  } else { Logger.log('--- No unlisted user found') }
}

// Move completed case to archive
function moveToArchive(selected_range) {
  let var_source = getVarSource();

  let forloop_start = selected_range.getRowIndex();
  let forloop_end = forloop_start + selected_range.getNumRows();
  for (let rowid = forloop_start; rowid < forloop_end; rowid++) {
    // Check if all are done
    let patient_info = getPatientInfo(rowid, var_source);
    let reten_epid = patient_info.reten_epid[0];
    let siasatan_status = patient_info.siasatan_status[0];
    // Set conditional value
    let isAllDone = Boolean();
    if (reten_epid == 'DONE' && siasatan_status == 'DONE') {
      isAllDone = true;
    } else { isAllDone = false }    

    // Move completed case to archive
    if (isAllDone) {
      Logger.log('-- Move case to archive for rowid: ' + rowid);
      let selected_row_range = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, SpreadsheetApp.getActiveSheet().getMaxColumns());
      let last_archive_range = var_source.sheet_kes_positif_archive.getRange(var_source.sheet_kes_positif_archive.getLastRow() + 1, 1);
      selected_row_range.copyTo(last_archive_range);
      selected_row_range.clear();
      // Mark with color gray
      selected_row_range.setBackground('#cccccc');
    } else { Logger.log('--- Skip rowid: ' + rowid); }
  }
}
