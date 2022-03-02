// Move completed case to archive
function moveToArchive(selected_range) {
  let var_source = getVarSource();

  let forloop_start = selected_range.getRowIndex();
  let forloop_end = forloop_start + selected_range.getNumRows();
  for (let rowid = forloop_start; rowid < forloop_end; rowid++) {
    // Check if all are done
    let patient_info = getPatientInfo(rowid, var_source);
    let status_siasatan_done = patient_info.status_siasatan[0] == 'DONE';
    let epid_daerah_done = patient_info.epid_daerah[0] == 'DONE';    

    // Move completed case to archive
    if (status_siasatan_done && epid_daerah_done) {
      moveCaseToArchive(rowid, var_source);
    } else { Logger.log('--- Skip rowid: ' + rowid); }
  }
}

function moveCaseToArchive(rowid, var_source) {
  Logger.log('-- Move case to archive for rowid: ' + rowid);
  let patient_row_range = var_source.sheet_kes_positif.getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns());
  let last_archive_range = var_source.sheet_kes_positif_archive.getRange(var_source.sheet_kes_positif_archive.getLastRow() + 1, 1);
  patient_row_range.copyTo(last_archive_range);
  patient_row_range.clear();
  patient_row_range.setBackground('#cccccc')
}

function addAllListedUser() {
  let var_source = getVarSource();
  let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let spreadsheet_archive = SpreadsheetApp.openById(var_source.spreadsheet_archive_id);
  let drive_folder = DriveApp.getFolderById(var_source.path_tlh_folder); 

  Logger.log('Get listed user in `User Access List` sheet');
  let sheet_uac = SpreadsheetApp.openById(var_source.spreadsheet_uac_id).getSheetByName('User Access List');
  let listed_editor_range = sheet_uac.getRange(2, 2, sheet_uac.getLastRow()).getValues();
  let listed_editor = listed_editor_range.map(user => {
    return user[0].toLowerCase().replace(/\s+/g, '');
  })
  listed_editor.pop(); // remove empty element at end of array (which always appear for unknown reason)
  Logger.log('--- Total user in list: ' + listed_editor.length);

  Logger.log('Processing user in awaiting list to grant permissions');
  Logger.log('--- Adding listed user to spreadsheet.');
  active_spreadsheet.addEditors(listed_editor);
  Logger.log('--- Adding listed user to spreadsheet archive.');
  spreadsheet_archive.addEditors(listed_editor);
  Logger.log('--- Adding listed user to Drive folder.');
  drive_folder.addEditors(listed_editor);
}

function addListedUser() {
  let var_source = getVarSource();
  let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let spreadsheet_archive = SpreadsheetApp.openById(var_source.spreadsheet_archive_id);
  let drive_folder = DriveApp.getFolderById(var_source.path_tlh_folder);

  Logger.log('Get actual user in spreadsheet.');
  let current_editor = active_spreadsheet.getEditors();
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
  Logger.log('--- User in list\n' + list_awaiting);

  if (list_awaiting.length != 0) {
    Logger.log('Processing user in awaiting list to grant permissions');
    Logger.log('--- Adding listed user to spreadsheet.');
    active_spreadsheet.addEditors(list_awaiting);
    Logger.log('--- Adding listed user to spreadsheet archive.');
    spreadsheet_archive.addEditors(list_awaiting);
    Logger.log('--- Adding listed user to Drive folder.');
    drive_folder.addEditors(list_awaiting);
  } else { Logger.log('--- No new user found') }
}

function removeUnlistedUser() {
  let var_source = getVarSource();
  let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let spreadsheet_archive = SpreadsheetApp.openById(var_source.spreadsheet_archive_id);
  let drive_folder = DriveApp.getFolderById(var_source.path_tlh_folder);

  Logger.log('Get actual user in spreadsheet.');
  let current_editor = active_spreadsheet.getEditors();
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
      active_spreadsheet.removeEditor(item);
      active_spreadsheet.removeViewer(item);
      spreadsheet_archive.removeEditor(item);
      spreadsheet_archive.removeViewer(item);
      drive_folder.removeViewer(item);
      drive_folder.removeEditor(item);
      Logger.log('--- User removed: ' + item);    
    })
  } else { Logger.log('--- No unlisted user found') }
}

// This removes all current user and adding them back
function refreshSharedUser() {
  let var_source = getVarSource();
  let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Remove user from spreadsheet
  Logger.log('Removing user from spreadsheet');
  let spreadsheet_editor = active_spreadsheet.getEditors();
  let spreadsheet_editor_list = spreadsheet_editor.map(user => { return user.getEmail() });
  spreadsheet_editor_list.forEach(user => {
    active_spreadsheet.removeEditor(user);
    active_spreadsheet.removeViewer(user);
    Logger.log('--- User removed: ' + user);
  })

  // Remove user from spreadsheet archive
  Logger.log('Removing user from spreadsheet archive');
  let spreadsheet_archive = SpreadsheetApp.openById(var_source.spreadsheet_archive_id);
  let spreadsheet_archive_editor = spreadsheet_archive.getEditors();
  let spreadsheet_archive_editor_list = spreadsheet_archive_editor.map(user => { return user.getEmail() });

  spreadsheet_archive_editor_list.forEach(user => {
    spreadsheet_archive.removeEditor(user);
    spreadsheet_archive.removeViewer(user);
    Logger.log('--- User removed: ' + user);
  })  

  // Remove user from drive folder
  Logger.log('Removing user from drive folder');
  let drive_folder = DriveApp.getFolderById(var_source.path_tlh_folder);
  let drive_editor = drive_folder.getEditors();
  let drive_editor_list = drive_editor.map(user => { return user.getEmail() });
  drive_editor_list.forEach(user => {
    drive_folder.removeEditor(user);
    Logger.log('--- User removed: ' + user);
  })

  // Add user from allowed list
  Logger.log('Adding allowed user to spreadsheet and drive folder')
  let sheet_uac = SpreadsheetApp.openById(var_source.spreadsheet_uac_id).getSheetByName('User Access List');
  let listed_editor = sheet_uac.getRange(2, 2, sheet_uac.getLastRow()).getValues();
  let listed_editor_list = listed_editor.map(user => { return user[0].toLowerCase().replace(/\s+/g, '') })
  listed_editor_list.pop(); // remove empty element at end of array (which always appear for unknown reason)
  active_spreadsheet.addEditors(listed_editor_list);
  spreadsheet_archive.addEditors(listed_editor_list);
  drive_folder.addEditors(listed_editor_list);
}
