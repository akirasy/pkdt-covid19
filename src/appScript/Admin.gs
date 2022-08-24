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
  // Logger.log('----- Select row range: Pending');
  let patient_row_range = var_source.sheet_kes_positif.getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns());
  // Logger.log('-------- Select row range: Done');
  Logger.log('----- Choose last range at archive: Pending');
  let last_archive_range = var_source.sheet_kes_positif_archive.getRange(var_source.sheet_kes_positif_archive.getLastRow() + 1, 1);
  Logger.log('-------- Choose last range at archive: Done');
  // Logger.log('----- Copy to archive: Pending');
  patient_row_range.copyTo(last_archive_range);
  // Logger.log('-------- Copy to archive: Done');
  // Logger.log('----- Clearing row range: Pending');  
  patient_row_range.clear();
  // Logger.log('-------- Clearing row range: Done');
  // Logger.log('----- Set gray background: Pending');
  patient_row_range.setBackground('#cccccc')
  // Logger.log('-------- Set gray background: Done');
}

function addUserForm() {
  let var_source = getVarSource();
  let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let spreadsheet_archive = SpreadsheetApp.openById(var_source.spreadsheet_archive_id);
  let drive_folder = DriveApp.getFolderById(var_source.path_tlh_folder);

  // === Get actual user
  Logger.log('Get actual user');
  let active_spreadsheet_user = active_spreadsheet.getEditors().map(user => { return user.getEmail().toLowerCase() });
  let spreadsheet_archive_user = spreadsheet_archive.getViewers().map(user => { return user.getEmail().toLowerCase() });
  let drive_folder_user = drive_folder.getEditors().map(user => { return user.getEmail().toLowerCase() });

  // === Get requesting user
  let listed_editor = new Array();
  let request_access_form = FormApp.openById(var_source.request_access_form_id);
  let form_responses = request_access_form.getResponses();
  form_responses.forEach(item => {
    listed_editor.push(item.getRespondentEmail());
  })

  // === Filter requesting user
  Logger.log('Check if user is listed and then add to awaiting list');
  let list_awaiting = new Array();
  listed_editor.forEach(user => {
    if ( !active_spreadsheet_user.includes(user) || !spreadsheet_archive_user.includes(user) || !drive_folder_user.includes(user) ) {
      list_awaiting.push(user);
    }
  })
  Logger.log('--- Number of user in awaiting list: ' + list_awaiting.length);
  Logger.log('--- New user listed\n' + list_awaiting.toString());

  // === Add user with email validation
  list_awaiting.map(user => {
    try {
      Logger.log('Adding user: ' + user);
      active_spreadsheet.addEditor(user);
      spreadsheet_archive.addViewer(user);
      drive_folder.addEditor(user);
    } catch(e) {
      Logger.log('Invalid email: ' + user);
    }
  })
}
