function addListedUser() {
  // Get listed user in `User Access List` sheet
  Logger.log('Get listed user in `User Access List` sheet')
  // var sheet_uac = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User Access List');
  var sheet_uac = SpreadsheetApp.openById(getVarSource().spreadsheet_uac_id).getSheetByName('User Access List');
  var listed_editor_range = sheet_uac.getRange(2, 2, sheet_uac.getLastRow()).getValues();  
  var listed_editor = new Array();
  for (let i = 0; i < listed_editor_range.length; i++) {
    listed_editor.push(listed_editor_range[i][0]);
  }

  // Add listed user
  // Logger.log('Adding listed user to spreadsheet. (Debug mode)')
  // for ( let i = 0; i < listed_editor.length; i++ ) {
  //   SpreadsheetApp.getActiveSpreadsheet().addEditor(listed_editor[i]);
  //   SpreadsheetApp.getActiveSpreadsheet().addEditor(listed_editor[i]);
  //   Logger.log('User added: ' + listed_editor[i]);
  // }
  // Logger.log('Done!')

  Logger.log('Adding listed user to spreadsheet. (Quick mode)')
  SpreadsheetApp.getActiveSpreadsheet().addEditors(listed_editor.filter(e =>  e));
  Logger.log('Done!')
}

function removeUnlistedUser() {
  // Get actual user in spreadsheet
  Logger.log('Get actual user in spreadsheet.')
  var current_editor = SpreadsheetApp.getActiveSpreadsheet().getEditors();

  // Get listed user in `User Access List` sheet
  Logger.log('Get listed user in `User Access List` sheet')
  // var sheet_uac = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User Access List');
  var sheet_uac = SpreadsheetApp.openById(getVarSource().spreadsheet_uac_id).getSheetByName('User Access List');
  var listed_editor_range = sheet_uac.getRange(2, 2, sheet_uac.getLastRow()).getValues();  
  var listed_editor = new Array();
  for (let i = 0; i < listed_editor_range.length; i++) {
    listed_editor.push(listed_editor_range[i][0]);
  }

  // Check if user is not listed and then add to removal list
  Logger.log('Check if user is not listed and then add to removal list')
  var list_remove = new Array();
  for ( let i = 0; i < current_editor.length; i++ ) {
    if (listed_editor.includes(current_editor[i].getEmail()) == false) {
      list_remove.push(current_editor[i]);
    }
  }

  // Remove user
  Logger.log('Removing unlisted user from spreadsheet.')
  for ( let i = 0; i < list_remove.length; i++ ) {
    SpreadsheetApp.getActiveSpreadsheet().removeEditor(list_remove[i]);
    SpreadsheetApp.getActiveSpreadsheet().removeViewer(list_remove[i]);
    Logger.log('User removed: ' + list_remove[i]);
  }
  Logger.log('Done!')
}
