function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('COVID19')
  .addItem('Set Google permission', 'aquireGooglePermission')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Formatting')
    .addItem('To upperCase', 'toUpperCase')
    .addItem('Trim Whitespaces', 'trimWhitespace')
    .addItem('Clean IC', 'cleanIc')
    .addItem('Set formatting & validation', 'setValidationAndFormatting')
  )
  .addSeparator()
  .addSubMenu(ui.createMenu('Peg Penyiasat')
    .addItem('Get info segera', 'menuInfoSegeraPenyiasat')
    .addItem('Generate borang siasatan', 'menuGenerateBorangSiasatan')
  )
  .addSubMenu(ui.createMenu('Peg Epid Daerah')
    .addItem('Kes Epid selesai', 'actionGenerateLaporanEpid')
    .addItem('Move selection to archive', 'actionMoveToArchive')
  )
  .addSeparator()
  .addSubMenu(ui.createMenu('Administrator')
    .addSubMenu(ui.createMenu('Trigger')
      .addItem('Borang Siasatan', 'menuTriggerGenerateBorangSiasatan')
      .addItem('Add listed user', 'menuTriggerGrantPermission')
      .addItem('Move to archive', 'triggerMoveToArchive')
      .addItem('Delete greyed row', 'deleteGreyedRow')
    )
    .addItem('Add listed user access', 'menuGrantPermission')
  )
  .addSubMenu(ui.createMenu('About')
    .addItem('Author', 'aboutAuthor')
    .addItem('License', 'aboutLicense')
    .addItem('Google AppScript', 'aboutGoogleAppScript')
  )
  .addToUi();
}

/**
 * UserMenu action to generate borang siasatan on single row/patient.
 */
function menuGenerateBorangSiasatan() {
  if (promptPassword()) {
    let projectVar = getProjectVariables();
    let headerKey = getHeaderKey(projectVar.sheetKesPositif);
    let rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
    generateBorangSiasatan(projectVar, headerKey, rowid);
  };
}

/**
 * UserMenu action to grant permission from listed GoogleForm. Password required.
 */
function menuGrantPermission() {
  if (promptPassword()) { grantPermission(); };
}

/**
 * UserMenu action to manually trigger grant permission/access. Password required.
 */
function menuTriggerGrantPermission() {
  if (promptPassword()) { grantPermission(); };
}

/**
 * UserMenu action to manually trigger generate borang siasatan. Password required.
 */
function menuTriggerGenerateBorangSiasatan() {
  if (promptPassword()) { triggerGenerateBorangSiasatan(); };
}

/**
 * UserMenu action to show infoSegera of single row/patient.
 */
function menuInfoSegeraPenyiasat() {
  let projectVar =  getProjectVariables();
  let headerKey = getHeaderKey(projectVar.sheetKesPositif);
  let ui = SpreadsheetApp.getUi();
  let rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  let patientInfo = projectVar.sheetKesPositif.getRange(rowid, 1, 1, projectVar.sheetKesPositif.getMaxColumns()).getValues()[0];
  ui.alert(infoSegeraPenyiasat(headerKey, patientInfo), ui.ButtonSet.OK);
}
