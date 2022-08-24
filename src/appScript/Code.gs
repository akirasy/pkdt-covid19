// Create topbar menu
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('COVID19 PKDT')
  .addItem('🕸 Set Google permission', 'aquireGooglePermission')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕸 Formatting')
    .addItem('⚪ To upperCase', 'toUpperCase')
    .addItem('⚪ To oneLine', 'toOneLine')
    .addItem('⚪ Clean IC', 'cleanIc')
    .addItem('⚪ Set formatting & validation', 'setValidationAndFormatting'))
  .addSeparator()
  // .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Penyiasat')
  //     .addItem('⚪ Get info segera', 'mainInfoSegeraPenyiasat')
  //     .addItem('⚪ Generate borang siasatan', 'mainGenerateBorangSiasatan'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Epid Daerah')
      .addItem('🚷 Kes Epid selesai', 'mainGenerateLaporanEpid')
      .addItem('🚷 Undo daftar kes', 'mainUndoLaporanEpid'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Developer')
      .addItem('🚷 Trigger Borang Siasatan', 'mainTriggerGenerateBorangSiasatan')
      .addItem('🚷 Trigger Add listed user', 'mainTriggerAddListedUser')
      .addItem('🚷 Trigger Move archive', 'mainTriggerMoveToArchive'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Administrator')
      .addItem('🚷 Add listed user access', 'addUserForm'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕊 About')
      .addItem('⚪ Google AppScript', 'aboutGoogleAppScript')
      .addItem('⚪ Author', 'aboutAuthor')
      .addItem('⚪ License', 'aboutLicense'))
  .addToUi();
}

function aquireGooglePermission() {
  SpreadsheetApp.getUi().alert(
    'Success',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function mainInfoSegeraPenyiasat() {
  let rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  let info_segera_txt = infoSegeraPenyiasat(rowid, getVarSource());
  // Show alert box
  SpreadsheetApp.getUi().alert('Info Segera Penyiasat.', info_segera_txt, SpreadsheetApp.getUi().ButtonSet.OK);
}

function mainGenerateBorangSiasatan() {
  if (promptPassword()) {
    let rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
    let var_source = getVarSource();
    generateBorangSiasatan(rowid, var_source);
  }
}

function mainGenerateLaporanEpid() {
  if (promptAdminUserOnly()) { 
    generateLaporanEpid() 
  }
}

function mainUndoLaporanEpid() {
  if (promptAdminUserOnly()) {
    undoLaporanEpid();
  }
}

function mainMoveToArchive() {
  if (promptAdminUserOnly()) {
    let selected_range = SpreadsheetApp.getActiveRange();
    moveToArchive(selected_range);    
  } 
}

function mainAddListedUser() { 
  if (promptAdminUserOnly()) { 
    addListedUser() 
  } 
}

function mainRemoveUnlistedUser() {
  if (promptAdminUserOnly()) {
    removeUnlistedUser();
  } 
}

function mainTriggerGenerateBorangSiasatan() {
  if (promptAdminUserOnly()) {
    triggerGenerateBorangSiasatan();
  }
}

function mainTriggerAddListedUser() {
  if (promptAdminUserOnly()) {
    triggerAddListedUser();
  }
}

function mainTriggerMoveToArchive() {
  if (promptAdminUserOnly()) {
    triggerMoveToArchive();
  }
}
