// Create topbar menu
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Admin')
  .addItem('🕸 Set Google permission', 'aquireGooglePermission')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕸 Text formatting')
  .addItem('⚪ To upperCase', 'toUpperCase')
  .addItem('⚪ To oneLine', 'toOneLine')
  .addItem('⚪ Clean IC', 'cleanIc'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Penyiasat')
      .addItem('⚪ Get info segera', 'mainInfoSegeraPenyiasat')
      .addItem('⚪ Generate borang siasatan', 'mainGenerateBorangSiasatan'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Epid Daerah')
      .addItem('🚷 Kes Epid selesai', 'mainGenerateLaporanEpid')
      .addItem('🚷 Undo daftar kes', 'mainUndoLaporanEpid'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Sorting')
      .addItem('⚪ Set formatting & validation', 'setValidationAndFormatting')
      .addItem('⚪ Select greyed row', 'selectGrayEmpty')
      .addItem('🚷 Move to archive', 'mainMoveToArchive'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Administrator')
      .addItem('🚷 Add listed user access', 'mainAddListedUser')
      .addItem('🚷 Remove unlisted user access', 'mainRemoveUnlistedUser'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Trigger')
      .addItem('🚷 Borang Siasatan', 'mainTriggerGenerateBorangSiasatan')
      .addItem('🚷 Add listed user', 'mainTriggerAddListedUser')
      .addItem('🚷 Move archive', 'mainTriggerMoveToArchive'))
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
  let rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  generateBorangSiasatan(rowid);
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
