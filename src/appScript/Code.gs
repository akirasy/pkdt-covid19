// Create topbar menu
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Admin')
  .addItem('🕸 Set Google permission', 'aquireGooglePermission')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕸 Text formatting')
  .addItem('To upperCase', 'toUpperCase')
  .addItem('To oneLine', 'toOneLine')
  .addItem('Clean IC', 'cleanIc'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Referral')
      .addItem('Get info segera', 'mainInfoSegeraReferral'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Penyiasat')
      .addItem('Get info segera', 'mainInfoSegeraPenyiasat')
      .addItem('Generate borang siasatan', 'mainGenerateBorangSiasatan'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🖋 Peg Epid Daerah')
      .addItem('🚷 Kes Epid selesai', 'mainGenerateLaporanEpid'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Sorting')
      .addItem('Set formatting & validation', 'setValidationAndFormatting')
      .addItem('Select greyed row', 'selectGrayEmpty')
      .addItem('🚷 Move to archive', 'mainMoveToArchive'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Administrator')
      .addItem('🚷 Add listed user access', 'mainAddListedUser')
      .addItem('🚷 Remove unlisted user access', 'mainRemoveUnlistedUser'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Trigger')
      .addItem('Borang Siasatan', 'triggerGenerateBorangSiasatan')
      .addItem('Add listed user', 'triggerAddListedUser')
      .addItem('Move archive', 'triggerMoveToArchive'))
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕊 About')
      .addItem('Google AppScript', 'aboutGoogleAppScript')
      .addItem('Author', 'aboutAuthor')
      .addItem('License', 'aboutLicense'))
  .addToUi();
}

function aquireGooglePermission() {
  SpreadsheetApp.getUi().alert(
    'Success.',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function mainInfoSegeraReferral() {
  let rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  let info_segera_txt = infoSegeraReferral(rowid, getVarSource());
  // Show alert box
  SpreadsheetApp.getUi().alert('Info Segera Referral.', info_segera_txt, SpreadsheetApp.getUi().ButtonSet.OK);
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
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Reserved function',
     'This is a reserved function for Peg Epid Daerah only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    generateLaporanEpid();
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }
}

function selectGrayEmpty() {
  let var_source = getVarSource();
  let greyed_A1_list = new Array();
  let sheet_max_column = var_source.sheet_kes_positif.getMaxColumns();
  let all_row_range = var_source.sheet_kes_positif.getRange(4, sheet_max_column, var_source.sheet_kes_positif.getLastRow());
  
  Logger.log('Searching for greyed row')
  let forloop_start = all_row_range.getRowIndex();
  let forloop_end = forloop_start + all_row_range.getNumRows();
  for (let rowid = forloop_start; rowid < forloop_end; rowid++) {
    let target = var_source.sheet_kes_positif.getRange(rowid,sheet_max_column);
    if (target.getBackground() == '#cccccc') {
      Logger.log('--- Found at row: ' + rowid);
      let greyed_row_A1 = var_source.sheet_kes_positif.getRange(rowid, 1, 1, sheet_max_column).getA1Notation();
      greyed_A1_list.push(greyed_row_A1);
    }
  }

  if (greyed_A1_list.length != 0) {
    var_source.sheet_kes_positif.getRangeList(greyed_A1_list).activate();
    Logger.log('Greyed row activated')
  } else { Logger.log('No greyed row found') }
}

function mainMoveToArchive() {
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Reserved function',
     'This is a reserved function for Peg PKD Daerah only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    let selected_range = SpreadsheetApp.getActiveRange();
    moveToArchive(selected_range);
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }  
}

function mainAddListedUser() {
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Reserved function',
     'This is a reserved function for Administrators only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    addListedUser();
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }  
}

function mainRemoveUnlistedUser() {
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Reserved function',
     'This is a reserved function for Administrators only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    removeUnlistedUser();
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }  
}

function setValidationAndFormatting() {
  // Format date to d mmm yyyy and d mmm where applicable
  let var_source = getVarSource();
  var_source.sheet_kes_positif.getRange(4,  1, var_source.sheet_kes_positif.getMaxRows(), 1).setNumberFormat('d mmm');      // date_notified
  var_source.sheet_kes_positif.getRange(4, 30, var_source.sheet_kes_positif.getMaxRows(), 3).setNumberFormat('d mmm yyyy'); // date_vaccine
  var_source.sheet_kes_positif.getRange(4, 35, var_source.sheet_kes_positif.getMaxRows(), 1).setNumberFormat('d mmm');      // date_onset
  var_source.sheet_kes_positif.getRange(4, 37, var_source.sheet_kes_positif.getMaxRows(), 2).setNumberFormat('d mmm');      // date_sampling
  var_source.sheet_kes_positif.getRange(4, 47, var_source.sheet_kes_positif.getMaxRows(), 1).setNumberFormat('d mmm');      // date_report

  // Data validation - dropdown
  let validation_collection_dropdown = [
    {'name'  : 'kk_refer'        , 'range' : 'B:B'  , 'validation_list' : ['KKBM','KKT', 'KKTL', 'KKL', 'KKS', 'KKK', 'KKKT', 'KKK', 'KKKK', 'LUAR DAERAH']},
    {'name'  : 'gender'          , 'range' : 'Q:Q'  , 'validation_list' : ['Lelaki','Perempuan']},
    {'name'  : 'covid_cat'       , 'range' : 'T:T'  , 'validation_list' : ['CAT 1', 'CAT 2 (mild)', 'CAT 2 (moderate)', 'CAT 3', 'CAT 4', 'CAT 5']},
    {'name'  : 'warganegara'     , 'range' : 'Y:Y'  , 'validation_list' : ['YA', 'TIDAK']},
    {'name'  : 'jenis_saringan'  , 'range' : 'AA:AA', 'validation_list' : ['BERGEJALA', 'KONTAK RAPAT', 'BERSASAR', 'KENDIRI', 'SARINGAN PEKERJAAN', 'SARINGAN PENGEMBARA', 'PRE-ADMISSION']},
    {'name'  : 'status_vaksin'   , 'range' : 'AC:AC', 'validation_list' : ['LENGKAP', 'TIDAK LENGKAP', 'TIADA VAKSIN']},
    {'name'  : 'jenis_vaksin'    , 'range' : 'AH:AH', 'validation_list' : ['TIADA', 'PFIZER', 'CANSINO', 'SINOVAC', 'ASTRA ZENECA']},
    {'name'  : 'epid_status'     , 'range' : 'AW:AW', 'validation_list' : ['HIDUP', 'MATI']},
    {'name'  : 'epid_mati'       , 'range' : 'AX:AX', 'validation_list' : ['N/A']},
    {'name'  : 'jenis_sampel'    , 'range' : 'AY:AY', 'validation_list' : ['RT-PCR', 'RAPID MOLECULAR', 'RTK-Ag']},
    {'name'  : 'sampel_kali'     , 'range' : 'AZ:AZ', 'validation_list' : ['PERTAMA', 'KEDUA', 'KETIGA']},
    {'name'  : 'punca_jangkitan' , 'range' : 'BA:BA', 'validation_list' : ['LOKAL', 'IMPORT A', 'IMPORT B', 'IMPORT C']},
    {'name'  : 'jangkitan_origin', 'range' : 'BB:BB', 'validation_list' : ['N/A']},
    {'name'  : 'generate_now'    , 'range' : 'BC:BC', 'validation_list' : ['YA']},
  ]
  Logger.log('Setting up data validations.');
  validation_collection_dropdown.forEach(item => {
    let target_range = SpreadsheetApp.getActiveSheet().getRange(item.range);
    let target_rule = SpreadsheetApp.newDataValidation().requireValueInList(item.validation_list, true).build();
    target_range.setDataValidation(target_rule);    
  })

  // Data validation - reject invalid date
  let validation_collection_date = [
    {'name' : 'date_notified' , 'range' : 'A:A'  },
    {'name' : 'date_vaccine'  , 'range' : 'AD:AF'},
    {'name' : 'date_onset'    , 'range' : 'AI:AI'},
    {'name' : 'date_sampling' , 'range' : 'AK:AL'},
    {'name' : 'date_report'   , 'range' : 'AU:AU'},
  ]
  Logger.log('Setting up date formatting and validations.');
  validation_collection_date.forEach(item => {
    let target_range = SpreadsheetApp.getActiveSheet().getRange(item.range);
    let target_rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Use this format -> d mmm yyyy (eg. 12 Sep 2021, 25 Aug 2021)').build();
    target_range.setDataValidation(target_rule);    
  })

  Logger.log('Clearing data validation at header.');
  let range_header = SpreadsheetApp.getActiveSheet().getRange('1:3');
  range_header.clearDataValidations();
}
