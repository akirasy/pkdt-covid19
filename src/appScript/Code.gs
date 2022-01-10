// Create topbar menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PKD Temerloh')
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
        .addItem('Set validation', 'daVaAll')
        .addItem('Select greyed row', 'selectGrayEmpty')
        .addItem('🚷 Move to archive', 'mainMoveToArchive'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠 Administrator')
        .addItem('🚷 Add listed user access', 'mainAddListedUser')
        .addItem('🚷 Remove unlisted user access', 'mainRemoveUnlistedUser'))
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

// Change lowercase to uppercase
function toUpperCase() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selected_range = activeSheet.getActiveRange();
  var data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().toUpperCase();
      }
    }
  }
  selected_range.setValues(data_list);
}

// Convert newline value to oneline only
function toOneLine() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selected_range = activeSheet.getActiveRange();
  var data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/\n/g, '  ');
      }
    }
  }
  selected_range.setValues(data_list).trimWhitespace();
}

// Removes dashes, star, spaces and apostrophy
function cleanIc() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selected_range = activeSheet.getActiveRange();
  var data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/[-|\'|\*|\s]/g,'');
      }
    }
  }
  selected_range.setValues(data_list);
}

function daVaAll() {
  // Format date to d mmm yyyy and d mmm where applicable
  var var_source = getVarSource();
  var_source.sheet_kes_positif.getRange(4,  1, var_source.sheet_kes_positif.getMaxRows(), 1).setNumberFormat('d mmm');      // date_notified
  var_source.sheet_kes_positif.getRange(4, 30, var_source.sheet_kes_positif.getMaxRows(), 3).setNumberFormat('d mmm yyyy'); // date_vaccine
  var_source.sheet_kes_positif.getRange(4, 35, var_source.sheet_kes_positif.getMaxRows(), 1).setNumberFormat('d mmm');      // date_onset
  var_source.sheet_kes_positif.getRange(4, 37, var_source.sheet_kes_positif.getMaxRows(), 2).setNumberFormat('d mmm');      // date_sampling
  var_source.sheet_kes_positif.getRange(4, 47, var_source.sheet_kes_positif.getMaxRows(), 1).setNumberFormat('d mmm');      // date_report

  // Data validation - dropdown
  var validation_collection_dropdown = [
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
  ]
  for (let i = 0; i < validation_collection_dropdown.length; i++) {
    let target_range = SpreadsheetApp.getActiveSheet().getRange(validation_collection_dropdown[i].range);
    let target_rule = SpreadsheetApp.newDataValidation().requireValueInList(validation_collection_dropdown[i].validation_list, true).build();
    target_range.setDataValidation(target_rule);
  }

  // Data validation - reject invalid date
  var validation_collection_date = [
    {'name' : 'date_notified' , 'range' : 'A:A'  },
    {'name' : 'date_vaccine'  , 'range' : 'AD:AF'},
    {'name' : 'date_onset'    , 'range' : 'AI:AI'},
    {'name' : 'date_sampling' , 'range' : 'AK:AL'},
    {'name' : 'date_report'   , 'range' : 'AU:AU'},
  ]
  for (let i = 0; i < validation_collection_date.length; i++) {
    let target_range = SpreadsheetApp.getActiveSheet().getRange(validation_collection_date[i].range);
    let target_rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Use this format -> d mmm yyyy (eg. 12 Sep 2021, 25 Aug 2021)').build();
    target_range.setDataValidation(target_rule);
  }

  // Clear data validation at header
  var range_header = SpreadsheetApp.getActiveSheet().getRange('1:3');
  range_header.clearDataValidations();
}

function mainInfoSegeraReferral() {
  infoSegeraReferral();
}

function mainInfoSegeraPenyiasat() {
  var rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  infoSegeraPenyiasat(rowid);
}

function mainGenerateBorangSiasatan() {
  generateBorangSiasatan();
}

// Get new TLH number
function mainGenerateTlhNumber() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Reserved function',
     'This is a reserved function for Peg Epid Daerah only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('TLH generated successfully.', 'TLH number is available at your current active cell, column F.', ui.ButtonSet.OK);
    var rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
    writeTlhNumber(rowid)
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }
}

function mainGenerateLaporanEpid() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
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

function mainGenerateLaporanCac() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Reserved function',
     'This is a reserved function for Peg CAC Daerah only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    generateLaporanCac();
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }
}

function selectGrayEmpty() {
  var var_source = getVarSource();
  var cursor = SpreadsheetApp.getCurrentCell();
  var selected_range = var_source.sheet_kes_positif.getRange(cursor.getRowIndex(), 1, 50);
  var rowid_range_gray = new Array();
  for (let i = 0; i < selected_range.getNumRows(); i++) {
    let rowid = selected_range.getRowIndex() + i;
    let target = var_source.sheet_kes_positif.getRange(rowid, getPatientInfo(rowid).reten_catatan[1]);
    if (target.getBackground() == '#cccccc') {
      let selected_row_range = var_source.sheet_kes_positif.getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns()).getA1Notation();
      rowid_range_gray.push(selected_row_range);
    }
  }
  if (rowid_range_gray.length) {
    var_source.sheet_kes_positif.getRangeList(rowid_range_gray).activate();
  } else {
    SpreadsheetApp.getUi().alert('All is done!', 'No greyed row found.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function mainMoveToArchive() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Reserved function',
     'This is a reserved function for Peg PKD Daerah only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    var selected_range = SpreadsheetApp.getActiveRange();
    moveToArchive(selected_range);
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  }  
}

function mainAddListedUser() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
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
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
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
