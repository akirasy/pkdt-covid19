// Select grayed row
function selectGrayEmpty() {
  let var_source = getVarSource();

  // Collect
  let sheet_max_column = var_source.sheet_kes_positif.getMaxColumns();
  let all_row_range = var_source.sheet_kes_positif.getRange(1, sheet_max_column, var_source.sheet_kes_positif.getMaxRows());

  // Select
  let greyed_A1 = new Array();
  let rowid_colours = all_row_range.getBackgrounds().map((item, index) => { return [(index+1), item[0]] });
  rowid_colours.map(item => {
    if (item[1] == '#cccccc' ) {
      let target_A1_value = var_source.sheet_kes_positif.getRange(item[0], 1, 1, sheet_max_column).getA1Notation();
      greyed_A1.push(target_A1_value);
    }
  })

  // Operate
  if (greyed_A1.length != 0) {
    Logger.log('Total greyed row found: ' + greyed_A1.length);
    Logger.log('--- Activating rows...');
    var_source.sheet_kes_positif.getRangeList(greyed_A1).activate();
    Logger.log('--- Greyed row activated');
  } else { Logger.log('No greyed row found') }
}

// Set data validation and formatting
function setValidationAndFormatting() {
  // Data validation - dropdown
  let validation_collection_dropdown = [
    {'name'  : 'kk_referral'      , 'range' : 'B:B'  , 'validation_list' : ['KKBM','KKT', 'KKTL', 'KKL', 'KKS', 'KKK', 'KKKT', 'KKK', 'KKKK', 'HOSHAS','LUAR DAERAH']},
    {'name'  : 'covid_category'   , 'range' : 'X:X'  , 'validation_list' : ['CAT 1', 'CAT 2A', 'CAT 2B', 'CAT 3', 'CAT 4', 'CAT 5']},
    {'name'  : 'status_vaksin'    , 'range' : 'Y:Y'  , 'validation_list' : ['TIDAK VAKSIN','TIDAK LENGKAP', 'LENGKAP', 'BOOSTER']},
    {'name'  : 'jenis_vaksin'     , 'range' : 'Z:Z'  , 'validation_list' : ['NA', 'PFIZER', 'CANSINO', 'SINOVAC', 'ASTRA ZENECA']},
    {'name'  : 'rawatan_paxlovid' , 'range' : 'AB:AB', 'validation_list' : ['YA', 'TIDAK', 'REFUSED']},
    {'name'  : 'warganegara'      , 'range' : 'AD:AD', 'validation_list' : ['MALAYSIA', 'BWN']},    
    {'name'  : 'jenis_saringan'   , 'range' : 'AF:AF', 'validation_list' : ['SARINGAN BERGEJALA', 'SARINGAN KONTAK RAPAT', 'BERSASAR', 'SARINGAN KENDIRI', 'SARINGAN PEKERJAAN', 'SARINGAN PENGEMBARA', 'SARINGAN PRE-ADMISSION']},
    {'name'  : 'punca_jangkitan'  , 'range' : 'AM:AM', 'validation_list' : ['LOKAL', 'IMPORT A', 'IMPORT B', 'IMPORT C']},
    {'name'  : 'generate_sekarang', 'range' : 'AQ:AQ', 'validation_list' : ['YA']},
  ]
  Logger.log('Setting up data validations.');
  validation_collection_dropdown.forEach(item => {
    let target_range = SpreadsheetApp.getActiveSheet().getRange(item.range);
    let target_rule = SpreadsheetApp.newDataValidation().requireValueInList(item.validation_list, true).build();
    target_range.setDataValidation(target_rule);    
  })

  // Data validation - reject invalid date
  let validation_collection_date = [
    {'name' : 'tarikh_notifikasi' , 'range' : 'A:A'  , 'number_format' : 'd mmm'},
    {'name' : 'tarikh_sampel'     , 'range' : 'M:M'  , 'number_format' : 'd mmm'},
    {'name' : 'tarikh_dinilai'    , 'range' : 'T:T'  , 'number_format' : 'd mmm'},
    {'name' : 'tarikh_onset'      , 'range' : 'AH:AH', 'number_format' : 'd mmm'},
    {'name' : 'tarikh_siasatan'   , 'range' : 'AN:AN', 'number_format' : 'd mmm'},
  ]
  Logger.log('Setting up date formatting and validations.');
  validation_collection_date.forEach(item => {
    let target_range = SpreadsheetApp.getActiveSheet().getRange(item.range);
    let target_rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Use this format -> d mmm yyyy (eg. 12 Sep 2021, 25 Aug 2021)').build();
    target_range.setDataValidation(target_rule);
    target_range.setNumberFormat(item.number_format);
  })

  Logger.log('Clearing data validation at header.');
  let range_header = SpreadsheetApp.getActiveSheet().getRange('1:4');
  range_header.clearDataValidations();
}

// Password protected function
function promptPassword() {
  Logger.log('Waiting for user input: Yes/No');
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Password protected function', 'Please enter password.', ui.ButtonSet.YES_NO);

  // Process the user's response.
  let allow_usage = new Boolean();
  let correct_password = response.getResponseText() == '123qwe';
  if (response.getSelectedButton() == ui.Button.YES && correct_password) {
    allow_usage = true;
  } else if (response.getSelectedButton() == ui.Button.NO) {
    allow_usage = false;
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
  } else {
    allow_usage = false;
    ui.alert('Access denied', 'Wrong password!', ui.ButtonSet.OK);;
  }
  return allow_usage
}

// Prompt user to verify identity
function promptAdminUserOnly() {
  Logger.log('Waiting for user input: Yes/No');
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Reserved function',
     'This is a reserved function for administrators only.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  let user_response = new Boolean();
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    user_response = true;
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
    user_response = false;
  }
  return user_response
}

// Check if date to parse toDateString()
function parseDate(arg) {
  let output = '';
  if (arg instanceof Date) { 
    let arg_input = new Date(arg);
    output = arg_input.getDate() + '/' + (arg_input.getMonth() + 1) + '/' + arg_input.getFullYear();
  } else { output = arg }
  return output;
}

// Change lowercase to uppercase
function toUpperCase() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
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
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
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
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/[-|\'|\*|\s]/g,'');
      }
    }
  }
  selected_range.setValues(data_list);
}
