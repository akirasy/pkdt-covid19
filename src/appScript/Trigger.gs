function triggerGenerateBorangSiasatan() {
  let var_source = getVarSource();
  let generate_now_column = getPatientInfo(1, var_source).generate_now[1];
  let all_row_range = var_source.sheet_kes_positif.getRange(1, generate_now_column, var_source.sheet_kes_positif.getLastRow());
  Logger.log('===== Searching for row to be generated =====');
  let to_generate = new Array();

  let all_row_range_values = all_row_range.getValues().map((item, index) => { return [(index+1),item[0]] });
  all_row_range_values.forEach(item => {
    if (item[1] == 'YA') {
      to_generate.push(item[0]);
    }
  })

  if (to_generate.length != 0) {
    Logger.log('Found row: ' + to_generate.length);
    Logger.log('===== Generating Borang Siasatan =====');
    to_generate.forEach(item => {
      Logger.log('--- Looking at rowid: ' + item);
      generateBorangSiasatan(item);
      Logger.log('--- Set value generate_now to empty')
      var_source.sheet_kes_positif.getRange(item, generate_now_column).setValue('');    
    })
  } else { Logger.log('--- No request for generate found') }
}

function triggerAddListedUser() {
  addListedUser();
}

function triggerRemoveUnlistedUser() {
  removeUnlistedUser();
}

function triggerMoveToArchive() {
  let var_source = getVarSource();
  let selected_range = var_source.sheet_kes_positif.getRange(4, 1, var_source.sheet_kes_positif.getLastRow());
  moveToArchive(selected_range);
}
