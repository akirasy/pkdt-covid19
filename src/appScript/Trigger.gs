function triggerGenerateBorangSiasatan() {
  let var_source = getVarSource();
  let generate_now_column = getPatientInfo(1, var_source).generate_now[1];
  let all_row_range = var_source.sheet_kes_positif.getRange(1, generate_now_column, var_source.sheet_kes_positif.getLastRow());
  Logger.log('===== Searching for row to be generated =====');
  let to_generate = new Array();
  for (let i = 0; i < all_row_range.getNumRows(); i++) {
    let rowid = all_row_range.getRowIndex() + i;
    if ( var_source.sheet_kes_positif.getRange(rowid, generate_now_column).getValue() == "YA" ) {
      to_generate.push(rowid);
    }
  }
  if (to_generate.length != 0) {
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
  let selected_range = var_source.sheet_kes_positif.getRange(1, 1, var_source.sheet_kes_positif.getLastRow());
  moveToArchive(selected_range);
}
