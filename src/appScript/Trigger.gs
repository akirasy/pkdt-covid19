function triggerGenerateBorangSiasatan() {
  let var_source = getVarSource();
  let patient_info = getPatientInfo(1, var_source);

  // Collect
  let generate_sekarang_values = var_source.sheet_kes_positif
    .getRange(1, patient_info.generate_sekarang[1], var_source.sheet_kes_positif.getLastRow()).getValues()
    .map((item, index) => { return [(index+1), item[0]] });

  // Select
  let selection_criteria = new Array();
  generate_sekarang_values.map((item, index) => {
    if (item[1] == 'YA') {
      selection_criteria.push(item[0]);
    }
  })
  Logger.log('--- Total case selected: ' + selection_criteria.length);

  // Operate with exec_limit
  selection_criteria.map((item, index) => {
    let exec_limit = 15;
    if (index < exec_limit) {
      generateBorangSiasatan(item);
      var_source.sheet_kes_positif.getRange(item, patient_info.generate_sekarang[1]).setValue('');
    }   
  })
}

function triggerAddListedUser() {
  addListedUser();
}

function triggerRemoveUnlistedUser() {
  removeUnlistedUser();
}

function triggerMoveToArchive() {
  let var_source = getVarSource();
  let patient_info = getPatientInfo(1, var_source);

  // Collect
  let status_siasatan_done_values = var_source.sheet_kes_positif
    .getRange(1, patient_info.status_siasatan[1], var_source.sheet_kes_positif.getLastRow()).getValues()
    .map((item, index) => { return [(index+1), item[0]] });
  let epid_daerah_done_values = var_source.sheet_kes_positif
    .getRange(1, patient_info.epid_daerah[1], var_source.sheet_kes_positif.getLastRow()).getValues()
    .map((item, index) => { return [(index+1), item[0]] });

  // Select
  let selection_criteria = new Array();
  status_siasatan_done_values.map((item, index) => {
    if (status_siasatan_done_values[index][1] == 'DONE' && epid_daerah_done_values[index][1] == 'DONE') {
      selection_criteria.push(item[0]);    
    }
  })
  Logger.log('--- Total case selected: ' + selection_criteria.length);

  // Operate with exec_limit
  selection_criteria.map((item, index) => {
    let exec_limit = 100;
    if (index < exec_limit) {
      moveCaseToArchive(item, var_source);
    }   
  })
}
