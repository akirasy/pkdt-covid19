function assignNewTlhNumber(rowid, current_sheet) {
  Logger.log('Initializing');
  let var_source = getVarSource();
  let generic_patient_info = getPatientInfo(1, var_source);

  Logger.log('Collecting data');
  let tlh_cell = current_sheet.getRange(rowid, generic_patient_info.id_kes[1]);
  let nama_cell = current_sheet.getRange(rowid, generic_patient_info.nama[1]);
  let url_cell = current_sheet.getRange(rowid, generic_patient_info.url_siasatan[1]);
  
  Logger.log('Reassign TLH number');
  let tlh_generated = var_source.range_tlh_max.getValue() + 1;
  let tlh_number = var_source.range_tlh_prefix.getValue() + tlh_generated;
  var_source.range_tlh_max.setValue(tlh_generated);
  tlh_cell.setValue(tlh_number);

  Logger.log('Rename file to new TLH number');
  let file_url = url_cell.getValue();
  let file_id = file_url.match(/[-\w]{25,}/);
  let doc_obj = DriveApp.getFileById(file_id);
  doc_obj.setName(tlh_number + ' ' + nama_cell.getValue());
}

function writeTlhNumber(rowid) {
  let var_source = getVarSource();
  let tlh_column = getPatientInfo(rowid, var_source).id_kes[1];

  let tlh_generated = 0;
  let reuse = reuseTlhNumber(var_source);
  if (reuse[0]) {
    // use un-used tlh number
    tlh_generated = reuse[1];
  } else {
    // increment tlh_max by 1
    tlh_generated = var_source.range_tlh_max.getValue() + 1;
    // update tlh_max
    var_source.range_tlh_max.setValue(tlh_generated);
  }

  // write to target
  let target = var_source.sheet_kes_positif.getRange(rowid, tlh_column);
  target.setValue(var_source.range_tlh_prefix.getValue() + tlh_generated);
  
  return var_source.range_tlh_prefix.getValue() + tlh_generated
}

function reuseTlhNumber(var_source) {
  let is_available = false;

  // get unused tlh number from appscript.gs
  let all_unused_number = var_source.range_unused_tlh.getValues();
  let unused_number = all_unused_number.shift();

  // if no unused number, just return false
  if (unused_number != '') {
    is_available = true;
    var_source.range_unused_tlh.clear();
    let target = var_source.sheet_var_source.getRange(var_source.range_unused_tlh.getRowIndex(), var_source.range_unused_tlh.getColumn(), all_unused_number.length);
    target.setValues(all_unused_number);
  } else { is_available = false }

  return [is_available, unused_number]
}

function undoLaporanEpid() {
  Logger.log('Waiting for user input: Yes/No')
  let ui = SpreadsheetApp.getUi();
  let heading = 'Make sure correct patient is choosen.';
  let message = 'This function will do these:\n 1) Rename borang siasatan to patient\'s name\n 2) Clear TLH number in its column\n 3) Remove EpidDone validation in its column\n\n Are you sure? ';
  let result = ui.alert(heading, message, ui.ButtonSet.YES_NO);

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

  if (user_response) {
    let var_source = getVarSource();
    let rowid = var_source.sheet_kes_positif.getCurrentCell().getRowIndex();
    let patient_info = getPatientInfo(rowid, var_source);

    Logger.log('Set value to cell');
    // Reset ID KES to AWAITING TLH NUMBER
    var_source.sheet_kes_positif.getRange(rowid, patient_info.id_kes[1]).setValue('AWAITING TLH NUMBER');
    // Clear validation at Epid Column
    var_source.sheet_kes_positif.getRange(rowid, patient_info.epid_daerah[1]).setValue('');

    // Rename file to patient's name only
    if (patient_info.url_siasatan[0] != '') {
      Logger.log('Renaming file to: ' + patient_info.nama[0]);
      let file_url = patient_info.url_siasatan[0];
      let file_id = file_url.match(/[-\w]{25,}/);
      let doc_obj = DriveApp.getFileById(file_id);
      doc_obj.setName(patient_info.nama[0]);
    }
  }
}

function generateLaporanEpid() {
  let var_source = getVarSource();
  let selected_range = SpreadsheetApp.getActiveRange();
  let destination_array = new Array();

  let forloop_start = selected_range.getRowIndex();
  let forloop_end = forloop_start + selected_range.getNumRows();
  for (let rowid = forloop_start; rowid < forloop_end; rowid++) {
    Logger.log('Processing row: ' + rowid);
    let patient_info = getPatientInfo(rowid, var_source);

    // STEP 1 : Validate all job done
    let status_siasatan_done = patient_info.status_siasatan[0] == 'DONE';
    let epid_daerah_done = patient_info.epid_daerah[0] == 'DONE';
    let status_siasatan_urlvalid = patient_info.url_siasatan[0] != '';
    let all_job_done = status_siasatan_done && !epid_daerah_done && status_siasatan_urlvalid;

    if (all_job_done) {
      // STEP 2 : Generate TLH number, rename file
      let tlh_number = writeTlhNumber(rowid);
      let file_url = patient_info.url_siasatan[0];
      let file_id = file_url.match(/[-\w]{25,}/);
      let doc_obj = DriveApp.getFileById(file_id);
      doc_obj.setName(tlh_number + ' ' + patient_info.nama[0]);

      // STEP 3: Collect data and add to data array
      destination_array.push([
        '',                                    // 1. epid week
        tlh_number,                            // 2. no TLH
        '',                                    // 3. no KKM
        '',                                    // 4. Tarikh daftar
        patient_info.nama[0],                  // 5. nama
        'PAHANG',                              // 6. negeri
        'TEMERLOH',                            // 7. daerah
        patient_info.mukim[0],                 // 8. mukim
        patient_info.admit[0],                 // 9. pusat rawatan
        patient_info.ic[0],                    // 10. IC
        patient_info.umur[0],                  // 11. umur
        patient_info.jantina[0],               // 12. jantina
        patient_info.bangsa[0],                // 13. kaum
        patient_info.warganegara[0],           // 14. warganegara
        patient_info.jenis_saringan[0],        // 15. kluster
        patient_info.pekerjaan[0],             // 16. pekerjaan
        patient_info.comorbid[0],              // 17. comorbid
        patient_info.alamat[0],                // 18. alamat
        patient_info.phone[0],                 // 19. telefon
        patient_info.tarikh_notifikasi[0],     // 20. tarikh admit
        patient_info.bilangan_kontak_rapat[0], // 21. bilangan kontak
        patient_info.tarikh_onset[0],          // 22. tarikh onset
        patient_info.jenis_gejala[0],          // 23. jenis simptom
        'HIDUP',                               // 24. status hidup
        'N/A',                                 // 25. sebab mati
        patient_info.fasiliti_makmal[0],       // 26. makmal
        patient_info.jenis_ujian[0],           // 27. jenis ujian
        patient_info.tarikh_notifikasi[0],     // 28. tarikh positif
        patient_info.kategori_jangkitan[0],    // 29. lokal import
        '',                                    // 30. origin import
        patient_info.nama_kes_indeks[0],       // 31. catatan
        '',                                    // 32. tarikh discharge
        patient_info.ctval_rdrp[0] + ', ' + 
        patient_info.ctval_n[0] + ', ' +  
        patient_info.ctval_orf[0],             // 33. ct value
        patient_info.tarikh_sampel[0],         // 34. tarikh sampel
        '',                                    // 35. kategori lain
        '',                                    // 36. vaksin satu
        '',                                    // 37. vaksin dua
        patient_info.covid_category[0],        // 38. covid category
        patient_info.status_vaksin[0],         // 39. status vaksin
        patient_info.jenis_vaksin[0],          // 40. jenis vaksin
        '',                                    // 41. vaksin tiga
        '',                                    // 42. tempoh daftar kes
        '',                                    // 43. tempoh vaksin elapsed
        patient_info.status_sampel[0],         // 44. sampel kali ke
        patient_info.catatan_epid[0]           // 45. catatan utk mo epid daerah
      ])

      // STEP 4 : Mark as epid done
      var_source.sheet_kes_positif.getRange(rowid, patient_info.epid_daerah[1]).setValue('DONE');      
    } else {
      Logger.log('Skip rowid: ' + rowid);
    }
  }
  // STEP 5 : Write array value to destination range
  let destination_range = var_source.sheet_laporan_epid.getRange(var_source.sheet_laporan_epid.getLastRow() + 1, 1, destination_array.length, destination_array[0].length);
  destination_range.setValues(destination_array);
}
