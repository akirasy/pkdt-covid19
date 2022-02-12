function writeTlhNumber(rowid) {
  let var_source = getVarSource();
  let tlh_column = getPatientInfo(4, var_source).pesakit_id[1];

  let tlh_generated = 0;
  let reuse = reuseTlhNumber();
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

function reuseTlhNumber() {
  let var_source = getVarSource();
  let is_available = false;

  // get unused tlh number from appscript.gs
  let all_unused_number = var_source.range_unused_tlh.getValues();
  let unused_number = all_unused_number.shift();

  // if no unused number, just return false
  if (unused_number != "") {
    is_available = true;
    var_source.range_unused_tlh.clear();
    let target = var_source.sheet_var_source.getRange(var_source.range_unused_tlh.getRowIndex(), var_source.range_unused_tlh.getColumn(), all_unused_number.length);
    target.setValues(all_unused_number);
  } else { is_available = false }

  return [is_available, unused_number]
}

function generateLaporanEpid() {
  let var_source = getVarSource();
  let selected_range = SpreadsheetApp.getActiveRange();
  let destination_array = new Array();

  for (let i = 0; i < selected_range.getNumRows(); i++) {
    let rowid = selected_range.getRowIndex() + i;
    let patient_info = getPatientInfo(rowid, var_source);

    // nest function in if-else to avoid null error
    if (patient_info.siasatan_status[0] == 'DONE') {
      Logger.log('Processing row: ' + rowid);
      // Generate TLH number
      let tlh_number = writeTlhNumber(rowid);
      // Rename file to prepend TLH number
      let file_url = patient_info.siasatan_url[0];
      let file_id = file_url.match(/[-\w]{25,}/);
      let doc_obj = DriveApp.getFileById(file_id);
      doc_obj.setName(tlh_number + ' ' + doc_obj.getName());

      // Collect data and arrange value to target column
      destination_array.push([
        patient_info.epid_minggu[0],                     // 1. epid week
        tlh_number,                                      // 2. no TLH
        ' ',                                             // 3. no KKM
        ' ',                                             // 4. Tarikh daftar
        patient_info.pesakit_nama[0],                    // 5. nama
        ' ',                                             // 6. negeri
        ' ',                                             // 7. daerah
        patient_info.demografi_mukim[0],                 // 8. mukim
        patient_info.logistik_admit[0],                  // 9. pusat rawatan
        patient_info.pesakit_ic[0],                      // 10. IC
        patient_info.logistik_umur[0],                   // 11. umur
        patient_info.logistik_jantina[0],                // 12. jantina
        patient_info.demografi_bangsa[0],                // 13. kaum
        patient_info.demografi_warganegara[0],           // 14. warganegara
        patient_info.demografi_saringan[0] + ' - ' + 
        patient_info.epid_nama_kluster[0],               // 15. kluster
        patient_info.demografi_pekerjaan[0],             // 16. pekerjaan
        patient_info.logistik_comorbid[0],               // 17. comorbid
        patient_info.pesakit_alamat[0],                  // 18. alamat
        patient_info.pesakit_phone[0],                   // 19. telefon
        patient_info.tindakan_tarikh[0],                 // 20. tarikh admit
        patient_info.epid_bil_kontak[0],                 // 21. bilangan kontak
        patient_info.demografi_gejala_tarikh[0],         // 22. tarikh onset
        patient_info.demografi_gejala_jenis[0],          // 23. jenis simptom
        patient_info.epid_status[0],                     // 24. status hidup
        patient_info.epid_sebab_mati[0],                 // 25. sebab mati
        patient_info.keputusan_makmal[0],                // 26. makmal
        patient_info.epid_jenis_ujian[0],                // 27. jenis ujian
        patient_info.tindakan_tarikh[0],                 // 28. tarikh positif
        patient_info.epid_lokal[0],                      // 29. lokal import
        patient_info.epid_origin[0],                     // 30. origin import
        'CC1 ' + patient_info.epid_nama_indeks[0] + 
        ' (' + patient_info.epid_id_indeks[0] + ')',     // 31. cc1 siapa
        ' ',                                             // 32. tarikh discharge
        patient_info.keputusan_rdrp[0] + ', ' + 
        patient_info.keputusan_n[0] + ', ' +  
        patient_info.keputusan_orf[0],                   // 33. ct value
        patient_info.demografi_sampel_satu[0] + ', ' +  
        patient_info.demografi_sampel_dua[0],            // 34. tarikh sampel
        ' ',                                             // 35. kategori lain
        patient_info.demografi_vaksin_satu[0],           // 36. vaksin satu
        patient_info.demografi_vaksin_dua[0],            // 37. vaksin dua
        patient_info.logistik_cat[0],                    // 38. covid category
        patient_info.demografi_vaksin_status[0],         // 39. status vaksin
        patient_info.demografi_vaksin_jenis[0],          // 40. jenis vaksin
        patient_info.demografi_vaksin_tiga[0],           // 41. vaksin tiga
        '',                                              // 42. tempoh daftar kes
        '',                                              // 43. tempoh vaksin elapsed
        patient_info.epid_sampel_kali[0],                // 44. sampel kali ke
        patient_info.reten_catatan[0]                    // 45. catatan utk mo epid daerah
      ])

      // mark as epid done
      var_source.sheet_kes_positif.getRange(selected_range.getRowIndex() + i, patient_info.reten_epid[1]).setValue('DONE');
    }
  }
  
  // set value to destination range
  let destination_range = var_source.sheet_laporan_epid.getRange(var_source.sheet_laporan_epid.getLastRow() + 1, 1, destination_array.length, destination_array[0].length);
  destination_range.setValues(destination_array);
}
