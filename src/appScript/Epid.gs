// return function to generate TLH number
function writeTlhNumber(rowid) {
  // init var from sheet: {appScript.gs}
  var var_source = getVarSource();

  // increment tlh_max by 1
  var tlh_generated = var_source.range_tlh_max.getValue() + 1;
  // write to target
  var target = var_source.sheet_kes_positif.getRange(rowid, 6);
  target.setValue(var_source.range_tlh_prefix.getValue() + tlh_generated);
  // update tlh_max
  var_source.range_tlh_max.setValue(tlh_generated);
  
  return var_source.range_tlh_prefix.getValue() + tlh_generated
}

// Get TLH number in bulk
function getBulkTlhNumber() {
  var selected_range = SpreadsheetApp.getActiveRange();
  for (let i = 0; i < selected_range.getNumRows(); i++) {
    let rowid = selected_range.getRowIndex() + i;
    writeTlhNumber(rowid);
  }
}

function generateLaporanEpid() {
  var var_source = getVarSource();
  var selected_range = SpreadsheetApp.getActiveRange();
  var destination_array = new Array();
  for (let i = 0; i < selected_range.getNumRows(); i++) {
    let rowid = selected_range.getRowIndex() + i;

    // nest function in if-else to avoid null error
    var is_done = var_source.sheet_kes_positif.getRange(rowid, 56);
    if (is_done.getValue() == 'DONE') {
      Logger.log('Processing row: ' + rowid);
      // Generate TLH number
      var tlh_number = writeTlhNumber(rowid);
      // Rename file to prepend TLH number
      var file_url = var_source.sheet_kes_positif.getRange(rowid, 55).getValue();
      var file_id = file_url.match(/[-\w]{25,}/);
      var doc_obj = DriveApp.getFileById(file_id);
      doc_obj.setName(tlh_number + ' ' + doc_obj.getName());

      // Collect data and arrange value to target column
      let patient_info = getPatientInfo(rowid);
      destination_array.push([
        patient_info.epid_minggu,                     // 1. epid week
        patient_info.pesakit_id,                      // 2. no TLH
        ' ',                                          // 3. no KKM
        ' ',                                          // 4. Tarikh daftar
        patient_info.pesakit_nama,                    // 5. nama
        ' ',                                          // 6. negeri
        ' ',                                          // 7. daerah
        patient_info.demografi_mukim,                 // 8. mukim
        patient_info.logistik_admit,                  // 9. pusat rawatan
        patient_info.pesakit_ic,                      // 10. IC
        patient_info.logistik_umur,                   // 11. umur
        patient_info.logistik_jantina,                // 12. jantina
        patient_info.demografi_bangsa,                // 13. kaum
        patient_info.demografi_warganegara,           // 14. warganegara
        patient_info.demografi_saringan + ' - ' + 
        patient_info.epid_nama_kluster,               // 15. kluster
        patient_info.demografi_pekerjaan,             // 16. pekerjaan
        patient_info.logistik_comorbid,               // 17. comorbid
        patient_info.pesakit_alamat,                  // 18. alamat
        patient_info.pesakit_phone,                   // 19. telefon
        patient_info.tindakan_tarikh,                 // 20. tarikh admit
        patient_info.epid_bil_kontak,                 // 21. bilangan kontak
        patient_info.demografi_gejala_tarikh,         // 22. tarikh onset
        patient_info.demografi_gejala_jenis,          // 23. jenis simptom
        patient_info.epid_status,                     // 24. status hidup
        patient_info.epid_sebab_mati,                 // 25. sebab mati
        patient_info.keputusan_makmal,                // 26. makmal
        patient_info.epid_jenis_ujian,                // 27. jenis ujian
        patient_info.tindakan_tarikh,                 // 28. tarikh positif
        patient_info.epid_lokal,                      // 29. lokal import
        patient_info.epid_origin,                     // 30. origin import
        'CC1 ' + patient_info.epid_nama_indeks + 
        ' (' + patient_info.epid_id_indeks + ')',     // 31. cc1 siapa
        ' ',                                          // 32. tarikh discharge
        patient_info.keputusan_rdrp + ', ' + 
        patient_info.keputusan_n + ', ' +  
        patient_info.keputusan_orf,                   // 33. ct value
        patient_info.demografi_sampel_satu + ', ' +  
        patient_info.demografi_sampel_dua,            // 34. tarikh sampel
        ' ',                                          // 35. kategori lain
        patient_info.demografi_vaksin_satu,           // 36. vaksin satu
        patient_info.demografi_vaksin_dua,            // 37. vaksin dua
        patient_info.logistik_cat,                    // 38. covid category
        patient_info.demografi_vaksin_status,         // 39. status vaksin
        patient_info.demografi_vaksin_jenis,          // 40. jenis vaksin
        patient_info.demografi_vaksin_tiga,           // 41. vaksin tiga
        '',                                           // 42. tempoh daftar kes
        '',                                           // 43. tempoh vaksin elapsed
        patient_info.epid_sampel_kali,                // 44. sampel kali ke
        patient_info.reten_catatan                    // 45. catatan utk mo epid daerah
      ])

      // mark as epid done
      var_source.sheet_kes_positif.getRange(selected_range.getRowIndex() + i, patient_info.reten_epid[1]).setValue('DONE');
    }
  }
  
  // set value to destination range
  var destination_range = var_source.sheet_laporan_epid.getRange(var_source.sheet_laporan_epid.getLastRow() + 1, 1, destination_array.length, destination_array[0].length);
  destination_range.setValues(destination_array);

  // backup and remove target row
  // moveToArchive(selected_range); // comment this out to give time for others finish their work first
}
