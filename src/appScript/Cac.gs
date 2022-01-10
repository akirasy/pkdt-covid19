function generateLaporanCac() {
  var var_source = getVarSource();
  var selected_range = SpreadsheetApp.getActiveRange();
  var destination_array = new Array();
  for (let i = 0; i < selected_range.getNumRows(); i++) {
    let patient_info = getPatientInfo(selected_range.getRowIndex() + i);

    // Arrange value to target column
    destination_array.push([
      patient_info.tindakan_tarikh,             // 1. tarikh positif
      patient_info.logistik_tarikh_dinilai,     // 2. tarikh dinilai
      patient_info.demografi_sampel_satu +
      ', ' + patient_info.demografi_sampel_dua, // 3. tarikh sampel
      patient_info.pesakit_nama,                // 4. nama
      patient_info.pesakit_ic,                  // 5. ic
      patient_info.pesakit_phone,               // 6. telefon
      patient_info.logistik_umur,               // 7. umur
      patient_info.logistik_jantina,            // 8. jantina
      patient_info.demografi_bangsa,            // 9. bangsa
      patient_info.logistik_comorbid,           // 10. comorbid
      patient_info.logistik_cat,                // 11. covid cat
      patient_info.logistik_admit               // 12. admission
    ])

    // mark as CAC done
    SpreadsheetApp.getActiveSheet().getRange(selected_range.getRowIndex() + i, patient_info.reten_cac[1]).setValue('DONE');
  }
  
  // set value to destination range
  var destination_range = var_source.sheet_laporan_cac.getRange(var_source.sheet_laporan_cac.getLastRow() + 1, 1, destination_array.length, destination_array[0].length);
  destination_range.setValues(destination_array);

  // backup and remove target row
  moveToArchive(selected_range);
}
