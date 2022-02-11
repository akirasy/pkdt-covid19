function generateBorangSiasatan(rowid) {
  // Logger.log('--- Init var from appScript.gs.');
  let var_source = getVarSource();
  let clerking_template_id = var_source.path_clerking_template;
  let target_folder_id = var_source.path_tlh_folder;

  // Logger.log('--- Retrieve user information from {kes positif}');
  let patient_info = getPatientInfo(rowid, var_source);
  
  // nest function in if-else to avoid duplication
  if (patient_info.siasatan_status[0] != 'DONE') {

    // Logger.log('--- Configure target folder id.');
    let today = new Date();
    let target_folder_today_id = '';
    if (today.getDate() == var_source.range_today_date.getValue()) {
      // use current day folder id
      target_folder_today_id = var_source.range_tlh_folder_today.getValue(); 
      } else {
      target_folder_today_id = DriveApp.getFolderById(target_folder_id).createFolder(today.toDateString()).getId();
      // update source_var : date & folder id
      var_source.range_today_date.setValue(today.getDate());
      var_source.range_tlh_folder_today.setValue(target_folder_today_id);
    }

    // ---- begin to generate file ----
    Logger.log('Begin to generate file');
    Logger.log('--- Generate filename');
    let doc_filename = patient_info.pesakit_nama[0];

    Logger.log('--- Make copy of the newly created file from template and get its url link');
    let doc_obj = DriveApp.getFileById(clerking_template_id).makeCopy(doc_filename, DriveApp.getFolderById(target_folder_today_id));

    Logger.log('--- Mark case as done generated and set URL Link to cell');
    var_source.sheet_kes_positif.getRange(rowid, patient_info.pesakit_id[1]).setValue('AWAITING TLH NUMBER');
    var_source.sheet_kes_positif.getRange(rowid, patient_info.siasatan_status[1]).setValue('DONE');
    var_source.sheet_kes_positif.getRange(rowid, patient_info.siasatan_url[1]).setValue(doc_obj.getUrl());

    Logger.log('--- Start writing to file');
    let body = DocumentApp.openById(doc_obj.getId()).getBody();
    let table_array = body.getTables();

    let tbl_1 = table_array[1];
    tbl_1.getCell(0, 2).setText(patient_info.pesakit_nama[0]);             // Nama
    tbl_1.getCell(1, 2).setText(patient_info.pesakit_ic[0]);               // IC
    tbl_1.getCell(2, 2).setText(patient_info.pesakit_phone[0]);            // Phone
    tbl_1.getCell(3, 2).setText(patient_info.pesakit_alamat[0]);           // Address
    tbl_1.getCell(5, 2).setText(patient_info.demografi_bangsa[0]);         // Race
    tbl_1.getCell(6, 2).setText(patient_info.logistik_umur[0] + ' TAHUN'); // Age
    tbl_1.getCell(7, 2).setText(patient_info.logistik_jantina[0]);         // Gender
    tbl_1.getCell(8, 2).setText(patient_info.demografi_warganegara[0]);    // Warganegara
    tbl_1.getCell(9, 2).setText(patient_info.demografi_mukim[0]);          // Mukim
    tbl_1.getCell(10, 2).setText(patient_info.epid_nama_kluster[0]);       // Kluster
    tbl_1.getCell(11, 2).setText(patient_info.demografi_pekerjaan[0]);     // Pekerjaan
    tbl_1.getCell(12, 2).setText(patient_info.demografi_vaksin_satu[0]);   // Vaksin satu
    tbl_1.getCell(13, 2).setText(patient_info.demografi_vaksin_dua[0]);    // Vaksin dua
    tbl_1.getCell(14, 2).setText(patient_info.demografi_vaksin_tiga[0]);   // Vaksin tiga 

    let tbl_2 = table_array[2];
    tbl_2.getCell(0, 2).setText(patient_info.tindakan_tarikh[0]);          // Date diagnosis
    tbl_2.getCell(1, 2).setText(patient_info.logistik_admit[0]);           // Admitting hospital
    tbl_2.getCell(2, 2).setText(
      patient_info.epid_nama_indeks[0] + ' (' +
      patient_info.epid_id_indeks[0] + ')');                               // Index case name
    tbl_2.getCell(3, 2).setText(patient_info.epid_hubungan[0]);            // Relation to index
    tbl_2.getCell(4, 2).setText(patient_info.epid_lokal[0]);               // Lokal atau import

    let tbl_4 = table_array[3];
    tbl_4.getCell(0, 2).setText(patient_info.demografi_gejala_tarikh[0]);  // Onset date
    tbl_4.getCell(1, 2).setText(patient_info.demografi_gejala_jenis[0]);   // Onset type

    let tbl_5 = table_array[4];
    tbl_5.getCell(0, 2).setText(patient_info.logistik_comorbid[0]);        // Comorbid

    let tbl_6 = table_array[5];
    tbl_6.getCell(0, 2).setText(patient_info.demografi_sampel_satu[0]);    // First sample
    tbl_6.getCell(1, 2).setText(patient_info.demografi_sampel_dua[0]);     // Second sample
    tbl_6.getCell(2, 2).setText(
      patient_info.keputusan_rdrp[0] + ', ' + 
      patient_info.keputusan_n[0] + ', ' +  
      patient_info.keputusan_orf[0]);                                      // Result CtV

    let tbl_10 = table_array[9];
    tbl_10.getCell(0, 0).setText(infoSegeraPenyiasat(rowid, var_source));  // Info segera

    let tbl_11 = table_array[10];
    tbl_11.getCell(0, 2).setText(patient_info.penyiasat_nama[0]);          // Investigator name
    tbl_11.getCell(1, 2).setText(patient_info.penyiasat_jawatan[0]);       // Investigator designation
    tbl_11.getCell(2, 2).setText(patient_info.penyiasat_tarikh[0]);        // Investigation date
    tbl_11.getCell(3, 2).setText(patient_info.penyiasat_telefon[0]);       // Investigator phone number
    
    if (doc_obj.getOwner().getEmail() != var_source.spreadsheet_owner) {
      Logger.log('--- Change ownership to Spreadsheet owner');
      doc_obj.setOwner(var_source.spreadsheet_owner);
    }
  }
}

function infoSegeraPenyiasat(rowid, var_source) {
  let patient_info = getPatientInfo(rowid, var_source);
  let info_segera_txt = 
      '*INFO SEGERA KES POSITIF*\n' +
      '\nNama: ' + patient_info.pesakit_nama[0] +
      '\nNo IC: ' + patient_info.pesakit_ic[0] +
      '\nUmur: '  + patient_info.logistik_umur[0] + ' tahun' +
      '\nPekerjaan: ' + patient_info.demografi_pekerjaan[0] +
      '\nCT-Value: ' + patient_info.keputusan_rdrp[0] + ', ' + patient_info.keputusan_n[0] + ', ' +  patient_info.keputusan_orf[0] +
      '\nSampel kali ke: ' + patient_info.epid_sampel_kali[0] +
      '\n\nAlamat: ' + patient_info.pesakit_alamat[0] +
      '\nMukim: '  + patient_info.demografi_mukim[0] +
      '\nTelefon: ' + patient_info.pesakit_phone[0] +
      '\nJenis saringan: ' + patient_info.demografi_saringan[0] +
      '\nKategori jangkitan: ' + patient_info.epid_lokal[0] +
      '\nNama Indeks: ' + patient_info.epid_nama_indeks[0] + ' (' + patient_info.epid_id_indeks[0] + ')' +
      '\nKluster: ' + patient_info.epid_nama_kluster[0];
  return info_segera_txt;
}

function infoSegeraReferral(rowid, var_source) {
  let patient_info = getPatientInfo(rowid, var_source);
  let info_segera_txt = 
      '*INFO SEGERA KES POSITIF*\n' +
      '\nNama: ' + patient_info.pesakit_nama[0] +
      '\nNo IC: ' + patient_info.pesakit_ic[0] +
      '\nUmur: '  + patient_info.logistik_umur[0] + ' tahun' +
      '\nJantina: ' + patient_info.logistik_jantina[0] +
      '\nBangsa: ' + patient_info.demografi_bangsa[0] +
      '\nAlamat: ' + patient_info.pesakit_alamat[0] +
      '\nNo Tel: '  + patient_info.pesakit_phone[0] +
      '\n\nPekerjaan: ' + patient_info.demografi_pekerjaan[0] +
      '\nComorbid: ' + patient_info.logistik_comorbid[0] +
      '\nTarikh vaksin pertama: ' + patient_info.demografi_vaksin_satu[0] +
      '\nTarikh vaksin kedua: ' + patient_info.demografi_vaksin_dua[0] +
      '\nCovid category: ' + patient_info.logistik_cat[0] +
      '\nTarikh onset gejala: ' + patient_info.demografi_gejala_tarikh[0] +
      '\nJenis gejala: ' + patient_info.demografi_gejala_jenis[0] +
      '\n\nJenis saringan: ' + patient_info.demografi_saringan[0] +
      '\nKluster: ' + patient_info.epid_nama_kluster[0] +
      '\nTarikh sampel pertama: ' + patient_info.demografi_sampel_satu[0] + 
      '\nTarikh sampel kedua: ' + patient_info.demografi_sampel_dua[0] +
      '\nTarikh positif: ' + patient_info.tindakan_tarikh[0] +
      '\nCT-Value: ' + patient_info.keputusan_rdrp[0] + ', ' + patient_info.keputusan_n[0] + ', ' +  patient_info.keputusan_orf[0] +
      '\nIsu sosial: ' + patient_info.logistik_sosial[0];
  return info_segera_txt;
}
