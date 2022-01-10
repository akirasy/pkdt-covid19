function generateBorangSiasatan() {
  // init var from sheet: {appScript.gs}
  var var_source = getVarSource();
  var clerking_template_id = var_source.path_clerking_template;
  var target_folder_id = var_source.path_tlh_folder;

  // retrieve user information from {kes positif}
  var rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  var patient_info = getPatientInfo(rowid);

  // check for duplication
  var is_duplicate = var_source.sheet_kes_positif.getRange(rowid, patient_info.siasatan_status[1]);
  
  // nest function in if-else to avoid duplication
  if (is_duplicate.getValue() == 'DONE') {
    SpreadsheetApp.getUi().alert(
      'Warning.',
      'The file already exist.\n\nContinue editing the existing file at:\n' + patient_info.siasatan_url[0] + '\n\nScript will abort.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    
    // create new folder for a fresh day
    var today = new Date();
    if (today.getDate() == var_source.range_today_date.getValue()) { 
      var target_folder_today_id = var_source.range_tlh_folder_today.getValue(); 
      } else {
      var target_folder_today_id = DriveApp.getFolderById(target_folder_id).createFolder(today.toDateString()).getId();
      // update source_var : date
      var_source.range_today_date.setValue(today.getDate());
      var_source.range_tlh_folder_today.setValue(target_folder_today_id);
    }

    // begin to generate file
    // var tlh_number = writeTlhNumber(rowid);  // remove user priviledge to generate TLH number

    // generate filename
    var doc_filename = patient_info.pesakit_nama;
    // var doc_filename = tlh_number + ' ' + patient_info.pesakit_nama; // remove user priviledge to generate TLH number

    // make copy of the newly created file from template and get its url link
    var doc_obj = DriveApp.getFileById(clerking_template_id).makeCopy(doc_filename, DriveApp.getFolderById(target_folder_today_id));

    // mark case as done generated and set URL Link to cell
    var_source.sheet_kes_positif.getRange(rowid, 6).setValue('AWAITING TLH NUMBER');
    is_duplicate.setValue('DONE');
    var_source.sheet_kes_positif.getRange(rowid, patient_info.siasatan_url[1]).setValue(doc_obj.getUrl());

    // Start writing to file
    var body = DocumentApp.openById(doc_obj.getId()).getBody();
    var table_array = body.getTables();

    var tbl_1 = table_array[1];
    tbl_1.getCell(0, 2).setText(patient_info.pesakit_nama);             // Nama
    tbl_1.getCell(1, 2).setText(patient_info.pesakit_ic);               // IC
    tbl_1.getCell(2, 2).setText(patient_info.pesakit_phone);            // Phone
    tbl_1.getCell(3, 2).setText(patient_info.pesakit_alamat);           // Address
    tbl_1.getCell(5, 2).setText(patient_info.demografi_bangsa);         // Race
    tbl_1.getCell(6, 2).setText(patient_info.logistik_umur + ' TAHUN'); // Age
    tbl_1.getCell(7, 2).setText(patient_info.logistik_jantina);         // Gender
    tbl_1.getCell(8, 2).setText(patient_info.demografi_warganegara);    // Warganegara
    tbl_1.getCell(9, 2).setText(patient_info.demografi_mukim);          // Mukim
    tbl_1.getCell(10, 2).setText(patient_info.epid_nama_kluster);       // Kluster
    tbl_1.getCell(11, 2).setText(patient_info.demografi_pekerjaan);     // Pekerjaan
    tbl_1.getCell(12, 2).setText(patient_info.demografi_vaksin_satu);   // Vaksin satu
    tbl_1.getCell(13, 2).setText(patient_info.demografi_vaksin_dua);    // Vaksin dua
    tbl_1.getCell(14, 2).setText(patient_info.demografi_vaksin_tiga);   // Vaksin tiga 

    var tbl_2 = table_array[2];
    tbl_2.getCell(0, 2).setText(patient_info.tindakan_tarikh);          // Date diagnosis
    tbl_2.getCell(1, 2).setText(patient_info.logistik_admit);           // Admitting hospital
    tbl_2.getCell(2, 2).setText(
      patient_info.epid_nama_indeks + ' (' +
      patient_info.epid_id_indeks + ')');                               // Index case name
    tbl_2.getCell(3, 2).setText(patient_info.epid_hubungan);            // Relation to index
    tbl_2.getCell(4, 2).setText(patient_info.epid_lokal);               // Lokal atau import

    var tbl_4 = table_array[3];
    tbl_4.getCell(0, 2).setText(patient_info.demografi_gejala_tarikh);  // Onset date
    tbl_4.getCell(1, 2).setText(patient_info.demografi_gejala_jenis);   // Onset type

    var tbl_5 = table_array[4];
    tbl_5.getCell(0, 2).setText(patient_info.logistik_comorbid);        // Comorbid

    var tbl_6 = table_array[5];
    tbl_6.getCell(0, 2).setText(patient_info.demografi_sampel_satu);    // First sample
    tbl_6.getCell(1, 2).setText(patient_info.demografi_sampel_dua);     // Second sample
    tbl_6.getCell(2, 2).setText(
      patient_info.keputusan_rdrp + ', ' + 
      patient_info.keputusan_n + ', ' +  
      patient_info.keputusan_orf);                                      // Result CtV

    var tbl_10 = table_array[9];
    tbl_10.getCell(0, 0).setText(infoSegeraPenyiasat(rowid));           // Info segera

    var tbl_11 = table_array[10];
    tbl_11.getCell(0, 2).setText(patient_info.penyiasat_nama);          // Investigator name
    tbl_11.getCell(1, 2).setText(patient_info.penyiasat_jawatan);       // Investigator designation
    tbl_11.getCell(2, 2).setText(patient_info.penyiasat_tarikh);        // Investigation date
    tbl_11.getCell(3, 2).setText(patient_info.penyiasat_telefon);       // Investigator phone number

    // Set permission to allow anyone with READ WRITE priviledge and change ownership to Spreadsheet owner
    doc_obj.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
    doc_obj.setOwner(var_source.spreadsheet_owner);
    }
}

function infoSegeraPenyiasat(rowid) {
  var patient_info = getPatientInfo(rowid);

  var info_segera_txt = 
      '*INFO SEGERA KES POSITIF*\n' +
      '\nNama: ' + patient_info.pesakit_nama +
      '\nNo IC: ' + patient_info.pesakit_ic +
      '\nUmur: '  + patient_info.logistik_umur + ' tahun' +
      '\nPekerjaan: ' + patient_info.demografi_pekerjaan +
      '\nCT-Value: ' + patient_info.keputusan_rdrp + ', ' + patient_info.keputusan_n + ', ' +  patient_info.keputusan_orf +
      '\nSampel kali ke: ' + patient_info.epid_sampel_kali +
      '\n\nAlamat: ' + patient_info.pesakit_alamat +
      '\nMukim: '  + patient_info.demografi_mukim +
      '\nTelefon: ' + patient_info.pesakit_phone +
      '\nJenis saringan: ' + patient_info.demografi_saringan +
      '\nKategori jangkitan: ' + patient_info.epid_lokal +
      '\nCC1: ' + patient_info.epid_nama_indeks + ' (' + patient_info.epid_id_indeks + ')' +
      '\nKluster: ' + patient_info.epid_nama_kluster;

  // Show alert box
  SpreadsheetApp.getUi().alert('Info Segera Penyiasat.', info_segera_txt, SpreadsheetApp.getUi().ButtonSet.OK);

  return info_segera_txt;
}

function infoSegeraReferral() {
  var rowid = SpreadsheetApp.getCurrentCell().getRowIndex();
  var patient_info = getPatientInfo(rowid);

  // Set text for info segera
  var info_segera_txt = 
      '*INFO SEGERA KES POSITIF*\n' +
      '\nNama: ' + patient_info.pesakit_nama +
      '\nNo IC: ' + patient_info.pesakit_ic +
      '\nUmur: '  + patient_info.logistik_umur + ' tahun' +
      '\nJantina: ' + patient_info.logistik_jantina +
      '\nBangsa: ' + patient_info.demografi_bangsa +
      '\nAlamat: ' + patient_info.pesakit_alamat +
      '\nNo Tel: '  + patient_info.pesakit_phone +
      '\n\nPekerjaan: ' + patient_info.demografi_pekerjaan +
      '\nComorbid: ' + patient_info.logistik_comorbid +
      '\nTarikh vaksin pertama: ' + patient_info.demografi_vaksin_satu +
      '\nTarikh vaksin kedua: ' + patient_info.demografi_vaksin_dua +
      '\nCovid category: ' + patient_info.logistik_cat +
      '\nTarikh onset gejala: ' + patient_info.demografi_gejala_tarikh +
      '\nJenis gejala: ' + patient_info.demografi_gejala_jenis +
      '\n\nJenis saringan: ' + patient_info.demografi_saringan +
      '\nKluster: ' + patient_info.epid_nama_kluster +
      '\nTarikh sampel pertama: ' + patient_info.demografi_sampel_satu + 
      '\nTarikh sampel kedua: ' + patient_info.demografi_sampel_dua +
      '\nTarikh positif: ' + patient_info.tindakan_tarikh +
      '\nCT-Value: ' + patient_info.keputusan_rdrp + ', ' + patient_info.keputusan_n + ', ' +  patient_info.keputusan_orf +
      '\nIsu sosial: ' + patient_info.logistik_sosial;
  
  // Show alert box
  SpreadsheetApp.getUi().alert('Info Segera Referral.', info_segera_txt, SpreadsheetApp.getUi().ButtonSet.OK);

  return info_segera_txt;
}
