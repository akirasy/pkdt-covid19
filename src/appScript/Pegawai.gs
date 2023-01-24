/**
 * Collect single patient/row data and generate borang siasatan in GoogleDoc.
 * @param {Object} projectVar Instance of `getProjectVariables()`.
 * @param {Object} headerKey Instance of `getHeaderKey(SheetKesPositif)`.
 * @param {Number} rowid Row Id.
 * @param {String} clerkingTemplateId Clerking template ID in string format (optional).
 */
function generateBorangSiasatan(projectVar, headerKey, rowid, clerkingTemplateId) {
  Logger.log('Begin to generate file for row: ' + rowid);
  let sheetKesPositif = projectVar.sheetKesPositif;

  // STEP 0 : Initializing variables
  if (!clerkingTemplateId) { clerkingTemplateId = projectVar.valueClerkingTemplateId };
  let patientInfo = sheetKesPositif.getRange(rowid, 1, 1, sheetKesPositif.getMaxColumns()).getValues()[0];

  if (patientInfo[headerKey.gen_status_siasatan] != 'DONE') {
    // STEP 1 : Prepare folder
    let today = new Date();
    let targetFolderId;
    if (today.getDate() == projectVar.rangeTodayDate.getValue()) {
      targetFolderId = projectVar.rangeGeneratedFolderToday.getValue();
    } else {
      targetFolderId = DriveApp.getFolderById(projectVar.valueGeneratedFolderMain).createFolder(today.toDateString()).getId();
      projectVar.rangeTodayDate.setValue(today.getDate());
      projectVar.rangeGeneratedFolderToday.setValue(targetFolderId);
    }

    // STEP 2 : Prepare file
    let docFilename = patientInfo[headerKey.nama];
    let docObj = DriveApp.getFileById(clerkingTemplateId).makeCopy(docFilename, DriveApp.getFolderById(targetFolderId));

    // STEP 3 : Write to file
    let body = DocumentApp.openById(docObj.getId()).getBody();
    let tableArray = body.getTables();
    let tbl1 = tableArray[1];
    tbl1.getCell( 0, 2).setText(patientInfo[headerKey.nama]);                  // Nama
    tbl1.getCell( 1, 2).setText(patientInfo[headerKey.id_ic]);                 // IC
    tbl1.getCell( 2, 2).setText(patientInfo[headerKey.phone]);                 // Phone
    tbl1.getCell( 3, 2).setText(patientInfo[headerKey.alamat]);                // Address
    tbl1.getCell( 5, 2).setText(patientInfo[headerKey.bangsa]);                // Race
    tbl1.getCell( 6, 2).setText(patientInfo[headerKey.umur] + ' TAHUN');       // Age
    tbl1.getCell( 7, 2).setText(patientInfo[headerKey.jantina]);               // Gender
    tbl1.getCell( 8, 2).setText(patientInfo[headerKey.warganegara]);           // Warganegara
    tbl1.getCell( 9, 2).setText(patientInfo[headerKey.mukim]);                 // Mukim
    tbl1.getCell(11, 2).setText(patientInfo[headerKey.pekerjaan]);             // Pekerjaan
    tbl1.getCell(12, 2).setText(patientInfo[headerKey.vaksin_status]);         // Vaksin satu
    tbl1.getCell(13, 2).setText(patientInfo[headerKey.vaksin_tarikh_booster]); // Vaksin dua
    tbl1.getCell(14, 2).setText(patientInfo[headerKey.vaksin_jenis]);          // Vaksin tiga 

    let tbl2 = tableArray[2];
    tbl2.getCell(0, 2).setText(patientInfo[headerKey.tarikh_notifikasi]);  // Date diagnosis
    tbl2.getCell(1, 2).setText(patientInfo[headerKey.rawat_admit]);        // Admitting hospital
    tbl2.getCell(2, 2).setText(patientInfo[headerKey.kes_indeks]);         // Index case name
    tbl2.getCell(3, 2).setText('');                                        // Relation to index
    tbl2.getCell(4, 2).setText(patientInfo[headerKey.kategori_jangkitan]); // Lokal atau import

    let tbl4 = tableArray[3];
    tbl4.getCell(0, 2).setText(patientInfo[headerKey.tarikh_onset]); // Onset date
    tbl4.getCell(1, 2).setText(patientInfo[headerKey.jenis_gejala]); // Onset type

    let tbl5 = tableArray[4];
    tbl5.getCell(0, 2).setText(patientInfo[headerKey.comorbid]); // Comorbid

    let tbl6 = tableArray[5];
    tbl6.getCell(0, 2).setText(patientInfo[headerKey.sampel_tarikh]); // First sample
    tbl6.getCell(1, 2).setText(patientInfo[headerKey.sampel_status]); // Second sample
    tbl6.getCell(2, 2).setText( //sampel_rdrp	sampel_n	sampel_orf
      patientInfo[headerKey.sampel_rdrp] + ', ' + 
      patientInfo[headerKey.sampel_n]    + ', ' +  
      patientInfo[headerKey.sampel_orf]);                             // Result CtV

    let tbl10 = tableArray[9];
    tbl10.getCell(0, 0).setText(infoSegeraPenyiasat(headerKey, patientInfo)); // Info segera

    let tbl11 = tableArray[10];
    tbl11.getCell(0, 2).setText(patientInfo[headerKey.penyiasat_nama]);    // Investigator name
    tbl11.getCell(1, 2).setText(patientInfo[headerKey.penyiasat_jawatan]); // Investigator designation
    tbl11.getCell(2, 2).setText(patientInfo[headerKey.siasatan_tarikh]);   // Investigation date
    tbl11.getCell(3, 2).setText('');                                       // Investigator phone number

    // STEP 4 : Mark as DONE
    sheetKesPositif.getRange(rowid, headerKey.id_covid+1).setValue('AWAITING REG NUMBER');
    sheetKesPositif.getRange(rowid, headerKey.gen_status_siasatan+1).setValue('DONE');
    sheetKesPositif.getRange(rowid, headerKey.gen_url+1).setValue(docObj.getUrl());
    Logger.log('-- Success!');
  };
}

/**
 * Generate text summary of selected case.
 * @param {Array} patientInfoValues values of selected patient.
 */
function infoSegeraPenyiasat(headerKey, patientInfoValues) {
  let output = 
      '*INFO SEGERA KES POSITIF*\n' +
      '\nNama: '      + patientInfoValues[headerKey.nama]      +
      '\nNo IC: '     + patientInfoValues[headerKey.id_ic]     +
      '\nUmur: '      + patientInfoValues[headerKey.umur]      + ' tahun' +
      '\nPekerjaan: ' + patientInfoValues[headerKey.pekerjaan] +
      '\nCT-Value: '  + 
        patientInfoValues[headerKey.sampel_rdrp] + ', ' + 
        patientInfoValues[headerKey.sampel_n]    + ', ' +  
        patientInfoValues[headerKey.sampel_orf]  +
      '\nSampel kali ke: '     + patientInfoValues[headerKey.sampel_status]  +
      '\n\nAlamat: '           + patientInfoValues[headerKey.alamat]         +
      '\nMukim: '              + patientInfoValues[headerKey.mukim]          +
      '\nTelefon: '            + patientInfoValues[headerKey.phone]          +
      '\nJenis saringan: '     + patientInfoValues[headerKey.jenis_saringan] +
      '\nKategori jangkitan: ' + patientInfoValues[headerKey.kategori_jangkitan];
  return output;
}

/**
 * Collect single patient/row data and generate laporan epid in sheet `Laporan Epid`.
 * @param {Object} projectVar Instance of `getProjectVariables()`.
 * @param {Object} headerKey Instance of `getHeaderKey(SheetKesPositif)`.
 * @param {Number} rowid Row Id.
 */
function generateLaporanEpid(projectVar, headerKey, rowid) {
  Logger.log('Begin to generate Laporan Epid for: ' + rowid);
  let sheetKesPositif = projectVar.sheetKesPositif;
  let sheetLaporanEpid = projectVar.sheetLaporanEpid;
  let patientInfo = sheetKesPositif.getRange(rowid, 1, 1, sheetKesPositif.getMaxColumns()).getValues()[0];

  // STEP 1 : Generate patientID number, rename file
  let generatedId = projectVar.rangePatientIdCurrent.getValue() + 1;
  projectVar.rangePatientIdCurrent.setValue(generatedId);
  let generatedIdName = projectVar.valuePatientIdPrefix + generatedId;
  let fileUrl = sheetKesPositif.getRange(rowid, headerKey.gen_url+1).getValue();
  let fileId = fileUrl.match(/[-\w]{25,}/);
  DriveApp.getFileById(fileId).setName(generatedIdName + ' ' + patientInfo[headerKey.nama]);

  // STEP 2 : Arrange values for laporan epid
  let targetLaporanEpid = sheetLaporanEpid.getRange(sheetLaporanEpid.getLastRow() + 1, 1, 1, sheetLaporanEpid.getMaxColumns());
  targetLaporanEpid.setValues([[
    '',                                           // 1. epid week
    generatedIdName,                              // 2. no TLH
    '',                                           // 3. no KKM
    '',                                           // 4. Tarikh daftar
    patientInfo[headerKey.nama],                  // 5. nama
    'PAHANG',                                     // 6. negeri
    'TEMERLOH',                                   // 7. daerah
    patientInfo[headerKey.mukim],                 // 8. mukim
    patientInfo[headerKey.rawat_admit],           // 9. pusat rawatan
    patientInfo[headerKey.id_ic],                 // 10. IC
    patientInfo[headerKey.umur],                  // 11. umur
    patientInfo[headerKey.jantina],               // 12. jantina
    patientInfo[headerKey.bangsa],                // 13. kaum
    patientInfo[headerKey.warganegara],           // 14. warganegara
    patientInfo[headerKey.jenis_saringan],        // 15. kluster
    patientInfo[headerKey.pekerjaan],             // 16. pekerjaan
    patientInfo[headerKey.comorbid],              // 17. comorbid
    patientInfo[headerKey.alamat],                // 18. alamat
    patientInfo[headerKey.phone],                 // 19. telefon
    patientInfo[headerKey.tarikh_notifikasi],     // 20. tarikh admit
    patientInfo[headerKey.bil_kontak_rapat],      // 21. bilangan kontak
    patientInfo[headerKey.tarikh_onset],          // 22. tarikh onset
    patientInfo[headerKey.jenis_gejala],          // 23. jenis simptom
    'HIDUP',                                      // 24. status hidup
    'N/A',                                        // 25. sebab mati
    patientInfo[headerKey.sampel_makmal],         // 26. makmal
    patientInfo[headerKey.sampel_jenis],          // 27. jenis ujian
    patientInfo[headerKey.tarikh_notifikasi],     // 28. tarikh positif
    patientInfo[headerKey.kategori_jangkitan],    // 29. lokal import
    '',                                           // 30. origin import
    patientInfo[headerKey.kes_indeks],            // 31. catatan
    '',                                           // 32. tarikh discharge
    patientInfo[headerKey.sampel_rdrp] + ', ' + 
    patientInfo[headerKey.sampel_n] + ', ' +  
    patientInfo[headerKey.sampel_orf],            // 33. ct value
    patientInfo[headerKey.sampel_tarikh],         // 34. tarikh sampel
    '',                                           // 35. kategori lain
    '',                                           // 36. vaksin satu
    '',                                           // 37. vaksin dua
    patientInfo[headerKey.covid_cat],             // 38. covid category
    patientInfo[headerKey.vaksin_status],         // 39. status vaksin
    patientInfo[headerKey.vaksin_jenis],          // 40. jenis vaksin
    patientInfo[headerKey.vaksin_tarikh_booster], // 41. tarikh vaksin booster
    '',                                           // 42. tempoh daftar kes
    '',                                           // 43. tempoh vaksin elapsed
    patientInfo[headerKey.sampel_status],         // 44. sampel kali ke
    patientInfo[headerKey.catatan_epid]           // 45. catatan utk mo epid daerah
  ]]);

  // STEP 2 : Mark row as DONE
  sheetKesPositif.getRange(rowid, headerKey.id_covid+1).setValue(generatedIdName);
  sheetKesPositif.getRange(rowid, headerKey.gen_status_epid+1).setValue('DONE');
}

/**
 * Generate Laporan Epid on selected range.
 */
function actionGenerateLaporanEpid() {
  let projectVar       = getProjectVariables();
  let headerKey        = getHeaderKey(projectVar.sheetKesPositif);
  let selectedRange    = projectVar.sheetKesPositif.getActiveRange();
  let selectedRowIndex = selectedRange.getRowIndex();

  let conditionalArray = projectVar.sheetKesPositif.getRange(selectedRowIndex, headerKey.gen_url+1, selectedRange.getNumRows(), 3).getValues();
  let selectedRowid = conditionalArray.map((item, index) => {
    if (item[0] != '' && item[1] == 'DONE' && item[2] != 'DONE') {
      let rowid = index + selectedRowIndex;
      return rowid
    };
  });
  selectedRowid.filter(item => item).forEach(rowid => {
    generateLaporanEpid(projectVar, headerKey, rowid);
  });
}
