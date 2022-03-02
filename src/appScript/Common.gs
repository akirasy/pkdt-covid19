function getVarSource() {
  // init var from sheet: {appScript.gs}
  let active_spreadsheet        = SpreadsheetApp.getActiveSpreadsheet();
  let sheet_var_source          = active_spreadsheet.getSheetByName('appScript.gs');

  let spreadsheet_owner         = sheet_var_source.getRange('B14').getValue();
  let spreadsheet_uac_id        = sheet_var_source.getRange('B15').getValue();
  let spreadsheet_archive_id    = sheet_var_source.getRange('B16').getValue();
  
  let sheet_kes_positif         = active_spreadsheet.getSheetByName(sheet_var_source.getRange('B4').getValue());
  let sheet_kes_positif_archive = active_spreadsheet.getSheetByName(sheet_var_source.getRange('B5').getValue());
  let sheet_laporan_epid        = active_spreadsheet.getSheetByName(sheet_var_source.getRange('B6').getValue());
  let sheet_laporan_cac         = active_spreadsheet.getSheetByName(sheet_var_source.getRange('B7').getValue());
  let sheet_cac_pending         = active_spreadsheet.getSheetByName(sheet_var_source.getRange('B8').getValue());

  let path_clerking_template    = sheet_var_source.getRange('B11').getValue();
  let path_tlh_folder           = sheet_var_source.getRange('B12').getValue();
  
  let range_tlh_prefix          = sheet_var_source.getRange('B13');
  let range_tlh_max             = sheet_var_source.getRange('B19');
  let range_tlh_folder_today    = sheet_var_source.getRange('B20');
  let range_today_date          = sheet_var_source.getRange('B21');
  let range_unused_tlh          = sheet_var_source.getRange('D4:D72');

  let var_source_json = {
    'sheet_var_source'          : sheet_var_source,
    'spreadsheet_owner'         : spreadsheet_owner,
    'spreadsheet_uac_id'        : spreadsheet_uac_id,
    'spreadsheet_archive_id'    : spreadsheet_archive_id,
    'sheet_kes_positif'         : sheet_kes_positif,
    'sheet_kes_positif_archive' : sheet_kes_positif_archive,
    'sheet_laporan_epid'        : sheet_laporan_epid,
    'sheet_laporan_cac'         : sheet_laporan_cac,
    'sheet_cac_pending'         : sheet_cac_pending,
    'path_clerking_template'    : path_clerking_template,
    'path_tlh_folder'           : path_tlh_folder,
    'range_tlh_prefix'          : range_tlh_prefix,
    'range_tlh_max'             : range_tlh_max,
    'range_tlh_folder_today'    : range_tlh_folder_today,
    'range_today_date'          : range_today_date,
    'range_unused_tlh'          : range_unused_tlh
  }
  
  return var_source_json;
}

function getPatientInfo(rowid, var_source) {
  let selected_patient = var_source.sheet_kes_positif.getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns()).getValues();
  let selected_patient_values = selected_patient[0].map(item => { return parseDate(item) });

  let patient_info_json = {
    'tarikh_notifikasi'       : [selected_patient_values[ 0],  1],
    'kk_referral'             : [selected_patient_values[ 1],  2],
    'pegawai_referral'        : [selected_patient_values[ 2],  3],
    'pegawai_penyiasat'       : [selected_patient_values[ 3],  4],
    'catatan_pencarian'       : [selected_patient_values[ 4],  5],
    'id_kes'                  : [selected_patient_values[ 5],  6],
    'nama'                    : [selected_patient_values[ 6],  7],
    'ic'                      : [selected_patient_values[ 7],  8],
    'umur'                    : [selected_patient_values[ 8],  9],
    'jantina'                 : [selected_patient_values[ 9], 10],
    'alamat'                  : [selected_patient_values[10], 11],
    'phone'                   : [selected_patient_values[11], 12],
    'tarikh_sampel'           : [selected_patient_values[12], 13],
    'status_sampel'           : [selected_patient_values[13], 14],
    'ctval_rdrp'              : [selected_patient_values[14], 15],
    'ctval_n'                 : [selected_patient_values[15], 16],
    'ctval_orf'               : [selected_patient_values[16], 17],
    'fasiliti_makmal'         : [selected_patient_values[17], 18],
    'jenis_ujian'             : [selected_patient_values[18], 19],
    'tarikh_dinilai'          : [selected_patient_values[19], 20],
    'catatan_umum'            : [selected_patient_values[20], 21],
    'comorbid'                : [selected_patient_values[21], 22],
    'bmi'                     : [selected_patient_values[22], 23],
    'covid_category'          : [selected_patient_values[23], 24],
    'status_vaksin'           : [selected_patient_values[24], 25],
    'jenis_vaksin'            : [selected_patient_values[25], 26],
    'admit'                   : [selected_patient_values[26], 27],
    'bangsa'                  : [selected_patient_values[27], 28],
    'warganegara'             : [selected_patient_values[28], 29],
    'mukim'                   : [selected_patient_values[29], 30],
    'jenis_saringan'          : [selected_patient_values[30], 31],
    'pekerjaan'               : [selected_patient_values[31], 32],
    'tarikh_onset'            : [selected_patient_values[32], 33],
    'jenis_gejala'            : [selected_patient_values[33], 34],
    'catatan_siasatan'        : [selected_patient_values[34], 35],
    'nama_kes_indeks'         : [selected_patient_values[35], 36],
    'bilangan_kontak_rapat'   : [selected_patient_values[36], 37],
    'kategori_jangkitan'      : [selected_patient_values[37], 38],
    'tarikh_siasatan'         : [selected_patient_values[38], 39],
    'nama_penyiasat'          : [selected_patient_values[39], 40],
    'jawatan_penyiasat'       : [selected_patient_values[40], 41],
    'generate_sekarang'       : [selected_patient_values[41], 42],
    'url_siasatan'            : [selected_patient_values[42], 43],
    'status_siasatan'         : [selected_patient_values[43], 44],
    'epid_daerah'             : [selected_patient_values[44], 45],
    'catatan_epid'            : [selected_patient_values[45], 46]
  }
  return patient_info_json
}
