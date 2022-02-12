function getVarSource() {
  // init var from sheet: {appScript.gs}
  let sheet_var_source          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('appScript.gs');

  let spreadsheet_owner         = sheet_var_source.getRange('B14').getValue();
  let spreadsheet_uac_id        = sheet_var_source.getRange('B15').getValue();
  
  let sheet_kes_positif         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B4').getValue());
  let sheet_kes_positif_archive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B5').getValue());
  let sheet_laporan_epid        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B6').getValue());
  let sheet_laporan_cac         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B7').getValue());
  let sheet_cac_pending         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B8').getValue());

  let path_clerking_template    = sheet_var_source.getRange('B11').getValue();
  let path_tlh_folder           = sheet_var_source.getRange('B12').getValue();
  
  let range_tlh_prefix          = sheet_var_source.getRange('B13');
  let range_tlh_max             = sheet_var_source.getRange('B18');
  let range_tlh_folder_today    = sheet_var_source.getRange('B19');
  let range_today_date          = sheet_var_source.getRange('B20');
  let range_unused_tlh          = sheet_var_source.getRange('D4:D25');

  let var_source_json = {
    'sheet_var_source'          : sheet_var_source,
    'spreadsheet_owner'         : spreadsheet_owner,
    'spreadsheet_uac_id'        : spreadsheet_uac_id,
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
  let selected_patient = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns()).getValues();
  let selected_patient_values = selected_patient[0].map(item => { return parseDate(item) });

  let patient_info_json = {
    'tindakan_tarikh'            : [selected_patient_values[0], 1],
    'tindakan_kk_referral'       : [selected_patient_values[1], 2],
    'tindakan_pegawai_referral'  : [selected_patient_values[2], 3],
    'tindakan_pegawai_penyiasat' : [selected_patient_values[3], 4],
    'tindakan_catatan'           : [selected_patient_values[4], 5],
    'pesakit_id'                 : [selected_patient_values[5], 6],
    'pesakit_nama'               : [selected_patient_values[6], 7],
    'pesakit_ic'                 : [selected_patient_values[7], 8],
    'pesakit_alamat'             : [selected_patient_values[8], 9],
    'pesakit_phone'              : [selected_patient_values[9], 10],
    'keputusan_rdrp'             : [selected_patient_values[10], 11],
    'keputusan_n'                : [selected_patient_values[11], 12],
    'keputusan_orf'              : [selected_patient_values[12], 13],
    'keputusan_makmal'           : [selected_patient_values[13], 14],
    'logistik_tarikh_dinilai'    : [selected_patient_values[14], 15],
    'logistik_umur'              : [selected_patient_values[15], 16],
    'logistik_jantina'           : [selected_patient_values[16], 17],
    'logistik_comorbid'          : [selected_patient_values[17], 18],
    'logistik_bmi'               : [selected_patient_values[18], 19],
    'logistik_cat'               : [selected_patient_values[19], 20],
    'logistik_admit'             : [selected_patient_values[20], 21],
    'logistik_sosial'            : [selected_patient_values[21], 22],
    'logistik_catatan'           : [selected_patient_values[22], 23],
    'demografi_bangsa'           : [selected_patient_values[23], 24],
    'demografi_warganegara'      : [selected_patient_values[24], 25],
    'demografi_mukim'            : [selected_patient_values[25], 26],
    'demografi_saringan'         : [selected_patient_values[26], 27],
    'demografi_pekerjaan'        : [selected_patient_values[27], 28],
    'demografi_vaksin_status'    : [selected_patient_values[28], 29],
    'demografi_vaksin_satu'      : [selected_patient_values[29], 30],
    'demografi_vaksin_dua'       : [selected_patient_values[30], 31],
    'demografi_vaksin_tiga'      : [selected_patient_values[31], 32],
    'demografi_vaksin_tempat'    : [selected_patient_values[32], 33],
    'demografi_vaksin_jenis'     : [selected_patient_values[33], 34],
    'demografi_gejala_tarikh'    : [selected_patient_values[34], 35],
    'demografi_gejala_jenis'     : [selected_patient_values[35], 36],
    'demografi_sampel_satu'      : [selected_patient_values[36], 37],
    'demografi_sampel_dua'       : [selected_patient_values[37], 38],
    'epid_nama_kluster'          : [selected_patient_values[38], 39],
    'epid_nama_indeks'           : [selected_patient_values[39], 40],
    'epid_id_indeks'             : [selected_patient_values[40], 41],
    'epid_hubungan'              : [selected_patient_values[41], 42],
    'epid_bil_kontak'            : [selected_patient_values[42], 43],
    'penyiasat_nama'             : [selected_patient_values[43], 44],
    'penyiasat_jawatan'          : [selected_patient_values[44], 45],
    'penyiasat_telefon'          : [selected_patient_values[45], 46],
    'penyiasat_tarikh'           : [selected_patient_values[46], 47],
    'epid_minggu'                : [selected_patient_values[47], 48],
    'epid_status'                : [selected_patient_values[48], 49],
    'epid_sebab_mati'            : [selected_patient_values[49], 50],
    'epid_jenis_ujian'           : [selected_patient_values[50], 51],
    'epid_sampel_kali'           : [selected_patient_values[51], 52],
    'epid_lokal'                 : [selected_patient_values[52], 53],
    'epid_origin'                : [selected_patient_values[53], 54],
    'generate_now'               : [selected_patient_values[54], 55],
    'siasatan_url'               : [selected_patient_values[55], 56],
    'siasatan_status'            : [selected_patient_values[56], 57],
    'reten_epid'                 : [selected_patient_values[57], 58],
    'reten_catatan'              : [selected_patient_values[58], 59],
  }
  return patient_info_json
}

// Check if date to parse toDateString()
function parseDate(arg) {
  let output = '';
  if (arg instanceof Date) { 
    let arg_input = new Date(arg);
    output = arg_input.getDate() + '/' + (arg_input.getMonth() + 1) + '/' + arg_input.getFullYear();
  } else { output = arg }
  return output;
}

// Change lowercase to uppercase
function toUpperCase() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().toUpperCase();
      }
    }
  }
  selected_range.setValues(data_list);
}

// Convert newline value to oneline only
function toOneLine() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/\n/g, '  ');
      }
    }
  }
  selected_range.setValues(data_list).trimWhitespace();
}

// Removes dashes, star, spaces and apostrophy
function cleanIc() {
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let selected_range = activeSheet.getActiveRange();
  let data_list = selected_range.getValues();
  for (let i = 0; i < data_list.length; i++) {
    for (let j = 0; j < data_list[i].length; j++) {
      if (!(data_list[i][j] instanceof Date)) {
        data_list[i][j] = data_list[i][j].toString().replace(/[-|\'|\*|\s]/g,'');
      }
    }
  }
  selected_range.setValues(data_list);
}
