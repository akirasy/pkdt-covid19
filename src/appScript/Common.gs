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

  let var_source_json = {
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
    'range_today_date'          : range_today_date
  }
  
  return var_source_json;
}

function getPatientInfo(rowid, var_source) {
  // retrieve patient_info from {kes positif}
  let selected_patient = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns()).getValues();
  let selected_patient_values = selected_patient[0].map(item => { return parseDate(item) });

  // make sure number of column is the same as the named_value = commented number at end of line is for reference
  let named_value = [
    'tindakan_tarikh'               , // 1 
    'tindakan_kk_referral'          , // 2
    'tindakan_pegawai_referral'     , // 3
    'tindakan_pegawai_penyiasat'    , // 4
    'tindakan_catatan'              , // 5
    'pesakit_id'                    , // 6
    'pesakit_nama'                  , // 7
    'pesakit_ic'                    , // 8
    'pesakit_alamat'                , // 9
    'pesakit_phone'                 , // 10
    'keputusan_rdrp'                , // 11
    'keputusan_n'                   , // 12
    'keputusan_orf'                 , // 13
    'keputusan_makmal'              , // 14
    'logistik_tarikh_dinilai'       , // 15
    'logistik_umur'                 , // 16
    'logistik_jantina'              , // 17
    'logistik_comorbid'             , // 18
    'logistik_bmi'                  , // 19
    'logistik_cat'                  , // 20
    'logistik_admit'                , // 21
    'logistik_sosial'               , // 22
    'logistik_catatan'              , // 23
    'demografi_bangsa'              , // 24
    'demografi_warganegara'         , // 25
    'demografi_mukim'               , // 26
    'demografi_saringan'            , // 27
    'demografi_pekerjaan'           , // 28
    'demografi_vaksin_status'       , // 29
    'demografi_vaksin_satu'         , // 30
    'demografi_vaksin_dua'          , // 31
    'demografi_vaksin_tiga'         , // 32
    'demografi_vaksin_tempat'       , // 33
    'demografi_vaksin_jenis'        , // 34
    'demografi_gejala_tarikh'       , // 35
    'demografi_gejala_jenis'        , // 36
    'demografi_sampel_satu'         , // 37
    'demografi_sampel_dua'          , // 38
    'epid_nama_kluster'             , // 39
    'epid_nama_indeks'              , // 40
    'epid_id_indeks'                , // 41
    'epid_hubungan'                 , // 42
    'epid_bil_kontak'               , // 43
    'penyiasat_nama'                , // 44
    'penyiasat_jawatan'             , // 45
    'penyiasat_telefon'             , // 46
    'penyiasat_tarikh'              , // 47
    'epid_minggu'                   , // 48
    'epid_status'                   , // 49
    'epid_sebab_mati'               , // 50
    'epid_jenis_ujian'              , // 51
    'epid_sampel_kali'              , // 52
    'epid_lokal'                    , // 53
    'epid_origin'                   , // 54
    'generate_now'                  , // 55
    'siasatan_url'                  , // 56
    'siasatan_status'               , // 57
    'reten_epid'                    , // 58
    'reten_catatan'                 , // 59
  ]

  // Create string to convert to JSON and set as function return value
  let patient_info = '{\n';
  named_value.forEach((name, i) => {
    patient_info += '"' + name + '" : ["' + selected_patient_values[i] + '", ' + (i+1).toString() + '],\n';
  })
  patient_info += '"end" : ["", 0]\n}';
  return JSON.parse(patient_info);
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
