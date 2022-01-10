function getVarSource() {
  // init var from sheet: {appScript.gs}
  var sheet_var_source          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('appScript.gs');

  var spreadsheet_owner         = sheet_var_source.getRange('B14').getValue();
  var spreadsheet_uac_id        = sheet_var_source.getRange('B15').getValue();
  
  var sheet_kes_positif         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B4').getValue());
  var sheet_kes_positif_archive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B5').getValue());
  var sheet_laporan_epid        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B6').getValue());
  var sheet_laporan_cac         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B7').getValue());
  var sheet_cac_pending         = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_var_source.getRange('B8').getValue());

  var path_clerking_template    = sheet_var_source.getRange('B11').getValue();
  var path_tlh_folder           = sheet_var_source.getRange('B12').getValue();
  
  var range_tlh_prefix          = sheet_var_source.getRange('B13');
  var range_tlh_max             = sheet_var_source.getRange('B18');
  var range_tlh_folder_today    = sheet_var_source.getRange('B19');
  var range_today_date          = sheet_var_source.getRange('B20');


  var var_source_json = {
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

function getPatientInfo(rowid) {
  // init var from sheet: {appScript.gs}
  var var_source = getVarSource();

  // retrieve information from {kes positif}
  var selected_patient = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, var_source.sheet_kes_positif.getMaxColumns());
  var patient_info_json = {
    'tindakan_tarikh'               : parseDate(selected_patient.getValues()[0][0]),
    'tindakan_kk_referral'          : selected_patient.getValues()[0][1],
    'tindakan_pegawai_referral'     : selected_patient.getValues()[0][2],
    'tindakan_pegawai_penyiasat'    : selected_patient.getValues()[0][3],
    'tindakan_catatan'              : selected_patient.getValues()[0][4],
    'pesakit_id'                    : selected_patient.getValues()[0][5],
    'pesakit_nama'                  : selected_patient.getValues()[0][6],
    'pesakit_ic'                    : selected_patient.getValues()[0][7],
    'pesakit_alamat'                : selected_patient.getValues()[0][8],
    'pesakit_phone'                 : selected_patient.getValues()[0][9],
    'keputusan_rdrp'                : selected_patient.getValues()[0][10],
    'keputusan_n'                   : selected_patient.getValues()[0][11],
    'keputusan_orf'                 : selected_patient.getValues()[0][12],
    'keputusan_makmal'              : selected_patient.getValues()[0][13],
    'logistik_tarikh_dinilai'       : parseDate(selected_patient.getValues()[0][14]),
    'logistik_umur'                 : selected_patient.getValues()[0][15],
    'logistik_jantina'              : selected_patient.getValues()[0][16],
    'logistik_comorbid'             : selected_patient.getValues()[0][17],
    'logistik_bmi'                  : selected_patient.getValues()[0][18],
    'logistik_cat'                  : selected_patient.getValues()[0][19],
    'logistik_admit'                : selected_patient.getValues()[0][20],
    'logistik_sosial'               : selected_patient.getValues()[0][21],
    'logistik_catatan'              : selected_patient.getValues()[0][22],
    'demografi_bangsa'              : selected_patient.getValues()[0][23],
    'demografi_warganegara'         : selected_patient.getValues()[0][24],
    'demografi_mukim'               : selected_patient.getValues()[0][25],
    'demografi_saringan'            : selected_patient.getValues()[0][26],
    'demografi_pekerjaan'           : selected_patient.getValues()[0][27],
    'demografi_vaksin_status'       : selected_patient.getValues()[0][28],
    'demografi_vaksin_satu'         : parseDate(selected_patient.getValues()[0][29]),
    'demografi_vaksin_dua'          : parseDate(selected_patient.getValues()[0][30]),
    'demografi_vaksin_tiga'         : parseDate(selected_patient.getValues()[0][31]),
    'demografi_vaksin_tempat'       : selected_patient.getValues()[0][32],
    'demografi_vaksin_jenis'        : selected_patient.getValues()[0][33],
    'demografi_gejala_tarikh'       : parseDate(selected_patient.getValues()[0][34]),
    'demografi_gejala_jenis'        : selected_patient.getValues()[0][35],
    'demografi_sampel_satu'         : parseDate(selected_patient.getValues()[0][36]),
    'demografi_sampel_dua'          : parseDate(selected_patient.getValues()[0][37]),
    'epid_nama_kluster'             : selected_patient.getValues()[0][38],
    'epid_nama_indeks'              : selected_patient.getValues()[0][39],
    'epid_id_indeks'                : selected_patient.getValues()[0][40],
    'epid_hubungan'                 : selected_patient.getValues()[0][41],
    'epid_bil_kontak'               : selected_patient.getValues()[0][42],
    'penyiasat_nama'                : selected_patient.getValues()[0][43],
    'penyiasat_jawatan'             : selected_patient.getValues()[0][44],
    'penyiasat_telefon'             : selected_patient.getValues()[0][45],
    'penyiasat_tarikh'              : parseDate(selected_patient.getValues()[0][46]),
    'epid_minggu'                   : selected_patient.getValues()[0][47],
    'epid_status'                   : selected_patient.getValues()[0][48],
    'epid_sebab_mati'               : selected_patient.getValues()[0][49],
    'epid_jenis_ujian'              : selected_patient.getValues()[0][50],
    'epid_sampel_kali'              : selected_patient.getValues()[0][51],
    'epid_lokal'                    : selected_patient.getValues()[0][52],
    'epid_origin'                   : selected_patient.getValues()[0][53],
    'siasatan_url'                  : [selected_patient.getValues()[0][54], 55],
    'siasatan_status'               : [selected_patient.getValues()[0][55], 56],
    'reten_epid'                    : [selected_patient.getValues()[0][56], 57],
    'reten_cac'                     : [selected_patient.getValues()[0][57], 58],
    'reten_catatan'                 : [selected_patient.getValues()[0][58], 59]
  }
  return patient_info_json
}

// Move completed case to archive
function moveToArchive(selected_range) {
  var var_source = getVarSource();
  for (let i = 0; i < selected_range.getNumRows(); i++) {
    let rowid = selected_range.getRowIndex() + i;
    Logger.log('Processing rowid: ' + rowid);

    // Check if all are done
    let patient_info = getPatientInfo(rowid);
    let reten_cac = patient_info.reten_cac[0];
    let reten_epid = patient_info.reten_epid[0];
    let siasatan_status = patient_info.siasatan_status[0];

/*  // Remove CAC function because it is not used
---------------------------------------------------------
    // Set conditional value
    let isAllDone = Boolean();
    if (reten_cac == 'DONE' && reten_epid == 'DONE' && siasatan_status == 'DONE') {
      isAllDone = true;
    } else { isAllDone = false; }

    // Set conditional value
    let isSiasatEpidDone = Boolean();
    if (reten_epid == 'DONE' && siasatan_status == 'DONE') {
      isSiasatEpidDone = true;
    } else { isSiasatEpidDone = false; }    

    // Move completed case to archive
    if (isAllDone) {
      Logger.log('All marked as DONE.')
      let selected_row_range = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, SpreadsheetApp.getActiveSheet().getMaxColumns());
      let last_archive_range = var_source.sheet_kes_positif_archive.getRange(var_source.sheet_kes_positif_archive.getLastRow() + 1, 1);
      selected_row_range.copyTo(last_archive_range);
      selected_row_range.clear();
      // Mark with color gray
      selected_row_range.setBackground('#cccccc');
    } else if (isSiasatEpidDone) {
      Logger.log('CAC not done yet')
      let selected_row_range = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, SpreadsheetApp.getActiveSheet().getMaxColumns());
      let last_cac_range     = var_source.sheet_cac_pending.getRange(var_source.sheet_cac_pending.getLastRow() + 1, 1);
      selected_row_range.copyTo(last_cac_range);
      selected_row_range.clear();
      // Mark with color gray
      selected_row_range.setBackground('#cccccc');
    }
--------------------------------------------------------- */

    // Set conditional value
    let isAllDone = Boolean();
    if (reten_epid == 'DONE' && siasatan_status == 'DONE') {
      isAllDone = true;
    } else { isAllDone = false; }    

    // Move completed case to archive
    if (isAllDone) {
      Logger.log('All marked as DONE.')
      let selected_row_range = SpreadsheetApp.getActiveSheet().getRange(rowid, 1, 1, SpreadsheetApp.getActiveSheet().getMaxColumns());
      let last_archive_range = var_source.sheet_kes_positif_archive.getRange(var_source.sheet_kes_positif_archive.getLastRow() + 1, 1);
      selected_row_range.copyTo(last_archive_range);
      selected_row_range.clear();
      // Mark with color gray
      selected_row_range.setBackground('#cccccc');
    }
  }
}

// Check if date to parse toDateString()
function parseDate(arg) {
  if (arg instanceof Date) { 
    var arg_input = new Date(arg);
    var output = arg_input.getDate() + '/' + (arg_input.getMonth() + 1) + '/' + arg_input.getFullYear();

  } else { 
    var output = arg;
  }
  return output;
}
