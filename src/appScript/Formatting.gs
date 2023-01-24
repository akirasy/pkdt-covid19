/**
 * Parse date object into desired `String` output as `dd/mm/yyyy`.
 * @param {Object} arg Any object to check and parse into `String` output.
 */
function parseDate(arg) {
  let output;
  if (arg instanceof Date) { 
    let argInput = new Date(arg);
    output = argInput.getDate() + '/' + (argInput.getMonth() + 1) + '/' + argInput.getFullYear();
  } else {
    output = arg
  };
  return output;
}

/**
 * Set data validation on cells and configure formatting.
 */
function setValidationAndFormatting() {
  let projectVar = getProjectVariables();
  let sheetKesPositif = projectVar.sheetKesPositif;
  let validationList = [
    {sheet:sheetKesPositif, range:'A6:A'  , dateFormat:'d mmm'},
    {sheet:sheetKesPositif, range:'B6:B'  , validationList:['KKBM','KKT', 'KKTL', 'KKL', 'KKS', 'KKK', 'KKKT', 'KKK', 'KKKK', 'HOSHAS','LUAR DAERAH']},
    {sheet:sheetKesPositif, range:'M6:M'  , dateFormat:'d mmm'},
    {sheet:sheetKesPositif, range:'T6:T'  , dateFormat:'d mmm'},
    {sheet:sheetKesPositif, range:'X6:X'  , validationList:['CAT 1', 'CAT 2A', 'CAT 2B', 'CAT 3', 'CAT 4A', 'CAT 4B', 'CAT 5A', 'CAT 5B']},
    {sheet:sheetKesPositif, range:'Y6:Y'  , validationList:['TIDAK VAKSIN','TIDAK LENGKAP', 'LENGKAP', 'BOOSTER']},
    {sheet:sheetKesPositif, range:'Z6:Z'  , dateFormat:'d mmm'},
    {sheet:sheetKesPositif, range:'AA6:AA', validationList:['NA', 'PFIZER', 'CANSINO', 'SINOVAC', 'ASTRA ZENECA']},
    {sheet:sheetKesPositif, range:'AC6:AC', validationList:['YA', 'TIDAK', 'REFUSED']},
    {sheet:sheetKesPositif, range:'AE6:AE', validationList:['MALAYSIA', 'BWN']},
    {sheet:sheetKesPositif, range:'AG6:AG', validationList:['SARINGAN BERGEJALA', 'SARINGAN KONTAK RAPAT', 'BERSASAR', 'SARINGAN KENDIRI', 'SARINGAN PEKERJAAN', 'SARINGAN PENGEMBARA', 'SARINGAN PRE-ADMISSION']},
    {sheet:sheetKesPositif, range:'AI6:AI', dateFormat:'d mmm'},
    {sheet:sheetKesPositif, range:'AN6:AN', validationList:['LOCAL', 'IMPORT A', 'IMPORT B', 'IMPORT C']},
    {sheet:sheetKesPositif, range:'AO6:AO', dateFormat:'d mmm'},
    {sheet:sheetKesPositif, range:'AR6:AR', validationList:['YA']}
  ];

  let newDataValidation = SpreadsheetApp.newDataValidation();
  validationList.forEach(item => {
    let range = item.sheet.getRange(item.range);
    if (item.validationList) {
      let rule = newDataValidation.requireValueInList(item.validationList, true).build();
      range.setDataValidation(rule);
    } else if (item.dateFormat) {
      let rule = newDataValidation.requireDate().setAllowInvalid(false).setHelpText('Use this format -> d mmm yyyy (eg. 12 Sep 2021, 25 Aug 2021)').build();
      range.setDataValidation(rule);
      range.setNumberFormat(item.dateFormat);
    };
  });
}

/**
 * Removes prefix and postfix of whitespaces.
 */
function trimWhitespace() {
  let selectedRange = SpreadsheetApp.getActiveRange();
  selectedRange.trimWhitespace();
}

/**
 * Convert selection to UPPERCASE value.
 */
function toUpperCase() {
  let selectedRange = SpreadsheetApp.getActiveRange();
  let dataList = selectedRange.getValues();
  for (let i=0; i<dataList.length; i++) {
    for (let j=0; j<dataList[i].length; j++) {
      let value = dataList[i][j];
      if (!(value instanceof Date)) {
        value = value.toString().toUpperCase();
      };
    };
  };
  selectedRange.setValues(dataList);
}

/**
 * Remove these special character `-, ', *` from ID/IC value.
 */
function cleanIc() {
  let selectedRange = SpreadsheetApp.getActiveRange();
  let dataList = selectedRange.getValues();
  for (let i=0; i<dataList.length; i++) {
    for (let j=0; j<dataList[i].length; j++) {
      let value = dataList[i][j];
      if (!(value instanceof Date)) {
        value = value.toString().replace(/[-|\'|\*|\s]/g, '');
      };
    };
  };
  selectedRange.setValues(dataList);
}
