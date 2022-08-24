function testFunction() {

}

function genBorSiaBulk() {
  let clerking_template_id_override = '12eToLFP13_VXtVge2cA2ZNsXMZYdiPqi8edqPOpe7II';
  let selected_range = SpreadsheetApp.getActiveRange();

  let var_source = getVarSource();
  let forloop_start = selected_range.getRowIndex();
  let forloop_end = forloop_start + selected_range.getNumRows();
  for (rowid = forloop_start; rowid < forloop_end; rowid++) {
    generateBorangSiasatan(rowid, var_source, clerking_template_id_override);
  }
}

function runOnce() {
  let active = SpreadsheetApp.getCurrentCell();
  let rowid = active.getRowIndex();
  let current_sheet = active.getSheet();
  // Logger.log(rowid + current_sheet);

  // let rowid = 234;
  // let current_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('(archive) Kes Positif');

  assignNewTlhNumber(rowid, current_sheet);
}

