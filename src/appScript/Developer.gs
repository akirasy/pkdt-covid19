function testFunction() {

}

function genBorSiaBulk() {
  let clerking_template_id_override = '12eToLFP13_VXtVge2cA2ZNsXMZYdiPqi8edqPOpe7II';
  let selected_range = SpreadsheetApp.getActiveRange();

  let forloop_start = selected_range.getRowIndex();
  let forloop_end = forloop_start + selected_range.getNumRows();
  for (rowid = forloop_start; rowid < forloop_end; rowid++) {
    generateBorangSiasatan(rowid, clerking_template_id_override);
  }
}

function runOnce() {
  let rowid = 194;
  let current_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('(archive) Kes Positif');
  assignNewTlhNumber(rowid, current_sheet);
}

