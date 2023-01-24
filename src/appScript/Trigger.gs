/**
 * Trigger function -> Generate borang siasatan at intervals.
 */
function triggerGenerateBorangSiasatan() {
  let initialTime = new Date();
  let projectVar = getProjectVariables();
  let headerKey = getHeaderKey(projectVar.sheetKesPositif);

  let sheetKesPositif = projectVar.sheetKesPositif;
  let genActionQue = sheetKesPositif.getRange(1, headerKey.gen_action+1, sheetKesPositif.getLastRow(), 1).getValues();
  let selectedRowid = genActionQue.map((item, index) => { if (item[0] == 'YA') { return index+1 } });

  selectedRowid.filter(item => item).forEach(rowid => {
    if (isEnoughTime(initialTime, 15)) {
      generateBorangSiasatan(projectVar, headerKey, rowid);
      sheetKesPositif.getRange(rowid, headerKey.gen_action+1).setValue('');
    };
  });
}

/**
 * Trigger function -> Move entry to archive at intervals.
 */
function triggerMoveToArchive() {
  let projectVar = getProjectVariables();
  let sheetKesPositif = projectVar.sheetKesPositif;
  let selectedRange = sheetKesPositif.getRange(1, 1, sheetKesPositif.getLastRow());
  moveToArchive(selectedRange);
}

/**
 * Trigger function -> Add requesting access to user at intervals.
 */
function triggerGrantPermission() {
  grantPermission();
}
