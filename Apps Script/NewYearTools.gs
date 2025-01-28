/*

  copySpreadSheet(finYear)
  setTrigers()
  copyRemainingFunds()
  clearDailyExpenses()
  clearPayments()
  clearMetersReadings()
  resetSettings()

*/

function copySpreadSheet(finYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSS = ss.copy("Финансы " + finYear);
  DriveApp
  .getFileById(newSS.getId())
  .moveTo(
    DriveApp
    .getFileById(ss.getId())
    .getParents()
    .next()
  );
  return newSS;
}

function setTrigers() {
  //
  // UpdateOnOpen(e)
  // onOnceAnHour()
  // onOnceADay()
}

function copyRemainingFunds() {
  //
}

function clearDailyExpenses() {
  //
}

function clearPayments() {
  //
}

function clearMetersReadings() {
  //
}

function resetSettings() {
  //
}
