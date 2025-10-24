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
  finYear = 2026;

  // Создаем копию и помещаем в ту же папку
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // В таблице создаем тригеры
  // UpdateOnOpen(e)
  ScriptApp.newTrigger('UpdateOnOpen').forSpreadsheet(ss).onOpen().create();
  // onOnceAnHour()
  ScriptApp.newTrigger('onOnceAnHour').timeBased().everyHours(1).create();
  // onOnceADay()
  ScriptApp.newTrigger('myFunction').timeBased().atHour(4).everyDays(1).create();
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
