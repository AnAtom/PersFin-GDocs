/*
 getFlag(fName) - Возвращает значение флага ФлfName
 InitVars(sss) - Присваивает значение переменным ss, rDGB, flgDbg и выставляет флаг isInited
 InitSS() - Создает начальную структуру
*/

var ss;             // текущая таблица
var flgDbg = false; // флаг отладки
var rDGB;           // место для вывода отладочной информации
var isInited = false;

function getFlag(fName) {
  const f = ss.getRangeByName('Фл' + fName);
  return f != undefined && f.getValue();
}

function InitVars(sss) {
  if (isInited)
    return;
  Logger.log("Устанавливаем переменные");
  ss = sss;
  rDGB = ss.getSheetByName('dbg').getRange(1, 1);
  flgDbg = getFlag('Отладка');
  isInited = true;
}

function InitSS() {
  InitVars(SpreadsheetApp.getActiveSpreadsheet());
  // Средства Расходы Операции Платежи i Магазины Чеки e-mail dbg
  const Sheets = ['Средства', 'Расходы', 'Операции', 'Платежи', 'i', 'Магазины', 'Чеки e-mail', 'dbg'];
  const SIndexes = [1, 2, 3, 4, 5, 6, 7, 8];
  var sn = '';
  for (let i = 0; i < 8; i++) {
    //
    sn = Sheets[i];
    let sss = ss.getSheetByName(sn);
    if (sheet == null) {
      Logger.log("Создаем лист " + sn);
      ss.insertSheet(sn, SIndexes[i]);
    }
  }
  for (let sn in Sheets) {
    // ss.insertSheet(sn, )
  }

}

function FillFundsSheet() {
  //
  // const range = sheet.getRange('C5:D6');
  // const style = SpreadsheetApp.newTextStyle().setFontSize(15).setUnderline(true).setBold(true).build();
  // range.setTextStyle(style);
  for (let i = 0; i < 4; i++) {
    //
  }

}

function FillSheet() {
  //
}
