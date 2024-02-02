/*

dbgGetDbgFlag(clearTest) - Возвращает значение флага ФлОтладка. Если true и аргумент true, то очищает лист Test.
dbgClearTestSheet() - Очищает и активирует лист Test.
dbgSplitLongString(sStr, maxLngth) - Разбивает длинную строку на набор строк длиной maxLngth.

*/

// 
function dbgGetDbgFlag(clearTest)
{
  const Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const Range = Spreadsheet.getRangeByName('ФлОтладка');

  if (Range != undefined && Range.getValue())
  {
    if (clearTest)
      Spreadsheet.getSheetByName('Test').clear();

    return true;
  }
  return false;
}

// Очистка листа отладки
function dbgClearTestSheet()
{
  SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName('Test')
  .clear()
  .activate();
}

// Разбиваем длинную строку ( >50000 ) на несколько строк по maxLngth символов
function dbgSplitLongString(sStr, maxLngth)
{
  let n = 0;
  let k = maxLngth;
  let sArr = [];
  do {
    sArr.push(sStr.slice(n, k));
    n += maxLngth;
    k += maxLngth;
  } while (sStr.length > n);

  return sArr;
}
