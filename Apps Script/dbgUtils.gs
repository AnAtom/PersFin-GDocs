/*

 dbgGetFlag(needClear) - Возвращает значение флага ФлОтладка. Если true и аргумент true, то очищает лист dbg.
 dbgClearSheet() - Очищает и активирует лист dbg.
 dbgSplitLongString(sStr, maxLngth) - Разбивает длинную строку на набор строк длиной maxLngth.
 dbgBillInfo(bBill) - Формирует строку с информацией о чеке для логирования.

*/

// 
function dbgGetFlag(needClear) {
  const Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const Range = Spreadsheet.getRangeByName('ФлОтладка');

  if (Range != undefined && Range.getValue())
  {
    if (needClear)
      Spreadsheet.getSheetByName('dbg').clear();

    return true;
  }
  return false;
}

// Очистка листа отладки
function dbgClearSheet() {
  SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName('dbg')
  .clear()
  .activate();
}

// Разбиваем длинную строку ( >50000 ) на несколько строк по maxLngth символов
function dbgSplitLongString(sStr, maxLngth) {
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

function dbgBillInfo(bBill) {
  const s =
    " от (" + bBill.date +
    ") магазин >" + bBill.name +
    "< на сумму [" + bBill.summ + 
    "] р. наличными {" + bBill.cash + "}";
    //"} ФН :" + bBill.jsonBill.fiscalDriveNumber +
    //" ФД :" + bBill.jsonBill.fiscalDocumentNumber +
    //" ФП :" + bBill.jsonBill.fiscalSign +
    //" товаров :" + bBill.jsonBill.items.length;
  return s
}
