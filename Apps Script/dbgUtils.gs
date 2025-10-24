/*

 dbgGetFlag(needClear) - Возвращает значение флага ФлОтладка. Если true и аргумент true, то очищает лист dbg.
 dbgGetLine(needInc) - Возвращает первую ячейку строки для вывода отладки. Если аргумент true, то переводит курсор на следующую.
 dbgClearSheet() - Очищает и активирует лист dbg.
 dbgSplitLongString(sStr, maxLngth) - Разбивает длинную строку на набор строк длиной maxLngth.
 dbgBillInfo(bBill) - Формирует строку с информацией о чеке для логирования.
 dbgPrintLongString(lStr) - Разбивает большую строку и выводит ее частями в строке отладки, переводит курсор на следующую.
 dbgPrintArr(sArr) - Выводит элементы массива в ячейках строки отладки, переводит курсор на следующую.

*/

// 
function dbgGetFlag(needClear) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const Range = SS.getRangeByName('ФлОтладка');

  if (Range != undefined && Range.getValue())
  {
    if (needClear)
      SS.getSheetByName('dbg').clear();

    return true;
  }
  return false;
}

//
function dbgGetLine(needInc) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const rLastErrorLine = SS.getRangeByName('LastErrorLine');
  let lastDbgLine = rLastErrorLine.getValue();
  if (lastDbgLine == "")
    lastDbgLine = 2;
  const rDBG = SS
    .getSheetByName('dbg')
    .getRange(lastDbgLine, 1);
  if (needInc)
    rLastErrorLine.setValue(lastDbgLine+1);
  return rDBG;
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

//
function dbgPrintArr(sArr) {
  //
  const rDBG = dbgGetLine(true);
  for(let i = 0; i<sArr.length; i++)
    rDBG.offset(0, 1+i).setValue(sArr[i]);
}

//
function dbgPrintLongString(lStr) {
  //
  const mm = dbgSplitLongString(lStr, 45000);
  dbgPrintArr(mm);
}
