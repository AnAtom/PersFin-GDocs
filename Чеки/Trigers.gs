/*

ScanMail()
ScanDrive()

*/

// Читает дату из ячейки. Если ячейка пуста, то возвращает дату ДатаЧек0.
function ReadLastDate(ss, rDate)
{
  let dLastDate = rDate.getValue();
  const sLastDate = dLastDate.toString();
  if (sLastDate == "") {
    dLastDate = ss.getRangeByName('ДатаЧек0').getValue();
    Logger.log("Принимаем дату последнего чека : " + dLastDate.toString());
  } else
    Logger.log("Дата последнего чека : " + sLastDate);
  return dLastDate;
}

function onOnceAnHour()
{
  Logger.log('Обрабатываем последние чеки.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fDBG = ss.getRangeByName('ФлагОтладки').getValue();
  const rDBG = ss.getSheetByName('DBG').getRange(1, 1);

  let newBills = [];

  // Сканируем диск
  const rLastDriveDate = ss.getRangeByName('ДатаЧекДиск');
  const dLastDriveDate = ReadLastDate(ss, rLastDriveDate);
  const newLastDriveDate = ScanDrive(ss, dLastDriveDate, newBills);

  // Сканируем почту
  const rLastMailDate = ss.getRangeByName('ДатаЧекПочта');
  const dLastMailDate = ReadLastDate(ss, rLastMailDate);
  const newLastMailDate = ScanMail(ss, dLastMailDate, newBills);

  if (newBills.length == 0) {
    Logger.log("Нет новых чеков.");
    return;
  }

  Logger.log('Обновляем даты.');
  if (newLastDriveDate > dLastDriveDate)
    rLastDriveDate.setValue(newLastDriveDate);
  if (newLastMailDate > dLastMailDate)
    rLastMailDate.setValue(newLastMailDate);

  Logger.log('Сохраняем ' + newBills.length + ' новых чеков.');
  // Сортируем перед записью
  if (newBills.length > 1)
    newBills.sort((a, b) => {return a.dtime - b.dtime});

  // Записываем новые чеки
  const sBills = ss.getSheetByName('Чеки');
  let n = ss.getRangeByName('НомерЧек').getValue();

  for (bill of newBills) {
    bill.number = ++n;
    let vals = [[bill.number, bill.sdate, bill.total, bill.cache, bill.fn, bill.fd, bill.fp, bill.name]];

    let tt = bill.dtime;
    let l = 4;
    let d = sBills.getRange(l, 2, 1, 1).getValue();
    if (d.toString() != "") {
      let dd = d.getTime();
      while (tt < dd) {
        d = sBills.getRange(++l, 2, 1, 1).getValue();
        if (d.toString() == "") break;
        dd = d.getTime();
      }
    }
    sBills.insertRowBefore(l);
    sBills.getRange(l, 1, 1, 8).setValues(vals);
  }
  Logger.log('Чеки сохранены.');

  Logger.log('Сохраняем товары.');
  let numGoods = 0;
  let newVals = [];
  for (bill of newBills) {
    for (product of bill.items) {
      newVals.unshift([bill.number, product.iname, product.iprice, product.iquantity, "", product.isum]);
      numGoods++;
    }
  }
  const sGoods = ss.getSheetByName('Товары');
  sGoods.insertRowsBefore(4, numGoods);
  sGoods.getRange(4, 1, numGoods, 6).setValues(newVals);
  Logger.log('Товары сохранены. ' + numGoods + ' новых товаров.');
}
