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
    newBills.sort((a, b) => a.dtime - b.dtime);

  // Записываем новые чеки
  const sBills = ss.getSheetByName('Чеки');
  let n = ss.getRangeByName('НомерЧек').getValue();
  let newStores = [];

  for (bill of newBills) {
    bill.number = ++n;
    let vals = [[bill.number, bill.sdate, bill.total, bill.cache, bill.fn, bill.fd, bill.fp, bill.name, bill.id]];

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
    sBills.getRange(l, 1, 1, 9).setValues(vals);
    if (! newStores.includes(bill.name))
      newStores.push(bill.name);
  }
  Logger.log('Чеки сохранены.');

  Logger.log('Сохраняем магазины. Новых : ' + newStores.length);
  if (newStores.length > 0) {
    const sStores = ss.getSheetByName('Магазины');
    sStores.insertRowsAfter(2, newStores.length);
    let storesVals = [];
    for (sStore of newStores)
      storesVals.push([sStore]);
    sStores.getRange(3, 4, newStores.length, 1).setValues(storesVals);
    Logger.log('Магазины сохранены.');
  }

  Logger.log('Сохраняем товары.');
  const fFitrUnqGoods = ss.getRangeByName('ФлагФильтрТовары').getValue();
  const sGoods = ss.getSheetByName('Товары');
  let lastRow = sGoods.getLastRow();
  let oldGoods = [[]];
  if (lastRow > 2)
    oldGoods = sGoods.getRange(4, 1, lastRow, 7).getValues();
  let numGoods = 0;
  let newVals = [];
  let chngdIds = [];
  // Заполняем список новых товаров для вставки
  for (bill of newBills) {
    for (product of bill.items) {
      // Предварительно фильтруем повторяющиеся товары
      if (fFitrUnqGoods) { // 0:Чек	1:Название	2:Цена	3:Количество	4:Единицы	5:Сумма	6:Покупок
        // В списке старых товаров
        let r = oldGoods.findIndex((element) => element[2] == product.iprice && element[1] == product.iname);
        if (~r) { // Добавляем количество покупок для этого товара в списке старых товаров
          oldGoods[r][6] += product.iquantity;
          chngdIds.push(r); // Запоминаем индекс для обновления в таблице
          continue;
        }
        // В списке старых товаров его нет. Ищем в списке новых добавленных товаров
        r = newVals.findIndex((element) => element[2] == product.iprice && element[1] == product.iname);
        if (~r) {
          newVals[r][6] += product.iquantity;
          continue;
        }
        // В списке новых тоже нет. Это первая покупка этого товара.
      }
      newVals.unshift([bill.number, product.iname, product.iprice, product.iquantity, "", product.isum, product.iquantity]);
      numGoods++;
    }
  }
  // Обновляем данные по дублированным товарам
  if (chngdIds.length > 0) {
    for (chngId of chngdIds) {
      sGoods.getRange(4 + chngId, 7, 1, 1).setValue(oldGoods[chngId][6]);
    }
  }
  // Вставляем новые товары
  sGoods.insertRowsBefore(4, numGoods);
  sGoods.getRange(4, 1, numGoods, 7).setValues(newVals);
  Logger.log('Товары сохранены. ' + numGoods + ' новых товаров.');
}
