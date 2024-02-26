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

  const cntBills = newBills.length;
  if (cntBills == 0) {
    Logger.log("<<< Нет новых чеков. >>>");
    return;
  }

  Logger.log('Обновляем даты.');
  if (newLastDriveDate > dLastDriveDate)
    rLastDriveDate.setValue(newLastDriveDate);
  if (newLastMailDate > dLastMailDate)
    rLastMailDate.setValue(newLastMailDate);

  // Записываем новые чеки
  Logger.log('>>> Сохраняем ' + cntBills + ' новых чеков.');
  let newRows = [];
  // Сортируем перед записью
  if (cntBills > 1)
    newBills.sort((a, b) => a.dtime - b.dtime);

  let n = ss.getRangeByName('НомерЧек').getValue();
  const sBills = ss.getSheetByName('Чеки');
  if (sBills.getLastRow() > 2) {
    // На листе есть старые чеки
    Logger.log('>>> Вставляем новые чеки в список старых на листе.');
    for (bill of newBills) {
      bill.number = ++n;
      let l = 4;
      let d = sBills.getRange(l, 2, 1, 1).getValue();
      let dt = d.getTime();
      let tt = bill.dtime;
      while (tt < dt) {
        d = sBills.getRange(++l, 2, 1, 1).getValue();
        if (d.toString() == "") break;
        dt = d.getTime();
      }
      sBills.insertRowBefore(l);
      sBills.getRange(l, 1, 1, 9).setValues([[bill.number, bill.sdate, bill.total, bill.cache, bill.fn, bill.fd, bill.fp, bill.name, bill.id]]);
    }
  } else {
    // На листе нет старых чеков
    Logger.log('>>> Просто добавляем чеки на лист.');
    for (bill of newBills) {
      bill.number = ++n;
      newRows.unshift([bill.number, bill.sdate, bill.total, bill.cache, bill.fn, bill.fd, bill.fp, bill.name, bill.id]);
    }
    sBills.insertRowsBefore(4, cntBills);
    sBills.getRange(4, 1, cntBills, 9).setValues(newRows);
  }
  Logger.log('<<< Чеки сохранены.');

  Logger.log('>>> Сохраняем товары из чеков.');
  newRows = [];
  const fFitrUnqGoods = ss.getRangeByName('ФлагФильтрТовары').getValue();
  if (fFitrUnqGoods)
    Logger.log('Фильтруем повторяющиеся товары в общем списке.');
  const sGoods = ss.getSheetByName('Товары');
  let chngdRows = []; // Массив индексов измененных записей на листе Товары
  let oldRows = []; // Массив старых записей на листе Товары, которые могут быть изменены
  let lastRow = sGoods.getLastRow();
  if (lastRow > 2)
    oldRows = sGoods.getRange(4, 1, lastRow-3, 4).getValues();
    // 0:Название	1:Цена	2:Количество	3:Сумма
  // Заполняем список новых товаров для вставки, фиксируем изменения в повторяющихся товарах
  for (bill of newBills) {
    for (product of bill.items) {
      // Отрезаем всякие акционные метки
      let pdctName = product.iname
        .replace(/^\<А\> /, '');
      // Предварительно фильтруем повторяющиеся товары
      if (fFitrUnqGoods) {
        // В списке старых товаров
        let r = oldRows.findIndex((element) => element[1] == product.iprice && element[0] == product.iname);
        if (~r) { // Добавляем количество покупок и сумму для этого товара в списке старых товаров
          oldRows[r][2] += product.iquantity;
          oldRows[r][3] += product.isum;
          chngdRows.push(r); // Запоминаем индекс для обновления в таблице
          continue;
        }
        // В списке старых товаров его нет. Ищем в списке новых добавленных товаров
        r = newRows.findIndex((element) => element[1] == product.iprice && element[0] == product.iname);
        if (~r) {
          newRows[r][2] += product.iquantity;
          newRows[r][3] += product.isum;
          continue;
        }
        // В списке новых тоже нет. Это первая покупка этого товара.
      }
      newRows.unshift([product.iname, product.iprice, product.iquantity, product.isum, bill.number]);
    }
  }
  // Обновляем данные по дублированным товарам
  if (chngdRows.length > 0)
    for (chngdRow of chngdRows)
      sGoods.getRange(4 + chngdRow, 3, 1, 2).setValues([[oldRows[chngdRow][2], oldRows[chngdRow][3]]]);
  // Вставляем новые товары
  let newLength = newRows.length;
  if (newLength > 0) {
    sGoods.insertRowsBefore(4, newLength);
    sGoods.getRange(4, 1, newLength, 5).setValues(newRows);
  }
  Logger.log('<<< Товары сохранены. Добавлено ' + newLength + ' новых товаров. Старых обновлено : ' + chngdRows.length);

  // Заготовка
  const fSortGoods = ss.getRangeByName('ФлагСортТовары').getValue();
  if (fSortGoods) {
    Logger.log('Сортируем товары на листе.');
    lastRow = sGoods.getLastRow();
    oldRows = sGoods.getRange(4, 1, lastRow-3, 7).getValues();
    // 0:Название	1:Цена	2:Количество	3:Сумма	4: Чек	5: Единицы	6: Проверка
    oldRows.sort((a, b) => a[0].localeCompare(b[0]));
    sGoods.getRange(4, 1, lastRow-3, 7).setValues(oldRows);
  }

  Logger.log('>>> Обнавляем информацию по магазинам.');
  newRows = [];
  const sStores = ss.getSheetByName('Магазины');
  chngdRows = []; // Массив индексов измененных записей на листе Магазины
  oldRows = []; // Массив старых записей на листе Магазины, которые могут быть изменены
  lastRow = sStores.getLastRow();
  if (lastRow > 2)
    oldRows = sStores.getRange(4, 4, lastRow-3, 3).getValues();
    // 0:Статья	1:Инфо	2:Примечание	3:Название	4:Чеков
  // Заполняем список новых магазинов для вставки, фиксируем изменения в повторяющихся магазинах
  for (bill of newBills) {
    const sStore = bill.name;
    // В списке старых магазинов
    let r = oldRows.findIndex((element) => element[0] == sStore);
    if (~r) { // Увеличиваем количество чеков и обновляем общую сумму для этого магазина
      oldRows[r][1] += 1;
      oldRows[r][2] += bill.total;
      chngdRows.push(r); // Запоминаем индекс для обновления в таблице
      continue;
    }
    // В списке старых магазинов его нет. Ищем в списке новых добавленных магазинов
    r = newRows.findIndex((element) => element[0] == sStore);
    if (~r) {
      newRows[r][1] += 1;
      newRows[r][2] += bill.total;
      continue;
    }
    // В списке новых тоже нет. Это первая покупка из этого магазина.
    newRows.unshift([sStore, 1, bill.total]);
  }

  Logger.log('--- Сохраняем магазины. Новых : ' + newRows.length + ' Изменено старых : ' + chngdRows.length);
  if (chngdRows.length > 0)
    for (chngdRow of chngdRows)
      sStores.getRange(4 + chngdRow, 5, 1, 2).setValues([[oldRows[chngdRow][1], oldRows[chngdRow][2]]]);
  // Вставляем новые магазины
  newLength = newRows.length;
  if (newLength > 0) {
    sStores.insertRowsBefore(4, newLength);
    sStores.getRange(4, 4, newLength, 3).setValues(newRows);
  }
  Logger.log('<<< Магазины сохранены.');
  Logger.log('Обработка завершена.');
}
