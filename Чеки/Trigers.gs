/*

ScanMail()
ScanDrive()

*/

function ResetData()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRangeByName('ДатаЧекДиск').setValue("");
  ss.getRangeByName('ДатаЧекПочта').setValue("");
  const sBills = ss.getSheetByName('Чеки');
  if (sBills.getLastRow() > 2)
    sBills.deleteRows(4, sBills.getLastRow()-3);
  const sGoods = ss.getSheetByName('Товары');
  if (sGoods.getLastRow() > 2)
    sGoods.deleteRows(4, sGoods.getLastRow()-3);
  const sStores = ss.getSheetByName('Магазины');
  if (sStores.getLastRow() > 2)
    sStores.getRange(4, 2, lastRow-3, 3).clearContent();
}

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

function getBillRow(theBill, sJSON)
{
  const jBill = theBill.jsonBill;
  // 0:№	1:Дата	2:Сумма	3:Магазин	4:ФН	5:ФД	6:ФП	7:Наличные	8:JSON	9:URL
  return [
    theBill.SN,
    jBill.dateTime,
    jBill.totalSum / 100.0,
    jBill.user,
    jBill.fiscalDriveNumber,
    jBill.fiscalDocumentNumber,
    jBill.fiscalSign,
    jBill.cashTotalSum / 100.0,
    sJSON,
    theBill.URL
  ];
}

function TakeIntoAccount(accnt, info)
{
  // Добавляем количество и сумму для этого товара
  accnt[2] += info.quantity;
  accnt[1] += (info.sum / 100.0);
  // Проверяем минимальную и максимальную цены
  if (info.price < accnt[4] * 100.0)
    accnt[4] = info.price / 100.0;
  else
    if (info.price > accnt[5] * 100.0)
      accnt[5] = info.price / 100.0;
}

function onOnceAnHour()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('Обрабатываем последние чеки.');
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
  // Сортируем перед записью
  if (cntBills > 1)
    newBills.sort((a, b) => a.dTime - b.dTime);

  const fSaveJSON = ss.getRangeByName('ФлагСохрJSON').getValue();
  const sBills = ss.getSheetByName('Чеки');
  let newRows = [];
  let sJSON = "";
  let n = ss.getRangeByName('НомерЧек').getValue();
  if (sBills.getLastRow() > 2) {
    // На листе есть старые чеки
    Logger.log('>>> Вставляем новые чеки в список старых на листе.');
    for (bill of newBills) {
      let l = 4;
      let d = sBills.getRange(l, 2, 1, 1).getValue();
      let dt = d.getTime();
      let tt = bill.dTime;
      while (tt < dt) {
        d = sBills.getRange(++l, 2, 1, 1).getValue();
        if (d.toString() == "") break;
        dt = d.getTime();
      }
      bill.SN = ++n;
      sBills.insertRowBefore(l);
      if (fSaveJSON) {
        sJSON = billFormatShort(bill.jsonBill);
        sBills.setRowHeightsForced(l, 1, 21);
      } else
        sJSON = "";
      sBills.getRange(l, 1, 1, 10).setValues([getBillRow(bill, sJSON)]);
    }
  } else {
    // На листе нет старых чеков
    Logger.log('>>> Просто добавляем чеки на лист.');
    for (bill of newBills) {
      bill.SN = ++n;
      if (fSaveJSON)
        sJSON = billFormatShort(bill.jsonBill);
      else
        sJSON = "";
      newRows.unshift(getBillRow(bill, sJSON));
    }
    sBills.insertRowsBefore(4, cntBills);
    sBills.getRange(4, 1, cntBills, 10).setValues(newRows);
    if (fSaveJSON)
      sBills.setRowHeightsForced(4, cntBills, 21);
  }
  Logger.log('<<< Чеки сохранены.');

  Logger.log('>>> Сохраняем товары из чеков.');
  Logger.log('+++ Отрезаем артикулы и метки акций из названий товаров.');
  Logger.log('+++ Фильтруем повторяющиеся товары в общем списке.');

  const sGoods = ss.getSheetByName('Товары');
  let chngdRows = []; // Массив индексов измененных записей на листе Товары
  let oldRows = []; // Массив старых записей на листе Товары, которые могут быть изменены
  let lastRow = sGoods.getLastRow();
  if (lastRow > 2)
    oldRows = sGoods.getRange(4, 1, lastRow-3, 6).getValues();

  // Заполняем список новых товаров для вставки, фиксируем изменения в повторяющихся товарах
  newRows = [];
  for (bill of newBills)
    for (product of bill.jsonBill.items) {
      const sProduct = cutPromoTag(product.name);

      // Ищем повторение в списке старых товаров
      const r = oldRows.findIndex((element) => element[0] == sProduct);
      if (~r) {
        // Запоминаем индекс для обновления в таблице
        chngdRows.push(r);
        // Добавляем количество и сумму для этого товара в списке старых товаров
        TakeIntoAccount(oldRows[r], product);
        continue;
      }

      // В списке старых товаров его нет. Ищем в списке новых добавленных товаров
      let elm = newRows.find((element) => element[0] == sProduct);
      if (elm != undefined) {
        TakeIntoAccount(elm, product);
        continue;
      }

      // В списке новых тоже нет. Это первая покупка этого товара.
      // 0:Название	1:Сумма	2:Количество	3:Единицы	4:Мин Цена	5:Макс Цена	6:Первый Чек	7:Проверка
      newRows.unshift([sProduct, product.sum / 100.0, product.quantity, product.unit, product.price / 100.0, product.price / 100.0, bill.SN]);
    }

  // Обновляем данные по дублированным товарам
  if (chngdRows.length > 0)
    for (chngdRow of chngdRows) {
      const newVal = oldRows[chngdRow];
      sGoods.getRange(4 + chngdRow, 2, 1, 2).setValues([[ newVal[1], newVal[2] ]]);
      sGoods.getRange(4 + chngdRow, 5, 1, 2).setValues([[ newVal[4], newVal[5] ]]);
    }

  // Вставляем новые товары
  let newLength = newRows.length;
  if (newLength > 0) {
    sGoods.insertRowsBefore(4, newLength);
    sGoods.getRange(4, 1, newLength, 7).setValues(newRows);
  }
  Logger.log('<<< Товары сохранены. Добавлено ' + newLength + ' новых товаров. Старых обновлено : ' + chngdRows.length);

  const fSortGoods = ss.getRangeByName('ФлагСортТовары').getValue();
  if (fSortGoods) {
    // Сортируем все товары на листе
    Logger.log('+++ Сортируем товары на листе.');
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
    oldRows = sStores.getRange(4, 1, lastRow-3, 4).getValues();
    // 0:Название	1:Чеков	2:Сумма	3:Последний чек	4:Статья	5:Инфо	6:Примечание
  // Заполняем список новых магазинов для вставки, фиксируем изменения в повторяющихся магазинах
  for (bill of newBills) {
    const sStore = bill.Shop;
    const nTotal = bill.jsonBill.totalSum / 100.0;
    const lastBill = bill.SN;

    // В списке старых магазинов
    let r = oldRows.findIndex((element) => element[0] == sStore);
    if (~r) { // Увеличиваем количество чеков и обновляем общую сумму для этого магазина
      oldRows[r][1] += 1;
      oldRows[r][2] += nTotal;
      oldRows[r][3] = lastBill;
      chngdRows.push(r); // Запоминаем индекс для обновления в таблице
      continue;
    }

    // В списке старых магазинов его нет. Ищем в списке новых добавленных магазинов
    let elm = newRows.find((element) => element[0] == sStore);
    if (elm != undefined) {
      elm[1] += 1;
      elm[2] += nTotal;
      elm[3] = lastBill;
      continue;
    }
    // В списке новых тоже нет. Это первая покупка из этого магазина.
    newRows.unshift([sStore, 1, nTotal, lastBill]);
  }

  Logger.log('--- Сохраняем магазины. Новых : ' + newRows.length + ' Изменено старых : ' + chngdRows.length);
  if (chngdRows.length > 0)
    for (chngdRow of chngdRows) 
      sStores.getRange(4 + chngdRow, 2, 1, 3).setValues([[ oldRows[chngdRow][1], oldRows[chngdRow][2], oldRows[chngdRow][3] ]]);
  // Вставляем новые магазины
  newLength = newRows.length;
  if (newLength > 0) {
    sStores.insertRowsBefore(4, newLength);
    sStores.getRange(4, 1, newLength, 4).setValues(newRows);
  }
  Logger.log('<<< Магазины сохранены.');
  Logger.log('Обработка завершена.');
}
