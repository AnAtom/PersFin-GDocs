/*

 onOpen(e) - Добавляем пункты меню
 onEdit(e) - Настраиваем списки выбора
 UpdateOnOpen(e)
 onOnceAnHour()
 onOnceADay()

 Редактирование на листе «Операции»
  SettingTrntnName - Устанавливаем доступные счета и Тип операции для выбранной из общего списка операции
  SettingTrntnType - Устанавливаем доступный список операций для выбранного Типа операции

 Редактирование на листе «Расходы»
  SettingCostInfo - Изменилась статья расходов, устанавливаем список расходов (Инфо)
  SettingCostNote - Изменился пункт статьи расходов (Инфо), устанавливаем список для Примечание
  SettingCostBill - Изменилась Заметка (парсим чек)

*/

function TestTest() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const newSpreadsheet = spreadsheet.copy("Копия Тест");
  const newId = newSpreadsheet.getId();
  const newFile = DriveApp.getFileById(newId);
  const oldId = spreadsheet.getId();
  const oldFile = DriveApp.getFileById(oldId);
  const oldFolders = oldFile.getParents();
  var folder;
  while (oldFolders.hasNext()) {
    folder = oldFolders.next();
    Logger.log(folder.getName());
  }
  newFile.moveTo(folder);
}

function MenuChangeYear(NewYear) {
  //
  var ui = SpreadsheetApp.getUi();

  NewYear = 2026;
  var result = ui.alert(
    "Изменение финансового года",
    "Создать новую копию для " + NewYear + " года?",
    ui.ButtonSet.YES_NO,
  );

  if (result == ui.Button.YES) {
    //
    Logger.log("Создаем копию таблицы");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var newSpreadsheet = spreadsheet.copy("Финансы " + NewYear);
  } else {
    Logger.log("Меняем год в этой таблице");
    result = ui.alert(
      "Изменение этого финансового года",
      "Очистить все данные по расходкам и платежам?",
      ui.ButtonSet.YES_NO,
    );
    if (result == ui.Button.YES) {
      //
      result = ui.alert(
        "ПРЕДУПРЕЖДЕНИЕ!!!",
        "Вы уверены, что хотите очистить все расходы и платежи?",
        ui.ButtonSet.OK_CANCEL,
      );
      if (result == ui.Button.CANCEL) return;
    }
  }
}

// Устанавливаем доступные счета и Тип операции для выбранной из общего списка операции
function SettingTrntnName(ss, br) {
  const NewVal = br.getValue();
  const OpAcc = br.offset(0,-2); // Счет
  const OpTrgt = br.offset(0,-1); // Цель

  const debit = 'Списание';
  const moving = 'Оборот';
  let i = findInRule(moving, NewVal);
  if (~i)
  {
    // Выбрана оборотная операция
    br.offset(0,1).setValue(moving);
    SetTargetRule(ss, OpAcc, 'СчетаДеб');

    if (NewVal == ss.getRangeByName('стрПеревод').getValue())
      SetTargetRule(ss, OpTrgt, 'СчетаДеб'); // Перевод
    else
    {
      OpTrgt.clearDataValidations();
      if (i == 0) // Снятие
        OpTrgt.clearContent();
    }
  }
  else if (~findInRule(debit, NewVal))
  { // Выбрана опреация списания
    br.offset(0,1).setValue(debit); // Списовать можем только с деьитовых счетов
    const credit = 'Кредиты';
    if (NewVal == ss.getRangeByName('стрПрцКрдт').getValue())
    { // Проценты по кредиту
      SetTargetRule(ss, OpAcc, credit);
      OpTrgt.clearDataValidations().clearContent();
    }
    else
    {
      SetTargetRule(ss, OpAcc, 'СчетаДеб');
      if (NewVal == ss.getRangeByName('стрПогКрдт').getValue()) // Погашение кредита
        SetTargetRule(ss, OpTrgt, credit);
      else
        if (NewVal == ss.getRangeByName('стрПлатеж').getValue()) // Платеж
          SetTargetRule(ss, OpTrgt, 'Платежи');
        else
          OpTrgt.clearDataValidations();
    }
  }
  else
  {
    const receipt = 'Начисление';
    i = findInRule(receipt, NewVal);
    if (~i) { // Выбрана операция начисления
      br.offset(0,1).setValue(receipt);
      SetTargetRule(ss, OpAcc, 'СчетаДеб');
      if (i < 4)
        OpAcc.setValue("ЗП");
    }
  }
}

// Устанавливаем соответствующий список операций для выбранного Типа операции
function SettingTrntnType(ss, br) {
  const NewVal = br.getValue();
  if (ss.getRangeByName(NewVal) == undefined)
    NewVal = 'Операция'; // Устанавливаем полный список операций для выбора если Тип неизвестен

  SetTargetRule(ss, br.offset(0,-1), NewVal)
}

// Устанавливаем список расходов для выбранной статьи расходов
function SettingCostInfo(ss, br) {
  const flgDbg = dbgGetFlag(false);

  if (flgDbg)
  {
    var rDGB = ss.getSheetByName('dbg').getRange(1, 1);
  }
  //br.setNote('Test :' + sTest + ' Range :' + rTest.getNumRows());

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  //br.setNote('Row :' + flgDbg + ' Val :' + NewVal);
  if (flgDbg) rDGB.offset(2, 1).setValue(NewVal);

  if (NewVal != '')
  {
    const range = ss.getRangeByName('СтРсх' + NewVal);

    if (range != undefined)
    {
      const rule = SpreadsheetApp
      .newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(range)
      .build();

      cell.setDataValidation(rule);
      return;
    }
  }
  cell.clearDataValidations();
}

// Устанавливаем список информации для выбранного расхода
function SettingCostNote(ss, br) {
  const flgDbg = dbgGetFlag(false);
  
  if (flgDbg)
  {
    // Лист для отладки
    var rDGB = ss.getSheetByName('dbg').getRange(1, 1);
  }

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  if (flgDbg) rDGB.offset(2, 1).setValue(NewVal);

  var range;
  //SpreadsheetApp.getActive().toast('Range :'+ range);

  switch(NewVal) {
    case 'Продукты':
      range = ss.getRangeByName('СтРсхЕдаМагаз');
      break;
    case 'Пиво':
      range = ss.getRangeByName('СтРсхАлкПиво');
      break;
    case 'Кабак':
      range = ss.getRangeByName('СтРсхАлкКабак');
      break;
    case 'Чаевые':
      // Используем дату, время и информацию с предыдущей строки
      br
      .offset(0, -5).setFormulaR1C1('=R[1]C[0]') // Дата
      .offset(0, 1).setFormulaR1C1('=R[0]C[-1]') // Время
      .offset(0, 2).setValue('Карман')           // Счет
      .offset(0, 3).setFormulaR1C1('=R[1]C[0]'); // Примечание (заведение)
      break;
  }

  if (range != undefined)
    cell.setDataValidation(
      SpreadsheetApp
      .newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(range)
      .build()
    );
  else cell.clearDataValidations();
}

// Читаем информацию о чеке из json строки
function SettingCostBill(ss, br) {
  const flgDbg = dbgGetFlag(false);

  if (flgDbg)
  {
    // Лист для отладки
    var rDGB = ss.getSheetByName('dbg').getRange(1, 1);
  }

  const NewVal = br.getValue();
  if (NewVal == "") return;

  if (flgDbg) rDGB.offset(2, 1).setValue(NewVal);

  const bill = billInfo(NewVal);

  if (flgDbg) {
    if (bill != undefined)
      rDGB.offset(3, 1, 1, 5).setValues([[bill.name, bill.summ, bill.date, bill.cash, bill.shop]]);
    else rDGB.offset(3, 1).setValue("UNDEFINED !!!");
  }

  if (bill == undefined) return;

  // Выставляем дату, время и сумму покупки
  // Формат ячеек "dd.mm", "HH:mm", "#,##0.00[$ ₽]"
  br
  .offset(0, -7).setValue(bill.date).setNumberFormat("dd.mm")
  .offset(0, 1).setFormulaR1C1('=R[0]C[-1]').setNumberFormat("HH:mm")
  .offset(0, 1).setValue(bill.summ).setNumberFormat("#,##0.00[$ ₽]");

  // Если наличные, то выставляем счет списания
  if (bill.cash != 0)
    br.offset(0,-4).setValue("Карман")

  // Выставляем Статью, Инфо и Примечание для магазина
  const lstStores = ss.getRangeByName('СпскМагазины');
  let shop = lstStores.getValues().find((element) => element[3] == bill.shop);
  if (shop == undefined) {
    // Добавляем в список новый магазин
    Logger.log("Новый магазин [" + bill.shop + "] (" + bill.name + ")");
    const sShop = ss.getSheetByName('Магазины');
    const newRow = lstStores.getNumRows() + 4;
    sShop.insertRowBefore(newRow);
    sShop.getRange(newRow, 4, 1, 2).setValues([[bill.shop, bill.name]]);
  } else
    br
    .offset(0, -3, 1, 3)
    .setValues([[shop[0], shop[1], shop[2]]]);
}

function onOpen(e) {
  const menuScan = [
    {name: "С новым годом!", functionName: 'MenuChangeYear'},
    null,
    {name: "Очистить отладку", functionName: 'dbgClearSheet'}
  ];
  Logger.log('Добавляем пункты меню.');
  e.source.addMenu("Действия", menuScan);
  Logger.log('Открылись.');
}

function UpdateOnOpen(e) {
  Logger.log('Первичное сканирование.');
  onOnceAnHour()
  Logger.log('Обновились.');
}

function onEdit(e) {
  const ss = e.source;
  Logger.log("Редактирование на листе <" + ss.getActiveSheet().getSheetName() + ">");
  // Читаем флаг "Использовать автосписки"
  let br = ss.getRangeByName('ФлАвтосписки');
  if (br == undefined || ! br.getValue()) return;

  br = e.range;
  Logger.log("<" + br + ">");
  Logger.log(">" + br.getA1Notation() + "<");
  if (br.getNumColumns() > 1) // Скопировали диапазон
    return;

  const ncol = br.getColumn();
  const sname = ss.getActiveSheet().getSheetName();
  let cname = ss.getActiveSheet().getRange(1, ncol).getValue();
  if (cname == undefined || cname == '')
    cname = ncol;
  Logger.log("Редактируем на листе [" + sname + "] в колонке (" + cname + ") строку :" + br.getRow());

  switch(sname) {
    case 'Операции':
      switch(ncol) {
        case 7: // Изменился Тип операции
          const v = e.value;
          if (v == undefined || v == '') // Устанавливаем полный список операций для выбора если Тип операции был очищен
            SetTargetRule(ss, br.offset(0,-1), 'Операция');
          else
            SettingTrntnType(ss, br);
          return; // -----------------------------------------------------------------------------------------------------------------------------------------------------------------
        case 6: // Изменилась Операция
          SettingTrntnName(ss, br);
        default:
          return;
      }
    case 'Расходы':
      switch(ncol) {
        case 5:
          // Изменилась статья расходов
          SettingCostInfo(ss, br);
          return;
        case 6:
          // Изменился пункт статьи расходов (Инфо)
          SettingCostNote(ss, br);
          return;
        case 8:
          // Изменилась заметка (вставили чек)
          SettingCostBill(ss, br);
          return;
      }
    case 'i':
      if (ncol == 8 && br.getRow() == 23) {
        Logger.log("Переключаем сохранение чеков.");
        //
      } else {
        //
      }
  }
}

function onOnceAnHour() {
  // Выполняется ежечасно
  Logger.log("Обрабатываем последние чеки");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dDate0 = ss.getRangeByName('День1').getValue();

  let newBills = [];

  // Сканируем диск
  const billsDrive = new DriveBillsScaner('ДнейРетроДиск');
  billsDrive.doScan(newBills);

  // Сканируем чеки в почте
  const billsMail = new MailTemplateScaner('ШаблоныЧеков');
  billsMail.doScan(billsMail.readData, newBills);

  // Сканируем покупки Ali
  //const rLastAliDate = getDateRangeDefault('ДатаЧекAli');

  // Сканируем поездки UBER
  const billsUBER = new MailLabelScaner('Uber');
  billsUBER.doScan(billUBER, newBills);

  // Сканируем поездки Яндекс Go
  const billsYandexGo = new MailLabelScaner('ЯндексGo');
  billsYandexGo.doScan(billYandexGo, newBills);

  const cntBills = newBills.length;
  if (cntBills == 0) {
    Logger.log("На текущий момент нет новых чеков.");
    return;
  }
  Logger.log(">>>>>>>> Обновляем " + cntBills + " чеков.");

  if (cntBills > 1)
    newBills.sort((a, b) => b.dTime - a.dTime); // Первым будет самый свежий чек
  Logger.log("-------------------------------------------------------------------------------------------------");
  for (const bb of newBills)
    Logger.log(" tDate: " + bb.tDate + " dTime: " + bb.dTime + " date: " + bb.date + " Чек Summ: " + bb.summ + " Shop: " + bb.shop);
  
  let nowDayTime = newBills[0].tDate; // Первый день в списке чеков.

  const lstStores = ss.getRangeByName('СпскМагазины').getValues();
  const lstIgnore = ss.getRangeByName('спскМагзныИгнор').getValues();
  const shops = ss.getSheetByName('Магазины');
  const costs = ss.getSheetByName("Расходы");
  //const costs = ss.getSheetByName('Расходы (копия)');
  const costsData = costs.getDataRange();
  let cdRows = costsData.getNumRows();

  // dTime: getDate(sDate).getTime(), // UNIX время чека
  // tDate: // UNIX время дня чека
  // date: sDate,  summ: nSumm,  cash: nCash,
  // name: sName,  shop: billFilterName(sName)};

  let firstDateSummRow = -1;
  let resetFrstDateSummRow = false;
  for (let i = 2; i < cdRows; i++) {
    const iDate = costsData.getCell(i, 1).getValue();
    if (iDate === '') continue; // Пустая строка. Продолжаем...
    let iDateDayTime = 0;
    try {
      iDateDayTime = iDate.getTime();
    } catch (err) {
      Logger.log(">>> !!! Ошибка чтения даты [" + iDate + "] в строке " + i + ". ", err);
      continue;
    }
    // Нашли на листе строку с датой
    let iInsrtBill = -1;
    let iDelBill = -1;
    const iSumm = costsData.getCell(i, 3).getValue();
    const isSumm = !(iSumm === '');
    //if (isSumm && !(~firstDateSummRow)) firstDateSummRow = i; // Запоминаем строку (если еще не запомнили) чтобы потом вставить чеки перед ней
    const iTime = costsData.getCell(i, 2).getValue();
    const isTime = !(iTime === '');
    const isPrevDay = iDateDayTime < nowDayTime;

    if (isTime) {
      if (isSumm) { // Указано и время, и сумма
        iDelBill = newBills.findIndex((bill) => bill.summ == iSumm && bill.dTime == iDateDayTime);
        if (~iDelBill) iInsrtBill = iDelBill-1; // Нашли такой чек
        else { // Ищем чеки старше этого времени
          iDelBill = newBills.findLastIndex((bill) => bill.dTime > iDateDayTime);
          if (!(~iDelBill)) continue; // Не нашли что вставлять
          iInsrtBill = iDelBill;
        }
      } else { // Указано время. Не указана сумма. Ищем чек с таким временем.
        iDelBill = newBills.findIndex((bill) => bill.dTime == iDateDayTime);
        if (~iDelBill) { // Нашли чек с таким временем. Обновляем строку
          setCostBill(costs.getRange(i, 3), newBills[iDelBill], getShopInfoRemarkNote(newBills[iDelBill].shop, newBills[iDelBill].name, lstStores, lstIgnore, shops));
          Logger.log('... В строке ' + i + ' Нашли чек с датой ' + iDate + ' Обновляем сумму ' + newBills[iDelBill].summ);
          costs.getRange(i, 15).setValue("updated");
          iInsrtBill = iDelBill-1;
        } else { // Не нашли чек с таким временем. Ищем все чеки старше для вставки
          iDelBill = newBills.findLastIndex((bill) => bill.dTime > iDateDayTime);
          if (!(~iDelBill)) continue; // Не нашли что вставлять
          iInsrtBill = iDelBill;
        }
      }
      // Сбросить запомниенную строку без времени
      resetFrstDateSummRow = true;
    } else // Не указано время
      if (isSumm) {
        if (!(~firstDateSummRow)) firstDateSummRow = i; // Запоминаем строку (если еще не запомнили) чтобы потом вставить чеки перед ней
        if (isPrevDay)
          iDelBill = newBills.findIndex((bill) => bill.summ == iSumm && bill.tDate == iDateDayTime); // Ищем вчера
        else
          iDelBill = newBills.findIndex((bill) => bill.summ == iSumm && bill.tDate == nowDayTime); // Ищем сегодня
        if (~iDelBill) { // Нашли чек с этой суммой в соответствующий день для обновления времени
          setCostBill(costs.getRange(i, 3), newBills[iDelBill], getShopInfoRemarkNote(newBills[iDelBill].shop, newBills[iDelBill].name, lstStores, lstIgnore, shops));
          Logger.log('... В строке ' + i + ' Нашли чек с суммой ' + iSumm + ' Обновляем дату ' + newBills[iDelBill].date);
          costs.getRange(i, 15).setValue("updated");
          iInsrtBill = iDelBill-1;
        } else { // Не нашли чек с этой суммой
          if (! isPrevDay) continue; // Сегодня переходим на следующую строку
          iDelBill = newBills.findLastIndex((bill) => bill.tDate == nowDayTime); // Если мы во вчера, то ищем все сегодняшние чеки чтобы закрыть день.
          if (!(~iDelBill)) continue; // Не нашли что вставлять
          iInsrtBill = iDelBill;
          Logger.log('Будем вставлять до' + iDelBill + ' чека.');
        }
      } else { // Не указана сумма
        // Не указано ни время, ни сумма
        if (! isPrevDay) continue; // Сегодня переходим на следующую строку
        iDelBill = newBills.findLastIndex((bill) => bill.tDate == nowDayTime); // Если мы во вчера, то ищем все сегодняшние чеки чтобы закрыть день.
        if (!(~iDelBill)) continue; // Не нашли что вставлять
        iInsrtBill = iDelBill;
      }

    // Вставляем чеки на лист
    if (~iInsrtBill) {
      iInsrtBill += 1;
      if (!(~firstDateSummRow)) firstDateSummRow = i; // Если небыло запомненной строки, то вставляем перед этой
      Logger.log('<<< Вставляем ' + iInsrtBill + ' чеков перед строкой ' + firstDateSummRow);
      // Определяем есть ли подчеркивание
      const underlinedRow = costs.getRange(firstDateSummRow - 1, 1, 1, 9);
      if (underlinedRow.getBorder() != null) {
        Logger.log('___ Вставляем над подчеркиванием.');
        const underlinedData = underlinedRow.getValues();         // Сохраняем данные из подчеркнутой строки
        costs.insertRowsBefore(firstDateSummRow - 1, iInsrtBill); // Вставляем строки над ней
        Logger.log('^^^ Перемещаем строку ' + underlinedData);
        costs.getRange(firstDateSummRow - 1, 1, 1, 9)
          .setValues(underlinedData)                              // Сохраняем данные из подчеркнутой строки в первую вставленную
          .offset(0, 0, 1, 1)                                     // потому, что эта строка будет надписана
          .setNumberFormat("dd.mm")                               // и выставляем правильные форматы для даты, времени и суммы
          .offset(0, 1)
          .setNumberFormat("HH:mm")
          .offset(0, 1)
          .setNumberFormat("#,##0.00[$ ₽]")   ; 
      } else costs.insertRowsBefore(firstDateSummRow, iInsrtBill);

      for (let j = 0; j < iInsrtBill; j++)
        setCostBill(costs.getRange(firstDateSummRow + j, 3), newBills[j], getShopInfoRemarkNote(newBills[j].shop, newBills[j].name, lstStores, lstIgnore, shops));
      i += iInsrtBill;
      cdRows += iInsrtBill;
      resetFrstDateSummRow = true;
    }

    // Сбрасываем запомненную строку
    if (resetFrstDateSummRow) {
      resetFrstDateSummRow = false;
      firstDateSummRow = -1;
    }

    // Удаляем обработанные чеки
    if (~iDelBill) {
      if (iDelBill == 0) {
        Logger.log('XXX Удаляем чек с суммой ' + newBills[0].summ + ' и датой ' + newBills[0].date + ' >>> ' + newBills[0].dTime + ' > ' + iDateDayTime);
        newBills.shift();
      } else {
        const lastDellBill = newBills[iDelBill];
        iDelBill += 1;
        Logger.log('XXX Удаляем ' + iDelBill + ' чеков. Последний с суммой ' + lastDellBill.summ + ' от ' + lastDellBill.date);
        newBills.splice(0, iDelBill);
      }
      if (newBills.length == 0) {
        Logger.log("<<< Все чеки обработаны. >>>");
        break;
      }
      // Обновляем время текущего дня
      if (isPrevDay) {
        Logger.log('+++ Новый день ' + newBills[0].date);
        nowDayTime = newBills[0].tDate;
      }
    }
  }
  if (newBills.length > 0)
    Logger.log(">>>>>>>>>> !!! " + newBills.length + "!!! <<<<<<<<<<");
  Logger.log("Пробежались по чекам на листе. В списке найденных осталось :" + newBills.length);

  Logger.log("Обновляем даты.");
  billsDrive.updateDate();
  billsMail.updateDate();
  billsUBER.updateDate();
  billsYandexGo.updateDate();
}

function onOnceADay() {
  // Выполняется ежежневно
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  onOnceAnHour();

  // Закрываем день.
  Logger.log("Закрываем день.");
  const Hour0 = ss.getRangeByName('Час0').getValue(); // Время отсечения закрываемого дня
  const Hours0 = Hour0.getHours();
  const Minutes0 = Hour0.getMinutes();

  let nowDate = new Date ();
  const nowDay = nowDate.getDate();
  const prevDay = nowDay - 1;
  const nowDateTime = nowDate.setHours(0, 0, 0, 0); // Сегодняшняя дата 00:00
  let prevDate = new Date ();
  prevDate.setHours(0, 0, 0, 0);
  const prevDateTime = prevDate.setDate(prevDay); // Предыдущая дата

  let CutOffDate = new Date ();
  const CutOffTime = CutOffDate.setHours(Hours0, Minutes0, 0, 0); // Полная дата отсечения закрываемого дня
  let prevCutOffDate = new Date ();
  prevCutOffDate.setDate(prevDay); // Полная дата отсечения дня предыдущего закрываемому
  const CutOffPrev = prevCutOffDate.setHours(Hours0, Minutes0, 0, 0);

  Logger.log("Час0 сегодня: " + CutOffDate + " --- 00:00 сегодня: " + nowDate);
  Logger.log("Час0 вчера: " + prevCutOffDate + " --- 00:00 вчера: " + prevDate);

  const costs = ss.getSheetByName("Расходы");
  const costsData = costs.getDataRange();
  const cdRows = costsData.getNumRows();

  let thisDayRow = 0; // Первая строка Сегодня
  let prevDayRow = 0; // Первая строка закрываемого дня
  let oldDayRow = 0; // Последняя строка перед закрываемым днем
  let isToday = true;

  for (let i = 2; i < cdRows; i++) {
    const cDate = costsData.getCell(i, 1);
    const iDate = cDate.getValue();
    if (iDate === '') {
      thisDayRow = i;
      continue;
    }
    // Нашли первую запись с датой
    let iDateDayTime = iDate.getTime();
    if (iDateDayTime < prevDateTime) {
      //firstPrevPrevDayRow = i;
      oldDayRow = i;
      break; // Курсор вышел в позавчера
    }
    const cTime = costsData.getCell(i, 2);
    const iTime = cTime.getValue();
    if (iTime === '') {
      // Запись с датой без времени
      if (iDateDayTime < nowDateTime) {
        // Курсор опустился во вчера
        prevDayRow = i;
      } else if (isToday) {
        // Курсор еще в сегодня
        // lastNowDayRow = i;
        thisDayRow = i;
      }
    } else {
      // Запись с датой и временем
      const iTimeDayTime = iTime.getTime();
      if (iTimeDayTime > CutOffPrev) {
        if (iTimeDayTime > CutOffTime) {
          // Курсор еще выше времени отсечения сегодня. Пока еще в сегодня
          // lastNowTimeRow = i;
          thisDayRow = i;
        } else {
          // Курсор опустился ниже времени отсечения сегодня. Уже в закрываемом дне.
          isToday = false;
        }
      } else {
        //firstTimePrevRow = i;
        oldDayRow = i;
        break; // Курсор вышел ниже времени отсечения вчера
      }
    }
  }

  const topRow = thisDayRow + 1;
  Logger.log("Сегодня: " + thisDayRow + " вчера: " + prevDayRow + " позавчера: " + oldDayRow + " всего строк: " + (oldDayRow-topRow));
  prevDayRow = oldDayRow - 1;

  if (prevDayRow - thisDayRow == 0) {
    Logger.log("Вчера небыло трат. Закрытие не требуется.");
    return;
  }
  let frmlSumm = "=C" + topRow;
  // Группируем строки дня
  if (prevDayRow - thisDayRow > 1) {
    const rDay = costs.getRange(topRow + 1, 1, prevDayRow - topRow, 10);
    rDay.shiftRowGroupDepth(1);
    frmlSumm = "=SUM(C" + topRow + ":C" + prevDayRow + ")";
  }

  // Суммируем расходы дня
  const rSumm = costs.getRange(topRow, 10);
  rSumm.setFormula(frmlSumm);

  if (nowDay != 1) {
    // Подчеркиваем снизу если не закончился месяц
    const rPrevDay = costs.getRange(thisDayRow, 1, 1, 10);
    rPrevDay.setBorder(null, null, true, null, null, null);
  } else {
    // Закрываем месяц.
    Logger.log("Закрываем месяц.");
    const rThisMonth = costs.getRange(thisDayRow, 1, 1, 11);
    //rThisMonth.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
  }

  Logger.log("Закрыли день.")
}
