/*

onOpen(e)
onEdit(e)
onOnceAnHour()
onOnceADay()

Редактирование на листе «Операции»
  SettingTrnctnName - Устанавливаем доступные счета и Тип операции для выбранной из общего списка операции
  SettingTrnctnType - Устанавливаем доступный список операций для выбранного Типа операции

Редактирование на листе «Расходы»
  SettingCostInfo - Изменилась статья расходов, устанавливаем список расходов (Инфо)
  SettingCostNote - Изменился пункт статьи расходов (Инфо), устанавливаем список для Примечание
  SettingCostBill - Изменилась Заметка (парсим чек)

*/

// Типовые ежемесячные операции
/*
  Дата  Сумма   Счет        Цель        Операция        Тип
  -     -       -           -           -               -
-	31.01																	Снятие					Оборот
-	30.01	0,00 ₽	Сбер										Проценты крдт		Списание
-	30.01	0,00 ₽	Сбер										Погашение крдт	Списание
-	30.01	0,00 ₽	Кредит ВТБ							Проценты крдт		Списание
-	30.01	0,00 ₽	ЗП					Кредит ВТБ	Погашение крдт	Списание
-	26.01	0,00 ₽	ЗП					Rostelecom	Платеж					Списание
+	25.01	0,00 ₽	ЗП											Аванс						Начисление
-	23.01	0,00 ₽	ЗП					Комуналка		Платеж					Списание
-	22.01	0,00 ₽							Милана			Помог/подарил		Списание
-	19.01	0,00 ₽	ЗП					Квартира		Платеж					Списание
-	15.01	0,00 ₽	Кредит ВБРР							Проценты крдт		Списание
-	15.01	0,00 ₽	ВБРР				Кредит ВБРР	Погашение крдт	Списание
-	13.01	0,00 ₽	ЗП					ВБРР				Перевод					Оборот
-	15.01	0,00 ₽	МИР					Х5.Пакет		Платеж					Списание
-	11.01	0,00 ₽	МИР					Такие дела	Помог/подарил		Списание
+	10.01	0,00 ₽	ЗП											ЗП							Начисление
-	10.01	0,00 ₽	МИР					Yota				Платеж					Списание
-	09.01	0,00 ₽	МИР					Я.Плюс			Платеж					Списание
-	01.01																	Снятие					Оборот
*/

function putBillsToExpenses(jsonBillsArr)
{
  //
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costs = ss.getSheetByName("Расходы");
  
  // costs.expandAllRowGroups();
  const costsData = costs.getDataRange();

  const cdRows = costsData.getNumRows();
  const cdColumns = costsData.getNumColumns();
  //Logger.log(costsData.getCell(cdRows, cdColumns).getValue());

  let theDate = jsonBillsArr[0].date;
  let aDate = new Date(jsonBillsArr[0].date);
  let theDay = aDate.getDate();
  let theMonth = aDate.getMonth();

  let prevDayRow = 0;
  let nextDayRow = 0;
  let insertRow = 0;

  // Находим начало дней
  // Сканируем день
  //
  // Находим окончание 
  // Находим окончание месяца
  for (var i = 2; i < cdRows; i++) {
    let n = 1;
    let cDate = costsData.getCell(i, 1);
    let iDate = cDate.getValue();
    if (iDate == "") continue;
    let dDate = new Date(iDate);
    let aDateDay = dDate.getDate();
    // Нашли первую запись
    let sDate = iDate.toISOString();
    if (costsData.getCell(i, 2).getValue() != "") {
      //
      Logger.log(sDate + " время " + costsData.getCell(i, 2).getValue());
    }
  }
}

function getUBERBillInfo(BillMail) {
  const sTripDate = BillMail.getSubject().slice(23);
  const spcPos = sTripDate.indexOf(" ");
  const spcPos2 = sTripDate.indexOf(" ", spcPos+2);
  const TripDate = sTripDate.slice(0, spcPos) + "."
    + getMonthNum(sTripDate.slice(spcPos+1, spcPos2)) + "."
    + sTripDate.slice(spcPos2+1, sTripDate.indexOf(" г.", spcPos2+2));

  const fBody = BillMail.getBody();

  let TripTime = between2(fBody, "From", "</tr>", "<td align", "</td>");
  TripTime = TripTime.slice(TripTime.indexOf(">")+1).trim();

  const TripDateTime = TripDate + " " + TripTime;
  const TripSumm = between2(fBody, "check__price", "</td>", ">", " ₽").trim();

  const bInfo = {summ: TripSumm, date: TripDateTime, name: '"ООО \"ЯНДЕКС.ТАКСИ\""', items: [{iname:"Перевозка пассажиров и багажа", iprice:TripSumm, isum:TripSumm, iquantity:1.0}]};
  Logger.log("UBER > ", bInfo);
  return bInfo;
}

// Пункт меню Сканировать - Чеки UBER
function MenuCheckUBER() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flgDbg = dbgGetDbgFlag(true);
  const sTest = ss.getSheetByName("Test"); // Лист для отладки

  let k = 1;
  const threads = GmailApp
    .getUserLabelByName("Моё/Мани/Такси")
    .getThreads();
  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      const message = messages[j];
      Logger.log( j + " > " + message.getSubject());
      const bInfo = getUBERBillInfo(message);
      if (flgDbg)
      {
        sTest.getRange(k, 2).setValue(bInfo.summ);
        sTest.getRange(k, 3).setValue(bInfo.date);
      }
      k++;
    } // Сообщения с чеками UBER
  } // Цепочки сообщений с чеками UBER
}

function getYandexGoBillInfo(BillMail) {
  const fSubject = BillMail.getSubject();
  let sTripDate = fSubject.slice(28);

  let spcPos = sTripDate.indexOf(" ");
  let sTripDay = sTripDate.slice(0, spcPos);
  let spcPos2 = sTripDate.indexOf(" ", spcPos+2);
  let sTripMonth = sTripDate.slice(spcPos+1, spcPos2);

  let TripMonth = getMonthNum(sTripMonth);
  let TripYear = sTripDate.slice(spcPos2+1, sTripDate.indexOf(" г.", spcPos2+2));

  var TripDate = sTripDay + "." + TripMonth + "." + TripYear;

  let fBody = BillMail.getBody();
  // finLib.between2();

  var TripTime = between2(fBody, "route__point-name", "</td>", "<p class=", "</p>");
  var j = TripTime.indexOf(">");
  TripTime = TripTime.slice(j+1).trim();

  var TripDateTime = TripDate + " " + TripTime;

  var TripSumm = between2(fBody, "report__value_main", "</td>", ">", " ₽").trim();

  var bInfo = {summ: TripSumm, date: TripDateTime, name: '"ООО \"ЯНДЕКС.ТАКСИ\""', items: [{iname:"Перевозка пассажиров и багажа", iprice:TripSumm, isum:TripSumm, iquantity:1.0}]};

  Logger.log("Yandex Go> ", bInfo);
  return bInfo;
}

function MenuCheckYandexGo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flgDbg = dbgGetDbgFlag(true);
  const sTest = ss.getSheetByName("Test"); // Лист для отладки

  let k = 1;

  const threads = GmailApp.getUserLabelByName("pers/отчеты/такси").getThreads();
  // if (flgDbg) SpreadsheetApp.getActive().toast(threads.length);

  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      Logger.log( j + " > " + message.getSubject());
      if (flgDbg) sTest.getRange(k, 1).setValue(message.getBody());

      var bInfo = getYandexGoBillInfo(message);

      if (flgDbg)
      {
        sTest.getRange(k, 2).setValue(bInfo.summ);
        sTest.getRange(k, 3).setValue(bInfo.date);
      }
      k++;
    } // Сообщения с чеками Яндекс Go
  } // Цепочки сообщений с Яндекс Go
}

function getAliExpressBillInfo(BillMail) {
  const fSubject = BillMail.getSubject();

}

function MenuCheckAliExpress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var flgDbg = dbgGetDbgFlag(true);
  
  // Лист для отладки
  var sTest = ss.getSheetByName("Test");
  var rTest = sTest.getRange(1, 1);

  var k = 0;

  var label = GmailApp.getUserLabelByName("Моё/Покупки/AliExpress");
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      Logger.log( j + " > " + subject + " >> " + subject.indexOf("Ваш номер заказа").toString());
      if (subject.indexOf("Ваш номер заказа") != -1) {
        var Body = message.getBody();
        Logger.log( j + " > " + subject + " [[[ "+ Body.length.toString() +" ]]]");

        if (flgDbg) rTest.offset(k, 0).setValue(" > " + subject + " [[[ "+ Body.length.toString() +" ]]]"); 
        if (flgDbg) dbgLongMailBody(rTest.offset(k, 1), Body);
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"Перевозка пассажиров и багажа",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      } // Тема сообщения "Ваш номер заказа ..."
      else
      {
        var Body = message.getBody();
        Logger.log( j + " > " + subject + " <<< "+ Body.length.toString() +" >>>");

        if (flgDbg) rTest.offset(k, 0).setValue(" # " + subject + " <<< "+ Body.length.toString() +" >>>"); 
        if (flgDbg) dbgLongMailBody(rTest.offset(k, 1), Body);
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"Перевозка пассажиров и багажа",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      }
    } // Сообщения с чеками AliExpress
  } // Цепочки сообщений с чеками AliExpress
}

// Устанавливаем доступные счета и Тип операции для выбранной из общего списка операции
function SettingTrnctnName(ss, br)
{
  const debit = 'Списание';
  const moving = 'Оборот';

  const NewVal = br.getValue();
  const OpAcc = br.offset(0,-2);
  const OpTrgt = br.offset(0,-1);

  let i = findInRule(moving, NewVal);
  if (~i)
  {
    // Выбрана оборотная операция
    br.offset(0,1).setValue(moving);

    SetTargetRule(ss, OpAcc, 'СчетаДеб');

    const Transfer = ss.getRangeByName('стрПеревод').getValue();
    if (NewVal == Transfer)
      SetTargetRule(ss, OpTrgt, 'СчетаДеб'); // Перевод
    else {
      OpTrgt.clearDataValidations();
      if (i == 0) // Снятие
        OpTrgt.clear();
    }
  }
  else if (~findInRule(debit, NewVal))
  {
    // Выбрана опреация списания
    br.offset(0,1).setValue(debit);

    const CredPersnt = ss.getRangeByName('стрПрцКрдт').getValue();
    if (NewVal == CredPersnt) {
      // Проценты по кредиту
      SetTargetRule(ss, OpAcc, 'Кредиты');
      OpTrgt.clear();
    }
    else
    {
      SetTargetRule(ss, OpAcc, 'СчетаДеб');

      const LoanPaymnt = ss.getRangeByName('стрПогКрдт').getValue();
      if (NewVal == LoanPaymnt) {
        // Погашение кредита
        SetTargetRule(ss, OpTrgt, 'Кредиты');
      }
      else
      {
        const Payment = ss.getRangeByName('стрПлатеж').getValue();
        if (NewVal == Payment) {
          // Платеж
          SetTargetRule(ss, OpTrgt, 'Платежи');
        }
        else OpTrgt.clearDataValidations();
      }
    }
  }
  else
  {
    const receipt = 'Начисление';
    i = findInRule(receipt, NewVal);
    if (~i) {
      // Выбрана операция начисления
      br.offset(0,1).setValue(receipt);

      SetTargetRule(ss, OpAcc, 'СчетаДеб');
      if (i < 4) OpAcc.setValue("ЗП");
    }
  }
}

// Устанавливаем соответствующий список операций для выбранного Типа операции
function SettingTrnctnType(ss, br)
{
  const NewVal = br.getValue();
  if (ss.getRangeByName(NewVal) == undefined)
    NewVal = 'Операция'; // Устанавливаем полный список операций для выбора если Тип неизвестен

  SetTargetRule(ss, br.offset(0,-1), NewVal)
}

// Устанавливаем список расходов для выбранной статьи расходов
function SettingCostInfo(ss, br)
{
  const flgDbg = dbgGetDbgFlag(false);
  
  if (flgDbg)
  {
    // Лист для отладки
    var sTest = ss.getSheetByName('Test');
    var rTest = sTest.getRange(1, 1);
  }
  //br.setNote('Test :' + sTest + ' Range :' + rTest.getNumRows());

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  //br.setNote('Row :' + flgDbg + ' Val :' + NewVal);
  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);
  
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
function SettingCostNote(ss, br)
{
  const flgDbg = dbgGetDbgFlag(false);
  
  if (flgDbg)
  {
    // Лист для отладки
    var rTest = ss.getSheetByName('Test').getRange(1, 1);
  }

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);

  var range;
  //SpreadsheetApp.getActive().toast('Range :'+ range);

  if (NewVal == 'Продукты') range = ss.getRangeByName('СтРсхЕдаМагаз');
  else if (NewVal == 'Пиво') range = ss.getRangeByName('СтРсхАлкПиво');
  else if (NewVal == 'Кабак') range = ss.getRangeByName('СтРсхАлкКабак');

  if (range != undefined)
  {
    const rule = SpreadsheetApp
    .newDataValidation()
    .setAllowInvalid(true)
    .requireValueInRange(range)
    .build();

    cell.setDataValidation(rule);
  }
  else cell.clearDataValidations();
}

// Читаем информацию о чеке из json строки
function SettingCostBill(ss, br)
{
  const flgDbg = dbgGetDbgFlag(false);

  if (flgDbg)
  {
    // Лист для отладки
    var rTest = ss.getSheetByName('Test').getRange(1, 1);
  }

  const NewVal = br.getValue();
  if (NewVal == "") return;

  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);

  const bill = billInfo(NewVal);

  if (flgDbg) {
    if (bill != undefined)
      rTest.offset(3, 1).setValue(bill.name)
      .offset(1, 0).setValue(bill.summ)
      .offset(1, 0).setValue(bill.date)
      .offset(1, 0).setValue(bill.cash)
      .offset(1, 0).setValue(bill.shop);
    else rTest.offset(3, 1).setValue("UNDEFINED !!!");
  }

  if (bill == undefined) return;
  // Формат ячеек
  // "dd.mm", "HH:mm", "#,##0.00[$ ₽]"

  // Выставляем сумму покупки
  br.offset(0,-5)
  .setValue(bill.summ)
  .setNumberFormat("#,##0.00[$ ₽]");

  // Выставляем дату покупки и получаем адрес ячейки с датой для выставления времени
  const A1date = br.offset(0,-7).setValue(bill.date).setNumberFormat("dd.mm").getA1Notation();

  if (flgDbg) rTest.offset(8, 1).setValue(A1date);

  // Выставляем время покупки
  br.offset(0,-6)
  .setValue("=" + A1date)
  .setNumberFormat("HH:mm");

  // Если наличные, то выставляем счет списания
  if (bill.cash != 0)
    br.offset(0,-4).setValue("Карман")

  // Выставляем Статью, Инфо и Примечание для магазина
  const lstStores = ss.getRangeByName('СпскМагазины');
  let shop = lstStores.getValues().find((element) => element[3] == bill.shop);
  if (shop != undefined) {
    if (flgDbg) rTest.offset(9, 1).setValue(shop.toString());
    for (let i = 0; i < 3; i++)
      br.offset(0,i-3).setValue(shop[i]);
  } else {
    // Добавляем в список новый магазин
    const sShop = ss.getSheetByName('Магазины');
    const newRow = lstStores.getNumRows() + 4;
    sShop.insertRowBefore(newRow);
    sShop.getRange(newRow, 4, 1, 2).setValues([[bill.shop, bill.name]]);
  }
}

function ScanAli(ss, dLastAliDate, arrBills)
{
  //

}

function ScanUber(ss, dLastUberDate, arrBills)
{
  //

}

function onEdit(e) 
{
  const ss = e.source;

  // Читаем флаг "Использовать автосписки"
  let br = ss.getRangeByName('ФлАвтосписки');
  if (br == undefined || ! br.getValue()) return;

  const TrnctnSheet = 'Операции';
  const CostsSheet = 'Расходы';

  br = e.range;
  if (br.getNumColumns() > 1) return; // Скопировали диапазон

  const ncol = br.getColumn();
  const sname = ss.getActiveSheet().getSheetName();
  //SpreadsheetApp.getActive().toast(sname);

  if (sname == TrnctnSheet)
  {
    if (ncol == 7)
    {
      let v = e.value;
      // Изменился тип операции
      if (v == undefined || v == '') // Устанавливаем полный список операций для выбора если Тип операции был очищен
        SetTargetRule(ss, br.offset(0,-1), 'Операция');
      else SettingTrnctnType(ss, br);
    }
    else if (ncol == 6)
    {
      // Изменилась операция
      SettingTrnctnName(ss, br);
    }
  }
  else if (sname == CostsSheet)
  {
    switch(ncol) {
    case 5:
      // Изменилась статья расходов
      SettingCostInfo(ss, br);
      break;
    case 6:
      // Изменился пункт статьи расходов (Инфо)
      SettingCostNote(ss, br);
      break;
    case 8:
      // Изменилась заметка (вставили чек)
      SettingCostBill(ss, br);
      break;
    }
  }
}

function MenuCloseDay()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var costs = ss.getSheetByName("Расходы");
  
  var flgDbg = dbgGetDbgFlag(true);
  
  // Лист для отладки
  var sTest = ss.getSheetByName("Test");
  var rTest = sTest.getRange(1, 1);

  var k = 1;

  //costs.expandAllRowGroups();
  var costsData = costs.getDataRange();

  var cdRows = costsData.getNumRows();
  var cdColumns = costsData.getNumColumns();
  if (flgDbg) 
  {
    rTest.offset(k, 1).setValue(cdRows)
    .offset(0, 1).setValue(cdColumns)
    .offset(0, 1).setValue(costsData.getValue());
    Logger.log(costsData.getCell(cdRows, cdColumns).getValue());
  }

  for (var i = 2; i < cdRows; i++) {
    var n = 1;
    var cData = costsData.getCell(i, 1);
    var iData = cData.getValue();

  }
}

/*
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('First item', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
*/

function onOpen(e)
{
  Logger.log('Добавляем пункты меню.');

  const menuScan = [
    {name: "Чеки UBER", functionName: 'MenuCheckUBER'},
    {name: "Чеки Яндекс Go", functionName: 'MenuCheckYandexGo'},
    {name: "Чеки AliExpress", functionName: 'MenuCheckAliExpress'},
    null,
    {name: "Очистить отладку", functionName: 'dbgClearTestSheet'}
  ];
  e.source.addMenu("Сканировать", menuScan);

  const menuFinance = [
    {name: "Закрыть день", functionName: 'MenuCloseDay'}

  ];
  e.source.addMenu("Финансы", menuFinance);

}

function onOnceAnHour()
{
  // Выполняется ежечасно
  Logger.log("Обрабатываем последние чеки");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let newBills = [];

  // Сканируем чеки в почте
  const rLastMailDate = ss.getRangeByName('ДатаЧекПочта');
  const dLastMailDate = ReadLastDate(ss, rLastMailDate);
  const newLastMailDate = ScanMail(ss, dLastMailDate, newBills);

  // Сканируем диск
  const rLastDriveDate = ss.getRangeByName('ДатаЧекДиск');
  const dLastDriveDate = ReadLastDate(ss, rLastDriveDate);
  const newLastDriveDate = ScanDrive(ss, dLastDriveDate, newBills);

  Logger.log("Обновляем " + newBills.length + " чеков.");

  Logger.log("Обновляем даты.");
}

function onOnceADay()
{
  // Выполняется ежежневно
  Logger.log("Закрываем день");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let newBills = [];

  Logger.log("Сканируем покупки Ali");
  const rLastAliDate = ss.getRangeByName('ДатаЧекДиск');
  const dLastAliDate = ReadLastDate(ss, rLastAliDate);
  const newLastAliDate = ScanAli(ss, dLastAliDate, newBills);

  Logger.log("Сканируем поездки Uber");
  const rLastUberDate = ss.getRangeByName('ДатаЧекДиск');
  const dLastUberDate = ReadLastDate(ss, rLastUberDate);
  const newLastUberDate = ScanUber(ss, dLastUberDate, newBills);

  Logger.log("Обновляем " + newBills.length + " чеков.");

  Logger.log("Обновляем даты.");

  // Закрываем день.
  Logger.log("Закрываем день.");

}
