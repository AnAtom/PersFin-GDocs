/*

onOpen(e)
onEdit(e)
onOnceAnHour()
onOnceADay()
onOnceAMonth()

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
    const sTest = ss.getSheetByName('Test');
    const rTest = sTest.getRange(1, 1);
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
  if (shop == undefined) {
    // Добавляем в список новый магазин
    Logger.log("Новый магазин [" + bill.shop + "] (" + bill.name + ")");
    const sShop = ss.getSheetByName('Магазины');
    const newRow = lstStores.getNumRows() + 4;
    sShop.insertRowBefore(newRow);
    sShop.getRange(newRow, 4, 1, 2).setValues([[bill.shop, bill.name]]);
  } else
    for (let i = 0; i < 3; i++)
      br.offset(0,i-3).setValue(shop[i]);
}

function ScanAli(ss, dLastAliDate, arrBills)
{
  // Раз в день сканируем заказы AliExpress
  let newLastAliDate = dLastAliDate;
  let NumBills = 0;
  let bBill = {};

  // Сканируем цепочки писем
  let thrd = 1;
  const mailThreads = mailGetThreadByRngName('ЧекиAli');
  for (messages of mailThreads) {
    if (!messages.getLastMessageDate() > dLastAliDate)
      continue;

    let m = 0;
    for (message of messages.getMessages()) {
      const dDate = message.getDate();
      if (dDate > dLastAliDate) {
        if (dDate > newLastAliDate)
          newLastAliDate = dDate;
      } else
        continue;

      const sBody = message.getBody();
      const mFrom = between(message.getFrom(), "<", ">");
      Logger.log( "Письмо " + thrd + "#" + ++m + " от " + dDate.toISOString() + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + mFrom + " ." );

      bBill = {dTime: dDate.getTime(), date: dDate.toISOString(), summ: 0, cash: 0, name: "a", shop: "A"}
      /*try {
        bBill = mailGenericGetInfo(theTmplt, sBody);
      } catch (err) {
        Logger.log(">>> !!! Ошибка чтения чека из письма.", err);
        continue;
      }*/
      arrBills.push(bBill);
      Logger.log("Чек N " + ++NumBills + dbgBillInfo(bBill));
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых чеков. Последнее письмо от " + newLastAliDate.toISOString());

  return newLastAliDate;
}

function ScanUber(ss, dLastUberDate, arrBills)
{
  // Раз в день сканируем поездки Uber

  return dLastUberDate;
}

function onEdit(e) 
{
  const ss = e.source;

  // Читаем флаг "Использовать автосписки"
  let br = ss.getRangeByName('ФлАвтосписки');
  if (br == undefined || ! br.getValue()) return;

  br = e.range;
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
            SettingTrnctnType(ss, br);
          return;
        case 6: // Изменилась Операция
          SettingTrnctnName(ss, br);
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
  const dDate0 = ss.getRangeByName('День1').getValue();

  let newBills = [];

  // Сканируем чеки в почте
  const rLastMailDate = ss.getRangeByName('ДатаЧекПочта');
  let dLastMailDate = rLastMailDate.getValue();
  if (dLastMailDate === "") {
    dLastMailDate = dDate0;
    Logger.log("Принимаем дату последнего письма : " + dLastMailDate);
  } else
    Logger.log("Дата последнего письма : " + dLastMailDate);
  const newLastMailDate = ScanMail(ss, dLastMailDate, newBills);

  // Сканируем диск
  const rLastDriveDate = ss.getRangeByName('ДатаЧекДиск');
  let dLastDriveDate = rLastDriveDate.getValue();
  if (dLastDriveDate === "") {
    dLastDriveDate = dDate0;
    Logger.log("Принимаем дату последнего файла : " + dLastDriveDate);
  } else
    Logger.log("Дата последнего файла : " + dLastDriveDate);
  const newLastDriveDate = ScanDrive(ss, dLastDriveDate, newBills);

  const cntBills = newBills.length;
  Logger.log("Обновляем " + cntBills + " чеков.");

  if (cntBills > 1)
    newBills.sort((a, b) => b.dTime - a.dTime); // Первым будет самый свежий чек

  let firstDay = newBills[0].tDate; // Первый день в списке чеков.
  // const lastDay = newBills[cntBills - 1].tDate; // Последний день в списке чеков. Более старые дни не смотрим.

  const lstStores = ss.getRangeByName('СпскМагазины').getValues();
  const lstIgnore = ss.getRangeByName('спскМагзныИгнор').getValues();
  const shops = ss.getSheetByName('Магазины');
  const costs = ss.getSheetByName("Расходы");
  // const costs = ss.getSheetByName("Расходы (Тест)");
  const costsData = costs.getDataRange();
  let cdRows = costsData.getNumRows();

  // dTime: getDate(sDate).getTime(), // UNIX время чека
  // tDate: // UNIX время дня чека
  // date: sDate,
  // summ: nSumm,
  // cash: nCash,
  // name: sName,
  // shop: billFilterName(sName)};

  for (var i = 2; i < cdRows; i++) {
    const iDate = costsData.getCell(i, 1).getValue();
    if (iDate === '') // Пустая строка. Продолжаем...
      continue;
    // Нашли на листе строку с датой
    const iDateDayTime = iDate.getTime();
    const iTime = costsData.getCell(i, 2).getValue();
    const iSumm = costsData.getCell(i, 3).getValue();

    if (iTime === '') {
      if (iSumm === '') { // Нет ни времени, ни суммы. Только дата. Проверяем не перешли ли мы во вчерашний день
        if (iDateDayTime < firstDay) { // Перешли во вчерашний день. Вставляем все сегодняшние чеки.
          const l = i - 1;
          Logger.log("После строки " + l + " и перед строкой с датой " + iDate + " пробуем вставить чеки в предыдущий день.");
          let k = -1;
          // Вставляем все чеки, которые старше этого времени
          for (bill of newBills) {
            if (bill.tDate < firstDay)
              break;
            costs.insertRowAfter(++k + l);
            cdRows++;
            setCostBill(costs.getRange(i++, 3), bill, getShopInfoRemarkNote(bill.shop, bill.name, lstStores, lstIgnore, shops));
          }
          // Удаляем эти чеки
          if (~k) {
            k++;
            Logger.log("В предыдущий день вставили " + k + " чеков. Удаляем эти чеки из найденных.");
            newBills.splice(0, k);
            const cnt = newBills.length;
            if (cnt == 0) {
              Logger.log("Все чеки обработаны.");
              break;
            }
            Logger.log("Осталось " + cnt + " чеков для обработки.");
          } else
            Logger.log("Не нашли чеки для вставки в предыдущий день!!!!");
          firstDay = newBills[0].tDate;
          // continue;
        }
      } else {
        // Если есть сумма и нет времени (забил расход в этот день без подробностей), то ищем чек для обновления строчки
        Logger.log("Ищем чек с суммой " + iSumm + " для уточнения времени и другой информации в строке " + i);
        let k = newBills.findIndex((bill) => bill.summ == iSumm && bill.tDate == iDateDayTime)
        if (~k) {
          setCostBill(costs.getRange(i, 3), bill, getShopInfoRemarkNote(bill.shop, bill.name, lstStores, lstIgnore, shops));
          costs.getRange(i, 10).setBackground("red");
          Logger.log("Обновили информацию в строке " + i + ". Удаляем чек " + k + ".");
          newBills.splice(k, 1);
          const cnt = newBills.length;
          if (cnt == 0) {
            Logger.log("Все чеки обработаны.");
            break;
          }
          Logger.log("Осталось " + cnt + " чеков для обработки.");
        } else
          Logger.log("Не нашли такой чек. Игнорируем строку " + i);
      }
    } else {
      // Если есть время и есть/нет суммы. Вставляем перед строкой все новые чеки iDateDayTime < bill.dTime
      const l = i - 1;
      Logger.log("После строки " + l + " и перед строкой с датой " + iDate + " пробуем вставить чеки с датой свежее.");
      let k = -1;
      // Вставляем все чеки, которые старше этого времени
      for (bill of newBills) {
        if (bill.dTime < iDateDayTime)
          break;
        costs.insertRowAfter(++k + l);
        cdRows++;
        setCostBill(costs.getRange(i++, 3), bill, getShopInfoRemarkNote(bill.shop, bill.name, lstStores, shops));
      }
      // Удаляем эти чеки
      if (~k) {
        k++;
        Logger.log("Перед датой " + iDate + " вставили " + k + " чеков. Удаляем эти чеки из найденных.");
        newBills.splice(0, k);
        const cnt = newBills.length;
        if (cnt == 0) {
          Logger.log("Все чеки обработаны.");
          break;
        }
        Logger.log("Осталось " + cnt + " чеков для обработки.");
      } else
        Logger.log("Не нашли чеки для вставки перед этой датой!");
      // continue;
    }
    // Указаны и дата и сумма и время покупки. Что еще можно добавить?
  }
  Logger.log("Пробежались по чекам на листе. В списке найденных осталось :" + newBills.length);

  Logger.log("Обновляем даты.");
  if (newLastDriveDate > dLastDriveDate)
    rLastDriveDate.setValue(newLastDriveDate);
  if (newLastMailDate > dLastMailDate)
    rLastMailDate.setValue(newLastMailDate);
}

function onOnceADay()
{
  // Выполняется ежежневно
  Logger.log("Сканируем поездки, покупки и закрываем день");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const Date0 = ss.getRangeByName('День1').getValue();

  let newBills = [];

  Logger.log("Сканируем покупки Ali");
  const rLastAliDate = ss.getRangeByName('ДатаЧекДиск');
  let dLastAliDate = rLastAliDate.getValue();
  if (dLastAliDate === "") {
    dLastAliDate = dDate0;
    Logger.log("Принимаем дату последней покупки : " + dLastAliDate);
  } else
    Logger.log("Дата последней покупки : " + dLastAliDate);
  const newLastAliDate = ScanAli(ss, dLastAliDate, newBills);

  Logger.log("Сканируем поездки Uber");
  const rLastUberDate = ss.getRangeByName('ДатаЧекДиск');
  let dLastUberDate = rLastUberDate.getValue();
  if (dLastUberDate === "") {
    dLastUberDate = dDate0;
    Logger.log("Принимаем дату последней покупки : " + dLastUberDate);
  } else
    Logger.log("Дата последней покупки : " + dLastUberDate);
  const newLastUberDate = ScanUber(ss, dLastUberDate, newBills);

  Logger.log("Обновляем " + newBills.length + " чеков.");

  Logger.log("Обновляем даты.");
/*
  if (newLastAliDate > dLastAliDate)
    rLastAliDate.setValue(newLastAliDate);
  if (newLastUberDate > dLastUberDate)
    rLastUberDate.setValue(newLastUberDate);
*/

  // Закрываем день.
  Logger.log("Закрываем день.");
  const Hour0 = ss.getRangeByName('Час0').getValue(); // Время отсечения закрываемого дня
  const Hours0 = Hour0.getHours();
  const Minutes0 = Hour0.getMinutes();

  //const dd = 0;

  let nowDate = new Date ();
  //  nowDate.setDate(nowDate.getDate()-dd);
  const prevDay = nowDate.getDate() - 1;
  const nowDateTime = nowDate.setHours(0, 0, 0, 0); // Сегодняшняя дата 00:00
  //const nowDateTime = nowDate.getTime();
  let prevDate = new Date ();
  //  prevDate.setDate(prevDate.getDate()-dd);
  prevDate.setHours(0, 0, 0, 0);
  const prevDateTime = prevDate.setDate(prevDay); // Предыдущая дата
  //const prevDateTime = prevDate.getTime();

  let CutOffDate = new Date ();
  //  CutOffDate.setDate(CutOffDate.getDate()-dd);
  const CutOffTime = CutOffDate.setHours(Hours0, Minutes0, 0, 0); // Полная дата отсечения закрываемого дня
  //const CutOffTime = CutOffDate.getTime(); // Время отсечения закрываемого дня
  let prevCutOffDate = new Date ();
  //  prevCutOffDate.setDate(CutOffDate.getDate()-dd);
  prevCutOffDate.setDate(prevDay); // Полная дата отсечения дня предыдущего закрываемому
  const CutOffPrev = prevCutOffDate.setHours(Hours0, Minutes0, 0, 0);
  //const CutOffPrev = prevCutOffDate.getTime(); // Время отсечения дня предыдущего закрываемому

  Logger.log("Час0 сегодня: " + CutOffDate);
  Logger.log("00:00 сегодня: " + nowDate);
  Logger.log("час0 вчера: " + prevCutOffDate);
  Logger.log("00:00 вчера: " + prevDate);

  const costs = ss.getSheetByName("Расходы");
  // const costs = ss.getSheetByName("Лист17");
  //const costs = ss.getSheetByName("Расходы (Тест)");
  const costsData = costs.getDataRange();
  const cdRows = costsData.getNumRows();

  //let lastNowTimeRow = 0; // Последняя строка сегодняшнего дня с указанным временем
  //let lastNowDayRow = 0; // Последняя строка сегодняшнего дня с пустым временем
  //let lastTimeRow = 0; // Последняя строка закрываемого дня с указанным временем
  //let lastDayRow = 0; // Последняя строка закрываемого дня с пустым временем
  //let firstTimePrevRow = 0; // Первая строка позапрошлого дня с указанным временем
  //let firstPrevPrevDayRow = 0; // Первая строка позапрошлого дня без указания времени

  let thisDayRow = 0; // Первая строка Сегодня
  let prevDayRow = 0; // Первая строка закрываемого дня
  let oldDayRow = 0; // Последняя строка перед закрываемым днем
  let isToday = true;

  for (var i = 2; i < cdRows; i++) {
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

  // Подчеркиваем снизу
  const rPrevDay = costs.getRange(thisDayRow, 1, 1, 10);
  rPrevDay.setBorder(null, null, true, null, null, null);

  Logger.log("Закрыли день.")
}

function onOnceAMonth()
{
  // Закрываем месяц.
  Logger.log("Закрываем месяц.");

}
