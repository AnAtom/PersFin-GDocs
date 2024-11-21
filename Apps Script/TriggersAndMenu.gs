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

/* UBER
 {"buyerPhoneOrAddress":"+79057685271","cashTotalSum":0,"code":3,"creditSum":0,

 "dateTime":"2024-11-10T06:19:00",
 "ecashTotalSum":101200,
 "fiscalDocumentFormatVer":4,"fiscalDocumentNumber":136327,"fiscalDriveNumber":"7380440801186965","fiscalSign":744264270,

 "fnsUrl":"www.nalog.gov.ru","internetSign":1,

 "items":[
  {"name":"Перевозка пассажиров и багажа","nds":6,"paymentAgentByProductType":64,"paymentType":4,"price":101200,"productType":1,"providerInn":"051302118203","quantity":1,"sum":101200}

 ],"kktRegId":"0000840547059265    ","machineNumber":"whitespirit2f","nds0":0,"nds10":0,"nds10110":0,"nds18":0,"nds18118":0,"ndsNo":101200,"operationType":1,"prepaidSum":0,
 "properties":{"propertyName":"psp_payment_id","propertyValue":"payment_3f8aa5a15e89465680f9510986ad40fd|authorization_0000"},
 "propertiesData":"ws:CNUJGVSRPH","provisionSum":0,"requestNumber":968,"retailPlace":"https://support-uber.com",
 "retailPlaceAddress":"248926, Россия, Калужская обл., г. Калуга, проезд 1-й Автомобильный, дом 8","sellerAddress":"support@support-uber.com","shiftNumber":137,"taxationType":1,"appliedTaxationType":1,

 "totalSum":101200,
 "user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"ЯНДЕКС.ТАКСИ\"","userInn":"7704340310  "}
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

  const sUberLabel = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRangeByName("ЧекиUber")
    .getValue();
  const threads = GmailApp
    .getUserLabelByName(sUberLabel)
    .getThreads();
  let k = 1;
  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      const message = messages[j];
      Logger.log( j + " > " + message.getSubject());
      const bInfo = getUBERBillInfo(message);
      if (flgDbg)
      {
        sTest.getRange(k, 1).setValue(message.getBody());
        sTest.getRange(k, 2).setValue(bInfo.summ);
        sTest.getRange(k, 3).setValue(bInfo.date);
      }
      k++;
    } // Сообщения с чеками UBER
  } // Цепочки сообщений с чеками UBER
}

/* Яндекс
 {"buyerPhoneOrAddress":"+79057685271","cashTotalSum":0,"code":3,"creditSum":0,

 "dateTime":"2024-10-19T03:15:00",
 "ecashTotalSum":55400,
 "fiscalDocumentFormatVer":4,"fiscalDocumentNumber":211161,"fiscalDriveNumber":"7386440800040048","fiscalSign":3663930572,

 "fnsUrl":"www.nalog.gov.ru","internetSign":1,
 "items":[
  {"name":"Перевозка пассажиров и багажа","nds":6,"paymentAgentByProductType":64,"paymentType":4,"price":55400,"productType":1,"providerInn":"504207820709","quantity":1,"sum":55400}

 ],"kktRegId":"0000840607026308    ","machineNumber":"whitespirit2f","nds0":0,"nds10":0,"nds10110":0,"nds18":0,"nds18118":0,"ndsNo":55400,"operationType":1,"prepaidSum":0,
 "properties":{"propertyName":"psp_payment_id","propertyValue":"payment_c9698b303b9347af89dfdb36bb4da522|authorization_0000"},
 "propertiesData":"ws:CICTKBVPRB","provisionSum":0,"requestNumber":877,"retailPlace":"taxi.yandex.ru",
 "retailPlaceAddress":"248926, Россия, Калужская обл., г. Калуга, проезд 1-й Автомобильный, дом 8","sellerAddress":"support@go.yandex.com","shiftNumber":233,"taxationType":1,"appliedTaxationType":1,

 "totalSum":55400,
 "user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"ЯНДЕКС.ТАКСИ\"","userInn":"7704340310  "}
*/

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

  //var TripTime = between2(fBody, "route__point-name", "</td>", "<p class=", "</p>");
  //var j = TripTime.indexOf(">");
  var TripTime = between(fBody, "route__point-name", "</table>");
  //Logger.log("Yandex Go>>>" + TripTime + "<<<");
  var j = TripTime.indexOf("route__point-name");
  TripTime = TripTime.slice(j+1);
  TripTime = between(TripTime, "<p class=", "</p>");
  j = TripTime.indexOf(">");
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

  const sYandexGoLabel = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRangeByName("ЧекиЯндексGo")
    .getValue();
  const threads = GmailApp.getUserLabelByName(sYandexGoLabel).getThreads();
  // if (flgDbg) SpreadsheetApp.getActive().toast(threads.length);

  let k = 1;

  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      Logger.log( j + " > " + message.getSubject());

      var bInfo = getYandexGoBillInfo(message);

      if (flgDbg)
      {
        sTest.getRange(k, 2).setValue(bInfo.summ);
        sTest.getRange(k, 3).setValue(bInfo.date);
        sTest.getRange(k, 4).setValue(message.getBody().length);
        sTest.getRange(k, 1).setValue(message.getBody());
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

  const flgDbg = dbgGetDbgFlag(true);
  
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
        if (flgDbg) {
          let s = dbgSplitLongString(Body, 4950);
          rTest.offset(k, 1, 1, s.length). setValues([s]);
        }
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"Перевозка пассажиров и багажа",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      } // Тема сообщения "Ваш номер заказа ..."
      else
      {
        var Body = message.getBody();
        Logger.log( j + " > " + subject + " <<< "+ Body.length.toString() +" >>>");

        if (flgDbg) rTest.offset(k, 0).setValue(" # " + subject + " <<< "+ Body.length.toString() +" >>>"); 
        if (flgDbg) {
          //dbgSplitLongString(rTest.offset(k, 1), Body);
          let s = dbgSplitLongString(Body, 49500);
          rTest.offset(k, 1, 1, s.length). setValues([s]);
        }

        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"Перевозка пассажиров и багажа",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      }
    } // Сообщения с чеками AliExpress
  } // Цепочки сообщений с чеками AliExpress
}

// Устанавливаем доступные счета и Тип операции для выбранной из общего списка операции
function SettingTrnctnName(ss, br) {
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
function SettingTrnctnType(ss, br) {
  const NewVal = br.getValue();
  if (ss.getRangeByName(NewVal) == undefined)
    NewVal = 'Операция'; // Устанавливаем полный список операций для выбора если Тип неизвестен

  SetTargetRule(ss, br.offset(0,-1), NewVal)
}

// Устанавливаем список расходов для выбранной статьи расходов
function SettingCostInfo(ss, br) {
  const flgDbg = dbgGetDbgFlag(false);

  if (flgDbg)
  {
    var rTest = ss.getSheetByName('Test').getRange(1, 1);
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
function SettingCostNote(ss, br) {
  const flgDbg = dbgGetDbgFlag(false);
  
  if (flgDbg)
  {
    // Лист для отладки
    var rTest = ss.getSheetByName('Test').getRange(1, 1);
  }

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);

  let range;
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
function SettingCostBill(ss, br) {
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

function ScanAli(ss, dLastAliDate, arrBills) {
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
      const mFrom = message.getFrom();
      Logger.log( "Письмо " + thrd + "#" + ++m + " от " + dDate.toISOString() + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + mFrom + " ." );

      bBill = {dTime: dDate.getTime(), date: dDate.toISOString(), summ: 0, cash: 0, name: "a", shop: "A"}
      /*try {
        Оформлен			17-04-2024, 19:13 UTC
        Сумма заказа	1396.24 ₽
        Номер заказа	5353566416757566

        bBill = mailGenericGetInfo(theTmplt, sBody);
      } catch (err) {
        Logger.log(">>> !!! Ошибка чтения чека из письма.", err);
        continue;
      }*/
      arrBills.push(bBill);
      Logger.log("Покупка N " + ++NumBills + dbgBillInfo(bBill));
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых покупок. Последнее письмо от " + newLastAliDate.toISOString());

  return newLastAliDate;
}

function ScanUber(ss, dLastUberDate, arrBills) {
  // Раз в день сканируем поездки Uber
  let newLastUberDate = dLastUberDate;
  let NumBills = 0;
  let bBill = {};

  // Сканируем цепочки писем
  let thrd = 1;
  const mailThreads = mailGetThreadByRngName('ЧекиUber');
  for (messages of mailThreads) {
    if (!messages.getLastMessageDate() > dLastUberDate)
      continue;

    let m = 0;
    for (message of messages.getMessages()) {
      const dDate = message.getDate();
      if (dDate > dLastUberDate) {
        if (dDate > newLastUberDate)
          newLastUberDate = dDate;
      } else
        continue;

      const sBody = message.getBody();
      const mFrom = message.getFrom();
      Logger.log( "Письмо " + thrd + "#" + ++m + " от " + dDate.toISOString() + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + mFrom + " ." );

      bBill = {dTime: dDate.getTime(), date: dDate.toISOString(), summ: 0, cash: 0, name: "u", shop: "U"}
      /*try {
                {"cashTotalSum":0,"dateTime":"2024-04-14T01:17:00",
                "fiscalDriveNumber":7281440701497667,"fiscalDocumentNumber":186536,"fiscalSign":4139559977,
                "items":[
                          {"name":"Перевозка пассажиров и багажа","price":94100,"quantity":1,"sum":94100,"unit":""}],
                "totalSum":94100,"user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"ЯНДЕКС.ТАКСИ\"","userInn":0}

          route__time
          check__price

        bBill = mailGenericGetInfo(theTmplt, sBody);
      } catch (err) {
        Logger.log(">>> !!! Ошибка чтения чека из письма.", err);
        continue;
      }*/
      arrBills.push(bBill);
      Logger.log("Покупка N " + ++NumBills + dbgBillInfo(bBill));
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых поездок. Последнее письмо от " + newLastUberDate.toISOString());

  return newLastUberDate;
}

function onOpen(e) {
  const menuScan = [
    {name: "Чеки UBER", functionName: 'MenuCheckUBER'},
    {name: "Чеки Яндекс Go", functionName: 'MenuCheckYandexGo'},
    {name: "Чеки AliExpress", functionName: 'MenuCheckAliExpress'},
    null,
    {name: "Очистить отладку", functionName: 'dbgClearTestSheet'}
  ];
  Logger.log('Добавляем пункты меню.');
  e.source.addMenu("Сканировать", menuScan);
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
            SettingTrnctnType(ss, br);
          return; // -----------------------------------------------------------------------------------------------------------------------------------------------------------------
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
    case 'i':
      if (ncol == 8 && br.getRow() == 23) {
        Logger.log("Переключаем сохранение чеков.");
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
  const rLastDriveDate = ss.getRangeByName('ДатаЧекДиск');
  let dLastDriveDate = rLastDriveDate.getValue();
  if (dLastDriveDate === "") {
    dLastDriveDate = dDate0;
    Logger.log("Принимаем дату последнего файла : " + dLastDriveDate);
  } else
    Logger.log("Дата последнего файла : " + dLastDriveDate);
  const newLastDriveDate = ScanDrive(ss, dLastDriveDate, newBills);

  // Сканируем чеки в почте
  const rLastMailDate = ss.getRangeByName('ДатаЧекПочта');
  let dLastMailDate = rLastMailDate.getValue();
  if (dLastMailDate === "") {
    dLastMailDate = dDate0;
    Logger.log("Принимаем дату последнего письма : " + dLastMailDate);
  } else
    Logger.log("Дата последнего письма : " + dLastMailDate);
  const newLastMailDate = ScanMail(ss, dLastMailDate, newBills);

  // Сканируем покупки Ali
  const rLastAliDate = getDateRangeDefault('ДатаЧекAli');

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
  let nowDayTime = newBills[0].tDate; // Первый день в списке чеков.

  const lstStores = ss.getRangeByName('СпскМагазины').getValues();
  const lstIgnore = ss.getRangeByName('спскМагзныИгнор').getValues();
  const shops = ss.getSheetByName('Магазины');
  const costs = ss.getSheetByName("Расходы");
  //const costs = ss.getSheetByName('Расходы (Тест)'); 
  const costsData = costs.getDataRange();
  let cdRows = costsData.getNumRows();

  // dTime: getDate(sDate).getTime(), // UNIX время чека
  // tDate: // UNIX время дня чека
  // date: sDate,  summ: nSumm,  cash: nCash,
  // name: sName,  shop: billFilterName(sName)};

  let firstDateSummRow = -1;
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

    if (isTime)
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
    else // Не указано время
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
      costs.insertRowsAfter(firstDateSummRow - 1, iInsrtBill);
      for (let j = 0; j < iInsrtBill; j++)
        setCostBill(costs.getRange(firstDateSummRow + j, 3), newBills[j], getShopInfoRemarkNote(newBills[j].shop, newBills[j].name, lstStores, lstIgnore, shops));
      i += iInsrtBill;
      cdRows += iInsrtBill;
    }
    // Удаляем обработанные чеки
    if (~iDelBill) {
      if (iDelBill == 0) {
        Logger.log('XXX Удаляем чек с суммой ' + newBills[0].summ + ' и датой ' + newBills[0].date);
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
      firstDateSummRow = -1; // Сбрасываем запомненную строку
    }
  }
  if (newBills.length > 0)
    Logger.log(">>>>>>>>>> !!! " + newBills.length + "!!! <<<<<<<<<<");
  Logger.log("Пробежались по чекам на листе. В списке найденных осталось :" + newBills.length);

  Logger.log("Обновляем даты.");
  if (newLastDriveDate > dLastDriveDate)
    rLastDriveDate.setValue(newLastDriveDate);
  if (newLastMailDate > dLastMailDate)
    rLastMailDate.setValue(newLastMailDate);
  //billsUBER.updateDate();
  //billsYandexGo.updateDate();
}

function onOnceADay() {
  // Выполняется ежежневно
  Logger.log("Сканируем поездки, покупки и закрываем день");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dDate0 = ss.getRangeByName('День1').getValue();

  onOnceAnHour();
  let newBills = [];

  // Сканируем в почте такси UBER - ЧекиUBER - ДатаЧекUber
  //ScanMailLabel(ss.getRangeByName("ЧекиUber").getValue(), getDateRangeDefault('ДатаЧекUber'), ScanUberMail, newBills);

  // Сканируем в почте такси Яндекс Go - ЧекиЯндексGo - ДатаЧекЯндексGo
  //ScanMailLabel(ss.getRangeByName("ЧекиЯндексGo").getValue(), getDateRangeDefault('ДатаЧекЯндексGo'), ScanYandexGoMail, newBills);

  // Сканируем в почте покупки Ali - ЧекиAli - ДатаЧекAli
  //ScanMailLabel(ss.getRangeByName("ЧекиAli").getValue(), getDateRangeDefault('ДатаЧекAli'), ScanAliMail, newBills);

  Logger.log("Сканируем покупки Ali");
  const rLastAliDate = ss.getRangeByName('ДатаЧекAli');
  let dLastAliDate = rLastAliDate.getValue();
  if (dLastAliDate === "") {
    dLastAliDate = dDate0;
    Logger.log("Принимаем дату последней покупки : " + dLastAliDate);
  } else
    Logger.log("Дата последней покупки : " + dLastAliDate);
  const newLastAliDate = ScanAli(ss, dLastAliDate, newBills);

  Logger.log("Сканируем поездки Uber");
  const rLastUberDate = ss.getRangeByName('ДатаЧекUber');
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
  }

  Logger.log("Закрыли день.")
}

function onOnceAMonth() {
  // Закрываем месяц.
  Logger.log("Закрываем месяц.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costs = ss.getSheetByName("Расходы");

  let thisMonthRow = 26;
  // Подчеркиваем снизу
  const rThisMonth = costs.getRange(thisMonthRow, 1, 1, 11);
  //rThisMonth.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);

}
