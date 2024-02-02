/*

onEdit(e)
onOpen(e)
onOnceAnHour()

ReadDriveOnTimer(rLastBillDate, dDay1)

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

// Читает последние чеки с диска
function ReadDriveOnTimer(rLastBillDate, dDay1)
{
  // Получаем Id папки с чеками из ячейки ЧекиДиск
  const folderId = GetGDriveFolderIdFromURL('ЧекиДиск');
  const folderBills = DriveApp.getFolderById(folderId);
  Logger.log("Папка с чеками: " + folderBills.getName() + " Id: " + folderId);

  // Читаем дату последнего обработанного чека
  let dLastBillDate = rLastBillDate.getValue();
  const sLastBillDate = dLastBillDate.toString();
  if (sLastBillDate == "") {
    // Ячейка с датой пуста
    dLastBillDate = dDay1;
    Logger.log("Принимаем дату последнего чека на диске: " + dLastBillDate.toString());
  }
  else
    Logger.log("Дата последнего чека на диске: " + sLastBillDate);

  const iLastMonth = dLastBillDate.getMonth();
  const iLastDay = dLastBillDate.getDate();

  let newLastBillDate = dLastBillDate;
  let NumBills = 0;
  let jsonBillsArr = [];
  let i = 0;

  // Сканируем вложенные папки
  const folders = folderBills.getFolders();
  while (folders.hasNext()) {
    let folder = folders.next();
    let nMonth = folder.getName().slice(3);
    
    Logger.log("Папка " + nMonth);
    let iMonth = getMonthNum(nMonth, true);

    if (iLastMonth > iMonth) continue; // Чеки в папке старше последнего обработанного чека

    // Сканируем файлы чеков
    let newBills = 0;
    let newBillsArr = [];
    let files = folder.getFiles();
    while (files.hasNext()) {
      let fBill = files.next();

      let sBill = fBill.getBlob().getDataAsString();
      if (sBill == undefined) continue;

      // Читаем дату чека
      i = sBill.indexOf("\"dateTime\":")+12;
      if (i < 12) continue;
      let sDate = sBill.slice(i, sBill.indexOf("\"", i+1));

      if (iLastMonth == iMonth && parseInt(sDate.slice(8, 10), 10) < iLastDay) continue;

      let dDate = new Date(sDate);
      if (dDate > dLastBillDate)
        Logger.log("Дата: " + sDate + " Файл: " + fBill.getName());
      else
        continue;

      // Сохраняем текст чека для переноса в файл месяца
      newBills = newBillsArr.push( {
        BillDate: dDate, 
        BillStr: billFormatText(sBill)
      } );

      // Сохраняем JSON чека для добавления в Расходы
      let jsonBill = billInfo(sBill);
      NumBills = jsonBillsArr.push( {
        billJSON: jsonBill,
        billStr: sBill
      } );

      if (dDate > newLastBillDate) newLastBillDate = dDate;
    } // Фафйлы чеков

    if (newBills > 0) {
      // Собираем новые чеки вместе через пустую строку
      let newBillsStr = "";
      for (i=0; i<newBillsArr.length; i++)
        newBillsStr += newBillsArr[i].BillStr + "\n\n";

      // Записываем чеки в файл
      let fMonthName = "Чеки " + nMonth + ".txt";
      files = folderBills.getFilesByName(fMonthName);
      if (files.hasNext()) {
        fMonth = files.next();
        Logger.log("Обновляем файл " + fMonthName);
        let sMonth = fMonth.getBlob().getDataAsString();
        if (sMonth == undefined) sMonth = "";
        fMonth.setContent(newBillsStr + sMonth);
      }
      else
      {
        Logger.log("Создаем файл " + fMonthName);
        fMonth = folderBills.createFile(fMonthName, newBillsStr);
      }

      Logger.log("В папке " + nMonth + " обработано " + newBills + " чеков");
    }
  } // Вложенные папки

  if (NumBills > 0) {
    // Переносим чеки в Расходы
    for (i = 0; i < NumBills; i++)
    {
      //
      let bAllInfo = billAllInfo(jsonBillsArr[i].billStr);
      Logger.log("Покупка у " + jsonBillsArr[i].billJSON.name + " всего " + bAllInfo.items.length + " товаров на сумму " + bAllInfo.summ / 100 + " руб.");
    }

    rLastBillDate.setValue(newLastBillDate);
    Logger.log("Добавлено " + NumBills + " чеков с диска. Последняя дата чека :" + newLastBillDate);
  }
}

// Читает последние чеки из почты
function ReadMailOnTimer(ss) {
  // Читаем метку, под которой собраны чеки, из ячейки ЧекиПочта
  const sLabel = ss.getRangeByName('ЧекиПочта').getValue();
  Logger.log("Чеки в почте собраны под меткой: " + sLabel);

  const mailThreads = GmailApp
                  .getUserLabelByName( sLabel )
                  .getThreads();

  // Читаем дату последнего обработанного письма с чеком
  const rLastBillDate = ss.getRangeByName('ДатаПочтаЧек');
  let dLastBillDate = rLastBillDate.getValue();
  const sLastBillDate = dLastBillDate.toString();
  if (sLastBillDate == "") {
    //
    dLastBillDate = ss.getRangeByName('День1').getValue();
    Logger.log("Принимаем дату последнего чека в почте: " + dLastBillDate.toString());
  }
  else
    Logger.log("Дата последнего чека в почте: " + sLastBillDate);

  let newLastBillDate = dLastBillDate;
  let NumBills = 0;

  for (var i = 0; i < mailThreads.length; i++) {
    //
    var stopReading = false;
    var messages = mailThreads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var dDate = messages[j].getDate();
      var sDate = dDate.toISOString();
      if (dDate > dLastBillDate)
      {
        //
        //var sLastDate = dLastBillDate.toString();
        if (dDate > newLastBillDate)
          newLastBillDate = dDate;

        NumBills++;
      }
      else
      {
        
        stopReading = true;
        continue;
      }

      var sBody = messages[j].getBody();
      Logger.log( sDate + " e-Mail " + j + " > " + messages[j].getSubject() + " [[[ "+ sBody.length.toString() +" ]]]");

      //if (flgDbg) dbgLongMailBody(rTest.offset(k, 0), sBody);

      var bInfo = {summ: "-", date: "-", name: " ", items: []};
      //bInfo = getMailBillInfo(messages[j]);
    } // Конец письма

    if (stopReading) break;
  } // Конец ветки обсуждения

  //
  if (NumBills > 0) {
    rLastBillDate.setValue(newLastBillDate);
    Logger.log("Добавлено " + NumBills + " чеков из писем. Последняя дата чека :" + newLastBillDate);
  }

}

function getTaxcomBillInfo(sBill) {
  var sDate = between2(sBill, 'КАССОВЫЙ ЧЕК', '</tr>', 'receipt-value-1012', '</span>');
  var j = sDate.indexOf(">");
  sDate = sDate.slice(j+1);

  var sTotal = between2(sBill, 'ИТОГО:', '</tr>', 'receipt-value-1020', '</span>');
  j = sTotal.indexOf(">");
  sTotal = sTotal.slice(j+1).replace(".", ",");

  var sName = between(sBill, '<text>Данный чек подтверждает совершение расчетов в <b>', '</b>.</text>');
  
  var bItems = [];
  var iName = "";
  var iPrice = "";
  var iQuantity = "";
  var iSum = "";
  var k = 0;

  j = sBill.indexOf('<div class="items">');
  if (j != -1)
    j = sBill.indexOf('<div class="item">', j+18);

  while (j != -1) {
    //
    j += 17;
    iName = finLib.betweenFrom(sBill, j, "<span class=", "</td>", ">", "</span>");
    k = sBill.indexOf('</table>', j)+7;
    iQuantity = finLib.betweenFrom(sBill, k, "<span class=", "x", ">", "</span>");
    j = sBill.indexOf('</span>', k)+6;
    iPrice = finLib.betweenFrom(sBill, j, "<span class=", "</td>", ">", "</span>");
    k = sBill.indexOf('<td class', j)+9;
    iSum = finLib.betweenFrom(sBill, k, "<span class=", "</td>", ">", "</span>");
    j = sBill.indexOf('<div class="item">', k);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  //return {summ: sTotal, date: sDate, name: sName, items: bItems};
  return {summ: sTotal, date: sDate, name: sName, items: []};
}

function getPlatformaOFDBillInfo(sBill) {
  var sName = between2(sBill, 'check-top', '/div', '<div style=', '<');
  var j = sName.indexOf(">");
  sName = sName.slice(j+1).replace('&quot;', '"').replace('&quot;', '"');

  var sDate = between2(sBill, 'Приход', 'check-row', 'check-col-right', '</div>');
  j = sDate.indexOf(">");
  sDate = sDate.slice(j+1).trim();

  var bItems = [];
  var iName = "";
  var iPrice = "";
  var iQuantity = "";
  var iSum = "";
  var iAll = "";
  var k = 0;

  j = sBill.indexOf('check-product-name', j);
  while (j != -1) {
    //
    j += 17;
    iName = finLib.betweenFrom(sBill, j, "style=", "/div", ">", "<");
    k = sBill.indexOf('check-col-right', j)+14;
    iAll = finLib.betweenFrom(sBill, k, "style=", "/div", ">", "<");
    j = iAll.indexOf("х");
    iQuantity = iAll.slice(0,j).trim();
    iPrice = iAll.slice(j+1).trim();
    j = sBill.indexOf('check-col-right', k)+14;
    iSum = finLib.betweenFrom(sBill, j, "style=", "/div", ">", "<");
    j = sBill.indexOf('check-product-name', j);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = finLib.between2(sBill, 'check-totals', 'check-row', 'check-col-right', '</div>');
  j = sTotal.indexOf(">");
  sTotal = sTotal.slice(j+1).trim().replace(".", ",");
  
  // return {summ: sTotal, date: sDate, name: sName, items: bItems};
  return {summ: sTotal, date: sDate, name: sName, items: []};
}

function getBeelineBillInfo(sBill) {
  var sName = finLib.between(sBill, '<p style="padding:0; margin: 0; color: #282828; font-size: 13px; line-height: normal;">', '/p').trim();

  var sDate = finLib.between2(sBill, 'Дата | Время', '</tr>', '"right">', '</td>').replace("|", "").trim();

  var bItems = [];
  var iName = "";
  var iPrice = "";
  var iQuantity = "";
  var iSum = "";
  var k = 0;
  const s = '<span style="line-height: 21px; color: #000000; font-weight: bold;">';

  var j = sBill.indexOf(s);
  while (j != -1) {
    j = sBill.indexOf(s, j+67);
    iName = finLib.betweenFrom(sBill, j, "style=", "/span", ">", "<");
    k = sBill.indexOf('Цена*Кол', j)+7;
    iPrice = finLib.betweenFrom(sBill, k, "<td width=", "/td", ">", "<");
    j = sBill.indexOf('<td align=', k)+9;
    iQuantity = finLib.betweenFrom(sBill, j, "right", "/td", ">", "<");
    k = sBill.indexOf('Сумма', j)+4;
    iSum = finLib.betweenFrom(sBill, k, "<td width", "/td", ">", "<");
    j = sBill.indexOf(s, k);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = finLib.between2(sBill, 'Итог:', '</tr>', '21px;">', '</span>').replace(".", ",");

  return {summ: sTotal, date: sDate, name: sName, items: []};
  // return {summ: sTotal, date: sDate, name: sName, items: bItems};
}

function getYandexBillInfo(sBill) {
  var sName = ""; // finLib.between(sBill, '<p style="padding:0; margin: 0; color: #282828; font-size: 13px; line-height: normal;">', '/p').trim();

  var sDate = ""; // finLib.between2(sBill, 'Дата | Время', '</tr>', '"right">', '</td>').replace("|", "").trim();

  var bItems = [];
  var iName = "";
  var iPrice = "";
  var iQuantity = "";
  var iSum = "";
  var k = 0;

  var sTotal = ""; // finLib.between2(sBill, 'Итог:', '</tr>', '21px;">', '</span>').replace(".", ",");

  return {summ: sTotal, date: sDate, name: sName, items: bItems};
}

function getUnicumBillInfo(sBill) {
  var sName = finLib.between2(sBill, '<!-- Details -->', '</tbody>', '<span style=', '</span>');
  var j = sName.indexOf(">");
  sName = sName.slice(j+1).replace('&quot;', '"').replace('&quot;', '"');

  var sDate = finLib.between2(sBill, 'ДАТА ВЫДАЧИ', '</tr>', '<span style=', '</span>');
  j = sDate.indexOf(">");
  sDate = sDate.slice(j+1).trim();

  var bItems = [];
  var iName = "";
  var iPrice = "";
  var iQuantity = "";
  var iSum = "";
  var iAll = "";
  var k = 0;

  j = sBill.indexOf('<!-- Products -->', j);
  while (j != -1) {
    j += 18;
    iName = finLib.betweenFrom(sBill, j, "<span style=", "</span>", "<b>", "</b>");
    k = sBill.indexOf('<table cellspacing=', j)+19;
    iAll = finLib.betweenFrom(sBill, k, "<span style=", "/div", ">", "</span>");
    j = iAll.indexOf("X");
    iQuantity = iAll.slice(0,j).trim();
    iPrice = iAll.slice(j+1).trim().replace(".", ",");
    j = sBill.indexOf('СУММА НДС', k)+9;
    iSum = finLib.betweenFrom(sBill, j, "<span style=", "/div", ">= ", "</span>").replace(".", ",");

    j = sBill.indexOf('<!-- Products -->', j);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = finLib.between2(sBill, 'ИТОГ', '</tr>', '<span style=', '</span>');
  j = sTotal.indexOf(">");
  sTotal = sTotal.slice(j+1).trim().replace(".", ",");

  return {summ: sTotal, date: sDate, name: sName, items: bItems};
}

function getFirstOFDBillInfo(sBill) {
  var sName = finLib.between(sBill, '<tr><td align="center" colspan="5">', '<br />').trim();
  var sDate = finLib.between2(finLib.between(sBill, '<td colspan="2">', '<td colspan="3" align="right">'), '<br />', '</td>', '<br />', '<br />');
  var bItems = [];
  var iName = "";
  var iPrice = "";
  var iQuantity = "";
  var iSum = "";
  var k = 0;

  var sTotal = "";

  return {summ: sTotal, date: sDate, name: sName, items: bItems};
}

function getMailBillInfo(BillMail) {
  var bInfo = {summ: "-", date: "-", name: " ", items: []};
  var fBody = BillMail.getBody();
  //var bDate = BillMail.getDataAsString();
  if (fBody.indexOf("platformaofd") != -1) {
    bInfo = getPlatformaOFDBillInfo(fBody);

  } else if (fBody.indexOf("taxcom") != -1) {
    bInfo = getTaxcomBillInfo(fBody);

  } else if (fBody.indexOf("ofd.beeline") != -1) {
    bInfo = getBeelineBillInfo(fBody);

  } else if (fBody.indexOf("plus@support.yandex.ru") != -1) {
    bInfo = getYandexBillInfo(fBody);

  } else if (fBody.indexOf("check.ofd.ru") != -1) {
    bInfo = getUnicumBillInfo(fBody);

  } else if (fBody.indexOf("1-ofd") != -1) {
    bInfo = getFirstOFDBillInfo(fBody);

  } else Logger.log("Что-то новое!");
  return bInfo;
}

// Пункт меню Сканировать - Почту
function MenuScanBillsFromMail() {

  // Таблица с которой работаем
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Флаг отладки
  const flgDbg = dbgGetDbgFlag(true);

  if (flgDbg) {
    // Лист для отладки
    var rTest = ss.getSheetByName("Test").getRange(1, 1);
  }

  var k = 0;
  var l = 1;

  var rLastDate = ss.getRangeByName("ДатаПочтаЧек");
  var dLastDate = rLastDate.getValue();

  var threads = GmailApp
                .getUserLabelByName("Моё/Мани/Чеки")
                .getThreads();

  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var dDate = messages[j].getDate();
      var sDate = dDate.toString();
      if (dDate > dLastDate) {
        //
        var sLastDate = dLastDate.toString();
      }
      var sBody = messages[j].getBody();
      Logger.log( j + " > " + messages[j].getSubject() + " [[[ "+ sBody.length.toString() +" ]]]");

      if (flgDbg) finLib.dbgLongMailBody(rTest.offset(k, 0), sBody);

      var bInfo = {summ: "-", date: "-", name: " ", items: []};
      bInfo = getMailBillInfo(messages[j]);

      if (flgDbg)
      {
        var c = 3;
        rTest.offset(k, c++).setValue(bInfo.date);
        rTest.offset(k, c++).setValue(bInfo.summ);
        rTest.offset(k, c++).setValue(bInfo.name);

        if (bInfo.items.length>0) {
          rTest.offset(k++, c).setValue(bInfo.items.length);

          // {iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum}
          bInfo.items.forEach(function(element) {
            rTest.offset(k, c++).setValue(element.isum);
            rTest.offset(k, c++).setValue(element.iquantity);
            rTest.offset(k, c++).setValue(element.iprice);
            rTest.offset(k++, c++).setValue(element.iname);
          });
        }
      } 

      k++;
      Logger.log("Чек >>> " + (l++).toString() + " <<<");
    } // Сообщения с чеками
  } // Цепочки сообщений с чеками
}

// Пункт меню Сканировать - Расходы
// Читает колонку Заметка со вкладки Расходы и пишет в лист отладки
function MenuScanBillsFromCosts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costs = ss.getSheetByName("Расходы");
  
  var flgDbg = dbgGetDbgFlag(true);
  
  // Лист для отладки
  var sTest = ss.getSheetByName("Test");
  var rTest = sTest.getRange(1, 1);

  var k = 1;

  costs.expandAllRowGroups();
  var costsData = costs.getDataRange();

  var cdRows = costsData.getNumRows();
  var cdColumns = costsData.getNumColumns();
  if (flgDbg) 
  {
    sTest.getRange(k, 1, 1, 1).setValue(cdRows);
    sTest.getRange(k, 2).setValue(cdColumns);
    sTest.getRange(k, 3).setValue(costsData.getValue());
    Logger.log(costsData.getCell(cdRows, cdColumns).getValue());
    sTest.getRange(k++, 4).setValue(costsData.getCell(cdRows, cdColumns).getValue());
  }

  for (var i = 2; i < cdRows; i++) {
    var n = 1;
    var cData = costsData.getCell(i, 1);
    var iData = cData.getValue();

    var cTime = costsData.getCell(i, 2);
    var iTime = cTime.getValue();

    var cSumm = costsData.getCell(i, 3);
    var iSumm = cSumm.getValue();

    var cJson = costsData.getCell(i, 8);
    var iJson = cJson.getValue();

    if (iData != "" && iSumm != "") {
      var sTime = iData.toString();
      var ssTime = cData.getDisplayValue();
      //var ssTime = costsData.getCell(i, 1).getDisplayValue();
      if (flgDbg)
      {
        sTest.getRange(k, n++).setValue(sTime);
        sTest.getRange(k, n++).setValue(ssTime);
      }
      var sSumm = iSumm.toString();
      var ssSumm = cSumm.getDisplayValue();

      var dFormat = cData.getNumberFormat();
      var tFormat = cTime.getNumberFormat();
      var sFormat = cSumm.getNumberFormat();
      // "dd.mm", "HH:mm", "#,##0.00[$ ₽]"

      if (flgDbg)
      {
        sTest.getRange(k, n++).setValue(sSumm);
        sTest.getRange(k, n++).setValue(dFormat);
        sTest.getRange(k, n++).setValue(tFormat);
        sTest.getRange(k, n++).setValue(sFormat);
        sTest.getRange(k, n++).setValue(ssSumm);
      }

      var aBill = undefined;
      var sBill = "";
      if (iJson != "") {
        if (flgDbg) sTest.getRange(k, 12).setValue(iJson);
        var ii = iJson.indexOf("\"receipt\":");
        if (ii > -1) {
          var jj = iJson.indexOf("],", ii);
          var kk = iJson.indexOf("}", jj) + 1;
          //
          if (iJson.indexOf("{", jj) != -1) {
            //
            jj = iJson.indexOf("}", iJson.indexOf("{", jj)) + 1;
            kk = iJson.indexOf("}", jj) + 1;
          }
          sBill = iJson.slice(ii + 10, kk);
          Logger.log("JSON >>>>" + sBill + "<<<<");
          if (flgDbg) sTest.getRange(k, 11).setValue(sBill);
          aBill = JSON.parse(sBill);
        }
      }

      if (aBill == undefined)
        aBill = {totalSum: "-1", dateTime: "-", user: "", items: {}};

      // name: sName, summ: sSumm, date: sDate, items
      if (flgDbg) {
        sTest.getRange(k, n++).setValue(aBill.user);
        sTest.getRange(k, n++).setValue(aBill.totalSum);
        sTest.getRange(k++, n++).setValue(aBill.dateTime);
      }
    }
  }
}

function getUBERBillInfo(BillMail) {
  let fSubject = BillMail.getSubject();
  let sTripDate = fSubject.slice(23);

  let spcPos = sTripDate.indexOf(" ");
  let sTripDay = sTripDate.slice(0, spcPos);
  let spcPos2 = sTripDate.indexOf(" ", spcPos+2);
  let sTripMonth = sTripDate.slice(spcPos+1, spcPos2);

  let TripMonth = getMonthNum(sTripMonth);
  let TripYear = sTripDate.slice(spcPos2+1, sTripDate.indexOf(" г.", spcPos2+2));

  var TripDate = sTripDay + "." + TripMonth + "." + TripYear;

  let fBody = BillMail.getBody();
  // finLib.between2();

  var TripTime = finLib.between2(fBody, "From", "</tr>", "<td align", "</td>");
  var j = TripTime.indexOf(">");
  TripTime = TripTime.slice(j+1).trim();

  var TripDateTime = TripDate + " " + TripTime;

  var TripSumm = finLib.between2(fBody, "check__price", "</td>", ">", " ₽").trim();

  var bInfo = {summ: TripSumm, date: TripDateTime, name: '"ООО \"ЯНДЕКС.ТАКСИ\""', items: [{iname:"Перевозка пассажиров и багажа", iprice:TripSumm, isum:TripSumm, iquantity:1.0}]};

  Logger.log("UBER > ", bInfo);
  return bInfo;
}

// Пункт меню Сканировать - Чеки UBER
function MenuCheckUBER() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var flgDbg = dbgGetDbgFlag(true);
  
  // Лист для отладки
  var sTest = ss.getSheetByName("Test");

  var k = 1;

  var label = GmailApp.getUserLabelByName("Моё/Мани/Такси");
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      Logger.log( j + " > " + subject);

      // var body = message.getBody();
      // if (flgDbg) sTest.getRange(k, 1).setValue(body);
      
      var bInfo = getUBERBillInfo(message);

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

  var TripTime = finLib.between2(fBody, "route__point-name", "</td>", "<p class=", "</p>");
  var j = TripTime.indexOf(">");
  TripTime = TripTime.slice(j+1).trim();

  var TripDateTime = TripDate + " " + TripTime;

  var TripSumm = finLib.between2(fBody, "report__value_main", "</td>", ">", " ₽").trim();

  var bInfo = {summ: TripSumm, date: TripDateTime, name: '"ООО \"ЯНДЕКС.ТАКСИ\""', items: [{iname:"Перевозка пассажиров и багажа", iprice:TripSumm, isum:TripSumm, iquantity:1.0}]};

  Logger.log("Yandex Go> ", bInfo);
  return bInfo;
}

function MenuCheckYandexGo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const flgDbg = dbgGetDbgFlag(true);
  
  // Лист для отладки
  const sTest = ss.getSheetByName("Test");

  let k = 1;

  var label = GmailApp.getUserLabelByName("pers/отчеты/такси");
  var threads = label.getThreads();
  if (flgDbg) SpreadsheetApp.getActive().toast(threads.length);

  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      Logger.log( j + " > " + subject);

      var body = message.getBody();
      if (flgDbg) sTest.getRange(k, 1).setValue(body);
      
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
        if (flgDbg) finLib.dbgLongMailBody(rTest.offset(k, 1), Body);
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"Перевозка пассажиров и багажа",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      } // Тема сообщения "Ваш номер заказа ..."
      else
      {
        var Body = message.getBody();
        Logger.log( j + " > " + subject + " <<< "+ Body.length.toString() +" >>>");

        if (flgDbg) rTest.offset(k, 0).setValue(" # " + subject + " <<< "+ Body.length.toString() +" >>>"); 
        if (flgDbg) finLib.dbgLongMailBody(rTest.offset(k, 1), Body);
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"Перевозка пассажиров и багажа",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      }
    } // Сообщения с чеками AliExpress
  } // Цепочки сообщений с чеками AliExpress
}

function SetTargetList(ss, c, l)
{
  const range = ss.getRangeByName(l);

  if (range != undefined) {
    const rule = range.getDataValidation();

    c.setDataValidation(rule);
  }
}

function SetTargetRule(ss, c, rn)
{
  const range = ss.getRangeByName(rn);

  if (range != undefined) {
    const rule = range.getDataValidation();

    c.setDataValidation(rule);
  }
}

// Устанавливаем доступные счета и Тип операции для выбранной из списка операции
function SettingTrnctnName(ss, br)
{
  const accrual = 'Начисление';
  const debit = 'Списание';
  const turnover = 'Оборот';

  const NewVal = br.getValue();
  const OpAcc = br.offset(0,-2);
  const OpTrgt = br.offset(0,-1);

  var i = findInRule(turnover, NewVal);
  if (i != -1)
  {
    // Выбрана оборотная операция
    br.offset(0,1).setValue(turnover);

    SetTargetRule(ss, OpAcc, 'СчетаДеб');

    const Transfer = ss.getRangeByName('стрПеревод').getValue();
    if (NewVal == Transfer)
    {
      // Перевод
      SetTargetRule(ss, OpTrgt, 'СчетаДеб');
    } else {
      OpTrgt.clearDataValidations();
      if (i == 0) {
        // Снятие
        OpTrgt.clear();
      }
    }
  }
  else if (findInRule(debit, NewVal) != -1)
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
    i = findInRule(accrual, NewVal);
    if (i != -1) {
      // Выбрана операция начисления
      br.offset(0,1).setValue(accrual);

      SetTargetRule(ss, OpAcc, 'СчетаДеб');
      if (i < 4) OpAcc.setValue("ЗП");
    }
  }
}

// Устанавливаем соответствующий список операций для выбранного Типа операции
function SettingTrnctnType(ss, br)
{
  const NewVal = br.getValue();

  if (ss.getRangeByName(NewVal) == undefined) {
    // Устанавливаем полный список операций для выбора если Тип неизвестен
    NewVal = 'Операция';
  }

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

  const bill = jsonBillInfo(NewVal);

  if (flgDbg) {
    if (bill != undefined)
      rTest.offset(3, 1).setValue(bill.name)
      .offset(1, 0).setValue(bill.summ)
      .offset(1, 0).setValue(bill.date)
      .offset(1, 0).setValue(bill.cash);
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

  if (flgDbg) rTest.offset(7, 1).setValue(A1date);

  // Выставляем время покупки
  br.offset(0,-6)
  .setValue("=" + A1date)
  .setNumberFormat("HH:mm");

  // Если наличные, то выставляем счет списания
  if (bill.cash != 0)
    br.offset(0,-4).setValue("Карман")

  // Выставляем Статью, Инфо и Примечание для магазина
  const storeList = ss.getRangeByName('СпскМагазины');

}

function TestSetBill () {
  //
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let br = ss.getSheetByName('Расходы').getRange(43, 8);
  var jj = JSON.parse(br.getValue());
  //Logger.log( " > " + jj.receipt.toString() + " <<< ");
  Logger.log( " >>> " + JSON.stringify(jj) + " < ");

  SettingCostBill(ss, br);
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
      // var v = e.value;
      // br.setNote(v);

    switch(ncol) {
    case 5:
      // Изменилась статья расходов
      SettingCostInfo(ss, br);
      break;
    case 6:
      // Изменился пункт статьи расходов
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

  costs.expandAllRowGroups();
  var costsData = costs.getDataRange();

  var cdRows = costsData.getNumRows();
  var cdColumns = costsData.getNumColumns();
  if (flgDbg) 
  {
    rTest.offset(k, 1).setValue(cdRows)
    .offset(0, 1).setValue(cdColumns)
    .offset(0, 1).setValue(costsData.getValue());
    Logger.log(costsData.getCell(cdRows, cdColumns).getValue());

    /*
    var st = [{
    _id: "659ed0afba091652867328c4",
    createdAt: "2024-01-10T17:15:27+00:00",
    ticket: {
      document: {
        receipt: {
          buyerPhoneOrAddress: "prodg@ya.ru",
          cashTotalSum: 0,
          code: 3,
          creditSum: 0,
          dateTime: "2024-01-10T20:14:00",
          ecashTotalSum: 126599,
          fiscalDocumentFormatVer: 4,
          fiscalDocumentNumber: 75447,
          fiscalDriveNumber: "7281440501036726",
          fiscalSign: 2290682911,
          fnsUrl: "www.nalog.gov.ru",
          items: [
            { name: "Пакет-майка Магнолия", nds: 1, ndsSum: 150, paymentType: 4, price: 899, productType: 1, quantity: 1, sum: 899 },
            { name: "Хлебцы скандинавские цельнозерн.ржаные 180г Бейкер Хаус", nds: 2, ndsSum: 1809, paymentType: 4, price: 19900, productType: 1, quantity: 1, sum: 19900 },
            { name: "Сухарики рж Три Корочки с сыром и семгой 40г", nds: 2, ndsSum: 679, paymentType: 4, price: 2490, productType: 1, quantity: 3, sum: 7470 }
          ],
          kktRegId: "0006533680025786    ",
          nds10: 9156,
          nds18: 4311,
          operationType: 1,
          operator: "Пулотов",
          prepaidSum: 0,
          provisionSum: 0,
          requestNumber: 586,
          retailPlace: "Магазин <Магнолия>",
          retailPlaceAddress: "115477, г. Москва, Пролетарский пр-т, дом № 31.",
          sellerAddress: "noreply@platformaofd.ru",
          shiftNumber: 162,
          taxationType: 1,
          appliedTaxationType: 1,
          totalSum: 126599,
          user: "ЗАО \"Т и К Продукты\"",
          userInn: "7731162754  "
        } } } } ];
    Logger.log(st);
    rTest.offset(k++, 1).setValue(st.toString());
    rTest.offset(k++, 1).setValue(st);
    */
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

  const menuScan = [
    {name: "Расходы", functionName: "MenuScanBillsFromCosts"},
    {name: "Чеки UBER", functionName: "MenuCheckUBER"},
    {name: "Чеки Яндекс Go", functionName: "MenuCheckYandexGo"},
    {name: "Чеки AliExpress", functionName: "MenuCheckAliExpress"},
    null,
    {name: "Очистить отладку", functionName: "finLib.ClearTestSheet"}
  ];
  e.source.addMenu("Сканировать", menuScan);

  const menuFinance = [
    {name: "Закрыть день", functionName: "MenuCloseDay"}

  ];
  e.source.addMenu("Финансы", menuFinance);

}

function onOnceAnHour()
{
  // Выполняется ежечасно
  Logger.log("Обрабатываем последние чеки");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dDay1Date = ss.getRangeByName('День1').getValue();

  Logger.log("Сканируем чеки на диске");
  try {
    //Block of code to try;
    ReadDriveOnTimer(ss.getRangeByName('ДатаДискЧек'), dDay1Date);
  }
  catch(err) {
    Logger.log(err);
  }
  finally {
    //    Block of code to be executed regardless of the try / catch result;
  }

  Logger.log("Сканируем чеки в почте");
  //try {
    //Block of code to try;
    ReadMailOnTimer(ss);
  //}
  //catch(err) {
  //  Logger.log(err);
  //}
  //finally {
    //    Block of code to be executed regardless of the try / catch result;
  //}

  // Надо вынести в другой скрипт и выполнять реже
  Logger.log("Сканируем покупки Ali");
}
