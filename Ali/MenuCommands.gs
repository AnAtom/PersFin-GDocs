/*

newOrderMail - читать из почты
newOrderURL - читать из ссылки
newOrder - добавить пустой заказ

*/

// 'История' 'Активные' " до d mmmm" .setNumberFormat("HH:mm");

function onEdit(e)
{
  const br = e.range;
  if (br.getNumColumns() > 1 && e.value === '') // Скопировали диапазон или очистили ячейку
    return;

  const ncol = br.getColumn();

  // SpreadSheet
  const ss = e.source;
  let cname = ss.getActiveSheet().getRange(1, ncol).getValue();
  if (cname == undefined || cname == '')
    cname = ncol;
  const nrow = br.getRow();
  const sname = ss.getActiveSheet().getSheetName();
  Logger.log("Редактируем на листе [" + sname + "] в колонке (" + cname + ") строку :" + nrow);
  Logger.log("Format [" + br.getNumberFormat() + "] value (" + e.value + ")");

  if (nrow == 3 && ncol == 2 && sname == 'Активные' && e.value.indexOf(' ') == -1) {
    // Редактируем номер заказа
    const orderNum = "'" + e.value.slice(0, 4) + ' ' + e.value.slice(4, 8) + ' ' + e.value.slice(8, 12) + ' ' + e.value.slice(12, 16);
    const orderURL = "https://aliexpress.ru/order-list/" + e.value;
    Logger.log("Номер заказа [" + orderNum + "] ссылка на заказ (" + orderURL + ")");
    br.setValue(orderNum);
    const valueURL = SpreadsheetApp.newRichTextValue()
    .setText("Товар")
    .setLinkUrl(orderURL)
    .build();
    ss.getSheetByName('История').getRange(3, 9).setRichTextValue(valueURL);
  }
}

// Разбиваем длинную строку ( >50000 ) на несколько строк по maxLngth символов
//          var rTest = sDBG.getRange(1, 1);
//          let s = dbgSplitLongString(Body, 4950);
//          rTest.offset(k, 1, 1, s.length). setValues([s]);
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

// https://mail.google.com/mail/u/0/#search/5454585585567566/FMfcgzQXJkXtQDQWbQcgGHhzgfQPLlzk
// https://mail.google.com/mail/u/0/#search/5454585585567566/FMfcgzQXJkXtQDPtrSrFQBBLBqjzSGKn
// https://mail.google.com/mail/u/0/popout?ver=mca2oziyi6fc&q=5454585585567566&qs=true&qid=F21EAB37-6ADA-46E6-816D-70055B3BA5B7&cmembership=1&search=query&silk=91494CEC-B592-473F-B52A-682D5C3CBE10&th=%23thread-f%3A1813691637966661266&qt=5454.1.5454585585567566.1.5855.1.7566.1.8556.1&cvid=1

function newOrderMail()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sDBG = ss.getSheetByName('dbg');
  sDBG.getRange(2, 1, sDBG.getLastRow(), sDBG.getLastColumn()).clear();
  var rTest = sDBG.getRange(1, 1);
  const showBody = false;
  const preDBG = false;

  var arrOrders = [];
  var k = 0;
  var l = 1;

  const rLastDate = ss.getRangeByName('ДатаОплачен'); // ДатаОформлен
  let dLastDate = rLastDate.getValue();
  let newLastDate = dLastDate;

  const mailThreads = GmailApp
    .getUserLabelByName(ss
        .getRangeByName('ПочтаОплачен') // ПочтаОформлен
        .getValue())
    .getThreads();
  for (messages of mailThreads) {
    for (message of messages.getMessages()) {
      const mDate = message.getDate();
      if (mDate > dLastDate) {
        if (mDate > newLastDate)
          newLastDate = mDate;
      } else
        continue;

      const subject = message.getSubject();
      const Body = message.getBody();
      Logger.log("_Оплачен_> " + subject + " [[[ "+ Body.length +" ]]] " + mDate);

      let s = between2(Body, 'Номер заказа', 'RRN', '<td align=', '</td>');
      var orderNum = s.slice(s.indexOf('>')+1).trim();

      s = between2(Body, 'Сумма', 'Комиссионное', '<td align=', '₽');
      var orderSum = s.slice(s.indexOf('>')+1).trim().replace(",", ".") * 1.0;

      s = between2(Body, 'Дата и время', 'Номер транзакции', '<td align=', '</td>');
      var orderDate = s.slice(s.indexOf('>')+1).trim();
      // 04/19/2024 08:29:17
      var isoDate = orderDate.slice(6, 10)  // Год
            .concat(
              "-", orderDate.slice(0, 2),   // месяц
              "-", orderDate.slice(3, 5),   // день
              "T", orderDate.slice(11))     // время

      s = between2(Body, 'Номер карты', '</table>', '<td align=', '</td>');
      var orderCard = s.slice(s.indexOf('>')+1).trim();

      Logger.log("Номер <" + orderNum + "> дата |" + orderDate + "| ISO " + isoDate + " сумма (" + orderSum + ') оплата [' + orderCard + ']');

      if (preDBG) { // Дата	Номер	Сумма	Карта	Сабж
        rTest.offset(l, 0, 1, 6).setValues([[isoDate, orderNum, orderSum, orderCard, subject, Body.length]]);
        if (showBody) {
          let m = dbgSplitLongString(Body, 4950);
          rTest.offset(l, 6, 1, m.length).setValues([m]);
          rTest.offset(l++, 6+m.length).setValue('### Оплачен ###');
        } else rTest.offset(l++, 4).setValue(subject);
      }

      // num: , paySum: , payDate: , payCard: , payURL: , placedSum: , placedDate: , placedURL: , closedSum: , closedDate: , closedURL:
      var Order = {num: orderNum,
            paySum: orderSum, payDate: new Date(isoDate), payCard: orderCard, payURL: messages.getPermalink(), 
            placedSum: 0, placedDate: '', placedURL: '', 
            closedSum: 0, closedDate: '', closedURL: ''};
      arrOrders.push(Order)
      k++;
    } // Сообщения с чеками AliExpress Оплачен
  } // Цепочки сообщений с чеками AliExpress Оплачен

  Logger.log(k + " ============================================================================");

  const rLastDate2 = ss.getRangeByName('ДатаОформлен');
  let dLastDate2 = rLastDate2.getValue();
  let newLastDate2 = dLastDate2;
  k = 0;

  const mailThreads2 = GmailApp
    .getUserLabelByName(ss
        .getRangeByName('ПочтаОформлен')
        .getValue())
    .getThreads();

  for (messages of mailThreads2) {
    for (message of messages.getMessages()) {
      const mDate = message.getDate();
      if (mDate > dLastDate2) {
        if (mDate > newLastDate2)
          newLastDate2 = mDate;
      } else
        continue;

      const subject = message.getSubject();
      const Body = message.getBody();
      Logger.log("_Оформлен_> " + subject + " [[[ "+ Body.length +" ]]] " + mDate);

      var orderNum = between2(Body, 'Ваш заказ', 'подтверждён', '>', '</a>'); // s.slice(s.indexOf('>')+1).trim();

      var orderDate = between2(Body, 'Оформлен', '</td>', '<b>', '</b>'); // s.slice(s.indexOf('>')+1).trim();
      // 23-10-2024, 08:11 UTC
      var isoDate = orderDate.slice(6, 10)         // Год
            .concat(
              "-", orderDate.slice(3, 5),          // месяц
              "-", orderDate.slice(0, 2),          // день
              "T", orderDate.slice(12, 17), 'Z');  // время

      let s = between2(Body, 'Сумма заказа', '</tr>', '<td align=', '₽');
      var orderSum = s.slice(s.indexOf('>')+1).trim().replace(",", ".") * 1.0;

      // Картинки https://ae01.alicdn.com/kf/A3e3967cb95cf4f799d005b6f3ce7a56bh.jpg_220x220.jpg

      Logger.log("Номер <" + orderNum + "> дата |" + orderDate + "| ISO " + isoDate + " сумма (" + orderSum + ')');

      if (preDBG) { // Дата	Номер	Сумма	Карта	Сабж
        rTest.offset(l, 0, 1, 6).setValues([[isoDate, orderNum, orderSum, '', subject, Body.length]]);
        if (showBody) {
          let m = dbgSplitLongString(Body, 4950);
          rTest.offset(l, 6, 1, m.length).setValues([m]);
          rTest.offset(l++, 6+m.length).setValue('### Оформлен ###');
        } else rTest.offset(l++, 4).setValue(subject);
      }

      var Order = arrOrders.find((element) => element.num == orderNum);
      if (Order == undefined) {
        Order = {num: orderNum,
          paySum: 0, payDate: '', payCard: '', payURL: '',
          placedSum: orderSum, placedDate: new Date(isoDate), placedURL: messages.getPermalink(),
          closedSum: 0, closedDate: '', closedURL: ''};
      } else {
        Order.placedSum = orderSum;
        Order.placedDate = new Date(isoDate);
        Order.placedURL = messages.getPermalink();
      }
      arrOrders.push(Order);
      k++;
    } // Сообщения с чеками AliExpress Оформлен
  } // Цепочки сообщений с чеками AliExpress Оформлен

  Logger.log(k + " ============================================================================");

  const rLastDate3 = ss.getRangeByName('ДатаЗавершен');
  let dLastDate3 = rLastDate3.getValue();
  let newLastDate3 = dLastDate3;
  k = 0;

  const mailThreads3 = GmailApp
    .getUserLabelByName(ss
        .getRangeByName('ПочтаЗавершен')
        .getValue())
    .getThreads();

  for (messages of mailThreads3) {
    for (message of messages.getMessages()) {
      const mDate = message.getDate();
      if (mDate > dLastDate3) {
        if (mDate > newLastDate3)
          newLastDate3 = mDate;
      } else
        continue;

      const subject = message.getSubject();
      const Body = message.getBody();
      Logger.log(l + "_Завершен_> " + subject + " [[[ "+ Body.length +" ]]] " + mDate);

      let s = between2(Body, 'Здравствуйте', '</table>', 'Заказ', '</a>');
      var orderNum = s.slice(s.indexOf('>')+1).trim();

      var orderDate = between2(Body, 'Оформлен', '</td>', '<b>', '</b>'); // s.slice(s.indexOf('>')+1).trim();
      var isoDate = orderDate.slice(6, 10)         // Год
            .concat(
              "-", orderDate.slice(3, 5),          // месяц
              "-", orderDate.slice(0, 2),          // день
              "T", orderDate.slice(12, 17), 'Z');  // время

      s = between2(Body, 'Сумма заказа', '</tr>', '<td align=', '₽');
      var orderSum = s.slice(s.indexOf('>')+1).trim().replace(",", ".") * 1.0;

      Logger.log("Номер <" + orderNum + "> дата |" + orderDate + "| ISO " + isoDate + " сумма (" + orderSum + ')');

      if (preDBG) { // Дата	Номер	Сумма	Карта	Сабж
        rTest.offset(l, 0, 1, 6).setValues([[isoDate, orderNum, orderSum, '', subject, Body.length]]);
        if (showBody) {
          let m = dbgSplitLongString(Body, 4950);
          rTest.offset(l, 6, 1, m.length).setValues([m]);
          rTest.offset(l++, 6+m.length).setValue('### Завершен ###');
        } else rTest.offset(l++, 4).setValue(subject);
      }
      var Order = arrOrders.find((element) => element.num == orderNum);
      if (Order == undefined) {
        Order = {num: orderNum,
          paySum: 0, payDate: '', payCard: '', payURL: '',
          placedSum: 0, placedDate: '', placedURL: '',
          closedSum: orderSum, closedDate: new Date(isoDate), closedURL: messages.getPermalink()};
      } else {
        Order.closedSum = orderSum;
        Order.closedDate = new Date(isoDate);
        Order.closedURL = messages.getPermalink();
      }
      arrOrders.push(Order);
      k++;
    } // Сообщения с чеками AliExpress Завершен
  } // Цепочки сообщений с чеками AliExpress Завершен

  Logger.log(k + " ============================================================================");

  for (Order of arrOrders) {
    // num: , paySum: , payDate: , payCard: , payURL: , placedSum: , placedDate: , placedURL: , closedSum: , closedDate: , closedURL:
    Logger.log('Заказ ' + Order.num + ' (' + Order.paySum + ') |' + Order.payDate + '| <' + Order.payURL + '> [' + Order.payCard + ']');
    Logger.log('      ................ (' + Order.placedSum + ') |' + Order.placedDate + '| <' + Order.placedURL + '>');
    Logger.log('      ________________ (' + Order.closedSum + ') |' + Order.closedDate + '| <' + Order.closedURL + '>');
  }
  //if (newLastDate > dLastDate)
  //  rLastDate.setValue(newLastDate);
  //if (newLastDate2 > dLastDate2)
  //  rLastDate2.setValue(newLastDate2);
  //if (newLastDate3 > dLastDate3)
  //  rLastDate3.setValue(newLastDate3);
}

function newOrderURL()
{
  const orderURL = Browser.inputBox('Введите URL покупки:');
  // https://aliexpress.ru/item/1005006072515663.html?spm=a2g2w.orderdetail.0.0.31354aa6t3c5Cj&sku_id=12000035602243964
  // https://aliexpress.ru/order-list/5353801443557566?spm=a2g2w.orderlist.0.0.6cf94aa6hZMUBo&filterName=active
  Logger.log("Добавляем новый заказ по URL " + orderURL);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sOrderHistory = ss.getSheetByName('История');
  const sActiveOrders = ss.getSheetByName('Активные');

}

function newOrder()
{
  Logger.log("Добавляем новый пустой заказ");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sOrderHistory = ss.getSheetByName('История');
  sOrderHistory
    .insertRowBefore(3)
    .getRange(3, 2)
    .setNumberFormat("dd.MM.yyyy")
    .setValue(new Date());

  const sActiveOrders = ss.getSheetByName('Активные');
  sActiveOrders
    .insertRowsAfter(2, 7)
    .getRange(3, 1, 7)
    .mergeVertically()
    .setFormula('=IMAGE("")');
  sActiveOrders
    .getRange(8, 2, 1, 2)
    .setBorder(null, null, true, null, null, null)
    .setFormulas([["='История'!C3", ""]]);
  sActiveOrders
    .getRange(5, 2)
    .setNumberFormat(" до d mmmm")
    .setFormula("='История'!D3");
  sActiveOrders
    .getRange(3, 2)
    .setFontWeight('bold');
}

function EmptyProc()
{
  Logger.log("Пустая трата времени");
}

function onOpen(e)
{
  Logger.log('Добавляем пункты меню.');
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Заказы')
      .addItem('Новый пустой', 'newOrder')
      .addItem('Новый из почты', 'newOrderMail')
      .addItem('Новый по URL', 'newOrderURL')
        .addSeparator()
      .addSubMenu(ui
        .createMenu('Хвост')
          .addItem('Голова', 'EmptyProc'))
  .addToUi();
}

function onOnceAnHour()
{
  //
}
