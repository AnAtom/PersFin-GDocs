// 'История' 'Активные' " до d mmmm" .setNumberFormat("HH:mm");

function onEdit(e)
{
  // SpreadSheet
  const ss = e.source;

  const br = e.range;
  if (br.getNumColumns() > 1) // Скопировали диапазон
    return;

  const ncol = br.getColumn();
  const sname = ss.getActiveSheet().getSheetName();
  let cname = ss.getActiveSheet().getRange(1, ncol).getValue();
  if (cname == undefined || cname == '')
    cname = ncol;
  Logger.log("Редактируем на листе [" + sname + "] в колонке (" + cname + ") строку :" + br.getRow());
}

function newOrderMail()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sDBG = ss.getSheetByName('dbg');

  const rLastDate = ss.getRangeByName('ПочтаДата');
  let dLastDate = rLastDate.getValue();
  let newLastDate = dLastDate;
  var k = 0;

  const mailThreads = GmailApp
    .getUserLabelByName(ss
        .getRangeByName('ПочтаМетка')
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

      var subject = message.getSubject();
      var Body = message.getBody();
      Logger.log(" > " + subject + " [[[ "+ Body.length +" ]]] " + mDate);

    } // Сообщения с чеками AliExpress
  } // Цепочки сообщений с чеками AliExpress
  //if (newLastDate > dLastDate)
  //  rLastDate.setValue(newLastDate);
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
      .addItem('Новый по URL', 'newOrderURL')
      .addItem('Новый из почты', 'newOrderMail')
      .addItem('Новый пустой', 'newOrder')
        .addSeparator()
      .addSubMenu(ui
        .createMenu('Хвост')
          .addItem('Голова', 'EmptyProc'))
  .addToUi();
}
