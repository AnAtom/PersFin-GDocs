/*
  getMarketByNumber(orderNum) определить площадку по номеру заказа
  getMarketStr(mrktId)        получить название площадки
  getOrderURL(orderNum)       получить ссылку на карточку заказа
  getReadableNumber(orderNum) получить читаемый номер заказа
  addNewOrder(orderNum)
*/

const mrktAli = 0;
const mrktOzon = 1;
const mrktYandex = 2;
//const mrktWB = 3;

function Test() {
  //
  const z1 = '5550677229097566';
  const z1_1 = '5550 6772 2909 7566';
  const z2 = '18044876-0173';
  const z3 = '60328348611';

  const s1 = getMarketStr(getMarketByNumber(z1));
  const s1_1 = getMarketStr(getMarketByNumber(z1_1));
  const s2 = getMarketStr(getMarketByNumber(z2));
  const s3 = getMarketStr(getMarketByNumber(z3));

  const u1 = getOrderURL(z1);
  const u1_1 = getOrderURL(z1_1);
  const u2 = getOrderURL(z2);
  const u3 = getOrderURL(z3);

  const r1 = getReadableNumber(z1);
  const r1_1 = getReadableNumber(z1_1);
  const r2 = getReadableNumber(z2);
  const r3 = getReadableNumber(z3);

  Logger.log('Заказ '+z1+' Магазин '+s1+' URL '+u1+' запишем <'+r1+'>');
  Logger.log('Заказ '+z1_1+' Магазин '+s1_1+' URL '+u1_1+' запишем <'+r1_1+'>');
  Logger.log('Заказ '+z2+' Магазин '+s2+' URL '+u2+' запишем <'+r2+'>');
  Logger.log('Заказ '+z3+' Магазин '+s3+' URL '+u3+' запишем <'+r3+'>');
}

function getMarketByNumber(orderNum) {
  // 5551 5791 3251 7566 / 5550677229097566
  // 18044876-0173
  // 60328348611
  if (orderNum.length == 11) { return mrktYandex }
  else if (~orderNum.indexOf("-")) { return mrktOzon }
  else { return mrktAli }
}

function getMarketStr(mrktId) {
  // Маркеты
  // OZON
  // Ali
  // Яндекс
  // WB
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vMarkets = ss.getRangeByName('Маркеты').getValues();
  return vMarkets[mrktId][0];
}

function getOrderURL(orderNum) {
  // Делаем ссылку на заказ по номеру
  // https://market.yandex.ru/my/order/57989200259
  // https://www.ozon.ru/my/orderdetails/?order=18044876-0131
  // https://aliexpress.ru/order-list/5550677229097566

  const yandexURL = "market.yandex.ru/my/order/"
  const ozonURL = "www.ozon.ru/my/orderdetails/?order="
  const aliURL = "aliexpress.ru/order-list/"
  var orderURL = "https://";

  switch(getMarketByNumber(orderNum)) {
    case mrktOzon: 
      orderURL += (ozonURL + orderNum);
      break;
    case mrktYandex: 
      orderURL += (yandexURL + orderNum);
      break;
    default:
      var ordrNum = orderNum;
      if (~ordrNum.indexOf(' ')) { ordrNum = ordrNum.replace(/\s/g, ''); }
      orderURL += (aliURL + ordrNum);
  }

  return orderURL;
}

function getReadableNumber(orderNum) {
  // 0    4    8    12   16
  // 5551 5791 3251 7566
  // 1804 4876-0173
  // 6032 8348 611
  var readableNum = "'";
  switch(getMarketByNumber(orderNum)) {
    case mrktOzon: 
      readableNum += (orderNum.slice(0, 4) + ' ' + orderNum.slice(4, 13));
      break;
    case mrktYandex: 
      readableNum += (orderNum.slice(0, 4) + ' '
        + orderNum.slice(4, 8) + ' '
        + orderNum.slice(8, 11));
      break;
    default:
      if (~orderNum.indexOf(' ')) { readableNum += orderNum }
      else {
        readableNum += (orderNum.slice(0, 4) + ' '
          + orderNum.slice(4, 8) + ' '
          + orderNum.slice(8, 12) + ' '
          + orderNum.slice(12, 16));
      }
  }

  return readableNum;
}

function addNewOrder(orderNum) {
  //
  Logger.log("Добавляем новый заказ с номером <" + orderNum + ">");
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
}
