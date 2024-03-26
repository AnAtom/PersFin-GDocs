/*

dbgGetDbgFlag(clearTest) - Возвращает значение флага ФлОтладка. Если true и аргумент true, то очищает лист Test.
dbgClearTestSheet() - Очищает и активирует лист Test.
dbgSplitLongString(sStr, maxLngth) - Разбивает длинную строку на набор строк длиной maxLngth.
dbgBillInfo(bBill) - Формирует строку с информацией о чеке для логирования.

*/

// 
function dbgGetDbgFlag(clearTest)
{
  const Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const Range = Spreadsheet.getRangeByName('ФлОтладка');

  if (Range != undefined && Range.getValue())
  {
    if (clearTest)
      Spreadsheet.getSheetByName('Test').clear();

    return true;
  }
  return false;
}

// Очистка листа отладки
function dbgClearTestSheet()
{
  SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName('Test')
  .clear()
  .activate();
}

// Разбиваем длинную строку ( >50000 ) на несколько строк по maxLngth символов
function dbgSplitLongString(sStr, maxLngth)
{
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

function dbgBillInfo(bBill)
{
  const s =
    " от (" + bBill.date +
    ") магазин >" + bBill.name +
    "< на сумму [" + bBill.summ + 
    "] р. наличными {" + bBill.cash + "}";
    //"} ФН :" + bBill.jsonBill.fiscalDriveNumber +
    //" ФД :" + bBill.jsonBill.fiscalDocumentNumber +
    //" ФП :" + bBill.jsonBill.fiscalSign +
    //" товаров :" + bBill.jsonBill.items.length;
  return s
}

// Отладка вставки чека на листе Расходы
function TestEditCell()
{
  //
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const br = ss.getSheetByName("Расходы").getRange(31,8,1,1);
  const sv = br.getValues(); 
  const s = br.getValue();
  SettingCostBill(ss, br);
}

function TestBInfo(){
  //
  const s = '{"cashTotalSum":0,"dateTime":"2024-03-08T18:58:00","fiscalDriveNumber":9960440503269952,"fiscalDocumentNumber":65146,"fiscalSign":25409474,"items":[{"name":"[М+] Вода СВЯТОЙ ИСТОЧНИК б/г   1.5л","price":4999,"quantity":1,"sum":4999,"unit":"шт."},{"name":"Пакет ПЯТЕРОЧКА 65х40см","price":849,"quantity":1,"sum":849,"unit":"шт."},{"name":"[М+] Вода СВЯТОЙ ИСТОЧНИК б/г   1.5л","price":4999,"quantity":1,"sum":4999,"unit":"шт."},{"name":"[М+] Вода СВЯТОЙ ИСТОЧНИК б/г   1.5л","price":4999,"quantity":1,"sum":4999,"unit":"шт."},{"name":"Яйцо СЕЛЯНОЧКА кур.С0 10шт","price":13499,"quantity":1,"sum":13499,"unit":"шт."}],"totalSum":54340,"user":"ООО \"Агроторг\"","userInn":"5036045205  "}';
  const b = billInfo(s);
}

function TestputBillsToExpenses()
{
  //
  const sBill = '{"cashTotalSum":0,"code":3,"creditSum":0,"dateTime":"2024-01-23T22:58:00","ecashTotalSum":169000,"fiscalDocumentFormatVer":2,"fiscalDocumentNumber":2266,"fiscalDriveNumber":"7282440700394281","fiscalSign":6132709,"items":[{"name":"ОПЯТА Светлое 0,5","nds":6,"paymentType":4,"price":25000,"productType":1,"quantity":2,"sum":50000},{"name":"ОПЯТА Светлое 0,5","nds":6,"paymentType":4,"price":25000,"productType":1,"quantity":1,"sum":25000},{"name":"Негрони","nds":6,"paymentType":4,"price":47000,"productType":1,"quantity":2,"sum":94000}],"kktRegId":"0001538015044333    ","ndsNo":169000,"operationType":1,"operator":"Елисеева Вика","prepaidSum":0,"provisionSum":0,"requestNumber":22,"retailPlace":"ресторан","shiftNumber":74,"taxationType":16,"appliedTaxationType":16,"totalSum":169000,"user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"НИКО\"","userInn":"7724009233  "}';

  const sBill2 = '{"cashTotalSum":0,"code":3,"creditSum":0,"dateTime":"2024-01-23T14:38:00","ecashTotalSum":47408,"fiscalDocumentFormatVer":4,"fiscalDocumentNumber":29867,"fiscalDriveNumber":"7281440501088639","fiscalSign":1608626468,"fnsUrl":"www.nalog.gov.ru","items":[{"name":"Борщ по-московски с мясом и сосисками","nds":6,"paymentType":4,"price":9800,"productType":1,"quantity":1,"sum":9800},{"name":"Тушеная квашеная капуста","nds":6,"paymentType":4,"price":5500,"productType":1,"quantity":0.5,"sum":2750},{"name":"Котлеты фаршированные карамелизированным луком","nds":6,"paymentType":4,"price":13700,"productType":1,"quantity":1,"sum":13700},{"name":"Майонез порционный","nds":6,"paymentType":4,"price":1100,"productType":1,"quantity":1,"sum":1100},{"name":"Салат Бар","nds":6,"paymentType":4,"price":79000,"productType":1,"quantity":0.202,"sum":15958},{"name":"Сметана","nds":6,"paymentType":4,"price":3500,"productType":1,"quantity":1,"sum":3500},{"name":"Черный хлеб","nds":6,"paymentType":4,"price":300,"productType":1,"quantity":2,"sum":600}],"kktRegId":"0005478468028273    ","ndsNo":47408,"operationType":1,"operator":"Тома Е.Н.","prepaidSum":0,"provisionSum":0,"requestNumber":366,"retailPlace":"Столовая","retailPlaceAddress":"125284,г. Москва, ул. Беговая д.3 стр.1","shiftNumber":74,"taxationType":2,"appliedTaxationType":2,"totalSum":47408,"user":"ИП Блинов М.Л.","userInn":"472002884724"}';

  const bBill = billInfo(sBill);
  const bBill2 = billInfo(sBill2);

  putBillsToExpenses([bBill, bBill2]);

/*
  let testBill = {date: "2024-01-23T22:58:00", summ: 169000, name: "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"НИКО\"", cash: 0};
  let testBill2 = {date: "2024-01-23T14:38:00", summ: 47408, name: "ИП Блинов М.Л.", cash: 0};
  putBillsToExpenses([testBill, testBill2]); */
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
