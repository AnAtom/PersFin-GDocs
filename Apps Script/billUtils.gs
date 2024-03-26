/* 

billFilterName(sName) - Вырезает название магазина и возвращает в верхнем регистре.
billInfo(sBill) - Возвращает Дату, Сумму и Магазин чека из json строки.
billAllInfo(sBill) - Возвращает информацию о чеке, включая список продуктов.
billInfoStr(bBill) - Формирует из JSON чека информационную строку для логирования.

*/

// JSON нового формата

/* [{"_id":"65b01a0cae17240837f960c2","createdAt":"2024-01-23T19:57:00+00:00","ticket":{"document":{"receipt":

{"cashTotalSum":0,"code":3,"creditSum":0,
"dateTime":"2024-01-23T22:58:00","ecashTotalSum":169000,"fiscalDocumentFormatVer":2,
"fiscalDocumentNumber":2266,              ФД
"fiscalDriveNumber":"7282440700394281",   ФН
"fiscalSign":6132709,                     ФПД

"items":[
{"name":"ОПЯТА Светлое 0,5","nds":6,"paymentType":4,"price":25000,"productType":1,"quantity":1,"sum":25000},
{
  "name":"ГИННЕС 0,5","nds":6,"paymentType":4,
  "price":139800,"productType":1,"productCodeDataError":"not supported product type 5",
  "quantity":0.5,
  "sum":69900
},
{"name":"Негрони","nds":6,"paymentType":4,"price":47000,"productType":1,"quantity":2,"sum":94000}
],"kktRegId":"0001538015044333    ","ndsNo":169000,"operationType":1,"operator":"Елисеева Вика","prepaidSum":0,"provisionSum":0,"requestNumber":22,"retailPlace":"ресторан","shiftNumber":74,"taxationType":16,"appliedTaxationType":16,

"totalSum":169000,
"user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"НИКО\"","userInn":"7724009233  "}

}}}] */

// Выделяет из названия организации конкретено название
function billFilterName(sName)
{
  const s = cutInsideQuotes(sName);
  const zs = '';

  if (s != zs)
    return s.toUpperCase();
  else
    return sName.toUpperCase()
      .replace("ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", zs)
      .replace("АКЦИОНЕРНОЕ ОБЩЕСТВО", zs)
      .replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", zs)
      .replace("ЗАО", zs)
      .replace("АО", zs)
      .replace("ООО", zs)
      .replace("ИП", zs)
      .replace(/\s\s+/g,' ')
      .trim();
}

// Возвращает Дату, Сумму и Магазин чека из json строки
function billInfo(sBill)
{
  // Проверка, что это json чека
  let i = sBill.indexOf("\"totalSum\":");
  if (i == -1)
    return undefined;

  // Сумма
  i += 11;
  const iSumm = parseInt(sBill.slice(i, sBill.indexOf(",", i)));

  // Дата
  i = sBill.indexOf("\"dateTime\":")+12;
  const sDate = sBill.slice(i, sBill.indexOf("\"", i+1));
  const dDate = new Date(sDate);
  const aDay = new Date(sDate.slice(0, sDate.indexOf("T")) + "T00:00:00");
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test").getRange(1,1).setValue(dDate);
  //Logger.log( "дата ["+ sDate +"] data " + dDate + " день {" + ss + "} day " + aDay);

  // Наличные
  i = sBill.indexOf("\"cashTotalSum\":")+15;
  const iCash = parseInt(sBill.slice(i, sBill.indexOf(",", i)));

  // Магазин
  i = sBill.indexOf("\"user\":\"")+8;
  const sName = sBill.slice(i, sBill.indexOf("\",", i+1))
    .replace(/\\\"/g,"\"")
    .trim();
  const sShop = billFilterName(sName);

  return {dTime: dDate.getTime(), tDate: aDay.getTime(), date: sDate, summ: iSumm / 100.0, cash: iCash / 100.0, name: sName, shop: sShop};
}

// Возвращает информацию о чеке, включая список продуктов.
function billAllInfo(sBill)
{
  let inf = billInfo(sBill);
  if (inf == undefined) return inf;

  let bItems = [];
  let iName = "";
  let iPrice = 0;
  let iSum = 0;
  let iQuantity = 0;

  // {"name":"Негрони","nds":6,"paymentType":4,"price":47000,"productType":1,"quantity":2,"sum":94000}

  let i = sBill.indexOf("\"name\":",sBill.indexOf("\"items\":[")+9);
  while (i != -1) {
    i += 8;
    let j = sBill.indexOf(",\"nds\":", i)-1;
    iName = sBill.slice(i,j).replace(/\\\"/g,"\"");

    i = sBill.indexOf(",\"price\":", j+8)+9;
    j = sBill.indexOf(",", i);
    iPrice = sBill.slice(i, j);

    i = sBill.indexOf(",\"quantity\":", j+1)+12;
    j = sBill.indexOf(",", i);
    iQuantity = sBill.slice(i, j);

    i = sBill.indexOf(",\"sum\":", j)+7;
    j = sBill.indexOf("}", i);
    iSum = sBill.slice(i, j);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});

    i = sBill.indexOf("\"name\":", j);
  }

  inf.items = bItems;
  return inf;
}
