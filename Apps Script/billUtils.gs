/* 

billFormatText(sBill) - Разбивает строку JSON на несколько строк для читаемости. Удаляет ненужное начало и хвост.
billInfo(sBill) - Возвращает Дату, Сумму и Магазин чека из json строки.
billAllInfo(sBill) - Возвращает информацию о чеке, включая список продуктов.

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

// Разбивает строку JSON на несколько строк для читаемости. Удаляем '[{"_id": ... "ticket":{"document":{"receipt":', '}}}]'
function billFormatText(sBill)
{
  return sBill.slice(sBill.indexOf("receipt\":{")+9, -4)
    .replace(/,\"dateTime/, ",\n\"dateTime")
    .replace(/\"fiscalDocumentNumber/, "\n\"fiscalDocumentNumber")
    .replace(/\"fiscalDriveNumber/, "\n\"fiscalDriveNumber")
    .replace(/\"fiscalSign/, "\n\"fiscalSign")
    .replace(/,\"items/, ",\n\"items")
    .replace(/\[{\"name/, "[\n\t{\"name")
    .replace(/,{\"name/g, ",\n\t{\"name")
    .replace(/}],\"kktRegId/, "}\n],\"kktRegId")
    .replace(/,\"totalSum/, ",\n\"totalSum")
    .replace(/,\"user\"/, ",\n\"user\"");
}

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
  const sSumm = sBill.slice(i, sBill.indexOf(",", i));

  // Дата
  i = sBill.indexOf("\"dateTime\":")+12;
  const sDate = sBill.slice(i, sBill.indexOf("\"", i+1));
  const dDate = new Date(sDate);

  // Наличные
  i = sBill.indexOf("\"cashTotalSum\":")+15;
  const sCash = sBill.slice(i, sBill.indexOf(",", i));

  // Магазин
  i = sBill.indexOf("\"user\":\"")+8;
  const sName = sBill.slice(i, sBill.indexOf("\",", i+1))
    .replace(/\\\"/g,"\"")
    .trim();
  const sShop = billFilterName(sName);

  //const jBill = {cashTotalSum: sCash / 100.0, dateTime: dDate, fiscalDriveNumber: sFN / 1.0, fiscalDocumentNumber: sFD / 1.0, fiscalSign: sFP / 1.0,
  //                items: [], totalSum: sSumm / 100.0, user: sName}

  return {dTime: dDate.getTime(), date: sDate, summ: sSumm / 100.0, cash: sCash / 100.0, name: sName, shop: sShop};
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

function billInfoStr(bBill)
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
