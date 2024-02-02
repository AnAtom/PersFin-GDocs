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
  .replace(/\[{\"name/, "[\n{\"name")
  .replace(/,{\"name/g, ",\n{\"name")
  .replace(/}],\"kktRegId/, "}\n],\"kktRegId")
  .replace(/,\"totalSum/, ",\n\"totalSum")
  .replace(/,\"user\"/, ",\n\"user\"");
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

  // Наличные
  i = sBill.indexOf("\"cashTotalSum\":")+15;
  const sCash = sBill.slice(i, sBill.indexOf(",", i));

  // Магазин
  i = sBill.indexOf("\"user\":\"")+8;
  const sName = sBill.slice(i, sBill.indexOf("\",", i+1)).replace(/\\\"/g,"\"");

  return {date: sDate, summ: sSumm, name: sName, cash: sCash};
}

// Возвращает информацию о чеке, включая список продуктов.
function billAllInfo(sBill)
{
  let inf = billInfo(sBill);

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

    i = sBill.indexOf(",\"sum\":", j)+6;
    j = sBill.indexOf("}", i);
    iSum = sBill.slice(i, j);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});

    i = sBill.indexOf("\"name\":", j);
  }

  inf.items = bItems;
  return inf;
}
