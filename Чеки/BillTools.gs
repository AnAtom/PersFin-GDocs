/*

billFormatShort(jBill)
billFormatText(sBill)
billInfo(sBill)
billAllInfo(sBill)
cutPromoTag(sProduct)
billInfoStr(pBill)
billFilterName(sName)

*/

function billFormatShort(jBill)
{
  return JSON.stringify(jBill)
    .replace(/\"fiscalDriveNumber/, "\n\"fiscalDriveNumber")
    .replace(/,\"items/, ",\n\"items")
    .replace(/\[{\"name/, "[\n\t{\"name")
    .replace(/,{\"name/g, ",\n\t{\"name")
    .replace(/,\"totalSum/, ",\n\"totalSum");
}

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
  const s = CutInsideQuotes(sName);
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

  // ФН
  i = sBill.indexOf("\"fiscalDriveNumber\":")+21;
  const sFN = sBill.slice(i, sBill.indexOf("\"", i));

  // ФД
  i = sBill.indexOf("\"fiscalDocumentNumber\":")+23;
  const sFD = sBill.slice(i, sBill.indexOf(",", i));

  // ФП
  i = sBill.indexOf("\"fiscalSign\":")+13;
  const sFP = sBill.slice(i, sBill.indexOf(",", i));

  const jBill = {cashTotalSum: sCash / 100.0, dateTime: dDate, fiscalDriveNumber: sFN / 1.0, fiscalDocumentNumber: sFD / 1.0, fiscalSign: sFP / 1.0,
                  items: [], totalSum: sSumm / 100.0, user: sName}

  return {dTime: dDate.getTime(), SN: 0, URL: '', Shop: sShop, jsonBill: jBill};
}

// Возвращает информацию о чеке, включая список продуктов.
function billAllInfo(sBill)
{
  let inf = billInfo(sBill);
  if (inf == undefined) return inf;

  // {"name":"Негрони","nds":6,"paymentType":4,"price":47000,"productType":1,"quantity":2,"sum":94000,"unit":"liter"}

  let iName = "";
  let iPrice = 0;
  let iSum = 0;
  let iQuantity = 0;
  let iUnit = "";

  let i = sBill.indexOf("\"items\":[")+9;
  i = sBill.indexOf("\"name\":", i);
  while (~i) {
    i += 8;
    let j = sBill.indexOf(",\"nds\":", i)-1;
    iName = sBill.slice(i,j).replace(/\\\"/g,"\"").trim();

    i = sBill.indexOf(",\"price\":", j+8)+9;
    j = sBill.indexOf(",", i);
    iPrice = sBill.slice(i, j);

    i = sBill.indexOf(",\"quantity\":", j+1)+12;
    j = sBill.indexOf(",", i);
    iQuantity = sBill.slice(i, j) / 1.0;

    i = sBill.indexOf(",\"sum\":", j)+7;
    j = sBill.indexOf("}", i);
    let m = sBill.indexOf(",", i);
    if (j > m) {
      i = sBill.indexOf("unit\":", m)+7;
      j = sBill.indexOf("}", i)-1;
      iUnit = sBill.slice(i, j);
    } else {
      iSum = sBill.slice(i, j);
      iUnit = "";
    }

    inf.jsonBill.items.push({name: iName, price: iPrice / 100, quantity: iQuantity / 1.0, sum: iSum / 100, unit: iUnit});

    i = sBill.indexOf("\"name\":", j);
  }

  return inf;
}

// Отрезаем артикулы и акционные метки в начале названия товара: 0000, *, <A>, [M], [M+]
function cutPromoTag(sProduct)
{
  const zs = '';
  return sProduct
    .replace(/^\<А\> /, zs)
    .replace(/^\[М\+?\] /, zs)
    .replace(/^[0-9]+ /, zs)
    .replace(/^\*/, zs);
}

function billInfoStr(bBill)
{
  const s =
    " от (" + bBill.jsonBill.dateTime.toISOString() +
    ") магазин >" + bBill.jsonBill.user +
    "< на сумму [" + bBill.jsonBill.totalSum + "] р. наличными {" + bBill.jsonBill.cashTotalSum +
    "} ФН :" + bBill.jsonBill.fiscalDriveNumber +
    " ФД :" + bBill.jsonBill.fiscalDocumentNumber +
    " ФП :" + bBill.jsonBill.fiscalSign +
    " товаров :" + bBill.jsonBill.items.length;
  return s
}
