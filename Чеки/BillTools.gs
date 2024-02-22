/*

billInfo(sBill)
billAllInfo(sBill)
filterUnqGoods(arrGoods)

*/

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
  const sName = sBill.slice(i, sBill.indexOf("\",", i+1)).replace(/\\\"/g,"\"");

  // ФН
  i = sBill.indexOf("\"fiscalDriveNumber\":")+21;
  const sFN = sBill.slice(i, sBill.indexOf("\"", i));

  // ФД
  i = sBill.indexOf("\"fiscalDocumentNumber\":")+23;
  const sFD = sBill.slice(i, sBill.indexOf(",", i));

  // ФП
  i = sBill.indexOf("\"fiscalSign\":")+13;
  const sFP = sBill.slice(i, sBill.indexOf(",", i));

  return {number: 0, dtime: dDate.getTime(), sdate: sDate, total: sSumm / 100, cache: sCash / 100, fn: sFN, fd: sFD, fp: sFP, name: sName};
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
    iQuantity = sBill.slice(i, j) / 1.0; // .replace(".", ",")

    i = sBill.indexOf(",\"sum\":", j)+7;
    j = Math.min(sBill.indexOf("}", i), sBill.indexOf(",", i));
    iSum = sBill.slice(i, j);

    bItems.push({iname: iName, iprice: iPrice / 100, iquantity: iQuantity, isum: iSum / 100});

    i = sBill.indexOf("\"name\":", j);
  }

  inf.items = filterUnqGoods(bItems);
  return inf;
}

function filterUnqGoods(arrGoods)
{
  let newGoods = [];
  for (itm of arrGoods) {
    let i = newGoods.findIndex((element) => element.iprice == itm.iprice && element.iname == itm.iname);
    if (~i) {
      newGoods[i].iquantity += itm.iquantity;
      newGoods[i].isum += itm.isum;
    } else
      newGoods.push(itm);
  }
  if (newGoods.length == arrGoods.length)
    return arrGoods;
  else
    return newGoods;
}
