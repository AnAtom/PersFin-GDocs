/*

GetTemplates - читает шаблоны для парсинга чеков в почте от различных ОФД
CutByTemplate - вырезает значение из сообщения по шаблону
getDateTime - Читает дату из строки и возвращает 
mailGenericGetInfo - универсальная процедура парсинга сообщения по шаблону

*/

/*

From	- адрес с которого пришел чек
Key		- Проверочная строка, которая должна присутствовать в чеке

Name pt		- Тип процедуры вырезки значения названия магазина
Name s1		Name e1		- строки начала и окончания первого уровня вырезки
Name s2		Name e2		- строки начала и окончания второго уровня вырезки

Date pt		Date s1		Date e1		Date s2		Date e2		- то же для даты чека
Total s1...		- то же для суммы чека
Cache s1...		- то же для суммы наличными
FN s1...		- то же для ФН
FD s1...		- то же для ФД
FP s1...		- то же для ФПД
Items s1
Item s1
iName s1
iQuantity s1
iPrice s1
iSum s1

*/

function GetTemplates(rTemplates)
{
  // Читаем значения полей шаблона из диапазона
  const v = rTemplates.getValues();
  let Tmplts = [];

  for (let i = 0; i < rTemplates.getNumColumns(); i++)
  {
    let j = 0;
    const sFrom = v[j++][i];
    const sKey = v[j++][i];
    const tName = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tDate = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tTotal = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tCache = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFN = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFD = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFP = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};

    //const tIName = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    //const tIQuantity = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    //const tIPrice = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    //const tISum = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};

    Tmplts.push({
      from: sFrom,
      key: sKey,
      name: tName,
      date: tDate,
      total: tTotal,
      cache: tCache,
      fn: tFN,
      fd: tFD,
      fp: tFP

      // iname: tIName,
      // iqntty: tIQuantity,
      // iprice: tIPrice,
      // isum: tISum
    });
  }
  Logger.log("Загружено " + Tmplts.length + " шаблонов.");
  return Tmplts;
}

function CutByTemplate(str, tmplt)
{
  switch(tmplt.pt) {
    case 0: return between(str, tmplt.s1, tmplt.e1).trim();
    case 1: return between2(str, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2).trim();
    case 2:
        let s = between2(str, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
        return s.slice(s.indexOf(">")+1).trim();
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
}

function getDateTime(s)
{
  const f = "20" + s.slice(6, 8)  // Год
    + "-" + s.slice(3, 5)         // месяц
    + "-" + s.slice(0, 2)         // день
    + "T" + s.slice(9);           // время
  const d = new Date(f);
  return d.getTime();
}

function mailGenericGetInfo(mailTmplt, email)
{
  // Вырезаем имя
  let sName = CutByTemplate(email, mailTmplt.name)
    .replace(/&quot;/g, '"');
  // Убираем обрамляющие кавычки
  if (sName.indexOf('"') == 0)
    sName = sName.slice(1, sName.length-1);

  const sDate = CutByTemplate(email, mailTmplt.date)
    .replace(" | ", " ")
    .replace(".202", ".2");
  const timeDate = getDateTime(sDate);

  const sTotal = CutByTemplate(email, mailTmplt.total)
    .replace(".", ",");

  let sCache = CutByTemplate(email, mailTmplt.cache);
  if (sCache == "") sCache = 0;
  else sCache = sCache.replace(".", ",");

  const sFN = CutByTemplate(email, mailTmplt.fn);
  const sFD = CutByTemplate(email, mailTmplt.fd);
  const sFP = CutByTemplate(email, mailTmplt.fp);

  let arrItems = [];

  return {number: 0, dtime: timeDate, sdate: sDate, total: sTotal, cache: sCache, fn: sFN, fd: sFD, fp: sFP, name: sName, items: arrItems};
}

/*

function mailScanOnTimer()
{
  //
}

*/

function getTaxcomBillInfo(sBill) {
  var sDate = between2(sBill, 'КАССОВЫЙ ЧЕК', '</tr>', 'receipt-value-1012', '</span>');
  var j = sDate.indexOf(">");
  sDate = sDate.slice(j+1);

  var sTotal = between2(sBill, 'ИТОГО:', '</tr>', 'receipt-value-1020', '</span>');
  j = sTotal.indexOf(">");
  sTotal = sTotal.slice(j+1).replace(".", ",");
  // >НАЛИЧНЫМИ:<
  // <span class="name receipt-name-1081">БЕЗНАЛИЧНЫМИ:</span>
  // <span class="name receipt-name-1031">НАЛИЧНЫМИ:</span>

  var sName = between(sBill, '<text>Данный чек подтверждает совершение расчетов в <b>', '</b>.</text>');
  var sFN = between(sBill, '<span class="value receipt-value-1041">', '</span>');
  var sFD = between(sBill, '<span class="value receipt-value-1040">', '</span>');
  var sFP = between(sBill, '<span class="value receipt-value-1077">', '</span>');
  
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
    iName = betweenFrom(sBill, j, "<span class=", "</td>", ">", "</span>");
    k = sBill.indexOf('</table>', j)+7;
    iQuantity = betweenFrom(sBill, k, "<span class=", "x", ">", "</span>");
    j = sBill.indexOf('</span>', k)+6;
    iPrice = betweenFrom(sBill, j, "<span class=", "</td>", ">", "</span>");
    k = sBill.indexOf('<td class', j)+9;
    iSum = betweenFrom(sBill, k, "<span class=", "</td>", ">", "</span>");
    j = sBill.indexOf('<div class="item">', k);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  //return {summ: sTotal, date: sDate, name: sName, FN: sFN, FD: sFD, FP: sFP, items: bItems};
  return {summ: sTotal, date: sDate, name: sName, FN: sFN, FD: sFD, FP: sFP, items: []};
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
    iName = betweenFrom(sBill, j, "style=", "/div", ">", "<");
    k = sBill.indexOf('check-col-right', j)+14;
    iAll = betweenFrom(sBill, k, "style=", "/div", ">", "<");
    j = iAll.indexOf("х");
    iQuantity = iAll.slice(0,j).trim();
    iPrice = iAll.slice(j+1).trim();
    j = sBill.indexOf('check-col-right', k)+14;
    iSum = betweenFrom(sBill, j, "style=", "/div", ">", "<");
    j = sBill.indexOf('check-product-name', j);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = between2(sBill, 'check-totals', 'check-row', 'check-col-right', '</div>');
  j = sTotal.indexOf(">");
  sTotal = sTotal.slice(j+1).trim().replace(".", ",");
  
  // return {summ: sTotal, date: sDate, name: sName, items: bItems};
  return {summ: sTotal, date: sDate, name: sName, items: []};
}

function getBeelineBillInfo(sBill) {
  var sName = between(sBill, '<p style="padding:0; margin: 0; color: #282828; font-size: 13px; line-height: normal;">', '/p').trim();

  var sDate = between2(sBill, 'Дата | Время', '</tr>', '"right">', '</td>').replace("|", "").trim();

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
    iName = betweenFrom(sBill, j, "style=", "/span", ">", "<");
    k = sBill.indexOf('Цена*Кол', j)+7;
    iPrice = betweenFrom(sBill, k, "<td width=", "/td", ">", "<");
    j = sBill.indexOf('<td align=', k)+9;
    iQuantity = betweenFrom(sBill, j, "right", "/td", ">", "<");
    k = sBill.indexOf('Сумма', j)+4;
    iSum = betweenFrom(sBill, k, "<td width", "/td", ">", "<");
    j = sBill.indexOf(s, k);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = between2(sBill, 'Итог:', '</tr>', '21px;">', '</span>').replace(".", ",");

  return {summ: sTotal, date: sDate, name: sName, items: []};
  // return {summ: sTotal, date: sDate, name: sName, items: bItems};
}
