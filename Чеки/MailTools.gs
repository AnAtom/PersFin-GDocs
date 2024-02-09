/*



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
  const vTmplts = rTemplates.getValues();
  let Tmplts = [];

  for (let i = 0; i < rTemplates.getNumColumns(); i++)
  {
    let j = 0;
    let Tmplt = {
      from: vTmplts[j++][i], key: vTmplts[j++][i],
      name_pt: vTmplts[j++][i], name_s1: vTmplts[j++][i], name_e1: vTmplts[j++][i], name_s2: vTmplts[j++][i], name_e2: vTmplts[j++][i],
      date_pt: vTmplts[j++][i], date_s1: vTmplts[j++][i], date_e1: vTmplts[j++][i], date_s2: vTmplts[j++][i], date_e2: vTmplts[j++][i],
      total_pt: vTmplts[j++][i], total_s1: vTmplts[j++][i], total_e1: vTmplts[j++][i], total_s2: vTmplts[j++][i], total_e2: vTmplts[j++][i],
      cache_pt: vTmplts[j++][i], cache_s1: vTmplts[j++][i], cache_e1: vTmplts[j++][i], cache_s2: vTmplts[j++][i], cache_e2: vTmplts[j++][i],
      fn_pt: vTmplts[j++][i], fn_s1: vTmplts[j++][i], fn_e1: vTmplts[j++][i], fn_s2: vTmplts[j++][i], fn_e2: vTmplts[j++][i],
      fd_pt: vTmplts[j++][i], fd_s1: vTmplts[j++][i], fd_e1: vTmplts[j++][i], fd_s2: vTmplts[j++][i], fd_e2: vTmplts[j++][i],
      fp_pt: vTmplts[j++][i], fp_s1: vTmplts[j++][i], fp_e1: vTmplts[j++][i], fp_s2: vTmplts[j++][i], fp_e2: vTmplts[j++][i]
    };
    Tmplts.push(Tmplt);
  }
  Logger.log("Загружено " + Tmplts.length + " шаблонов.");
  return Tmplts;
}

function FindInTemplates(arrTmplts, mailFrom)
{
  for (let i = 0; i < arrTmplts.length; i++)
    if (arrTmplts[i].from == mailFrom) return i;
  return -1;
}

function CutByTemplate(str, pt, s1, e1, s2, e2)
{
  switch(pt) {
    case 0: return between(str, s1, e1).trim();
    case 1: return between2(str, s1, e1, s2, e2).trim();
    case 2:
        let s = between2(str, s1, e1, s2, e2);
        return s.slice(s.indexOf(">")+1).trim();
    default:
        Logger.log("Неизвестный Preprocessing Type " + pt + " ");
  }
  return "";
}

function mailGenericGetInfo(mailTmplt, email)
{
  // mailTmplt.name_s1
  let sName = "";
  sName = CutByTemplate(email, mailTmplt.name_pt, mailTmplt.name_s1, mailTmplt.name_e1, mailTmplt.name_s2, mailTmplt.name_e2);
  //sName = between2(email, mailTmplt.name_s1, mailTmplt.name_e1, mailTmplt.name_s2, mailTmplt.name_e2).trim();
  sName = sName.replace(/&quot;/g, '"');
  if (sName.indexOf('"') == 0)
    sName = sName.slice(1, sName.length-1);

  let sDate = "";
  sDate = CutByTemplate(email, mailTmplt.date_pt, mailTmplt.date_s1, mailTmplt.date_e1, mailTmplt.date_s2, mailTmplt.date_e2);
  //sDate = between2(email, mailTmplt.date_s1, mailTmplt.date_e1, mailTmplt.date_s2, mailTmplt.date_e2).trim();
  sDate = sDate.replace("|", "");

  let sTotal = "";
  sTotal = CutByTemplate(email, mailTmplt.total_pt, mailTmplt.total_s1, mailTmplt.total_e1, mailTmplt.total_s2, mailTmplt.total_e2);
  //sTotal = between2(email, mailTmplt.total_s1, mailTmplt.total_e1, mailTmplt.total_s2, mailTmplt.total_e2).trim();
  sTotal = sTotal.replace(".", ",");

  let sCache = "";
  sCache = CutByTemplate(email, mailTmplt.cache_pt, mailTmplt.cache_s1, mailTmplt.cache_e1, mailTmplt.cache_s2, mailTmplt.cache_e2);
  if (sCache == "") sCache = 0;
  else sCache = sCache.replace(".", ",");

  let sFN = "";
  sFN = CutByTemplate(email, mailTmplt.fn_pt, mailTmplt.fn_s1, mailTmplt.fn_e1, mailTmplt.fn_s2, mailTmplt.fn_e2);

  let sFD = "";
  sFD = CutByTemplate(email, mailTmplt.fd_pt, mailTmplt.fd_s1, mailTmplt.fd_e1, mailTmplt.fd_s2, mailTmplt.fd_e2);

  let sFP = "";
  sFP = CutByTemplate(email, mailTmplt.fp_pt, mailTmplt.fp_s1, mailTmplt.fp_e1, mailTmplt.fp_s2, mailTmplt.fp_e2);

  return {date: sDate, total: sTotal, cache: sCache, name: sName, fn: sFN, fd: sFD, fp: sFP};
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
