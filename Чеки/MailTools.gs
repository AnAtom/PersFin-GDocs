/*



*/

function mailScanOnTimer()
{
  //
}

function mailGenericGetInfo()
{
  //
}

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
    iName = finLib.betweenFrom(sBill, j, "<span class=", "</td>", ">", "</span>");
    k = sBill.indexOf('</table>', j)+7;
    iQuantity = finLib.betweenFrom(sBill, k, "<span class=", "x", ">", "</span>");
    j = sBill.indexOf('</span>', k)+6;
    iPrice = finLib.betweenFrom(sBill, j, "<span class=", "</td>", ">", "</span>");
    k = sBill.indexOf('<td class', j)+9;
    iSum = finLib.betweenFrom(sBill, k, "<span class=", "</td>", ">", "</span>");
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
    iName = finLib.betweenFrom(sBill, j, "style=", "/div", ">", "<");
    k = sBill.indexOf('check-col-right', j)+14;
    iAll = finLib.betweenFrom(sBill, k, "style=", "/div", ">", "<");
    j = iAll.indexOf("х");
    iQuantity = iAll.slice(0,j).trim();
    iPrice = iAll.slice(j+1).trim();
    j = sBill.indexOf('check-col-right', k)+14;
    iSum = finLib.betweenFrom(sBill, j, "style=", "/div", ">", "<");
    j = sBill.indexOf('check-product-name', j);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = finLib.between2(sBill, 'check-totals', 'check-row', 'check-col-right', '</div>');
  j = sTotal.indexOf(">");
  sTotal = sTotal.slice(j+1).trim().replace(".", ",");
  
  // return {summ: sTotal, date: sDate, name: sName, items: bItems};
  return {summ: sTotal, date: sDate, name: sName, items: []};
}

function getBeelineBillInfo(sBill) {
  var sName = finLib.between(sBill, '<p style="padding:0; margin: 0; color: #282828; font-size: 13px; line-height: normal;">', '/p').trim();

  var sDate = finLib.between2(sBill, 'Дата | Время', '</tr>', '"right">', '</td>').replace("|", "").trim();

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
    iName = finLib.betweenFrom(sBill, j, "style=", "/span", ">", "<");
    k = sBill.indexOf('Цена*Кол', j)+7;
    iPrice = finLib.betweenFrom(sBill, k, "<td width=", "/td", ">", "<");
    j = sBill.indexOf('<td align=', k)+9;
    iQuantity = finLib.betweenFrom(sBill, j, "right", "/td", ">", "<");
    k = sBill.indexOf('Сумма', j)+4;
    iSum = finLib.betweenFrom(sBill, k, "<td width", "/td", ">", "<");
    j = sBill.indexOf(s, k);

    bItems.push({iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum});
  }

  var sTotal = finLib.between2(sBill, 'Итог:', '</tr>', '21px;">', '</span>').replace(".", ",");

  return {summ: sTotal, date: sDate, name: sName, items: []};
  // return {summ: sTotal, date: sDate, name: sName, items: bItems};
}
