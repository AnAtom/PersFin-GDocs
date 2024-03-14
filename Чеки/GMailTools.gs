/*

CutByTemplate - вырезает строку из сообщения по шаблону
CutFromPosByTemplate - вырезает строку из сообщения после заданной позиции по шаблону
getDateTime - читает дату из строки и возвращает UNIX время
mailMagnitGetGoods - процедура чтения списка товаров Магнит ОФД
mailFirstGetGoods - процедура чтения списка товаров Первый ОФД
mailGenericGetInfo - универсальная процедура парсинга сообщения по шаблону
GetTemplates - читает шаблоны для парсинга чеков в почте от различных ОФД
ScanMail - читает чеки из новых писем

*/

function CutByTemplate(str, tmplt)
{
  switch(tmplt.pt) {
    case 0: return between(str, tmplt.s1, tmplt.e1).trim();
    case 1: return between2(str, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2).trim();
    case 2:
        let s = between2(str, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
        return s.slice(s.indexOf(">")+1).trim();
    case 11: return "";
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
}

function CutFromPosByTemplate(str, fpos, tmplt)
{
  switch(tmplt.pt) {
    case 0: return cutfrom(str, fpos, tmplt.s1, tmplt.e1);
    case 1: return between2from(str, fpos, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
    case 2:
        let s = between2from(str, fpos, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
        return s.slice(s.indexOf(">")+1);
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
}

function mailMagnitGetGoods(fPos, email)
{
  let arrItems = [];
  let iName = "";
  let iPrice = 0;
  let iQuantity = 0;
  let iSum = 0;
  let iUnit = "";
  let i = 0;
  let j = 0;
  const ePos = email.indexOf("</tbody>", fPos);
  let cPos = email.indexOf("<td bgcolor=", fPos);
  while (cPos < ePos) {
    i = email.indexOf("<td class=", cPos);
    if (i > ePos)
      break;
    i = email.indexOf(">", i) + 1;
    j = email.indexOf("</td>", i);
    iName = email.slice(i, j).trim();

    i = email.indexOf(">", j + 5) + 1;
    j = email.indexOf("</td>", i);
    iQuantity = email.slice(i, j).trim() / 1.0;

    i = email.indexOf(">", j + 5) + 1;
    j = email.indexOf("</td>", i);
    iPrice = Math.round(email.slice(i, j).trim() * 100.0);

    i = email.indexOf(">", j + 5) + 1;
    j = email.indexOf("</td>", i);
    iSum = Math.round(email.slice(i, j).trim() * 100.0);
    //
    arrItems.push({name: iName, price: iPrice, quantity: iQuantity, sum: iSum, unit: iUnit});
    cPos = email.indexOf("<td bgcolor=", j + 5);
  }
  return arrItems;
}

function mailFirstGetGoods(fPos, email)
{
  let arrItems = [];
  let iName = "";
  let iPrice = 0;
  let iQuantity = 0;
  let iSum = 0;
  let iUnit = "";
  let i = 0;
  let j = 0;
  let cPos = email.indexOf("<td valign=\"top\" width=", fPos);
  while (~cPos) {
    i = email.indexOf(">", cPos + 23) + 1;
    j = email.indexOf("</td>", i);
    iName = email.slice(i, j).trim();

    i = email.indexOf(">", j + 5) + 1;
    j = email.indexOf("</td>", i);
    iPrice = Math.round(email.slice(i, j).replace(/\s/g,'').replace(",", ".") * 100.0);

    i = email.indexOf(">", j + 5) + 1;
    j = email.indexOf("</td>", i);
    iQuantity = email.slice(i, j).replace(/\s/g,'').replace(",", ".") / 1.0;

    i = email.indexOf(">", j + 5) + 1;
    j = email.indexOf("</td>", i);
    iSum = Math.round(email.slice(i, j).replace(/\s/g,'').replace(",", ".") * 100.0);

    arrItems.push({name: iName, price: iPrice, quantity: iQuantity, sum: iSum, unit: iUnit});
    cPos = email.indexOf("<td valign=\"top\" width=", j + 5);
  }
  return arrItems;
}

function mailGenericGetGoods(fPos, mailTmplt, email)
{
  let arrItems = [];
  let n = 1;
  let iName = "";
  let sQuantity = "";
  let iUnit = "";
  let iQuantity = 0;
  let iPrice = 0;
  let iSum = 0;
  const is = mailTmplt.item;
  const il = is.length;
  let j = email.indexOf(is, fPos);
  while (~j) {
    iName = CutFromPosByTemplate(email, j, mailTmplt.iname).trim();
    // Отрезаем нумерацию позиций NN:
    let k = iName.indexOf(':');
    if (~k && iName.slice(0, k) == n++)
      iName = iName.slice(k+1).trim();

    sQuantity = CutFromPosByTemplate(email, j, mailTmplt.iqntty).trim();
    // Отрезаем единицы измерения (шт.)
    k = sQuantity.indexOf(' ');
    if (~k) {
      iUnit = sQuantity.slice(k+1);
      sQuantity = sQuantity.slice(0, k);
    } else
      iUnit = "";
    iQuantity = sQuantity / 1.0;
    iPrice = Math.round(CutFromPosByTemplate(email, j, mailTmplt.iprice) * 100.0);
    iSum = Math.round(CutFromPosByTemplate(email, j, mailTmplt.isum) * 100.0);

    arrItems.push({name: iName, price: iPrice, quantity: iQuantity, sum: iSum, unit: iUnit});
    j = email.indexOf(is, j + il);
  }
  return arrItems;
}

function getDate(s)
{
  const d = "20" + s.slice(6, 8)  // Год
    + "-" + s.slice(3, 5)         // месяц
    + "-" + s.slice(0, 2)         // день
    + "T" + s.slice(9);           // время
  return new Date(d);
}

function mailGenericGetInfo(mailTmplt, email)
{
  // Вырезаем имя
  let sName = CutByTemplate(email, mailTmplt.name)
    .replace(/&quot;/g, '"');
  // Убираем обрамляющие кавычки
  if (sName.indexOf('"') == 0)
    sName = CutOuterQuotes(sName);

  const sShop = billFilterName(sName);

  let sDate = "";
  if (mailTmplt.date.pt == "D") {
    let di = email.indexOf(mailTmplt.date.s1) + mailTmplt.date.s1.length;
    const ds = mailTmplt.date.s2;
    const dl = ds.length;
    for (let ii = 0; ii < 5; ii++)
      di = email.indexOf(ds, di)+dl;
    di = email.indexOf(">", di) + 1;
    sDate = email.slice(di, email.indexOf(mailTmplt.date.e2, di)).replace(".202", ".2");
  } else if (mailTmplt.date.pt == "d") {
    let di = email.indexOf(mailTmplt.date.s1) + mailTmplt.date.s1.length;
    const ds = mailTmplt.date.s2;
    const dl = ds.length;
    for (let ii = 0; ii < 2; ii++)
      di = email.indexOf(ds, di)+dl;
    sDate = email.slice(di, email.indexOf(mailTmplt.date.e2, di)).trim().replace(".202", ".2");
  } else {
    sDate = CutByTemplate(email, mailTmplt.date)
      .replace(" | ", " ")
      .replace(".202", ".2");
  }
  const dDate = getDate(sDate);

  const sSumm = Math.round(CutByTemplate(email, mailTmplt.total).replace(/\s/g,'').replace(",", ".") * 100.0);

  let sCach = CutByTemplate(email, mailTmplt.cache);
  if (sCach == "") sCach = 0;
  else sCach = Math.round(sCach * 100.0);

  const iFN = parseInt(CutByTemplate(email, mailTmplt.fn));
  const iFD = parseInt(CutByTemplate(email, mailTmplt.fd));
  const iFP = parseInt(CutByTemplate(email, mailTmplt.fp));

  let arrItems = [];
  let i = email.indexOf(mailTmplt.items);
  if (~i) {
    i += mailTmplt.items.length;

    if (mailTmplt.item.slice(0, 7) != "intProc")
      arrItems = mailGenericGetGoods(i, mailTmplt, email);
    else // Невозможно выделить однозначные маркеры. Для получения данных используется специальная процедура.
      if (mailTmplt.item == "intProcM")
        arrItems = mailMagnitGetGoods(i, email);
      else if (mailTmplt.item == "intProc1")
        arrItems = mailFirstGetGoods(i, email);
  }
  const jBill = {cashTotalSum: sCach, dateTime: dDate, fiscalDriveNumber: iFN, fiscalDocumentNumber: iFD, fiscalSign: iFP,
                  items: arrItems, totalSum: sSumm, user: sName, userInn: 0}
  return {dTime: dDate.getTime(), SN: 0, URL: "", Shop: sShop, jsonBill: jBill};
}

/* Структура шаблона

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
Items s				- Строка после которой начинается перечисление товаров
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

    const sItemsFrom = v[j++][i];
    const sItemFrom = v[j++][i];
    const tIName = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tIQuantity = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tIPrice = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tISum = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};

    Tmplts.push({
      from: sFrom,
      key: sKey,
      name: tName,
      date: tDate,
      total: tTotal,
      cache: tCache,
      fn: tFN,
      fd: tFD,
      fp: tFP,

      items: sItemsFrom,
      item: sItemFrom,
      iname: tIName,
      iqntty: tIQuantity,
      iprice: tIPrice,
      isum: tISum
    });
  }
  Logger.log("Загружено " + Tmplts.length + " шаблонов.");
  return Tmplts;
}

function ScanMail(ss, dLastMailDate, arrBills)
{
  let newLastMailDate = dLastMailDate;
  let NumBills = 0;
  let bBill = {};

  // Читаем метку, под которой собраны чеки, из ячейки ЧекиПочта
  const sLabel = ss.getRangeByName('ЧекиПочта').getValue();
  Logger.log("Читаем чеки из почты с меткой: " + sLabel);

  // Читаем шаблоны для сканера
  const eTmplts = GetTemplates(ss.getRangeByName('ШаблоныЧеков'));

  // Сканируем цепочки писем
  const mailThreads = GmailApp.getUserLabelByName(sLabel).getThreads();
  let thrd = 1;
  let mURL = "";
  for (messages of mailThreads) {
    if (messages.getLastMessageDate() > dLastMailDate)
      mURL = messages.getPermalink();
    else
      continue;
    let m = 0;
    for (message of messages.getMessages()) {
      const dDate = message.getDate();
      if (dDate > dLastMailDate) {
        if (dDate > newLastMailDate)
          newLastMailDate = dDate;
      } else
        continue;

      const sBody = message.getBody();
      const sFrom = message.getFrom();
      let mFrom = sFrom;
      if (~sFrom.indexOf("<"))
        mFrom = between(sFrom, "<", ">");
      const theTmplt = eTmplts.find((element) => element.from == mFrom);
      if (theTmplt == undefined)
      {
        Logger.log(">>> !!! Неизвестный источник чека :" + sFrom + " Пропускаем письмо [" + sBody.length + "] от " + dDate.toISOString() + " >>> ");
        //ss.getSheetByName('DBG').getRange(1, 1).setValue(sBody);
        continue;
      }

      Logger.log( "Письмо " + thrd + "#" + ++m + " от " + dDate.toISOString() + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + sFrom + " ." );
      //let doc = XmlService.parse(between(sBody, '<body>', '</body>'));

      try {
        bBill = mailGenericGetInfo(theTmplt, sBody);
      } catch (err) {
        Logger.log(">>> !!! Ошибка чтения чека из письма.", err);
        continue;
      }
      Logger.log("Чек N " + ++NumBills + billInfoStr(bBill));

      bBill.URL = mURL;
      arrBills.push(bBill);
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых чеков. Последнее письмо от " + newLastMailDate.toISOString());
  return newLastMailDate;
}
