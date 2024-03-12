/*

GetTemplates - читает шаблоны для парсинга чеков в почте от различных ОФД
CutByTemplate - вырезает значение из сообщения по шаблону
getDateTime - читает дату из строки и возвращает UNIX время
mailGenericGetInfo - универсальная процедура парсинга сообщения по шаблону
ScanMail - читает чеки из новых писем

*/

/*

From	- адрес с которого пришел чек
Key		- Проверочная строка, которая должна присутствовать в чеке

Name pt		- Тип процедуры вырезки значения названия магазина
Name s1		Name e1		- строки начала и окончания первого уровня вырезки
Name s2		Name e2		- строки начала и окончания второго уровня вырезки

Date pt		Date s1		Date e1		Date s2		Date e2		- то же для даты чека
Total s1...		- то же для суммы чека
Cash s1...		- то же для суммы наличными
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
    const tCash = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
/*
    const tFN = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFD = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFP = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};

    const sItemsFrom = v[j++][i];
    const sItemFrom = v[j++][i];
    const tIName = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tIQuantity = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tIPrice = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tISum = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
*/
    Tmplts.push({
      from: sFrom,
      key: sKey,
      name: tName,
      date: tDate,
      total: tTotal,
      cash: tCash
/*
      fn: tFN,
      fd: tFD,
      fp: tFP,

      items: sItemsFrom,
      item: sItemFrom,
      iname: tIName,
      iqntty: tIQuantity,
      iprice: tIPrice,
      isum: tISum
*/
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
    case 11: return "";
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
}

function CutFromPosByTemplate(str, pos, tmplt)
{
  switch(tmplt.pt) {
    case 0: return cutfrom(str, pos, tmplt.s1, tmplt.e1);
    case 1: return between2from(str, pos, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
    case 2:
        let s = between2from(str, pos, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
        return s.slice(s.indexOf(">")+1);
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
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
    sName = cutOuterQuotes(sName);

  const sShop = billFilterName(sName);

  const sDate = CutByTemplate(email, mailTmplt.date)
    .replace(" | ", " ")
    .replace(".202", ".2");
  const dDate = getDate(sDate);

  const sSumm = CutByTemplate(email, mailTmplt.total) / 1.0;

  let sCash = CutByTemplate(email, mailTmplt.cash);
  if (sCash == "") sCash = 0.0;
  else sCash = sCash / 1.0;
/*
  const iFN = parseInt(CutByTemplate(email, mailTmplt.fn));
  const iFD = parseInt(CutByTemplate(email, mailTmplt.fd));
  const iFP = parseInt(CutByTemplate(email, mailTmplt.fp));

  let arrItems = [];
  let i = email.indexOf(mailTmplt.items);
  if (~i) {
    i += mailTmplt.items.length;

    let n = 1;
    let iName = "";
    let sQuantity = "";
    let iUnit = "";
    let iQuantity = 0;
    let iPrice = 0;
    let iSum = 0;
    let j = email.indexOf(mailTmplt.item, i);
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
      iPrice = CutFromPosByTemplate(email, j, mailTmplt.iprice) / 1.0;
      iSum = CutFromPosByTemplate(email, j, mailTmplt.isum) / 1.0;

      arrItems.push({name: iName, price: iPrice, quantity: iQuantity, sum: iSum, unit: iUnit});
      i = j + mailTmplt.item.length;
      j = email.indexOf(mailTmplt.item, i);
    }
  } */
  //const jBill = {cashTotalSum: sCach, dateTime: dDate, fiscalDriveNumber: sFN / 1.0, fiscalDocumentNumber: sFD / 1.0, fiscalSign: sFP / 1.0,
  //                items: arrItems, totalSum: sSumm, user: sName}
  //return {dTime: dDate.getTime(), SN: 0, URL: "", jsonBill: jBill};
  return {dTime: dDate.getTime(), date: sDate, summ: sSumm, cash: sCash, name: sName, shop: sShop};
}

function ScanMail(ss, dLastMailDate, arrBills)
{
  // Читаем метку, под которой собраны чеки, из ячейки ЧекиПочта
  const sLabel = ss.getRangeByName('ЧекиПочта').getValue();
  Logger.log("Читаем чеки из почты с меткой: " + sLabel);

  // Читаем шаблоны для сканера
  const eTmplts = GetTemplates(ss.getRangeByName('ШаблоныЧеков'));

  let newLastMailDate = dLastMailDate;
  let NumBills = 0;
  let bBill = {};

  // Сканируем цепочки писем
  const mailThreads = GmailApp.getUserLabelByName(sLabel).getThreads();
  let thrd = 1;
  for (messages of mailThreads) {
    if (!messages.getLastMessageDate() > dLastMailDate)
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
      const mFrom = between(sFrom, "<", ">");
      /* if (mFrom == "noreply@chek.pofd.ru") {
        Logger.log("Новый чек");
      } */
      const theTmplt = eTmplts.find((element) => element.from == mFrom);
      if (theTmplt == undefined)
      {
        Logger.log(">>> !!! Неизвестный источник чека :" + sFrom + " Пропускаем письмо [" + sBody.length + "] от " + dDate.toISOString() + " >>> ");
        // ss.getSheetByName('DBG').getRange(1, 1).setValue(sBody);
        continue;
      }

      Logger.log( "Письмо " + thrd + "#" + ++m + " от " + dDate.toISOString() + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + sFrom + " ." );

      try {
        bBill = mailGenericGetInfo(theTmplt, sBody);
      } catch (err) {
        Logger.log(">>> !!! Ошибка чтения чека из письма.", err);
        continue;
      }
      arrBills.push(bBill);
      Logger.log("Чек N " + ++NumBills + dbgBillInfo(bBill));
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых чеков. Последнее письмо от " + newLastMailDate.toISOString());
  return newLastMailDate;
}
