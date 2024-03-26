/*

mailGetThreadByRngName - возвращает цепочки писем из метки
CutByTemplate - вырезает значение из сообщения по шаблону
mailGenericGetInfo - универсальная процедура парсинга сообщения по шаблону
GetTemplates - читает шаблоны для парсинга чеков в почте от различных ОФД
ScanMail - читает чеки из новых писем

*/

function mailGetThreadByRngName(rName)
{
  // Возвращает цепочки писем из метки
  const sLabel = SpreadsheetApp
    .getActiveSpreadsheet()
    .getRangeByName(rName)
    .getValue();
  Logger.log("Читаем из почты с меткой: " + sLabel);

  return GmailApp.getUserLabelByName(sLabel).getThreads();
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
/*
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
*/
function getDateByTemplate(email, tmplt)
{
  let s = '';
  let l = 2;
  switch (tmplt.pt) {
    case 'D': l += 3;
    case 'd':
      let di = email.indexOf(tmplt.s1) + tmplt.s1.length;
      const ds = tmplt.s2;
      const dl = ds.length;
      for (let ii = 0; ii < l; ii++)
        di = email.indexOf(ds, di)+dl;
      if (tmplt.pt === 'D')
        di = email.indexOf('>', di) + 1;
      s = email.slice(di, email.indexOf(tmplt.e2, di));
      break;
    default:
      s = CutByTemplate(email, tmplt)
          .replace(" | ", " ")
  }
  if (tmplt.pt === 'd')
    s = s.trim();
  return s.replace(".202", ".2");
}

function mailGenericGetInfo(mailTmplt, email)
{
  // Вырезаем имя
  let sName = CutByTemplate(email, mailTmplt.name)
              .replace(/&quot;/g, '"');
  // Убираем обрамляющие кавычки
  if (sName[0] == '"')
    sName = cutOuterQuotes(sName);

  const sDate = getDateByTemplate(email, mailTmplt.date);
  const sd = "20" + sDate.slice(6, 8)  // Год
    + "-" + sDate.slice(3, 5)         // месяц
    + "-" + sDate.slice(0, 2)         // день
  const tDate = new Date(sd + "T" + sDate.slice(9)); // Дата чека
  const tDay = new Date(sd + "T00:00:00"); // Дата дня чека
  // Logger.log( "дата ["+ sd + "T" + sDate.slice(9) +"] data " + tDate + " день {" + sd + "T00:00:00" + "} day " + tDay);

  const nSumm = CutByTemplate(email, mailTmplt.total).replace(/\s/g,'').replace(",", ".") * 1.0;

  let nCash = CutByTemplate(email, mailTmplt.cash);
  if (nCash === "") nCash = 0;
  else nCash = nCash * 1.0;
/*
  const iFN = parseInt(CutByTemplate(email, mailTmplt.fn));
  const iFD = parseInt(CutByTemplate(email, mailTmplt.fd));
  const iFP = parseInt(CutByTemplate(email, mailTmplt.fp));
*/
  return {dTime: tDate.getTime(), tDate: tDay.getTime(), date: sDate, summ: nSumm, cash: nCash, name: sName, shop: billFilterName(sName)};
}

/* Структура шаблона
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
*/
    });
  }
  Logger.log("Загружено " + Tmplts.length + " шаблонов.");
  return Tmplts;
}

function ScanMail(ss, dLastMailDate, arrBills)
{
  // Читаем шаблоны для сканера
  const eTmplts = GetTemplates(ss.getRangeByName('ШаблоныЧеков'));

  let newLastMailDate = dLastMailDate;
  let NumBills = 0;
  let bBill = {};

  // Сканируем цепочки писем
  let thrd = 1;
  const mailThreads = mailGetThreadByRngName('ЧекиПочта');
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
      let mFrom = sFrom;
      if (~sFrom.indexOf("<"))
        mFrom = between(sFrom, "<", ">");
      const theTmplt = eTmplts.find((element) => element.from == mFrom);
      if (theTmplt == undefined)
      {
        Logger.log(">>> !!! Неизвестный источник чека :" + sFrom + " Пропускаем письмо [" + sBody.length + "] от " + dDate.toISOString() + " >>> ");
        // ss.getSheetByName('DBG').getRange(1, 1).setValue(sBody);
        continue;
      }

      Logger.log( "Письмо " + thrd + "#" + ++m + " от " + dDate.toISOString() + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + sFrom + " ." );

      //try {
        bBill = mailGenericGetInfo(theTmplt, sBody);
      /*} catch (err) {
        Logger.log(">>> !!! Ошибка чтения чека из письма.", err);
        continue;
      }*/
      arrBills.push(bBill);
      Logger.log("Чек N " + ++NumBills + dbgBillInfo(bBill));
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых чеков. Последнее письмо от " + newLastMailDate.toISOString());
  return newLastMailDate;
}
