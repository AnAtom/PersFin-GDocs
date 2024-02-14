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

function ScanMail(ss, dLastMailDate, arrBills)
{
  const fDBG = ss.getRangeByName('ФлагОтладки').getValue();
  const rDBG = ss.getSheetByName('DBG').getRange(1, 1);

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
  for (messages of mailThreads) {
    let m = 0;
    for (message of messages.getMessages()) {
      const dDate = message.getDate();
      if (dDate > dLastMailDate) {
        if (dDate > newLastMailDate)
          newLastMailDate = dDate;
      } else
        continue;

      const sFrom = message.getFrom();
      const mFrom = between(sFrom, "<", ">");
      const theTmplt = eTmplts.find((element) => element.from == mFrom);
      if (theTmplt == undefined)
      {
        Logger.log("Неизвестный источник чека :" + sFrom + " Пропускаем письмо от " + dDate.toISOString());
        continue;
      }

      const sBody = message.getBody();

      Logger.log( dDate.toISOString() + " e-Mail " + thrd + "#" + ++m + " > " + message.getSubject() + " ["+ sBody.length +"] From: " + sFrom + " ." );
      //let doc = XmlService.parse(between(sBody, '<body>', '</body>'));

      try {
        bBill = mailGenericGetInfo(theTmplt, sBody);
      } catch (err) {
        Logger.log("Ошибка чтения чека из письма.", err);
        continue;
      }
      arrBills.push(bBill);
      NumBills++;

      Logger.log(
        "Чек N " + NumBills + 
        " от (" + bBill.sdate + 
        ") магазин >" + bBill.name + 
        "< на сумму [" + bBill.total + "] р. наличными {" + bBill.cache + 
        "} ФН :" + bBill.fn + 
        " ФД :" + bBill.fd + 
        " ФП :" + bBill.fp
      );
    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Считано " + NumBills + " новых чеков. Последнее письмо от " + newLastMailDate.toISOString());
  return newLastMailDate;
}
