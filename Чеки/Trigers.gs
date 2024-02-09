/*



*/

function ScanMail()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fDBG = ss.getRangeByName('ФлагОтладки').getValue();
  const rDBG = ss.getSheetByName('DBG').getRange(1, 1);

  // Читаем дату последнего обработанного письма с чеком
  const rLastMailDate = ss.getRangeByName('ДатаЧекПочта');
  let dLastMailDate = rLastMailDate.getValue();
  const sLastMailDate = dLastMailDate.toString();
  if (sLastMailDate == "") {
    dLastMailDate = ss.getRangeByName('ДатаЧек0').getValue();
    Logger.log("Принимаем дату последнего чека в почте: " + dLastMailDate.toString());
  } else
    Logger.log("Дата последнего чека в почте: " + sLastMailDate);

  let newLastMailDate = dLastMailDate;
  let NumBills = 0;

  // Читаем метку, под которой собраны чеки, из ячейки ЧекиПочта
  const sLabel = ss.getRangeByName('ЧекиПочта').getValue();
  Logger.log("Читаем чеки из почты с меткой: " + sLabel);

  // Читаем шаблоны для сканера
  const eTmplts = GetTemplates(ss.getRangeByName('ШаблоныЧеков'));

  const mailThreads = GmailApp.getUserLabelByName(sLabel).getThreads();
  let thrd = 0;
  for (messages of mailThreads) {

    let m = 1;
    for (message of messages.getMessages()) {
      let dDate = message.getDate();
      if (dDate < dLastMailDate) continue;

      let sFrom = message.getFrom();
      let sFromMail = between(sFrom, "<", ">");
      let iTmplt = FindInTemplates(eTmplts, sFromMail);
      if (iTmplt == -1)
      {
        Logger.log("Неизвестный источник чека :" + sFrom);
        continue;
      }

      let sSubject = message.getSubject();
      let sBody = message.getBody();

      /*
      const arrS = dbgSplitLongString(sBody, 49000);
      let n = 0;
      for (S of arrS) rDBG.offset(l, n++).setValue(S);
      l++;
      */

      Logger.log( dDate.toISOString() + " e-Mail " + thrd + "#" + m++ + " > " + sSubject + " [[[ "+ sBody.length.toString() +" ]]] From: " + sFrom + " <" );

      /*
      if (sFromMail == "noreply@chek.pofd.ru")
        Logger.log(" Платформа " + l);
      if (sFromMail == "ofdreceipt@beeline.ru")
        Logger.log(" Beeline " + l);
      */

      let bBill = mailGenericGetInfo(eTmplts[iTmplt], sBody);
      NumBills++;

      Logger.log("Чек " + NumBills + " от (" + bBill.date + ") магазин >" + bBill.name + "< на сумму [" + bBill.total + "] р. наличными {" + bBill.cache 
        + "} ФН :" + bBill.fn + " ФД :" + bBill.fd + " ФП :" + bBill.fp);

      // ОФД Такском <noreply@taxcom.ru>
      // билайн ОФД <ofdreceipt@beeline.ru>
      // Платформа ОФД <noreply@chek.pofd.ru>

    } // Письма в цепочке
    thrd++;
  } // Цепочки писем

  Logger.log("Добавлено " + NumBills + " чеков. Последнее письмо от " + newLastMailDate.toISOString());

}

function onOnceAnHour()
{
  //
  Logger.log("Обрабатываем последние чеки.");
}
