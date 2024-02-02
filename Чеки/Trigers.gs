/*



*/

// Разбиваем длинную строку ( >50000 ) на несколько строк по maxLngth символов
function dbgSplitLongString(sStr, maxLngth)
{
  let n = 0;
  let k = maxLngth;
  let sArr = [];
  do {
    sArr.push(sStr.slice(n, k));
    n += maxLngth;
    k += maxLngth;
  } while (sStr.length > n);

  return sArr;
}

/*

From
Key

Name t
Name s1
Name e1
Name s2
Name e2

Date s1
Total s1
Cache s1
FN s1
FD s1
FP s1
Items s1
Item s1
iName s1
iQuantity s1
iPrice s1
iSum s1

*/

function GetTemplates()
{
  //
}

function ScanMail()
{
  //
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Читаем дату последнего обработанного письма с чеком
  const rLastBillDate = ss.getRangeByName('ДатаЧекПочта');
  let dLastBillDate = rLastBillDate.getValue();
  const sLastBillDate = dLastBillDate.toString();
  if (sLastBillDate == "") {
    //
    dLastBillDate = ss.getRangeByName('ДатаЧек0').getValue();
    Logger.log("Принимаем дату последнего чека в почте: " + dLastBillDate.toString());
  }
  else
    Logger.log("Дата последнего чека в почте: " + sLastBillDate);

  let newLastBillDate = dLastBillDate;
  let NumBills = 0;

  // Читаем метку, под которой собраны чеки, из ячейки ЧекиПочта
  const sLabel = ss.getRangeByName('ЧекиПочта').getValue();
  Logger.log("Читаем чеки из почты с меткой: " + sLabel);

  const mailThreads = GmailApp.getUserLabelByName(sLabel).getThreads();
  let i = mailThreads.length - 1;
  for (let i = 0; i < mailThreads.length; i++) {

    let messages = mailThreads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      let dDate = messages[j].getDate();

      let sFrom = messages[j].getFrom();
      let sFromMail = between(sFrom, "<", ">");
      let sSubject = messages[j].getSubject();
      let sBody = messages[j].getBody();
      if (j == 1) {
        let arrS = dbgSplitLongString(sBody, 49000);
        ss.getSheetByName('DBG').getRange(1, 1).setValue(arrS[0]);
      }

      //Logger.log( dDate.toISOString() + " e-Mail " + i + "#" + j + " > " + sSubject + " [[[ "+ sBody.length.toString() +" ]]] From: " + sFrom + " <" );

      // ОФД Такском <noreply@taxcom.ru>

      // билайн ОФД <ofdreceipt@beeline.ru>

      // Платформа ОФД <noreply@chek.pofd.ru>

    } // Письма в цепочке
    //

  } // Цепочки писем

}

function onOnceAnHour()
{
  //
  Logger.log("Обрабатываем последние чеки.");
}
