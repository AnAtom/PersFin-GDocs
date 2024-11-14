/*

ScanDrive
ScanMail
ScanTaxi
ScanUBER
ScanYandexGo
ScanAli

*/

function ScanDrive(ss, dLastDriveDate, arrBills)
{
  // Читаем папку, в которой собраны чеки, из ячейки ЧекиДиск
  const folderId = GetGDriveFolderIdFromURL('ЧекиДиск');
  const folderBills = DriveApp.getFolderById(folderId);
  Logger.log("Читаем чеки на диске из папки: " + folderBills.getName() + " Id: " + folderId);
  const bFolders = folderBills.getFolders();

  const monthToday = ss
    .getRangeByName('Сегодня')
    .getValue()
    .getMonth();
  let monthPrev = dLastDriveDate.getMonth();
  if (dLastDriveDate.getDate() < ss.getRangeByName('ДнейРетроДиск').getValue()) {
    monthPrev--;
    if (monthPrev < 0) monthPrev = 0;
  }

  let newLastDriveDate = dLastDriveDate;
  let NumBills = 0;

  // Сканируем вложенные папки
  while (bFolders.hasNext()) {
    let bFolder = bFolders.next();
    const nMonth = bFolder.getName().slice(3);
    const iMonth = getMonthNum(nMonth, true);
    // Пропускаем будушие месяцы и месяцы предшествующие предпоследнему обработанному
    if (iMonth > monthToday || iMonth < monthPrev)
      continue;
    
    Logger.log("Папка " + nMonth);

    let aFiles = bFolder.getFiles();
    while (aFiles.hasNext()) {
      const fBill = aFiles.next();
      const bFileDate = fBill.getDateCreated();
      if (bFileDate > dLastDriveDate) {
        if (bFileDate > newLastDriveDate)
          newLastDriveDate = bFileDate;
      } else continue;

      const sBill = fBill.getBlob().getDataAsString();
      if (sBill == undefined) continue;

      const bBill = billInfo(sBill);
      arrBills.push(bBill);
      Logger.log("Чек N " + ++NumBills + dbgBillInfo(bBill));
    } // цикл файлов в папке
  } // цикл вложенных папок по месяцам

  Logger.log("Считано " + NumBills + " новых чеков. Последний файл от " + newLastDriveDate.toISOString());
  return newLastDriveDate;
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

function GetMailBill(mailBody)
{
  var bBill = {};
  //return bBill;
  /*
  const jBill = {cashTotalSum: sCash, dateTime: dDate, fiscalDriveNumber: iFN, fiscalDocumentNumber: iFD, fiscalSign: iFP,
                  items: arrItems, totalSum: sSumm, user: sName, userInn: 0}
  return {dTime: dDate.getTime(), SN: 0, URL: "", Shop: sShop, jsonBill: jBill};
  */

  //return {dTime: dDate.getTime(), tDate: aDay.getTime(), date: sDate, summ: iSumm / 100.0, cash: iCash / 100.0, name: sName, shop: sShop};
}

function ScanUberMail(sBody)
{
  //
}

function ScanYandexGoMail(sBody)
{
  //
}

function ScanAliMail(sBody)
{
  //
}

function ScanMailLabel(sLabelName, rLastDate, scanFunc, arrBills)
{
  var dLastMailDate = rLastDate.getValue();
  var newLastMailDate = dLastMailDate;

  // Сканируем цепочки писем
  const mailThreads = GmailApp.getUserLabelByName(sLabelName).getThreads();
  if (mailThreads == null) {
    Logger.log('В почте нет метки "' + sLabelName + '"');
    return [];
  }
  var newBill;
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
      newBill = scanFunc(sBody);
      arrBills.push(newBill);
    }
  }

  if (newLastMailDate > dLastMailDate)
    rLastDate.setValue(newLastMailDate);
}

function ScanTaxi(ss, dLastTaxiDate, arrBills)
{
  let newLastTaxiDate = dLastTaxiDate;
  let NumBills = 0;
  let bBill = {};

  const rLastMailDate = ss.getRangeByName('ДатаЧекиUBER');
  ScanMailLabel("Моё/Такси/Uber", rLastMailDate, GetMailBill)
  Logger.log("Считано " + NumBills + " новых поездок. Последняя поездка от " + newLastTaxiDate.toISOString());
  return newLastTaxiDate;
}

function TestScan()
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //
  ScanTaxi(ss, "", []);
  Logger.log("Тест пройден. ");
}
