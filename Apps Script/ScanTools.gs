/*

ScanDrive
ScanMail
ScanTaxi
ScanUBER
ScanYandexGo
ScanAli

*/

class MailLabelScaner {

  constructor(sName) {
    this.sName = sName;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const NameLabel = ss
      .getRangeByName('Чеки' + sName)
      .getValue();
    Logger.log('Читаем [' + sName + '] из метки ' + NameLabel);
    this.mailThread = GmailApp
      .getUserLabelByName(NameLabel)
      .getThreads();
    Logger.log('Для [' + sName + '] в метке ' + NameLabel + ' ' + this.mailThread.length + ' цепочек.');

    this.newDate = ss.getRangeByName('День1').getValue();
    this.rLastDate = ss.getRangeByName('ДатаЧек' + sName);
    this.dLastDate = this.rLastDate.getValue();
    if (this.dLastDate === '') {
      this.dLastDate = this.newDate;
      Logger.log('Принимаем последнюю дату для [' + sName + '] : ' + this.dLastDate);
    } else
      Logger.log('Последняя дата для [' + sName + '] : ' + this.dLastDate);
  }

  doScan(readBill, arrBills) {
    let mURL = '';
    let newBill = {};
    for (const messages of this.mailThread) {
      if (messages.getLastMessageDate() > this.dLastDate)
        mURL = messages.getPermalink();
      else
        continue;
      let m = 0;
      for (const message of messages.getMessages()) {
        const dDate = message.getDate();
        if (dDate > this.dLastDate) {
          if (dDate > this.newDate)
            this.newDate = dDate;
        } else
          continue;

        newBill = readBill(message);
        arrBills.push(newBill);
      }
    }
  }

  updateDate() {
    if (this.newDate > this.dLastDate) {
      Logger.log('Новая последняя дата для [' + this.sName + '] : ' + this.dLastDate);
      this.rLastDate.setValue(this.newDate);
    }
  }

}

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
