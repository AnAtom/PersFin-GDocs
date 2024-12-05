/*

 class BillScaner
 class MailLabelScaner extends BillScaner
 ScanDrive

*/

class BillScaner {

  constructor(sName) {
    this.sName = sName;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.newDate = ss.getRangeByName('День1').getValue();
    this.rLastDate = ss.getRangeByName('ДатаЧек' + sName);
    this.dLastDate = this.rLastDate.getValue();
    if (this.dLastDate === '') {
      this.dLastDate = this.newDate;
      this.doLog('Принимаем последнюю дату : ' + this.dLastDate);
    } else
      this.doLog('Последняя дата : ' + this.dLastDate);
  }

  updateDate() {
    if (this.newDate > this.dLastDate) {
      this.doLog('Обновляем последнюю дату : ' + this.newDate);
      this.rLastDate.setValue(this.newDate);
    } else
      this.doLog('Дата не изменилась.');
  }

  doLog(msg) {
    Logger.log('> Сканер > [' + this.sName + '] : ' + msg);
  }

}

class MailLabelScaner extends BillScaner {

  constructor(sName) {
    super(sName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const NameLabel = ss
      .getRangeByName('Чеки' + sName)
      .getValue();
    this.doLog('Читаем из метки ' + NameLabel);
    this.mailThread = GmailApp
      .getUserLabelByName(NameLabel)
      .getThreads();
    this.doLog('В метке ' + NameLabel + ' ' + this.mailThread.length + ' цепочек.');
  }

  extractData(eMail) {
    return null;
  }

  doScan(readBill, arrBills) {
    let mt = 0;
    let mc = 0;
    let bc = 0;
    let newBill = {};
    for (const messages of this.mailThread) {
      if (messages.getLastMessageDate() > this.dLastDate) {
        for (const message of messages.getMessages()) {
          const dDate = message.getDate();
          if (dDate > this.dLastDate) {
            if (dDate > this.newDate)
              this.newDate = dDate;
          } else
            continue;

          const eData = this.extractData(message);
          if (eData == undefined)
            continue;

          newBill = readBill(message, eData);
          if (newBill != null) {
            bc++;
            arrBills.push(newBill);
          }
          mc++;
        }
        mt++;
      } else
        break;
    }
    this.doLog(bc + ' чеков в ' + mc + ' письмах из ' + mt + ' цепочек добавлено.');
  }

}

class MailTemplateScaner extends MailLabelScaner {

  constructor(sTmplts) {
    super('Почта');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.eTmplts = GetTemplates(ss.getRangeByName(sTmplts));
    this.doLog('Загружено ' + this.eTmplts.length + ' шаблонов.');
  }

  extractData(eMail) {
    const sFrom = eMail.getFrom();
    let mFrom = sFrom;
    if (~sFrom.indexOf("<"))
      mFrom = between(sFrom, "<", ">");
    const theTmplt = this.eTmplts.find((element) => element.from == mFrom);
    if (theTmplt == undefined)
      this.doLog("!!! Неизвестный источник чека :" + sFrom + " Пропускаем письмо !!! " + eMail.getSubject());

    return theTmplt;
  }

  readData(MSG, Tmplt) {
    return mailGenericGetInfo(Tmplt, MSG.getBody());
  }

}

class DriveBillsScaner extends BillScaner {

  constructor(sRetro) {
    super('Диск');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const nURL = 'ЧекиДиск';
    const folderId = Sheets.Spreadsheets.get(ss.getId(), {ranges: nURL, fields: 'sheets.data.rowData.values.hyperlink'})
      .sheets[0]
      .data[0]
      .rowData[0]
      .values[0]
      .hyperlink
    .substring(39);
    const folderBills = DriveApp.getFolderById(folderId);

    const NameFolder = folderBills.getName();
    this.doLog("Читаем чеки из папки: " + NameFolder + " Id: " + folderId);

    this.bFolders = folderBills.getFolders();

    this.monthToday = ss
      .getRangeByName('Сегодня')
      .getValue()
      .getMonth();

    const nRetro = ss.getRangeByName(sRetro).getValue();
    this.monthPrev = this.dLastDate.getMonth();
    if (this.dLastDate.getDate() < nRetro) {
      this.monthPrev--;
      if (this.monthPrev < 0) this.monthPrev = 0;
    }
  }

  doScan(arrBills) {
    //
    let NumBills = 0;
    let NumFiles = 0;

    // Сканируем вложенные папки
    while (this.bFolders.hasNext()) {
      const bFolder = this.bFolders.next();
      const nMonth = bFolder.getName().slice(3);
      const iMonth = getMonthNum(nMonth, true);
      // Пропускаем будушие месяцы и месяцы предшествующие предпоследнему обработанному
      if (iMonth > this.monthToday || iMonth < this.monthPrev)
        continue;
    
      this.doLog("Папка " + nMonth);

      let aFiles = bFolder.getFiles();
      while (aFiles.hasNext()) {
        const fBill = aFiles.next();
        NumFiles++;
        const bFileDate = fBill.getDateCreated();
        if (bFileDate > this.dLastDate) {
          if (bFileDate > this.newDate)
            this.newDate = bFileDate;
        } else continue;

        const sBill = fBill.getBlob().getDataAsString();
        if (sBill == undefined) continue;

        const bBill = billInfo(sBill);
        arrBills.push(bBill);
        this.doLog("Чек N " + ++NumBills + " в файле (" + NumFiles + ") " + fBill.getName() + dbgBillInfo(bBill));
      } // цикл файлов в папке
    } // цикл вложенных папок по месяцам

    this.doLog("Считано " + NumBills + " новых чеков из " + NumFiles + " файлов. Последний файл от " + this.newDate.toISOString());
  }

}

function ScanDrive(ss, dLastDriveDate, arrBills) {
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
