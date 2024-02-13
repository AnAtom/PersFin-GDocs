/*

ScanMail()
ScanDrive()

*/

function ScanMail(arrBills, ss)
{
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

  Logger.log("Добавлено " + NumBills + " чеков. Последнее письмо от " + newLastMailDate.toISOString());
  if (newLastMailDate > dLastMailDate)
    rLastMailDate.setValue(newLastMailDate);
}

function ScanDrive(arrBills, ss)
{
  const fDBG = ss.getRangeByName('ФлагОтладки').getValue();
  const rDBG = ss.getSheetByName('DBG').getRange(1, 1);

  // Читаем папку, в которой собраны чеки, из ячейки ЧекиДиск
  const folderId = Sheets.Spreadsheets.get(
    ss.getId(),
    {
      ranges: 'ЧекиДиск',
      fields: 'sheets.data.rowData.values.hyperlink'
    })
    .sheets[0]
    .data[0]
    .rowData[0]
    .values[0]
    .hyperlink.substring(39); // Отрезаем https://drive.google.com/drive/folders/
  const folderBills = DriveApp.getFolderById(folderId);
  Logger.log("Читаем чеки на диске из папки: " + folderBills.getName() + " Id: " + folderId);
  const bFolders = folderBills.getFolders();

  // Читаем дату последнего обработанного файла с чеком
  const rLastDriveDate = ss.getRangeByName('ДатаЧекДиск');
  let dLastDriveDate = rLastDriveDate.getValue();
  const sLastDriveDate = dLastDriveDate.toString();
  if (sLastDriveDate == "") {
    dLastDriveDate = ss.getRangeByName('ДатаЧек0').getValue();
    Logger.log("Принимаем дату последнего чека на диске: " + dLastDriveDate.toString());
  } else
    Logger.log("Дата последнего чека на диске: " + sLastDriveDate);

  let newLastDriveDate = dLastDriveDate;
  let NumBills = 0;
  let bBill = {};

  // Сканируем вложенные папки
  while (bFolders.hasNext()) {
    let bFolder = bFolders.next();
    const nMonth = bFolder.getName().slice(3);
    
    Logger.log("Папка " + nMonth);

    let bFiles = bFolder.getFiles();
    while (bFiles.hasNext()) {
      let fBill = bFiles.next();
      let bFileDate = fBill.getDateCreated();
      if (bFileDate > dLastDriveDate) {
        if (bFileDate > newLastDriveDate)
          newLastDriveDate = bFileDate;
      } else continue;

      let sBill = fBill.getBlob().getDataAsString();
      if (sBill == undefined) continue;

    } // цикл файлов в папке
  } // цикл вложенных папок по месяцам

}

function onOnceAnHour()
{
  Logger.log('Обрабатываем последние чеки.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fDBG = ss.getRangeByName('ФлагОтладки').getValue();
  const rDBG = ss.getSheetByName('DBG').getRange(1, 1);

  let newBills = [];

  ScanDrive(newBills, ss);

  ScanMail(newBills, ss);

  if (newBills.length > 0) {
    Logger.log('Сохраняем ' + newBills.length + ' чеков.');
    // Сортировка
    if (newBills.length > 1)
      newBills.sort((a, b) => {return a.dtime - b.dtime});

    const sBills = ss.getSheetByName('Чеки');
    let n = ss.getRangeByName('НомерЧек').getValue();
    let k = 2;
    for (bill of newBills) {
      bill.number = ++n;
      let vals = [[bill.number, bill.sdate, bill.total, bill.cache, bill.fn, bill.fd, bill.fp, bill.name]];
      sBills.getRange(3, 1, 1, 8).setValues(vals);
      sBills.insertRowBefore(3);
    }

    for (let i = 0; i < newBills.length; i++) {
      let bill = newBills[i];
      Logger.log(bill);
    }
  }
}
