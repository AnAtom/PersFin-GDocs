/*



*/

function MonthNum(sMonth)
{
  switch(sMonth) {
    case 'Январь':  return 0;
    case 'Февраль': return 1;
    case 'Март': return 2;
    case 'Апрель': return 3;
    case 'Май': return 4;
    case 'Июнь': return 5;
    case 'Июль': return 6;
    case 'Август': return 7;
    case 'Сентябрь': return 8;
    case 'Октябрь': return 9;
    case 'Ноябрь': return 10;
    case 'Декабрь': return 11;
    default: return -1;
  }
}

function ScanDrive(ss, dLastDriveDate, arrBills)
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

  let monthPrev = dLastDriveDate.getMonth();
  const daysPrevMonth = ss.getRangeByName('ДнейРетроЧекДиск').getValue();
  if (daysPrevMonth > 0) {
    monthPrev--;
    if (monthPrev < 0) monthPrev = 0;
  }

  const monthToday = ss.getRangeByName('ДатаСегодня').getValue().getMonth();
  let monthBefore = monthToday - 1;
  if (monthBefore < 0) monthBefore = 0;

  let newLastDriveDate = dLastDriveDate;
  let NumBills = 0;
  let bBill = {};

  // Сканируем вложенные папки
  while (bFolders.hasNext()) {
    let bFolder = bFolders.next();
    const nMonth = bFolder.getName().slice(3);
    const iMonth = MonthNum(nMonth);
    // Пропускаем будушие месяцы и месяцы предшествующие предпоследнему обработанному
    if (iMonth > monthToday || iMonth < monthPrev) continue;
    
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

      let bBill = billInfo(sBill);
      arrBills.push(bBill);
      NumBills++;

    } // цикл файлов в папке
  } // цикл вложенных папок по месяцам

  Logger.log("Считано " + NumBills + " новых чеков. Последний файл от " + newLastDriveDate.toISOString());
  return newLastDriveDate;
}
