/*

ScanDrive(ss, dLastDriveDate, arrBills)

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

  const monthToday = ss.getRangeByName('ДатаСегодня').getValue().getMonth();
  let monthPrev = dLastDriveDate.getMonth();
  if (dLastDriveDate.getDate() < ss.getRangeByName('ДнейРетроЧекДиск').getValue()) {
    monthPrev--;
    if (monthPrev < 0) monthPrev = 0;
  }

  let newLastDriveDate = dLastDriveDate;
  let NumBills = 0;

  // Сканируем вложенные папки
  const bFolders = folderBills.getFolders();
  while (bFolders.hasNext()) {
    let bFolder = bFolders.next();
    const nMonth = bFolder.getName().slice(3);
    const iMonth = MonthNum(nMonth);
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

      let bBill = billAllInfo(sBill);
      bBill.URL = fBill.getUrl();
      arrBills.push(bBill);
      Logger.log("Чек N " + ++NumBills + billInfoStr(bBill));
    } // цикл файлов в папке
  } // цикл вложенных папок по месяцам

  Logger.log("Считано " + NumBills + " новых чеков. Последний файл от " + newLastDriveDate.toISOString());
  return newLastDriveDate;
}
