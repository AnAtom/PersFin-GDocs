/*

 findInRange(rangeName, s) - Поиск в именованном диапазоне.
 findInRule(ruleName, s) - Поиск в списке выбора значения именованного диапазона.
 getDateRangeDefault(rangeName) - Возвращает именованный диапазон с датой. Если дата пуста, то подставляет День1
 getMonthNum(sMonth, capitalLetter) - Возвращает номер месяца по названию.
 getMonthName(dDate) - Возвращает название месяца по дате
 GetGDriveFolderIdFromURL(rng) - Достает URL из Именованной ячейки таблицы.
 getShopInfoRemarkNote(sShop, sUser, lstStores, ssShop) - Возвращает Инфо-Примечание-Заметка для Магазина с листа Магазины или добавляет новый магазин на этот лист
 setCostBill(rSumm, bBill, arrInfoRemarkNote) - Выставляет в строке расходов информацию по чеку

*/

// Поиск в именованном диапазоне. Возвращает индекс строки в списке занчений именованного диапазона или -1
function findInRange(rangeName, s)
{
  const rangeValues = SpreadsheetApp.getActiveSpreadsheet()
  .getRangeByName(rangeName)
  .getValues();

  for(let i = 0; i<rangeValues.length; i++)
    if (rangeValues[i][0] == s)
      return i;

  return -1;
}

// Поиск в списке выбора значения. Возвращает индекс строки в списке занчений диапазона проверки данных или -1
function findInRule(ruleName, s)
{
  const rangeValues = SpreadsheetApp.getActiveSpreadsheet()
  .getRangeByName(ruleName)
  .getDataValidation()
  .getCriteriaValues()[0]
  .getValues();

  for(let i = 0; i<rangeValues.length; i++)
    if (rangeValues[i] == s)
      return i;

  return -1;
}

// Возвращает именованный диапазон с датой. Если дата пуста, то подставляет День1
function getDateRangeDefault(rangeName)
{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const r = ss.getRangeByName(rangeName);
  let d = r.getValue();
  if (d === "") {
    const Date0 = ss.getRangeByName('День1').getValue();
    Logger.log('Последняя дата "' + rangeName + '" не определена. Принимаем дату : ' + Date0);
    r.setValue(Date0);
  }
  return r;
}

// Присваивает ячейке именованный список значений из именованной ячейки
function SetTargetRule(ss, c, rn)
{
  const range = ss.getRangeByName(rn);

  if (range == undefined)
    return;

  const rule = range.getDataValidation();
  c.setDataValidation(rule);
}

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

// Возвращает номер месяца по названию
function getMonthNum(sMonth, capitalLetter)
{
  if (capitalLetter)
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
  else
    switch(sMonth) {
      case 'января': return 1;
      case 'февраля': return 2;
      case 'марта': return 3;
      case 'апреля': return 4;
      case 'мая': return 5;
      case 'июня': return 6;
      case 'июля': return 7;
      case 'августа': return 8;
      case 'сентября': return 9;
      case 'октября': return 10;
      case 'ноября': return 11;
      case 'декабря': return 11;
      default: return -1;
    }
}

// Возвращает название месяца по дате
function getMonthName(dDate)
{
  const names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return names[dDate.getMonth()];    
}

// Достает URL из Именованной ячейки таблицы.
function GetGDriveFolderIdFromURL(rgn)
{
  const url = Sheets.Spreadsheets.get(
    SpreadsheetApp.getActiveSpreadsheet().getId(),
    {
      ranges: rgn,
      fields: 'sheets.data.rowData.values.hyperlink'
    })
    .sheets[0]
    .data[0]
    .rowData[0]
    .values[0]
    .hyperlink;

  // Отрезаем https://drive.google.com/drive/folders/
  return url.substring(39);
}

function getShopInfoRemarkNote(sShop, sUser, lstStores, lstIgnore, ssShop)
{
  // Ищем магазин в списке
  const shop = lstStores.find((element) => element[3] == sShop);
  if (shop != undefined) // Нашли.
    return [shop[0], shop[1], shop[2]]; // Возвращаем Статья-Инфо-Примечание для этого магазина.

  const notFound = ["", "", ""];
  // Не нашли в известных
  if (~lstIgnore.findIndex((element) => element[0] == sShop)) // Нашли в игнорируемых
    return notFound;

  // Не нашли ни в известных, ни в игнорируемых. Добавляем в список новый магазин
  Logger.log("Новый магазин [" + sShop + "] (" + sUser + ")");
  const newRow = lstStores.length + 4;
  ssShop.insertRowBefore(newRow)
        .getRange(newRow, 4, 1, 2)
        .setValues([[sShop, sUser]]);
  return notFound;
}

function setCostBill(rSumm, bBill, arrInfoRemarkNote)
{
  // Выставляем сумму, Статья-Инфо-Примечание, дату и для времени покупки получаем адрес ячейки с датой
  const A1date = rSumm
    .setValue(bBill.summ) // Сумма
    .setNumberFormat("#,##0.00[$ ₽]")
    .offset(0, 2, 1, 3)
    .setValues([arrInfoRemarkNote]) // Статья-Инфо-Примечание
    .offset(0, -4, 1, 1)
    .setValue(bBill.date) // Дата
    .setNumberFormat("dd.mm")
    .getA1Notation();

  // Выставляем время покупки
  rSumm.offset(0,-1)
    .setFormula("=" + A1date)
    .setNumberFormat("HH:mm")
    .offset(0, 14)
    .setFormula("=НАЗВМЕС(" + A1date + ")");

  // Если наличные, то выставляем счет списания
  if (bBill.cash != 0)
    rSumm.offset(0, 1).setValue("Карман");
}
