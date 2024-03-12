/*

findInRange(rangeName, s) - Поиск в именованном диапазоне.
findInRule(ruleName, s) - Поиск в списке выбора значения именованного диапазона.
getMonthNum(sMonth, capitalLetter) - Возвращает номер месяца по названию.
getMonthName(dDate) - Возвращает название месяца по дате
GetGDriveFolderIdFromURL(rng) - Достает URL из Именованной ячейки таблицы.
ReadLastDate(ss, rDate) - Читает дату из ячейки. Если ячейка пуста, то возвращает дату День1
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
function GetGDriveFolderIdFromURL(rng)
{
  const url = Sheets.Spreadsheets.get(
    SpreadsheetApp.getActiveSpreadsheet().getId(),
    {
      ranges: rng,
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

// Читает дату из ячейки. Если ячейка пуста, то возвращает дату День1.
function ReadLastDate(ss, rDate)
{
  let dLastDate = rDate.getValue();
  const sLastDate = dLastDate.toString();
  if (sLastDate == "") {
    dLastDate = ss.getRangeByName('День1').getValue();
    Logger.log("Принимаем дату последнего чека : " + dLastDate.toString());
  } else
    Logger.log("Дата последнего чека : " + sLastDate);
  return dLastDate;
}

