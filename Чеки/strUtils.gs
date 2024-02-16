/*

between(str, start, end) - Возвращает подстроку между двумя строками
between2(str, start1, end1, start2, end2) - Возвращает подстроку из подстроки
between2from(sStr, fPos, start1, end1, start2, end2) - Возвращает подстроку из подстроки после заданной позиции

dbgSplitLongString(sStr, maxLngth) - Разбиваем длинную строку ( >50000 ) на массив строк по maxLngth символов

*/

//Возвращает подстроку между двумя строками
function between(str, start, end) {
  const zs = '';
  let startAt = str.indexOf(start);
  if (startAt == -1)
    return zs;
  startAt += start.length;
  const endAt = str.indexOf(end, startAt);
  if (endAt == -1)
    return zs;
  return str.slice(startAt, endAt);
}

// Поиск подстроки в подстроке
function between2(str, start1, end1, start2, end2) {
  const zs = '';
  let startAt1 = str.indexOf(start1);
  if (startAt1 == -1)
    return zs;
  startAt1 += start1.length;
  const endAt1 = str.indexOf(end1, startAt1);
  if (endAt1 == -1)
    return zs;
  const s = str.slice(startAt1, endAt1);
  let startAt2 = s.indexOf(start2);
  if (startAt2 == -1)
    return zs;
  startAt2 += start2.length;
  const endAt2 = s.indexOf(end2, startAt2);
  if (endAt2 == -1)
    return zs;
  return s.slice(startAt2, endAt2).trim();
}

// Поиск подстроки в строке начиная с позиции
function cutfrom(sStr, fPos, start, end) {
  const zs = '';
  let startAt = sStr.indexOf(start, fPos);
  if (startAt == -1)
    return zs;
  startAt += start.length;
  const endAt = sStr.indexOf(end, startAt);
  if (endAt == -1)
    return zs;
  return sStr.slice(startAt, endAt);
}

// Поиск подстроки в подстроке начиная с позиции
function between2from(sStr, fPos, start1, end1, start2, end2) {
  const zs = '';
  let startAt1 = sStr.indexOf(start1, fPos);
  if (startAt1 == -1)
    return zs;
  startAt1 += start1.length;
  const endAt1 = sStr.indexOf(end1, startAt1);
  if (endAt1 == -1)
    return zs;
  const s = sStr.slice(startAt1, endAt1);
  let startAt2 = s.indexOf(start2);
  if (startAt2 == -1)
    return zs;
  startAt2 += start2.length;
  const endAt2 = s.indexOf(end2, startAt2);
  if (endAt2 == -1)
    return zs;
  return s.slice(startAt2, endAt2).trim();
}

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
