/*

between(str, start, end)
between2(str, start1, end1, start2, end2)
between2from(sStr, fPos, start1, end1, start2, end2)

*/

//Возвращает подстроку между двумя строками
function between(str, start, end) {
  let startAt = str.indexOf(start);
  if (startAt == -1)
    return undefined;
  startAt += start.length;
  const endAt = str.indexOf(end, startAt);
  if (endAt == -1)
    return undefined;
  return str.slice(startAt, endAt);
}

// Поиск подстроки в подстроке
function between2(str, start1, end1, start2, end2) {
  let startAt1 = str.indexOf(start1);
  if (startAt1 == -1)
    return undefined;
  startAt1 += start1.length;
  const endAt1 = str.indexOf(end1, startAt1);
  if (endAt1 == -1)
    return undefined;
  const s = str.slice(startAt1, endAt1);
  let startAt2 = s.indexOf(start2);
  if (startAt2 == -1)
    return undefined;
  startAt2 += start2.length;
  const endAt2 = s.indexOf(end2, startAt2);
  if (endAt2 == -1)
    return undefined;
  return s.slice(startAt2, endAt2).trim();
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
