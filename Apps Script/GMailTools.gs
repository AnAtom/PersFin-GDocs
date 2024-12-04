/*

 CutByTemplate - вырезает значение из сообщения по шаблону
 mailGenericGetInfo - универсальная процедура парсинга сообщения по шаблону
 GetTemplates - читает шаблоны для парсинга чеков в почте от различных ОФД
 ScanMail - читает чеки из новых писем

*/

function CutByTemplate(str, tmplt) {
  switch(tmplt.pt) {
    case 0: return between(str, tmplt.s1, tmplt.e1).trim();
    case 1: return between2(str, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2).trim();
    case 2:
        let s = between2(str, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
        return s.slice(s.indexOf(">")+1).trim();
    case 11: return "";
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
}

/*
 function CutFromPosByTemplate(str, pos, tmplt) {
  switch(tmplt.pt) {
    case 0: return cutfrom(str, pos, tmplt.s1, tmplt.e1);
    case 1: return between2from(str, pos, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
    case 2:
        let s = between2from(str, pos, tmplt.s1, tmplt.e1, tmplt.s2, tmplt.e2);
        return s.slice(s.indexOf(">")+1);
  }
  Logger.log("Неизвестный Preprocessing Type " + tmplt.pt + " ");
  return "";
 }
*/
function getDateByTemplate(email, tmplt) {
  let s = '';
  let l = 2;
  switch (tmplt.pt) {
    case 'D': l += 3;
    case 'd':
      let di = email.indexOf(tmplt.s1) + tmplt.s1.length;
      const ds = tmplt.s2;
      const dl = ds.length;
      for (let ii = 0; ii < l; ii++)
        di = email.indexOf(ds, di)+dl;
      if (tmplt.pt === 'D')
        di = email.indexOf('>', di) + 1;
      s = email.slice(di, email.indexOf(tmplt.e2, di));
      break;
    default:
      s = CutByTemplate(email, tmplt)
          .replace(" | ", " ")
  }
  if (tmplt.pt === 'd')
    s = s.trim();
  return s.replace(".202", ".2");
}

function mailGenericGetInfo(mailTmplt, email) {
  // Вырезаем имя
  let sName = CutByTemplate(email, mailTmplt.name)
              .replace(/&quot;/g, '"');
  // Убираем обрамляющие кавычки
  if (sName[0] == '"')
    sName = cutOuterQuotes(sName);

  const sDate = getDateByTemplate(email, mailTmplt.date);
  const isoDate = "20" + sDate.slice(6, 8)  // Год
    + "-" + sDate.slice(3, 5)         // месяц
    + "-" + sDate.slice(0, 2)         // день
  
  const isoDateTime = isoDate + "T" + sDate.slice(9);
  let aBill = billDate(isoDateTime);

  const nSumm = CutByTemplate(email, mailTmplt.total).replace(/\s/g,'').replace(",", ".") * 1.0;
  aBill.summ = nSumm;

  let nCash = CutByTemplate(email, mailTmplt.cash);
  if (nCash === "") nCash = 0;
  else nCash = nCash * 1.0;
  aBill.cash = nCash;
 /*
  const iFN = parseInt(CutByTemplate(email, mailTmplt.fn));
  const iFD = parseInt(CutByTemplate(email, mailTmplt.fd));
  const iFP = parseInt(CutByTemplate(email, mailTmplt.fp));
 */

  aBill.name = sName;
  aBill.shop = billFilterName(sName);

  return aBill;
  //return {dTime: tDate.getTime(), tDate: tDay.getTime(), date: sDate, summ: nSumm, cash: nCash, name: sName, shop: billFilterName(sName)};
}

/* Структура шаблона
 From	- адрес с которого пришел чек
 Key	- Проверочная строка, которая должна присутствовать в чеке

 Name pt	- Тип процедуры вырезки значения названия магазина
 Name s1		Name e1		- строки начала и окончания первого уровня вырезки
 Name s2		Name e2		- строки начала и окончания второго уровня вырезки

 Date pt		Date s1		Date e1		Date s2		Date e2		- то же для даты чека
 Total s1...	- то же для суммы чека
 Cash s1...		- то же для суммы наличными

 FN s1...		- то же для ФН
 FD s1...		- то же для ФД
 FP s1...		- то же для ФПД
*/

function GetTemplates(rTemplates) {
  // Читаем значения полей шаблона из диапазона
  const v = rTemplates.getValues();
  let Tmplts = [];

  for (let i = 0; i < rTemplates.getNumColumns(); i++)
  {
    let j = 0;
    const sFrom = v[j++][i];
    const sKey = v[j++][i];
    const tName = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tDate = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tTotal = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tCash = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
 /*
    const tFN = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFD = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
    const tFP = {pt: v[j++][i], s1: v[j++][i], e1: v[j++][i], s2: v[j++][i], e2: v[j++][i]};
 */
    Tmplts.push({
      from: sFrom,
      key: sKey,
      name: tName,
      date: tDate,
      total: tTotal,
      cash: tCash
 //     fn: tFN, fd: tFD, fp: tFP
    });
  }
  return Tmplts;
}

/* Яндекс

 [{"_id":"6713afe1b206204a5ff995a3","createdAt":"2024-10-19T13:10:57+00:00","ticket":{"document":{"receipt":
 {"buyerPhoneOrAddress":"+79057685271","cashTotalSum":0,"code":3,"creditSum":0,

 "dateTime":"2024-10-19T03:15:00",
 "ecashTotalSum":55400,
 "fiscalDocumentFormatVer":4,"fiscalDocumentNumber":211161,"fiscalDriveNumber":"7386440800040048","fiscalSign":3663930572,

 "fnsUrl":"www.nalog.gov.ru","internetSign":1,

 "items":[
  {"name":"Перевозка пассажиров и багажа","nds":6,"paymentAgentByProductType":64,"paymentType":4,"price":55400,"productType":1,"providerInn":"504207820709","quantity":1,"sum":55400}

 ],"kktRegId":"0000840607026308    ","machineNumber":"whitespirit2f","nds0":0,"nds10":0,"nds10110":0,"nds18":0,"nds18118":0,"ndsNo":55400,"operationType":1,"prepaidSum":0,

 "properties":{"propertyName":"psp_payment_id","propertyValue":"payment_c9698b303b9347af89dfdb36bb4da522|authorization_0000"},
 "propertiesData":"ws:CICTKBVPRB","provisionSum":0,"requestNumber":877,"retailPlace":"taxi.yandex.ru",
 "retailPlaceAddress":"248926, Россия, Калужская обл., г. Калуга, проезд 1-й Автомобильный, дом 8","sellerAddress":"support@go.yandex.com","shiftNumber":233,"taxationType":1,"appliedTaxationType":1,

 "totalSum":55400,
 "user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"ЯНДЕКС.ТАКСИ\"","userInn":"7704340310  "}}}}]

*/

function billYandexGo(eMail) {
  const sTripDate = eMail.getSubject().slice(28);
  const spcPos = sTripDate.indexOf(" ");
  const spcPos2 = sTripDate.indexOf(" ", spcPos+2);
  const sDate = sTripDate.slice(spcPos2+1, sTripDate.indexOf(" г.", spcPos2+2)) + '-'
    + getMonthNum(sTripDate.slice(spcPos+1, spcPos2)).toString().padStart(2, '0') + '-'
    + sTripDate.slice(0, spcPos).padStart(2, '0');

  const fBody = eMail.getBody();

  var TripTime = between(fBody, "route__point-name", "</table>");
  //Logger.log("Yandex Go>>>" + TripTime + "<<<");
  var j = TripTime.indexOf("route__point-name");
  TripTime = TripTime.slice(j+1);
  TripTime = between(TripTime, "<p class=", "</p>");
  j = TripTime.indexOf(">");
  TripTime = TripTime.slice(j+1).trim();

  const TripDateTime = sDate + "T" + TripTime;
  let aBill = billDate(TripDateTime);

  const TripSumm = between2(fBody, "report__value_main", "</td>", ">", " ₽").trim();
  //
  aBill.summ = TripSumm;
  aBill.cash = 0.0;
  aBill.name = 'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "ЯНДЕКС.ТАКСИ"';
  aBill.shop = 'ЯНДЕКС.ТАКСИ';
  return aBill;
}

/* UBER

 [{"_id":"673190cc8359dbbcb6c3f179","createdAt":"2024-11-11T05:06:20+00:00","ticket":{"document":{"receipt":
 {"buyerPhoneOrAddress":"+79057685271","cashTotalSum":0,"code":3,"creditSum":0,

 "dateTime":"2024-11-10T06:19:00",
 "ecashTotalSum":101200,
 "fiscalDocumentFormatVer":4,"fiscalDocumentNumber":136327,"fiscalDriveNumber":"7380440801186965","fiscalSign":744264270,

 "fnsUrl":"www.nalog.gov.ru","internetSign":1,

 "items":[

  {"name":"Перевозка пассажиров и багажа","nds":6,"paymentAgentByProductType":64,"paymentType":4,"price":101200,"productType":1,"providerInn":"051302118203","quantity":1,"sum":101200}

 ],"kktRegId":"0000840547059265    ","machineNumber":"whitespirit2f","nds0":0,"nds10":0,"nds10110":0,"nds18":0,"nds18118":0,"ndsNo":101200,"operationType":1,"prepaidSum":0,

 "properties":{"propertyName":"psp_payment_id","propertyValue":"payment_3f8aa5a15e89465680f9510986ad40fd|authorization_0000"},
 "propertiesData":"ws:CNUJGVSRPH","provisionSum":0,"requestNumber":968,"retailPlace":"https://support-uber.com",
 "retailPlaceAddress":"248926, Россия, Калужская обл., г. Калуга, проезд 1-й Автомобильный, дом 8","sellerAddress":"support@support-uber.com","shiftNumber":137,"taxationType":1 ,"appliedTaxationType":1,

 "totalSum":101200,
 "user":"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"ЯНДЕКС.ТАКСИ\"","userInn":"7704340310  "}

}}}]

*/

function billUBER(eMail) {
  const sTripDate = eMail.getSubject().slice(23);
  const spcPos = sTripDate.indexOf(" ");
  const spcPos2 = sTripDate.indexOf(" ", spcPos+2);

  const sDate = sTripDate.slice(spcPos2+1, sTripDate.indexOf(" г.", spcPos2+2)) + '-'
    + getMonthNum(sTripDate.slice(spcPos+1, spcPos2)).toString().padStart(2, '0') + '-'
    + sTripDate.slice(0, spcPos).padStart(2, '0');

  const fBody = eMail.getBody();

  let TripTime = between2(fBody, "From", "</tr>", "<td align", "</td>");
  TripTime = TripTime.slice(TripTime.indexOf(">")+1).trim();

  const TripDateTime = sDate + "T" + TripTime;
  let aBill = billDate(TripDateTime);

  const TripSumm = between2(fBody, "check__price", "</td>", ">", " ₽").trim();
  //
  aBill.summ = TripSumm;
  aBill.cash = 0.0;
  aBill.name = 'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "ЯНДЕКС.ТАКСИ"';
  aBill.shop = 'ЯНДЕКС.ТАКСИ';
  return aBill;
  //return {dTime: tDate.getTime(), tDate: tDay.getTime(), date: sDate, summ: nSumm, cash: 0.0, name: 'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "ЯНДЕКС.ТАКСИ"', shop: 'ЯНДЕКС.ТАКСИ'};
}

/*

  const sDate = getDateByTemplate(email, mailTmplt.date);
  const sd = "20" + sDate.slice(6, 8)  // Год
    + "-" + sDate.slice(3, 5)         // месяц
    + "-" + sDate.slice(0, 2)         // день

  const tDate = new Date(sd + "T" + sDate.slice(9)); // Дата чека
  const tDay = new Date(sd + "T00:00:00"); // Дата дня чека

  // Logger.log( "дата ["+ sd + "T" + sDate.slice(9) +"] data " + tDate + " день {" + sd + "T00:00:00" + "} day " + tDay);

  return {

 dTime: tDate.getTime(), 

 tDate: tDay.getTime(), 

 date: sDate, 

 summ: nSumm, cash: nCash, name: sName, shop: billFilterName(sName)};

 "dateTime":"2024-11-18T22:05:00"

  const sDate = sBill.slice(i, sBill.indexOf("\"", i+1));
  const dDate = new Date(sDate);
  const aDay = new Date(sDate.slice(0, sDate.indexOf("T")) + "T00:00:00");
  
*/

//  return {dTime: tDate.getTime(), tDate: tDay.getTime(), date: sDate, summ: nSumm, cash: nCash, name: sName, shop: billFilterName(sName)};
//  return {dTime: dDate.getTime(), tDate: aDay.getTime(), date: sDate, summ: iSumm / 100.0, cash: iCash / 100.0, name: sName, shop: sShop};

function billAli(eMail) {
  //
}
