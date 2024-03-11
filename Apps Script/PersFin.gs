/*

onEdit(e)
onOpen(e)
onOnceAnHour()

*/

// "8?>2K5 565<5AOG=K5 >?5@0F88
/*
  0B0  !C<<0   !G5B        &5;L        ?5@0F8O        "8?
  -     -       -           -           -               -
-	31.01																	!=OB85					1>@>B
-	30.01	0,00 ½	!15@										@>F5=BK :@4B		!?8A0=85
-	30.01	0,00 ½	!15@										>30H5=85 :@4B	!?8A0=85
-	30.01	0,00 ½	@548B "							@>F5=BK :@4B		!?8A0=85
-	30.01	0,00 ½						@548B "	>30H5=85 :@4B	!?8A0=85
-	26.01	0,00 ½						Rostelecom	;0B56					!?8A0=85
+	25.01	0,00 ½												20=A						0G8A;5=85
-	23.01	0,00 ½						><C=0;:0		;0B56					!?8A0=85
-	22.01	0,00 ½							8;0=0			><>3/?>40@8;		!?8A0=85
-	19.01	0,00 ½						20@B8@0		;0B56					!?8A0=85
-	15.01	0,00 ½	@548B   							@>F5=BK :@4B		!?8A0=85
-	15.01	0,00 ½	  				@548B   	>30H5=85 :@4B	!?8A0=85
-	13.01	0,00 ½						  				5@52>4					1>@>B
-	15.01	0,00 ½	 					%5.0:5B		;0B56					!?8A0=85
-	11.01	0,00 ½	 					"0:85 45;0	><>3/?>40@8;		!?8A0=85
+	10.01	0,00 ½																			0G8A;5=85
-	10.01	0,00 ½	 					Yota				;0B56					!?8A0=85
-	09.01	0,00 ½	 					/.;NA			;0B56					!?8A0=85
-	01.01																	!=OB85					1>@>B
*/

function putBillsToExpenses(jsonBillsArr)
{
  //
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costs = ss.getSheetByName(" 0AE>4K");
  
  // costs.expandAllRowGroups();
  const costsData = costs.getDataRange();

  const cdRows = costsData.getNumRows();
  const cdColumns = costsData.getNumColumns();
  //Logger.log(costsData.getCell(cdRows, cdColumns).getValue());

  let theDate = jsonBillsArr[0].date;
  let aDate = new Date(jsonBillsArr[0].date);
  let theDay = aDate.getDate();
  let theMonth = aDate.getMonth();

  let prevDayRow = 0;
  let nextDayRow = 0;
  let insertRow = 0;

  // 0E>48< =0G0;> 4=59
  // !:0=8@C5< 45=L
  //
  // 0E>48< >:>=G0=85 
  // 0E>48< >:>=G0=85 <5AOF0
  for (var i = 2; i < cdRows; i++) {
    let n = 1;
    let cDate = costsData.getCell(i, 1);
    let iDate = cDate.getValue();
    if (iDate == "") continue;
    let dDate = new Date(iDate);
    let aDateDay = dDate.getDate();
    // 0H;8 ?5@2CN 70?8AL
    let sDate = iDate.toISOString();
    if (costsData.getCell(i, 2).getValue() != "") {
      //
      Logger.log(sDate + " 2@5<O " + costsData.getCell(i, 2).getValue());
    }
  }

  // 

}

// C=:B <5=N !:0=8@>20BL - >GBC
function MenuScanBillsFromMail() {

  // "01;8F0 A :>B>@>9 @01>B05<
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // $;03 >B;04:8
  const flgDbg = dbgGetDbgFlag(true);

  if (flgDbg) {
    // 8AB 4;O >B;04:8
    var rTest = ss.getSheetByName("Test").getRange(1, 1);
  }

  var k = 0;
  var l = 1;

  var rLastDate = ss.getRangeByName("0B0>GB0'5:");
  var dLastDate = rLastDate.getValue();

  var threads = GmailApp
                .getUserLabelByName(">Q/0=8/'5:8")
                .getThreads();

  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var dDate = messages[j].getDate();
      var sDate = dDate.toString();
      if (dDate > dLastDate) {
        //
        var sLastDate = dLastDate.toString();
      }
      var sBody = messages[j].getBody();
      Logger.log( j + " > " + messages[j].getSubject() + " [[[ "+ sBody.length.toString() +" ]]]");

      if (flgDbg) {
        // rTest.offset(k, 0)
        let arrBody = dbgSplitLongString(sBody, 4500);

      }

      var bInfo = {summ: "-", date: "-", name: " ", items: []};
      bInfo = getMailBillInfo(messages[j]);

      if (flgDbg)
      {
        var c = 3;
        rTest.offset(k, c++).setValue(bInfo.date);
        rTest.offset(k, c++).setValue(bInfo.summ);
        rTest.offset(k, c++).setValue(bInfo.name);

        if (bInfo.items.length>0) {
          rTest.offset(k++, c).setValue(bInfo.items.length);

          // {iname: iName, iprice: iPrice, iquantity: iQuantity, isum: iSum}
          bInfo.items.forEach(function(element) {
            rTest.offset(k, c++).setValue(element.isum);
            rTest.offset(k, c++).setValue(element.iquantity);
            rTest.offset(k, c++).setValue(element.iprice);
            rTest.offset(k++, c++).setValue(element.iname);
          });
        }
      } 

      k++;
      Logger.log("'5: >>> " + (l++).toString() + " <<<");
    } // !>>1I5=8O A G5:0<8
  } // &5?>G:8 A>>1I5=89 A G5:0<8
}

function getUBERBillInfo(BillMail) {
  let fSubject = BillMail.getSubject();
  let sTripDate = fSubject.slice(23);

  let spcPos = sTripDate.indexOf(" ");
  let sTripDay = sTripDate.slice(0, spcPos);
  let spcPos2 = sTripDate.indexOf(" ", spcPos+2);
  let sTripMonth = sTripDate.slice(spcPos+1, spcPos2);

  let TripMonth = getMonthNum(sTripMonth);
  let TripYear = sTripDate.slice(spcPos2+1, sTripDate.indexOf(" 3.", spcPos2+2));

  var TripDate = sTripDay + "." + TripMonth + "." + TripYear;

  let fBody = BillMail.getBody();
  // finLib.between2();

  var TripTime = between2(fBody, "From", "</tr>", "<td align", "</td>");
  var j = TripTime.indexOf(">");
  TripTime = TripTime.slice(j+1).trim();

  var TripDateTime = TripDate + " " + TripTime;

  var TripSumm = between2(fBody, "check__price", "</td>", ">", "/½").trim();

  var bInfo = {summ: TripSumm, date: TripDateTime, name: '" \"/!."!\""', items: [{iname:"5@52>7:0 ?0AA068@>2 8 103060", iprice:TripSumm, isum:TripSumm, iquantity:1.0}]};

  Logger.log("UBER > ", bInfo);
  return bInfo;
}

// C=:B <5=N !:0=8@>20BL - '5:8 UBER
function MenuCheckUBER() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var flgDbg = dbgGetDbgFlag(true);
  
  // 8AB 4;O >B;04:8
  var sTest = ss.getSheetByName("Test");

  var k = 1;

  var label = GmailApp.getUserLabelByName(">Q/0=8/"0:A8");
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      Logger.log( j + " > " + subject);

      // var body = message.getBody();
      // if (flgDbg) sTest.getRange(k, 1).setValue(body);
      
      var bInfo = getUBERBillInfo(message);

      if (flgDbg)
      {
        sTest.getRange(k, 2).setValue(bInfo.summ);
        sTest.getRange(k, 3).setValue(bInfo.date);
      }

      k++;

    } // !>>1I5=8O A G5:0<8 UBER
  } // &5?>G:8 A>>1I5=89 A G5:0<8 UBER
}

function getYandexGoBillInfo(BillMail) {
  const fSubject = BillMail.getSubject();
  let sTripDate = fSubject.slice(28);

  let spcPos = sTripDate.indexOf(" ");
  let sTripDay = sTripDate.slice(0, spcPos);
  let spcPos2 = sTripDate.indexOf(" ", spcPos+2);
  let sTripMonth = sTripDate.slice(spcPos+1, spcPos2);

  let TripMonth = getMonthNum(sTripMonth);
  let TripYear = sTripDate.slice(spcPos2+1, sTripDate.indexOf(" 3.", spcPos2+2));

  var TripDate = sTripDay + "." + TripMonth + "." + TripYear;

  let fBody = BillMail.getBody();
  // finLib.between2();

  var TripTime = between2(fBody, "route__point-name", "</td>", "<p class=", "</p>");
  var j = TripTime.indexOf(">");
  TripTime = TripTime.slice(j+1).trim();

  var TripDateTime = TripDate + " " + TripTime;

  var TripSumm = between2(fBody, "report__value_main", "</td>", ">", "/½").trim();

  var bInfo = {summ: TripSumm, date: TripDateTime, name: '" \"/!."!\""', items: [{iname:"5@52>7:0 ?0AA068@>2 8 103060", iprice:TripSumm, isum:TripSumm, iquantity:1.0}]};

  Logger.log("Yandex Go> ", bInfo);
  return bInfo;
}

function MenuCheckYandexGo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const flgDbg = dbgGetDbgFlag(true);
  
  // 8AB 4;O >B;04:8
  const sTest = ss.getSheetByName("Test");

  let k = 1;

  var label = GmailApp.getUserLabelByName("pers/>BG5BK/B0:A8");
  var threads = label.getThreads();
  if (flgDbg) SpreadsheetApp.getActive().toast(threads.length);

  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      Logger.log( j + " > " + subject);

      var body = message.getBody();
      if (flgDbg) sTest.getRange(k, 1).setValue(body);
      
      var bInfo = getYandexGoBillInfo(message);

      if (flgDbg)
      {
        sTest.getRange(k, 2).setValue(bInfo.summ);
        sTest.getRange(k, 3).setValue(bInfo.date);
      }

      k++;

    } // !>>1I5=8O A G5:0<8 /=45:A Go
  } // &5?>G:8 A>>1I5=89 A /=45:A Go
}

function getAliExpressBillInfo(BillMail) {
  const fSubject = BillMail.getSubject();

}

function MenuCheckAliExpress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var flgDbg = dbgGetDbgFlag(true);
  
  // 8AB 4;O >B;04:8
  var sTest = ss.getSheetByName("Test");
  var rTest = sTest.getRange(1, 1);

  var k = 0;

  var label = GmailApp.getUserLabelByName(">Q/>:C?:8/AliExpress");
  var threads = label.getThreads();
  for (var i = 0; i < threads.length; i++) {
    Logger.log(threads[i].getFirstMessageSubject());
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      Logger.log( j + " > " + subject + " >> " + subject.indexOf("0H =><5@ 70:070").toString());
      if (subject.indexOf("0H =><5@ 70:070") != -1) {
        var Body = message.getBody();
        Logger.log( j + " > " + subject + " [[[ "+ Body.length.toString() +" ]]]");

        if (flgDbg) rTest.offset(k, 0).setValue(" > " + subject + " [[[ "+ Body.length.toString() +" ]]]"); 
        if (flgDbg) dbgLongMailBody(rTest.offset(k, 1), Body);
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"5@52>7:0 ?0AA068@>2 8 103060",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      } // "5<0 A>>1I5=8O "0H =><5@ 70:070 ..."
      else
      {
        var Body = message.getBody();
        Logger.log( j + " > " + subject + " <<< "+ Body.length.toString() +" >>>");

        if (flgDbg) rTest.offset(k, 0).setValue(" # " + subject + " <<< "+ Body.length.toString() +" >>>"); 
        if (flgDbg) dbgLongMailBody(rTest.offset(k, 1), Body);
                
        //var bInfo = {summ: "-", date: "-", name: '"AliExpress"', items: [{name:"5@52>7:0 ?0AA068@>2 8 103060",price:62500,sum:62500,quantity:1.0}];
        var bInfo = getAliExpressBillInfo(message);

        k++;
      }
    } // !>>1I5=8O A G5:0<8 AliExpress
  } // &5?>G:8 A>>1I5=89 A G5:0<8 AliExpress
}

function SetTargetRule(ss, c, rn)
{
  const range = ss.getRangeByName(rn);

  if (range == undefined)
    return;

  const rule = range.getDataValidation();
  c.setDataValidation(rule);
}

// #AB0=02;8205< 4>ABC?=K5 AG5B0 8 "8? >?5@0F88 4;O 2K1@0==>9 87 A?8A:0 >?5@0F88
function SettingTrnctnName(ss, br)
{
  const accrual = '0G8A;5=85';
  const debit = '!?8A0=85';
  const turnover = '1>@>B';

  const NewVal = br.getValue();
  const OpAcc = br.offset(0,-2);
  const OpTrgt = br.offset(0,-1);

  var i = findInRule(turnover, NewVal);
  if (i != -1)
  {
    // K1@0=0 >1>@>B=0O >?5@0F8O
    br.offset(0,1).setValue(turnover);

    SetTargetRule(ss, OpAcc, '!G5B051');

    const Transfer = ss.getRangeByName('AB@5@52>4').getValue();
    if (NewVal == Transfer)
    {
      // 5@52>4
      SetTargetRule(ss, OpTrgt, '!G5B051');
    } else {
      OpTrgt.clearDataValidations();
      if (i == 0) {
        // !=OB85
        OpTrgt.clear();
      }
    }
  }
  else if (findInRule(debit, NewVal) != -1)
  {
    // K1@0=0 >?@50F8O A?8A0=8O
    br.offset(0,1).setValue(debit);

    const CredPersnt = ss.getRangeByName('AB@@F@4B').getValue();
    if (NewVal == CredPersnt) {
      // @>F5=BK ?> :@548BC
      SetTargetRule(ss, OpAcc, '@548BK');
      OpTrgt.clear();
    }
    else
    {
      SetTargetRule(ss, OpAcc, '!G5B051');

      const LoanPaymnt = ss.getRangeByName('AB@>3@4B').getValue();
      if (NewVal == LoanPaymnt) {
        // >30H5=85 :@548B0
        SetTargetRule(ss, OpTrgt, '@548BK');
      }
      else
      {
        const Payment = ss.getRangeByName('AB@;0B56').getValue();
        if (NewVal == Payment) {
          // ;0B56
          SetTargetRule(ss, OpTrgt, ';0B568');
        }
        else OpTrgt.clearDataValidations();
      }
    }
  }
  else
  {
    i = findInRule(accrual, NewVal);
    if (i != -1) {
      // K1@0=0 >?5@0F8O =0G8A;5=8O
      br.offset(0,1).setValue(accrual);

      SetTargetRule(ss, OpAcc, '!G5B051');
      if (i < 4) OpAcc.setValue("");
    }
  }
}

// #AB0=02;8205< A>>B25BAB2CNI89 A?8A>: >?5@0F89 4;O 2K1@0==>3> "8?0 >?5@0F88
function SettingTrnctnType(ss, br)
{
  const NewVal = br.getValue();

  if (ss.getRangeByName(NewVal) == undefined) {
    // #AB0=02;8205< ?>;=K9 A?8A>: >?5@0F89 4;O 2K1>@0 5A;8 "8? =58725AB5=
    NewVal = '?5@0F8O';
  }

  SetTargetRule(ss, br.offset(0,-1), NewVal)
}

// #AB0=02;8205< A?8A>: @0AE>4>2 4;O 2K1@0==>9 AB0BL8 @0AE>4>2
function SettingCostInfo(ss, br)
{
  const flgDbg = dbgGetDbgFlag(false);
  
  if (flgDbg)
  {
    // 8AB 4;O >B;04:8
    var sTest = ss.getSheetByName('Test');
    var rTest = sTest.getRange(1, 1);
  }
  //br.setNote('Test :' + sTest + ' Range :' + rTest.getNumRows());

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  //br.setNote('Row :' + flgDbg + ' Val :' + NewVal);
  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);
  
  if (NewVal != '')
  {
    const range = ss.getRangeByName('!B AE' + NewVal);

    if (range != undefined)
    {
      const rule = SpreadsheetApp
      .newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(range)
      .build();

      cell.setDataValidation(rule);
      return;
    }
  }

  cell.clearDataValidations();
}

// #AB0=02;8205< A?8A>: 8=D>@<0F88 4;O 2K1@0==>3> @0AE>40
function SettingCostNote(ss, br)
{
  const flgDbg = dbgGetDbgFlag(false);
  
  if (flgDbg)
  {
    // 8AB 4;O >B;04:8
    var rTest = ss.getSheetByName('Test').getRange(1, 1);
  }

  const cell = br.offset(0,1);
  const NewVal = br.getValue();

  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);

  var range;
  //SpreadsheetApp.getActive().toast('Range :'+ range);

  if (NewVal == '@>4C:BK') range = ss.getRangeByName('!B AE400307');
  else if (NewVal == '82>') range = ss.getRangeByName('!B AE;:82>');
  else if (NewVal == '010:') range = ss.getRangeByName('!B AE;:010:');

  if (range != undefined)
  {
    const rule = SpreadsheetApp
    .newDataValidation()
    .setAllowInvalid(true)
    .requireValueInRange(range)
    .build();

    cell.setDataValidation(rule);
  }
  else cell.clearDataValidations();
}

// '8B05< 8=D>@<0F8N > G5:5 87 json AB@>:8
function SettingCostBill(ss, br)
{
  const flgDbg = dbgGetDbgFlag(false);

  if (flgDbg)
  {
    // 8AB 4;O >B;04:8
    var rTest = ss.getSheetByName('Test').getRange(1, 1);
  }

  const NewVal = br.getValue();
  if (NewVal == "") return;

  if (flgDbg) rTest.offset(2, 1).setValue(NewVal);

  const bill = billInfo(NewVal);

  if (flgDbg) {
    if (bill != undefined)
      rTest.offset(3, 1).setValue(bill.name)
      .offset(1, 0).setValue(bill.summ)
      .offset(1, 0).setValue(bill.date)
      .offset(1, 0).setValue(bill.cash);
    else rTest.offset(3, 1).setValue("UNDEFINED !!!");
  }

  if (bill == undefined) return;
  // $>@<0B OG55:
  // "dd.mm", "HH:mm", "#,##0.00[$ ½]"

  // KAB02;O5< AC<<C ?>:C?:8
  br.offset(0,-5)
  .setValue(bill.summ)
  .setNumberFormat("#,##0.00[$ ½]");

  // KAB02;O5< 40BC ?>:C?:8 8 ?>;CG05< 04@5A OG59:8 A 40B>9 4;O 2KAB02;5=8O 2@5<5=8
  const A1date = br.offset(0,-7).setValue(bill.date).setNumberFormat("dd.mm").getA1Notation();

  if (flgDbg) rTest.offset(7, 1).setValue(A1date);

  // KAB02;O5< 2@5<O ?>:C?:8
  br.offset(0,-6)
  .setValue("=" + A1date)
  .setNumberFormat("HH:mm");

  // A;8 =0;8G=K5, B> 2KAB02;O5< AG5B A?8A0=8O
  if (bill.cash != 0)
    br.offset(0,-4).setValue("0@<0=")

  // KAB02;O5< !B0BLN, =D> 8 @8<5G0=85 4;O <03078=0
  const storeList = ss.getRangeByName('!?A:03078=K');

}

function onEdit(e) 
{
  const ss = e.source;

  // '8B05< D;03 "A?>;L7>20BL 02B>A?8A:8"
  let br = ss.getRangeByName('$;2B>A?8A:8');
  if (br == undefined || ! br.getValue()) return;

  const TrnctnSheet = '?5@0F88';
  const CostsSheet = ' 0AE>4K';

  br = e.range;
  if (br.getNumColumns() > 1) return; // !:>?8@>20;8 480?07>=

  const ncol = br.getColumn();
  const sname = ss.getActiveSheet().getSheetName();
  //SpreadsheetApp.getActive().toast(sname);

  if (sname == TrnctnSheet)
  {
    if (ncol == 7)
    {
      let v = e.value;
      // 7<5=8;AO B8? >?5@0F88
      if (v == undefined || v == '') // #AB0=02;8205< ?>;=K9 A?8A>: >?5@0F89 4;O 2K1>@0 5A;8 "8? >?5@0F88 1K; >G8I5=
        SetTargetRule(ss, br.offset(0,-1), '?5@0F8O');
      else SettingTrnctnType(ss, br);
    }
    else if (ncol == 6)
    {
      // 7<5=8;0AL >?5@0F8O
      SettingTrnctnName(ss, br);
    }
  }
  else if (sname == CostsSheet)
  {
      // var v = e.value;
      // br.setNote(v);

    switch(ncol) {
    case 5:
      // 7<5=8;0AL AB0BLO @0AE>4>2
      SettingCostInfo(ss, br);
      break;
    case 6:
      // 7<5=8;AO ?C=:B AB0BL8 @0AE>4>2
      SettingCostNote(ss, br);
      break;
    case 8:
      // 7<5=8;0AL 70<5B:0 (2AB028;8 G5:)
      SettingCostBill(ss, br);
      break;
    }
  }
}

function MenuCloseDay()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var costs = ss.getSheetByName(" 0AE>4K");
  
  var flgDbg = dbgGetDbgFlag(true);
  
  // 8AB 4;O >B;04:8
  var sTest = ss.getSheetByName("Test");
  var rTest = sTest.getRange(1, 1);

  var k = 1;

  //costs.expandAllRowGroups();
  var costsData = costs.getDataRange();

  var cdRows = costsData.getNumRows();
  var cdColumns = costsData.getNumColumns();
  if (flgDbg) 
  {
    rTest.offset(k, 1).setValue(cdRows)
    .offset(0, 1).setValue(cdColumns)
    .offset(0, 1).setValue(costsData.getValue());
    Logger.log(costsData.getCell(cdRows, cdColumns).getValue());
  }

  for (var i = 2; i < cdRows; i++) {
    var n = 1;
    var cData = costsData.getCell(i, 1);
    var iData = cData.getValue();

  }
}

/*
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('First item', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
*/

function onOpen(e)
{
  Logger.log('>102;O5< ?C=:BK <5=N.');

  const menuScan = [
    {name: "'5:8 UBER", functionName: 'MenuCheckUBER'},
    {name: "'5:8 /=45:A Go", functionName: 'MenuCheckYandexGo'},
    {name: "'5:8 AliExpress", functionName: 'MenuCheckAliExpress'},
    null,
    {name: "G8AB8BL >B;04:C", functionName: 'dbgClearTestSheet'}
  ];
  e.source.addMenu("!:0=8@>20BL", menuScan);

  const menuFinance = [
    {name: "0:@KBL 45=L", functionName: 'MenuCloseDay'}

  ];
  e.source.addMenu("$8=0=AK", menuFinance);

}

function onOnceAnHour()
{
  // K?>;=O5BAO 565G0A=>
  Logger.log("1@010BK205< ?>A;54=85 G5:8");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let newBills = [];

  // !:0=8@C5< 48A:
  const rLastDriveDate = ss.getRangeByName('0B0'5:8A:');
  const dLastDriveDate = ReadLastDate(ss, rLastDriveDate);
  const newLastDriveDate = ScanDrive(ss, dLastDriveDate, newBills);

  const rLastMailDate = ss.getRangeByName('0B0'5:>GB0');
  const dLastMailDate = ReadLastDate(ss, rLastMailDate);
  const newLastMailDate = ScanMail(ss, dLastMailDate, newBills);

  Logger.log("!:0=8@C5< G5:8 2 ?>GB5");
  //ReadMailOnTimer(ss);

  // 04> 2K=5AB8 2 4@C3>9 A:@8?B 8 2K?>;=OBL @565
  Logger.log("!:0=8@C5< ?>:C?:8 Ali");

}
   
