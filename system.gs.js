var sheetfile = SpreadsheetApp.openById("19jTTXOhwUyNusvq2gWEkD4XCAvS_TzDQMrbNRS2yO0A");

function fileOpen()
{ id = DriveApp.getFilesByName("#Current");
  id = sheetfile.next().getId();
  sheetfile = SpreadsheetApp.openById("19jTTXOhwUyNusvq2gWEkD4XCAvS_TzDQMrbNRS2yO0A");
}

function fileInitalize()
{ var j = ["PZ", "HO", "AG", "URGENT", "Master"];
  var w = [215, 180, 285, 230];
  for(i=0;i<j.length;++i)
  { sheetfile.insertSheet(j[i], 0);
    //DEVELOPER.REFERANCE.....................|1....................|2.....|3.......|4.......|5......|6...........|7...........|8............|9.............|10...........|11..........|12
    sheetfile.getSheetByName(j[i]).appendRow(["Date/Time Signed In","Name","E-Mail","Reason","Reply","Add To Cal","Meet With?","¿sentEMail?","¿movedSheet?","¿sentReply?","¿addedCal?","¿finalized?"]);
    sheetfile.getSheetByName(j[i]).setFrozenRows(1);
    sheetfile.getSheetByName(j[i]).getRange(2, 6, 999).setNumberFormat("mmmm\" \"d\", \"yyyy");
    for(c=1;c<=w.length;++c) sheetfile.getSheetByName(j[i]).setColumnWidth(c, w[c-1]);
} }

function LOOPmaster()
{ var master = sheetfile.getSheetByName("Master");
  for(var r=2;r<=master.getLastRow();++r)
  { ACTIONsentEMail(r);
    ACTIONmovedSheet(r);
} }

function LOOPothers()
{ var s = ["URGENT","AG","HO","PZ"];
  for(var i=0;i<s.length;++i)
  { for(var r=2;r<=sheetfile.getSheetByName(s[i]).getLastRow();++r)
    { ACTIONsentReply(r,s[i]);
      ACTIONaddedCal(r,s[i]);
      ACTIONfinalized(r,s[i]);
} } }

function ACTIONsentEMail(r)
{ if(sheetfile.getSheetByName("Master").getRange(r, 8).getDisplayValue()!="sentEMail")
  { GmailApp.sendEmail(sheetfile.getSheetByName("Master").getRange(r, 3).getDisplayValue(), "Subject", "Body");
    sheetfile.getSheetByName("Master").getRange(r, 8).setValue("sentEMail");
} }

function ACTIONmovedSheet(r)
{ if(sheetfile.getSheetByName("Master").getRange(r, 9).getDisplayValue()!="movedSheet")
  { 
    var num=sheetfile.getSheetByName("Master").getRange(r, 2).getDisplayValue().toUpperCase().charCodeAt(0);
    var cou = "ERROR";
    if(num >= 65) {if(num <= 71) {cou="AG";}}
    if(num >= 72) {if(num <= 79) {cou="HO";}}
    if(num >= 80) {if(num <= 90) {cou="PZ";}}
    if(sheetfile.getSheetByName("Master").getRange(r, 2).getDisplayValue().indexOf("[URGENT]")!=-1) {cou="URGENT"}
    if(cou=="ERROR")
    { sheetfile.getSheetByName("Master").getRange(r, 9).setValue(cou);}
    else
    { sheetfile.getSheetByName("Master").getRange(r, 9).setValue("movedSheet");
      sheetfile.getSheetByName("Master").getRange(r, 1, 1, 25).copyValuesToRange(sheetfile.getSheetByName(cou), 1, 25, sheetfile.getSheetByName(cou).getLastRow()+1, sheetfile.getSheetByName(cou).getLastRow()+1);
} } }

function ACTIONsentReply(r,s)
{ if(sheetfile.getSheetByName(s).getRange(r, 10).getDisplayValue()!="sentReply")
  { if(sheetfile.getSheetByName(s).getRange(r, 5).getDisplayValue()!="")
    { GmailApp.sendEmail(sheetfile.getSheetByName(s).getRange(r, 3).getDisplayValue(),"Reply from the Guidance Office",sheetfile.getSheetByName(s).getRange(r, 5).getDisplayValue());
      sheetfile.getSheetByName(s).getRange(r, 10).setValue("sentReply");
} } }

function ACTIONaddedCal(r,s)
{ if(sheetfile.getSheetByName(s).getRange(r, 11).getDisplayValue()!="addedCal")
  { if(sheetfile.getSheetByName(s).getRange(r, 6).getDisplayValue()!="")
    { sheetfile.getSheetByName(s).getRange(r, 6).setNumberFormat("mmmm\" \"d\", \"yyyy");
      var cal = CalendarApp.getCalendarsByName("Guidance "+s)[0].createAllDayEvent(sheetfile.getSheetByName(s).getRange(r, 2).getDisplayValue(), sheetfile.getSheetByName(s).getRange(r,6).getValue());
      sheetfile.getSheetByName(s).getRange(r, 11).setValue("addedCal");
} } }

function ACTIONfinalized(r,s)
{ if(sheetfile.getSheetByName(s).getRange(r,7).getValue()!="")
  { var newRow = sheetfile.getSheetByName("Master").getRange(sheetfile.getLastRow()+1, 1);
    sheetfile.getSheetByName(s).getRange(r, 1, 1, 25).copyTo(newRow);
    sheetfile.getSheetByName("Master").getRange(sheetfile.getLastRow(), 1, 1, 25).protect().setWarningOnly(true);
    sheetfile.getSheetByName(s).deleteRow(r);
} }

function DEVELOPERtest()
{ GmailApp.sendEmail("patrick.hirsch28@gmail.com", "123", sheetfile.getSheetByName("PZ").getRange(5, 6).getNumberFormat())
}
//{ var s="HO";
//  var r=2;
//  ACTIONsentEMail(r);
//  ACTIONmovedSheet(r);
//  ACTIONsentReply(r,s);
//  ACTIONaddedCal(r,s);
//  ACTIONfinalized(r,s);
//  }

//{ var event = CalendarApp.getDefaultCalendar().createAllDayEvent('Apollo 11 Landing',
//    new Date('July 20, 1969'));
//Logger.log('Event ID: ' + event.getId());
//}

//function DEVELOPERreset
//{ for(i=0;i<sheetfile.getNumSheets();++i)
//  { if(sheetfile.getSh)
//  }
//}