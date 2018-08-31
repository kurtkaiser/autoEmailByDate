// Kurt Kaiser
// Date Messaging System
// All Rights Reserved, 2018

var ss = SpreadsheetApp.getActiveSpreadsheet();
var msgSheet = ss.getSheetByName("Messages");
var infoSheet = ss.getSheetByName("Info");
var infoLastColumn = infoSheet.getLastColumn();
var infoLastRow = infoSheet.getLastRow();
var msgLastRow = msgSheet.getLastRow();

// ----------------- Formating -----------------
function makeMessageSheet() {
  if (msgSheet != null) return false;
  ss.insertSheet("Messages");
  msgSheet = ss.getSheetByName("Messages");
  msgSheet.getRange('B1:F1').mergeAcross();
  msgSheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  msgSheet.getRange('B1:F1').activate();
  var currentCell = msgSheet.getCurrentCell();
  msgSheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  msgSheet.getRange('B1:F1').copyTo(msgSheet.getActiveRange(),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  msgSheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  msgSheet.getRange('A1').setValue('Title');
  msgSheet.getRange('B1').setValue('Message');
  msgSheet.getRange('B1:F1').mergeAcross();
  msgSheet.getActiveRangeList().setHorizontalAlignment('center');
  msgSheet.getRange('1:1').setFontWeight('bold');
  msgSheet.setFrozenRows(1);
  msgSheet.getRange('A2').activate();
}

function checkFormating() {
  var sheets = ss.getSheets();
  if (sheets.length != 2) {
    SpreadsheetApp.getUi().alert('Error: Spreadsheet must a' +
      ' sheet named "Info" and a sheet named "Messages".');
  }
}

// Updates column one of the date sheet with current message title list
function dataValidation() {
  ss.getRange('Info!A2:A1000').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireValueInRange(ss.getRange('Messages!$A$2:$A'), true)
    .build());
  // Sets date column to not allow past dates to be entered
  infoSheet.getRange(infoLastRow, 3, 1000, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireDateAfter(new Date())
    .setHelpText('Date must be in the future. Any date today or ealier is not valid.')
    .build());
}

// ----------------- Tasks -----------------

/* Creates object from the information sheet, dates must
be in the future and Status cell must be blank */
function getInfoObject() {
  // Get first three columns of info sheet into an array
  var info = [];
  var msgTitles = infoSheet.getRange(2, 1, infoLastRow, 1).getValues();
  var statuses = infoSheet.getRange(2, 2, infoLastRow, 1).getValues();
  var dates = infoSheet.getRange(2, 3, infoLastRow, 1).getValues();
  var emails = infoSheet.getRange(2, 4, infoLastRow, 1).getValues();
  // Convert column objects into arrays
  msgTitles = [].concat.apply([], msgTitles);
  statuses = [].concat.apply([], statuses);
  dates = [].concat.apply([], dates);
  emails = [].concat.apply([], emails);
  for (i = 0; i < infoLastRow; i++) {
    info[i] = {
      row: i + 2,
      title: msgTitles[i],
      status: statuses[i],
      date: dates[i],
      email: emails[i]
    }
  }
  return info;
}

// Creates object from the message sheet
function getMessagesObject() {
  var messages = [];
  var titles = msgSheet.getRange(2, 1, msgLastRow, 1).getValues();
  var days = msgSheet.getRange(2, 2, msgLastRow, 1).getValues();
  var subjects = msgSheet.getRange(2, 3, msgLastRow, 1).getValues();
  var msgs = msgSheet.getRange(2, 6, msgLastRow, 1).getValues();
  // Convert column objects into arrays
  titles = [].concat.apply([], titles);
  days = [].concat.apply([], days);
  subjects = [].concat.apply([], subjects);
  msgs = [].concat.apply([], msgs);
  for (i = 0; i < msgLastRow - 1; i++) {
    messages[i] = {
      row: i + 2,
      title: titles[i],
      day: days[i],
      subject: subjects[i],
      msg: msgs[i]
    }
  }
  return messages;
}

//Convert dates to show how many days away they are
function getDaysAway(info, msgs) {
  var removeIndex = [];
  var today = new Date();
  today.setHours(0,0,0,0);
  /* Interate through data, if status blank and date not passed, 
     subtract today from date, convert from miliseconds to days */
  for (c = 0; c < infoLastRow; c++) {
    if (info[c].status == "" && today < info[c].date ) {
      info[c].daysAway = (Math.abs(info[c].date - today)) / 86400000;
      info[c].daysAway = parseInt(info[c].daysAway);
    } else {
      removeIndex.push(c);
      continue;
    }
  }
  // Remove info from array that won't be used
  var count = 0;
  for (n = 0; n < removeIndex.length; n++) {
    info.splice(removeIndex[n] - count, 1);
    count++;
  }
  return info;
}

// Alphabetizes objects by the titles
function alphabetize(obj) {
  obj.sort(function(a, b) {
    return (a.title > b.title) ? 1 : ((b.title > a.title) ? -1 : 0);
  });
  return obj;
}

// Create an object of data that needs to be sent
function sendInfo(info, msgs) {
  var addInfo;
  var header;
  var i = 0;
  var m = 0;
	var otherInfo = "";
  header = infoSheet.getRange(1, 5, 1, infoLastColumn - 5).getValues();
  header = [].concat.apply([], header);
	/* Get any other columns on info page, check if a date or a time
	    and format accordingly */
  while (info[i]) {// Add info from message sheet to info object
    while (true) {
      info[i].button = "";
      if (info[i].title == msgs[m].title) {
        if (info[i].daysAway != msgs[m].day){
          info.splice(i, 1);
          continue;
        }
        info[i].subject = msgs[m].subject;
        info[i].msg = msgs[m].msg;
        if (msgs[m].buttonLink) makeButtons(info[i], msgs[m]);

        break;
      } else {
        m++;
        if (m == msgs.length) {
           SpreadsheetApp.getUi().alert('Error: Row ' + info[i].row + 
                                        ' has an unknow message title.');
          return;
        }
      }
    }
    info[i].other = infoSheet.getRange(info[i].row, 5, 1, infoLastColumn - 5).getValues();
    info[i].other = [].concat.apply([], info[i].other);
    var otherInfo = "<br />";
    for (c = 0; c < info[i].other.length; c++) {
      if (Date.parse(info[i].other[c])) {
        if (info[i].other[c].getYear() == 1899) {
          info[i].other[c] = info[i].other[c].getTime();
          info[i].other[c] = info[i].other[c].toLocaleTimeString();
          info[i].other[c].slice(0, info[i].other[c].length - 4);
        } else {
          info[i].other[c] = info[i].other[c].toLocaleDateString();
        }
      }
      otherInfo = otherInfo + "<br />" + header[c] + ":   " + info[i].other[c];
    }
    info[i].other = otherInfo;
    sendEmail(info[i]);
    i++;
  }
}

function makeButton(infoI, msgsM){
    info[i].button =
      '<br /><br /><a href="' +
    msgsM.buttonLink +
    '" class="btn" style="-webkit-border-radius: 28;' +
    "-moz-border-radius: 5;border-radius: 5px;font-family: Arial; color: #ffffff;font-size: 15px;" +
    'background: #ff7878;padding:8px 20px 8px 20px;text-decoration: none;">' +
    msgsM.buttonText +
    '</a>';
}

function sendEmail(infoI) {
  MailApp.sendEmail({
    to: infoI.email,
    subject: infoI.subject,
    htmlBody: makeEmail(infoI)
  })
}

function makeEmail(infoI) {
  return (
    '<!DOCTYPE html><html><head><base target="_top"></head><body><div style="text-align: center;' +
    'font-family: Arial;"><div id="center" style="width:300px;border: 2px dotted grey;background:' +
    '#ececec; margin:25px;margin-left:auto; margin-right:auto;padding:15px;"><div style=" border: 2px dotted grey;' +
    'background:white;margin-right:auto; margin-left:auto; padding:10px;"><h2>' +
    infoI.title +
    "</h2><h3>" +
    infoI.msg +
    infoI.other +
    infoI.button +
    '<br /><br /></div></div><div><p style="font-size:12px">' +
    'Created by<a href="https://www.linkedin.com/in/kurtkaiser/">' + 
    'Kurt Kaiser</a><br />2a546573543a44</p></div></body></html>'
  );
}

// ----------------- Main -----------------
function main() {
  var info = getInfoObject();
  var messages = getMessagesObject();
  info = alphabetize(info);
  messages = alphabetize(messages);
  getDaysAway(info, messages);
  sendInfo(info, messages);
}
