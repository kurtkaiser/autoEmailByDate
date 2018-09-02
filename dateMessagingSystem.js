// Kurt Kaiser
// Date Emailing System
// All Right Reserved, 2018

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();

function format(){
  sheet.getRange('A1').setValue('Email');
  sheet.getRange('B1').setValue('Send Date');
  sheet.getRange('C1').setValue('Status');
  sheet.getRange('D1').setValue('Subject');
  sheet.getRange('E1').setValue('Header');
  sheet.getRange('F1').setValue('Message');
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A2:A1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireTextIsEmail()
    .build());
  sheet.getRange('B2:B1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireDate()
    .build());
  sheet.getRange('C2:C1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .setHelpText('Program uses this column, leave cells blank.')
    .requireTextEqualTo('Email Sent')
    .build());
}

// Get rows of messages that need to be sent
function getRowsToSend(){
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var rowsToSend = [];
  var dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  dates = [].concat.apply([], dates);
  for (i = 0; i < dates.length; i++){
    if (dates[i].getTime() === today.getTime()){
      rowsToSend.push(i + 2);
    }
  }
  return rowsToSend;
}

// Object Constructor for needed rows
function Info(row){
  var values = sheet.getRange(row, 1, 1, lastColumn).getValues();
  values = [].concat.apply([], values);
  this.email = values[0];
  this.date = values[1];
  this.subject = values[3];
  this.title = values[4];
  this.message = values[5];
  this.row = row;
  if (values[values.length - 1] != ""){
    this.addedInfo = getAddedInfo(values.slice(6, lastColumn));
  }
}

// If sheet has additional info, loop and format it
function getAddedInfo(addedInfo){
  var header = sheet.getRange(1, 7, 1, lastColumn - 6).getValues();
  header = [].concat.apply([], header);
  var formatAddedInfo = '<br />';
  for (i = 0; i < addedInfo.length; i++){
    if (addedInfo[i] != ""){
      if (Date.parse(addedInfo[i])) {
        if (addedInfo[i].getYear() == 1899) {
          addedInfo[i] = addedInfo[i].toLocaleTimeString();
          addedInfo[i].slice(0, addedInfo[i].length - 4);
        } else {
          addedInfo[i] = addedInfo[i].toLocaleDateString();
        }
      }
      formatAddedInfo = formatAddedInfo + '<br />' +
        header[i] + ':   ' + addedInfo[i];
    }
  }
  return formatAddedInfo;
}

// Prepare body of email
function makeEmail(info) {
  return (
    '<!DOCTYPE html><html><head><base target="_top"></head><body><div style="text-align: center;' +
    'font-family: Arial;"><div id="center" style="width:300px;border: 2px dotted grey;background:' +
    '#ececec; margin:25px;margin-left:auto; margin-right:auto;padding:15px;"><br /><div style=" ' +
    'border: 2px dotted grey;background:white;margin-right:auto; margin-left:auto; padding:10px;"><h2>' +
    info.title +
    '</h2><h3>' +
    info.message +
    info.addedInfo +
    '<br /><br /></div></div><div><p style="font-size:12px">' +
    'Created by<a href="https://www.linkedin.com/in/kurtkaiser/">' +
    'Kurt Kaiser</a><br /> 2a546573543a44 </p></div></body></html>'
  );
}

function sendEmail(info) {
  MailApp.sendEmail({
    to: info.email,
    subject: info.subject,
    htmlBody: makeEmail(info)
  })
  sheet.getRange(info.row, 3).setValue('Email Sent');
}
                             
// ----------------- Main -----------------
function main(){
  var rowsToSend = getRowsToSend();
  var info;
  for (c = 0; c < rowsToSend.length; c++){
    info = new Info(rowsToSend[c]);
    sendEmail(info);
  }
}
