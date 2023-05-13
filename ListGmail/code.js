var sheetName= 'ListEmails'

function myFunction() {
  initialHeaderSheet()
  var ui= SpreadsheetApp.getUi();
  ui.createMenu("Hosting Group's Emails")
    .addItem("Restart Table","initialHeaderSheet")
    .addSeparator()
    .addSubMenu(ui.createMenu("Get Emails")
      .addItem("Group-hosting-fw Lable","gmail_Group_hosting_fw")
      .addItem("Group-hosting Lable","gmail_Group_hosting"))
    .addToUi();
}

function initialHeaderSheet() {

  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  sheet = sheets.getSheetByName(sheetName);

  if (  sheet === null ) {
    
    // If sheet does not exist, so create it
    sheets.insertSheet(sheetName);

  } else {

    sheets.deleteSheet(sheet);
    sheet = sheets.insertSheet(sheetName);

  }

  var maxCols = sheet.getMaxColumns();
  if ( maxCols > 4 ) {
    sheet.deleteColumns(1,maxCols - 4);
  }


  var startRow = 10
  


  var headerRow = sheet.getRange(1, 1 , startRow - 1, 4);
  // add Header values
  // headerRow.setValues([['Subject','Date and Time', 'Sender Derails', 'Body Contents']]);
  headerRow.setFontSize(10);
  headerRow.setFontColor('black');
  headerRow.setFontWeight('bold');
  headerRow.setHorizontalAlignment('center');
  headerRow.setVerticalAlignment('middle');
  headerRow.setBackground('#efefef'); // light gray 2
  // sheet.setRowHeight(startRow, 28);
  // sheet.setColumnWidths(1, 4, 500);

  headerRow = sheet.getRange('A3:A5');
  headerRow.setValue('search');
  headerRow.setFontSize(14);
  headerRow.merge();
  headerRow = sheet.getRange('B3');
  headerRow.setValue('Lable:')
  headerRow.setFontSize(10);
  listGmailLabels();
  headerRow = sheet.getRange('B4');
  headerRow.setValue('Subject:')
  headerRow.setFontSize(10);
  headerRow = sheet.getRange('B5');
  headerRow.setValue('Email:')
  headerRow.setFontSize(10);
  headerRow = sheet.getRange(3,2,3,2);
  headerRow.setBackground('white');
  headerRow.setBorder(true, true, true, true, true, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  headerRow = sheet.getRange(3,4,3,1);
  headerRow.merge();
  headerRow = sheet.getRange(1,1,2,4);
  sheet.setRowHeightsForced(1,2,10);
  headerRow.merge();
  headerRow = sheet.getRange(6,1,4,4);
  headerRow.merge();
  sheet.setRowHeightsForced(6,4,5);

  var bodyRow = sheet.getRange(startRow, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  bodyRow.applyRowBanding() 

  var bodyHeaderRow = sheet.getRange(startRow, 1, 1, 4);
  // add Boady Header values
  bodyHeaderRow.setValues([['Subject','Date and Time', 'Sender Derails', 'Body Contents']]);
  bodyHeaderRow.setFontSize(14);
  bodyHeaderRow.setFontColor('white');
  bodyHeaderRow.setFontWeight('bold');
  bodyHeaderRow.setHorizontalAlignment('center');
  bodyHeaderRow.setVerticalAlignment('middle');
  bodyHeaderRow.setBackground('grey');
  sheet.setRowHeight(startRow, 25);
  sheet.setColumnWidths(1, 4, 500);


  var otherRow = sheet.getRange(startRow + 1 , 1, sheet.getMaxRows(), sheet.getMaxColumns());
  otherRow.setHorizontalAlignment('left');
  otherRow.setVerticalAlignment('top');
  otherRow.setFontSize(12);
  sheet.setRowHeightsForced(startRow + 1,sheet.getMaxRows()-startRow,25);

  //sheet.setRowHeight(sheet.getMaxRows(), 10);
  var dataRow = sheet.getRange(startRow + 1, 2, sheet.getMaxRows(), 2);
  dataRow.setHorizontalAlignment('center');
  sheet.setColumnWidths(2,1,180);




}


function resetSheet() {

  var startRow = 10
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  sheet = sheets.getSheetByName(sheetName);

  // var data = sheet.getDataRange().getValues();
  var howManyToDelete = sheet.getLastRow() - startRow; 

  if (howManyToDelete > 0){
     sheet.deleteRows(startRow + 1 , howManyToDelete );
  } //else {
  //     Browser.msgBox('The sheet is empty anyway!');  
  //}

}


function gmail_Group_hosting_fw() {
  resetSheet();
  gmail('Group-hosting/Group-hosting-fw');
  mergeSameCell();
}

function gmail_Group_hosting() {
  resetSheet();
  gmail();
  mergeSameCell();
}

function gmail() {
  if ( Session.getActiveUser().getEmail() != 'narges.ahmadi@mediamonks.com' ){
    Browser.msgBox('Please login as the Hosting gmail member gmail account.');
  }

  // var labels = GmailApp.getUserLabels();
  // for (var i = 0; i < labels.length; i++) {
  // Logger.log(labels[i].getName());
  // }

  const today = new Date();
  const today_2 = new Date();
  today_2.setDate(new Date().getDate()+2);
  // today.setDate(new Date().getDate()-2);
  const td = Utilities.formatDate(today, 'GMT+1', "yyyy/MM/dd");
  const td_2 = Utilities.formatDate(today_2, 'GMT+1', "yyyy/MM/dd");
  const mylabel = 'unread';

  var lableName =  listGmailLabels();
  var subjectName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(lableName).getRange('C4').getValue(); 
  var emailName =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(lableName).getRange('C5').getValue();
  //if {  }
  const queryString = `label: ${lableName} after: ${td} before: ${td_2}`;

  Logger.log(queryString); // check the output in the View -> Logs
  const threads = GmailApp.search(queryString)

  //var lable = GmailApp.getUserLabelByName(lableName);
  //var threads = lable.getThreads();
  for (var i = threads.length - 1 ; i>=0 ; i--){
    var messages= threads[i].getMessages();
    for (var j = 0; j < messages.length ; j++){
      var message = messages[j]
      extractDetails(message);
    }
  }
}


function extractDetails(message){
  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var senderDetails = message.getFrom();
  var bodyContents = message.getPlainBody();

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.appendRow([subjectText,dateTime,senderDetails,bodyContents]);

}

function listGmailLabels() {

  var label = [];
  var sheets = SpreadsheetApp.getActiveSpreadsheet();    
  var menuSheet = sheets.getSheetByName(sheetName);
 
  var labels = GmailApp.getUserLabels();
  for (var i = 0; i < labels.length; i++) {
  label.push( labels[i].getName());

  }

  // PART DROPDOWN
  var partCell = menuSheet.getRange('C3'); 
  var partRange = label;
  var partRule = SpreadsheetApp.newDataValidation().requireValueInList(partRange).build();
  partCell.setDataValidation(partRule);

  return menuSheet.getRange('C3').getValue();

}

function mergeSameCell() {

  var start = 10; // Start row number for values.
  var c = {};
  var k = "";
  var offset = 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Retrieve values of column B.
  var data = ss.getRange(start, 1, ss.getLastRow()-start , 1).getValues();
  
  //Logger.log(data);
  var a = {}
  // Retrieve the number of duplication values.
  var i=1;
  data.forEach(function(e){Logger.log(i);Logger.log(e[0]); i++ });  
  data.forEach(function(e){c[e[0]] = c[e[0]] ? c[e[0]] + 1 : 1;});

  // Merge cells.
  data.forEach(function(e){
    if (k != e[0]) {
      ss.getRange(start + offset, 1, c[e[0]], 1).merge();
      offset += c[e[0]];
      //Logger.log(ss.getRange(c[e[0]]).getValues());
      Logger.log(offset);
    }
    k = e[0];
  });

}

