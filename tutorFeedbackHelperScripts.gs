function createMasterSheet() {
  reset();
  createNamedLocalSheets();
}

// You can run this function if you need to clear all the tutor tab you've made
function reset() {
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Set Up Instructions').activate();
  var sheets = SpreadsheetApp.getActive().getSheets();

  for (var i = sheets.length - 1; i > 2; i--) {
    ss.setActiveSheet(sheets[i]);
    ss.deleteActiveSheet();
  }
}

// Use this to create tabs with all our tutors' names from the list that should be in the 'Tutors' tab.
function createNamedLocalSheets() {
  var queryString = "=QUERY(\'All data\'!A2:K1000, \"select * where Col1 contains \'";
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Tutors').activate();
  var range = SpreadsheetApp.getActive().getRangeByName("Tutors");
  var names = range.getValues();
  ss.insertSheet('All data');

  for (row in names) {
    ss.getSheetByName('All data').activate();
    ss.insertSheet(names[row][0]);
    var activeSheet = SpreadsheetApp.getActiveSheet();
    activeSheet.getRange('A1').setValue("Tutor");
    activeSheet.getRange('C1').setValue("Course");
    activeSheet.getRange('D1').setValue("Mode");
    activeSheet.getRange('E1').setValue("It was easy to find an available appt.");
    activeSheet.getRange('F1').setValue("I felt welcomed upon arrival.");
    activeSheet.getRange('G1').setValue("The time spent with my tutor was enough.");
    activeSheet.getRange('H1').setValue("My tutor encouraged me to actively participate.");
    activeSheet.getRange('I1').setValue("I enjoyed the vibe of the session.");
    activeSheet.getRange('J1').setValue("I feel comfortable to come back.");
    activeSheet.getRange('K1').setValue("Feedback");
    activeSheet.getRange('A2').setValue(queryString + names[row][0].toString().split(" ")[0] + "\'" + "\")");
    activeSheet.deleteColumn(2);
    activeSheet.setFrozenRows(1);
    activeSheet.getRange('A1:K1000').setWrap(true);
  }
}

function checkAddresses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var range = ss.getRangeByName("Addresses");
  var addresses = range.getValues();

  for (var i = 4; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    var address = addresses[range.getLastRow() - (i - 3)];

    console.log("Sending " + sheetName + " to " + address);
  }
}

function generateAndSend() { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var folder = DriveApp.getFileById(ss.getId()).getParents().next();
  var newFolder = folder.createFolder('Sheets for tutors ' + getDate());
  var range = ss.getRangeByName("Addresses");
  var addresses = range.getValues();

  for (var i = 4; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    var pdfBlob = createPDF(ss, sheet);
    newFolder.createFile(pdfBlob).setName(sheetName + '.pdf');
    var address = addresses[range.getLastRow() - (i - 3)].toString();
    //var address = "kstadler@ucsd.edu";
    sendPdf(pdfBlob, address);
  }
}

function createPDF(ss, sheet) {
  var url = ss.getUrl().replace(/edit$/, '');
  var sheetId = sheet.getSheetId();

  var exportOptions = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'letter',
    portrait: false,
    fitw: true,
    sheetnames: false,
    pagenumbers: false,
    gridlines: true,
    fzr: false,
    gid: sheetId,
  };

  var params = [];
  for (var key in exportOptions) {
    params.push(key + '=' + exportOptions[key]);
  }

  var queryString = params.join('&');

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url + 'export?' + queryString, {
    headers: {
      'Authorization': 'Bearer' + token,
    },
  });

  return response.getBlob().setName(sheet.getName() + '.pdf');
}

function sendPdf(pdfBlob, address) {
  var message = {
    to: address,
    subject: "Student Survey Responses for the CT Center",
    body: "Hello!\n\nBelow, find your survey responses updated as of this week! Note that the first row shows the prompts, responses follow below. As always, let us know if you have any questions and we will get back to you asap. Please note: we are trying out a new method of distributing survey results, so if there are any mistakes or inconsistencies in what you receive please let us know. If your sheet shows #NA, that means you have no responses yet; continue encouraging your students to submit! \n\nThank you,\nKaty Stadler, CT Ops Assistant",
    attachments: [pdfBlob],
    bcc: "tlc-contenttutoring@ucsd.edu"
  }
  MailApp.sendEmail(message);
}

function getDate() {
  var today = new Date();
  today.setDate(today.getDate());
  today = Utilities.formatDate(today, 'GMT+08:00', "MM.dd.yyyy' 'HH:mm");
  return today.toString();
}
