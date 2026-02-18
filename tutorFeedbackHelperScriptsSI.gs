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
  var queryString = "=QUERY(\'Survey data\'!A2:J500, \"select * where Col1 contains \'";
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Tutors').activate();
  var range = SpreadsheetApp.getActive().getRangeByName("Tutors");
  var names = range.getValues();
  ss.insertSheet('Raw data');
  ss.insertSheet('Survey data');

  var headers = [
    "Tutor",
    "Course",
    "Mode",
    "The SI/PLL session helped me feel more confident with the course material.",
    "I believe that attending SI/PLL will help me to succeed/reach my goals in this course.",
    "The activities (e.g., practice problems, discussions, group work) supported my learning.",
    "I felt comfortable participating in the session.",
    "I felt encouraged throughout the session.",
    "Overall, this session was a valuable use of my time.",
    "Feedback"
  ];

  // populate Survey data sheet 
  surveySheet = ss.getSheetByName('Survey data');
  surveySheet.getRange('A1:J1').setValues([headers]);
  surveySheet.setFrozenRows(1);
  surveySheet.getRange('A1:J1').setWrap(true);
  
  // === STEP 2: Set the formulas in the second row ===
  // Use double quotes for outer JS string, single quotes inside formula are fine.
  surveySheet.getRange('A2').setFormula(
    "=FILTER(FLATTEN(INDIRECT(\"'Raw data'!B2:AC500\")), FLATTEN(INDIRECT(\"'Raw data'!B2:AC500\")) <> \"\")"
  );
  surveySheet.getRange('B2').setFormula(
    "=FILTER(FLATTEN(INDIRECT(\"'Raw data'!A2:A999\")), FLATTEN(INDIRECT(\"'Raw data'!A2:A999\")) <> \"\")"
  );
  surveySheet.getRange('C2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AD2:AD999\"))"
  );
  surveySheet.getRange('D2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AE2:AE999\"))"
  );
  surveySheet.getRange('E2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AF2:AF999\"))"
  );
  surveySheet.getRange('F2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AG2:AG999\"))"
  );
  surveySheet.getRange('G2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AH2:AH999\"))"
  );
  surveySheet.getRange('H2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AI2:AI999\"))"
  );
  surveySheet.getRange('I2').setFormula(
    "=FLATTEN(INDIRECT(\"\'Raw data\'!AJ2:AJ999\"))"
  );
  surveySheet.getRange('J2').setFormula(
    "=FLATTEN(INDIRECT(\"'Raw data'!AK2:AK999\"))"
  );

  for (row in names) {
    // Create a sheet for each tutor
    ss.insertSheet(names[row][0]);
    var activeSheet = ss.getSheetByName(names[row][0]);

    // Copy headers into the new sheet
    activeSheet.getRange('A1:J1').setValues([headers]);
    activeSheet.setFrozenRows(1);
    activeSheet.getRange('A1:J1000').setWrap(true);

    // Add your custom formula or value in A2
    activeSheet.getRange('A2').setValue(queryString + names[row][0].toString().split(" ")[0] + "\'" + "\")");
  }
}

function checkAddresses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var range = ss.getRangeByName("Addresses");
  var addresses = range.getValues().flat(); // flatten into 1D array

  for (var i = 5; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    var address = addresses[i - 5]; // match address index to sheet index offset

    console.log("Sending " + sheetName + " to " + address);
  }
}


function generateAndSendBatchedSingleRun(batchSize = 5) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var range = ss.getRangeByName("Addresses");
  var addresses = range.getValues();

  var startIndex = 5; // start at sheet index 5
  var totalSheets = sheets.length;

  while (startIndex < totalSheets) {
    var endIndex = Math.min(startIndex + batchSize, totalSheets);

    Logger.log(`Processing sheets ${startIndex} to ${endIndex - 1}...`);

    for (var i = startIndex; i < endIndex; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();

      try {
        var pdfBlob = createPDF(ss, sheet);
        var address = addresses[i - 5][0];
        sendPdf(pdfBlob, address);

        Logger.log("✅ Sent " + sheetName + " to " + address);
      } catch (e) {
        Logger.log("❌ Failed to send " + sheetName + ": " + e);
      }

      // Wait 6–8 seconds between individual exports
      Utilities.sleep(3000 + Math.random() * 2000);
    }

    startIndex += batchSize;

    if (startIndex < totalSheets) {
      // Wait 30 seconds between batches
      Logger.log("⏳ Waiting 25 seconds before next batch...");
      Utilities.sleep(25000);
    }
  }

  Logger.log("✅ All sheets sent!");
}

function createPDF(ss, sheet) {
  var fileId = ss.getId(); // spreadsheet ID
  var gid = sheet.getSheetId(); // sheet (tab) ID
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?';

  var exportOptions = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'letter',
    portrait: false,
    fitw: true,
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: true,
    fzr: false,
    gid: gid, // correct gid!
  };

  var queryString = Object.keys(exportOptions)
    .map(key => key + '=' + exportOptions[key])
    .join('&');

  var token = ScriptApp.getOAuthToken();

  var response = UrlFetchApp.fetch(url + queryString, {
    headers: {
      'Authorization': 'Bearer ' + token, // fixed header
    },
  });

  return response.getBlob().setName(sheet.getName() + '.pdf');
}

function sendPdf(pdfBlob, address) {
  var message = {
    to: address,
    subject: "Student Survey Responses for Supplemental Instruction",
    body: "Hello SI Leader!\n\nBelow, find your updated student survey results here! Note that the first row shows the prompts, responses follow below. If your sheet shows #NA, that just means you have no survey responses yet! This is a new method of distributing your survey results, so if there are any mistakes, formatting issues, or inconsistencies in what you recieve please email aah-si@ucsd.edu. Thank you!", 
    attachments: [pdfBlob],
    bcc: "tlc-contenttutoring@ucsd.edu, aah-si@ucsd.edu"
  }
  MailApp.sendEmail(message);
}

function getDate() {
  var today = new Date();
  today.setDate(today.getDate());
  today = Utilities.formatDate(today, 'GMT+03:00', "MM.dd.yyyy' 'HH:mm");
  return today.toString();
}
