function generateObservationFolders() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var tutorNames;

  // use the tabs in the master sheet to get a list of tutors (not hidden)
  for (var i = 1; i < sheets.length; i++) {
    var currSheet = sheets[i];
    if (!currSheet.isSheetHidden()) {
      tutorNames.push(currSheet.getName());
    }
  }

  // create folders in correct quarter folder based on tutor names
  const quarterFolderName = "WI 25 Observations";
  const quarterFolder = DriveApp.getFoldersByName(quarterFolderName).next();
  for (var j = 0; j < tutorNames.length; j++) {
    var currName = tutorNames[j];
    quarterFolder.createFolder(currName);
  }

  // now that folders are created, populate with observation form
  const fileName = "Copy WI 25 Content Tutor Observation Form";
  const tutorFolders = quarterFolder.getFolders();
  const file = quarterFolder.getFilesByName(fileName).next();

  while (tutorFolders.hasNext()) {
    const currFolder = tutorFolders.next();
    var toBeNamed = `${currFolder.getName()} WI 25 CT Observation Form`;

    // check if the file already exists
    const existingFiles = currFolder.getFilesByName(toBeNamed);
    if (existingFiles.hasNext()) {
      continue;
    }

    const copiedFile = file.makeCopy(toBeNamed, currFolder);
  }
}
