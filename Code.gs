function doGet(e) {
  console.log(e)
  if (e.parameter.param == "doc") {
    let html = HtmlService.createTemplateFromFile("index")
      .evaluate().setTitle("Select Folder");
    return html;
  } else if (e.parameter.param == "sheet") {
    let html = HtmlService.createTemplateFromFile("sheet")
      .evaluate().setTitle("Select Folder");
    return html;
  }
}

//  Called when picker is initialised.
function pickerConfig() {
  DriveApp.getRootFolder();
  return {
    oauthToken: ScriptApp.getOAuthToken(),
    developerKey: "AIzaSyBKdhXo3-OLXghFyPpHrHXrEGSqhvmifzI"
  }
}

function storeDriveSelections(fileId) {
  // Append current list of files and folders.
  let storedDocs = JSON.parse(PropertiesService.getUserProperties()
    .getProperty("files"));

  let updateArray = () => {
    //Combine current list with incoming and remove duplicates.
    return [...new Map([...fileId, ...storedDocs].map(item => [item.id, item])).values()]

  };

  // IF not stored ids just input the fileId otherwise add both to array.
  let docsAll = (storedDocs === null) ? fileId : updateArray();

  //Add storedDocs to selected docs;
  PropertiesService.getUserProperties()
    .setProperty("files", JSON.stringify(docsAll))

  // Allows us to only keep these properties when using is working on saved properties.
  let file = PropertiesService.getUserProperties()
    .setProperty("filePick", JSON.stringify(true));

  console.log("file", file.getProperties());
}

function storeDriveSelectionsForSheet(fileId) {
  // Append current list of files and folders.
  let storedSheet = JSON.parse(PropertiesService.getUserProperties()
    .getProperty("sheetFiles"));

  let updateArray = () => {
    //Combine current list with incoming and remove duplicates.
    return [...new Map([...fileId, ...storedSheet].map(item => [item.id, item])).values()]

  };

  // IF not stored ids just input the fileId otherwise add both to array.
  let docsAll = (storedSheet === null) ? fileId : updateArray();

  //Add storedSheet to selected docs;
  PropertiesService.getUserProperties()
    .setProperty("sheetFiles", JSON.stringify(docsAll))

  // Allows us to only keep these properties when using is working on saved properties.
  let file = PropertiesService.getUserProperties()
    .setProperty("sheetFilePick", JSON.stringify(true));

  console.log("sheetFiles", file.getProperties());
}

function clearFilesFromPropServ() {
  PropertiesService.getUserProperties()
    .deleteProperty("files");
}

function clearSheetFilesFromPropServ() {
  PropertiesService.getUserProperties()
    .deleteProperty("sheetFiles");
}

// Fetching Headers of the sheet
function fetchHead(url) {
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var hRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues();
  // console.log(hRow.flat(1));

  return hRow.flat(1);
}

function createNotice(e){
  console.log("e",e)
  let templateNotice = DriveApp.getFileById('1cmVZoLQVKAkSn7cT-1hyXqFUTrUWtcbUQfJgi_LeWcQ');
  let destinationFolder = DriveApp.getFolderById('17bGaatlPAtVrVaXn9xeMx1roEacWE7oP');

  let copy = templateNotice.makeCopy(e.formInput.name,destinationFolder);
  let doc = DocumentApp.openById(copy.getId());
  let body = doc.getBody();
  body.replaceText('{{Date}}',Utilities.formatDate(new Date(),"IST","yyyy-MM-dd"));
  body.replaceText('{{Company_Name}}',e.formInput.name);
  body.replaceText('{{Job_Profiles}}',e.formInput.profiles);
  body.replaceText('{{Package}}',e.formInput.package);
  let date = Utilities.formatDate(new Date(e.formInput.dateTime.msSinceEpoch),"IST","yyyy-MM-dd HH:mm")
  body.replaceText('{{Recruitment_Date}}',date);

  doc.saveAndClose();
  let docID = doc.getId();
  // Converting doc to pdf
  let file = DriveApp.getFileById(docID);
  pdfFile = file.getAs('application/pdf');

  destinationFolder.createFile(pdfFile);

 //delete the original doc file
  let docFile = DriveApp.getFileById(docID);
  docFile.setTrashed(true);
}
