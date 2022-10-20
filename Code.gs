function doGet(e) {
  console.log(e)
  let html = HtmlService.createTemplateFromFile("index");
  let mime = e.parameter.param;
  html.mime = mime;
  return html.evaluate().setTitle("Select File");
}

//  Called when picker is initialised.
function pickerConfig() {
  DriveApp.getRootFolder();
  return {
    oauthToken: ScriptApp.getOAuthToken(),
    developerKey: "AIzaSyBI_fVDuIk2sZfNdKi_BFxpRGeyXDQyyPo"
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

  onFormCard();
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
  onFormCard();
}

function storeDriveSelectionsForForm(fileId) {
  // Append current list of files and folders.
  let storeFormData = JSON.parse(PropertiesService.getUserProperties()
    .getProperty("formFiles"));

  let updateArray = () => {
    //Combine current list with incoming and remove duplicates.
    return [...new Map([...fileId, ...storeFormData].map(item => [item.id, item])).values()]

  };

  // IF not stored ids just input the fileId otherwise add both to array.
  let docsAll = (storeFormData === null) ? fileId : updateArray();

  //Add storedSheet to selected docs;
  PropertiesService.getUserProperties()
    .setProperty("formFiles", JSON.stringify(docsAll))

  // Allows us to only keep these properties when using is working on saved properties.
  let file = PropertiesService.getUserProperties()
    .setProperty("formFilePick", JSON.stringify(true));

  console.log("formFiles", file.getProperties());
  onFormCard();

}

function clearFilesFromPropServ() {
  PropertiesService.getUserProperties()
    .deleteProperty("files");
}

function clearSheetFilesFromPropServ() {
  PropertiesService.getUserProperties()
    .deleteProperty("sheetFiles");
}
function clearFormFilesFromPropServ() {
  PropertiesService.getUserProperties()
    .deleteProperty("formFiles");
}

// Fetching Headers of the sheet
function fetchHead(url) {
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getActiveSheet();
  var hRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues();
  // console.log(hRow.flat(1));

  return hRow.flat(1);
}

function saveData(e) {
  PropertiesService.getUserProperties().setProperties({ "inputData": JSON.stringify(e.formInput) });
  const inputDetails = JSON.parse(PropertiesService.getUserProperties().getProperty("inputData"));
  const inputDetailsLength = Object.keys(inputDetails).length;

  if ((inputDetailsLength < 4) || (inputDetails === "{}")) {
    const error = "error";
    return onhomepage(error);
  }

  return formCard();
}

function createNotice(docDetail, inputDetails) {
  let docID = docDetail[0].id;
  let file = DriveApp.getFileById(docID);
  let parentFolder = file.getParents();
  let parentId = parentFolder.next().getId();
  let templateNotice = DriveApp.getFileById(docID);
  let destinationFolder = DriveApp.getFolderById(parentId);

  let copy = templateNotice.makeCopy(inputDetails.name, destinationFolder);
  let doc = DocumentApp.openById(copy.getId());
  let body = doc.getBody();
  body.replaceText('{{Date}}', Utilities.formatDate(new Date(), "IST", "yyyy-MM-dd"));
  body.replaceText('{{Company_Name}}', inputDetails.name);
  body.replaceText('{{Job_Profiles}}', inputDetails.profiles);
  body.replaceText('{{Package}}', inputDetails.package);
  let date = Utilities.formatDate(new Date(inputDetails.dateTime.msSinceEpoch), "IST", "yyyy-MM-dd HH:mm")
  body.replaceText('{{Recruitment_Date}}', date);

  doc.saveAndClose();
  // Converting doc to pdf
  let pdfFile = doc.getAs('application/pdf');
  let createdFile = destinationFolder.createFile(pdfFile);

  let pdfID = createdFile.getId();
  // PropertiesService.getUserProperties().setProperty("Id", JSON.stringify(pdfID));
  //delete the original doc file
  let delFile = DriveApp.getFileById(doc.getId());
  delFile.setTrashed(true);
  return pdfID;
}

function sendNotice(e) {
  deleteProp();
  console.log("e", e);
  console.log("form response = ", e.formInput.select1);
  let docDetail = JSON.parse(PropertiesService.getUserProperties().getProperty("files"));
  let sheetDetail = JSON.parse(PropertiesService.getUserProperties().getProperty("sheetFiles"));
  let inputDetails = JSON.parse(PropertiesService.getUserProperties().getProperty("inputData"));
  let formDetail = JSON.parse(PropertiesService.getUserProperties().getProperty("formFiles"));
  console.log(sheetDetail);
  let pdfId = createNotice(docDetail, inputDetails);
  let pdf = DriveApp.getFileById(pdfId);

  let formId = formDetail[0].id;
  let frmDescription = e.formInput.description;
  console.log(JSON.stringify(frmDescription));
  let formFile = DriveApp.getFileById(formId);
  let parentFolder = formFile.getParents();
  let parentId = parentFolder.next().getId();
  let destinationFolder = DriveApp.getFolderById(parentId);

  let newForm = formFile.makeCopy(inputDetails.name, destinationFolder);

  let editForm = FormApp.openById(newForm.getId());
  editForm.setDescription(frmDescription);
  let formLink = editForm.getPublishedUrl();
  console.log(formLink);

  let sheetUrl = sheetDetail[0].url;

  let sheet = SpreadsheetApp.openByUrl(sheetUrl).getActiveSheet();

  let row1 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  let fRow = row1.flat(1);
  let emailColumn = fRow.findIndex(element => {
    console.log("elements -", element);
    console.log("formInppur -", e.formInput.select1);
    if (element == e.formInput.select1)
      return true;
    else
      return false;
  })

  let placedColumn = fRow.findIndex(element => {
    if (element == e.formInput.select2)
      return true;
    else
      return false;
  })

  let colValues = sheet.getRange(2, placedColumn + 1, sheet.getLastRow() - 1, 1).getValues().flat(1);
  let emailValues = sheet.getRange(2, emailColumn + 1, sheet.getLastRow() - 1, 1).getValues().flat(1);

  console.log("emailCol = ", fRow, emailColumn, placedColumn);
  console.log("Col = ", colValues);
  let emailArr = [];
  colValues.forEach((element, index) => {
    if (element == '') {
      emailArr.push(emailValues[index]);
    }
  })
  console.log("Array for Email", emailArr)

  let body = `We are elated to informyou that ${inputDetails.name} will be coming for On-Campus Recruitnment.
   Find all the information in the notice attached below and fill the form if interested 
   ${formLink}`;

  let subject = `Notice for ${inputDetails.name} Recruitment`;

  emailArr.forEach((element) => {
    MailApp.sendEmail({
      to: element,
      subject: subject,
      htmlBody: body,
      attachments: pdf
    });
  })
  PropertiesService.getUserProperties().deleteProperty("inputData");
  // PropertiesService.getUserProperties().setProperty("filePick", JSON.stringify(false));
  // PropertiesService.getUserProperties().setProperty("sheetFilePick", JSON.stringify(false));
  return CardService.newNavigation().updateCard(onhomepage(e));

}