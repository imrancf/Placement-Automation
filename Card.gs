const deploy_link = "https://script.google.com/a/macros/cloudfort.in/s/AKfycbwpnfPkC1DY4WEOC4HO9CCR95deYlW4wroZiyCJASE/dev?param=";

function onhomepage(error) {
  let status = PropertiesService.getUserProperties().getProperty("status");
  console.log(status);

  if (status) {
    return formCard();
  }

  let cardHeader1 = CardService.newCardHeader()
    .setTitle('Generate Notice')
    .setImageStyle(CardService.ImageStyle.CIRCLE);

  let cardSection2TextInput1 = CardService.newTextInput();
  let cardSection2TextInput2 = CardService.newTextInput();
  let cardSection2TextInput3 = CardService.newTextInput();
  let cardSection2DatePTimePicker1 = CardService.newDateTimePicker();

  cardSection2TextInput1 = cardSection2TextInput1
    .setFieldName('name')
    .setTitle('Name of the Company')
    .setMultiline(false);

  cardSection2TextInput2 = cardSection2TextInput2
    .setFieldName('profiles')
    .setTitle('Job Profiles')
    .setHint('Enter the Single or Multiple Profiles separated by " , ".')
    .setMultiline(false);

  cardSection2TextInput3 = cardSection2TextInput3
    .setFieldName('package')
    .setTitle('Package')
    .setMultiline(false);

  cardSection2DatePTimePicker1 = cardSection2DatePTimePicker1
    .setFieldName('dateTime')
    .setTitle('Recruitment Date & Time');

  let cardSection2Divider1 = CardService.newDivider();

  let cardSection1ButtonList1Button1Action1 = CardService.newAction()
    .setFunctionName('saveData');

  let cardSection1ButtonList1Button1 = CardService.newTextButton()
    .setText('Next')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(cardSection1ButtonList1Button1Action1);

  let cardSection1ButtonList1 = CardService.newButtonSet()
    .addButton(cardSection1ButtonList1Button1);

  let cardSection2 = CardService.newCardSection();

  cardSection2 = cardSection2
    .addWidget(cardSection2TextInput1)
    .addWidget(cardSection2TextInput2)
    .addWidget(cardSection2TextInput3)
    .addWidget(cardSection2DatePTimePicker1)
    .addWidget(cardSection2Divider1)
    .addWidget(cardSection1ButtonList1);

  if (error) {
    let errorText = CardService.newTextParagraph()
      .setText('Please enter all the fields.');

    cardSection2 = cardSection2
      .addWidget(errorText);
  }

  let card = CardService.newCardBuilder()
    .setHeader(cardHeader1)
    .addSection(cardSection2)
    .build();
  return card;
}

function formCard() {
  let param1 = "doc";
  let param2 = "sheet"
  let param3 = "form";

  //  Have to set the Function in palce of setOpenLink 

  var action = CardService.newAction().setFunctionName('openLinkCallback').setParameters({ "param": param1 });
  let cardSection1DecoratedText1Button1 = CardService.newImageButton().setOnClickAction(action)
    .setIconUrl("https://i.ibb.co/56ykKjx/211608-folder-icon.png")
    .setAltText('Select Template');

  let cardSection1DecoratedText1 = CardService.newDecoratedText()
    .setText('Select the Template')
    .setBottomLabel('Choose a Doc file')
    .setButton(cardSection1DecoratedText1Button1);

  let filePick = JSON.parse(PropertiesService.getUserProperties().getProperty("filePick"));
  let prop = PropertiesService.getUserProperties().getProperty("files");

  if (prop == null || !filePick) {
    clearFilesFromPropServ();

  } else {
    let docs = JSON.parse(prop);
    console.log("docouter", docs[0])

    cardSection1DecoratedText1 = CardService.newDecoratedText().setText(docs[0].name)
      .setButton(cardSection1DecoratedText1Button1);
    console.log("Picker for document stopped running");
  }

  var action1 = CardService.newAction().setFunctionName('openLinkCallback').setParameters({ "param": param2 });
  let cardSection1DecoratedText2Button1 = CardService.newImageButton().setOnClickAction(action1)
    .setIconUrl("https://i.ibb.co/56ykKjx/211608-folder-icon.png")
    .setAltText('Select Sheet');

  let cardSection1DecoratedText2 = CardService.newDecoratedText()
    .setText('Select the Sheet with Email')
    .setButton(cardSection1DecoratedText2Button1);

  let sheetFilePick = JSON.parse(PropertiesService.getUserProperties().getProperty("sheetFilePick"));
  let sheetProp = PropertiesService.getUserProperties().getProperty("sheetFiles");

  if (sheetProp == null || !sheetFilePick) {
    clearSheetFilesFromPropServ();
  } else {
    let docs = JSON.parse(sheetProp);
    console.log("sheetcounter", docs[0])
    let url = docs[0].url;
    console.log("url=", url);
    var col = fetchHead(url);
    cardSection1DecoratedText2 = CardService.newDecoratedText().setText(docs[0].name)
      .setButton(cardSection1DecoratedText2Button1);
  }

  // PropertiesService.getUserProperties().setProperty("sheetFilePick", JSON.stringify(false));
  let condition = false;

  let cardSection1SelectionInput1 = CardService.newSelectionInput()
    .setFieldName('select1')
    .setTitle('Select Email Column')
    .setType(CardService.SelectionInputType.DROPDOWN);

  if (!col) {
    col = [""]
  }

  col.forEach((element) => {
    cardSection1SelectionInput1 = cardSection1SelectionInput1.addItem(element, element, condition);
    console.log("Email Coloumn working")
  })

  let cardSection1SelectionInput2 = CardService.newSelectionInput()
    .setFieldName('select2')
    .setTitle('Select Placed Column')
    .setType(CardService.SelectionInputType.DROPDOWN);
  col.forEach((element) => {
    cardSection1SelectionInput2 = cardSection1SelectionInput2.addItem(element, element, condition);
  })

  var action2 = CardService.newAction().setFunctionName('openLinkCallback').setParameters({ "param": param3 });
  let cardSection1DecoratedText3Button1 = CardService.newImageButton().setOnClickAction(action2)
    .setIconUrl("https://i.ibb.co/56ykKjx/211608-folder-icon.png")
    .setAltText('Select Form');

  let cardSection1DecoratedText3 = CardService.newDecoratedText()
    .setText('Select the Form')
    .setButton(cardSection1DecoratedText3Button1);

  let formFilePick = JSON.parse(PropertiesService.getUserProperties().getProperty("formFilePick"));
  let formProp = PropertiesService.getUserProperties().getProperty("formFiles");

  if (formProp == null || !formFilePick) {
    clearFormFilesFromPropServ();
  } else {
    let form = JSON.parse(formProp);
    console.log("formcounter", form[0])
    cardSection1DecoratedText3 = CardService.newDecoratedText().setText(form[0].name)
      .setButton(cardSection1DecoratedText3Button1);
  }

  let cardSection1Divider1 = CardService.newDivider();

  let cardSection1TextInput2 = CardService.newTextInput()
    .setFieldName('description')
    .setTitle('Input the Form Description')
    .setMultiline(false);

  let cardSection1ButtonList1Button1Action1 = CardService.newAction()
    .setFunctionName('sendNotice');

  let cardSection1ButtonList1Button1 = CardService.newTextButton()
    .setText('Send Notice')
    .setBackgroundColor('#1b9655ff')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(cardSection1ButtonList1Button1Action1);

  let cardSection1ButtonList1Button1Action2 = CardService.newAction()
    .setFunctionName('goBack')

  let cardSection1ButtonList1Button2 = CardService.newTextButton()
    .setText('Back')
    .setBackgroundColor('#1b9655ff')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(cardSection1ButtonList1Button1Action2);

  let cardSection1ButtonList1 = CardService.newButtonSet()
    .addButton(cardSection1ButtonList1Button2)
    .addButton(cardSection1ButtonList1Button1);

  let cardSection1 = CardService.newCardSection()
    .addWidget(cardSection1DecoratedText1)
    .addWidget(cardSection1Divider1)
    .addWidget(cardSection1DecoratedText3)
    .addWidget(cardSection1Divider1)
    .addWidget(cardSection1TextInput2)
    .addWidget(cardSection1DecoratedText2)
    .addWidget(cardSection1SelectionInput1)
    .addWidget(cardSection1SelectionInput2)
    .addWidget(cardSection1Divider1)
    .addWidget(cardSection1ButtonList1);

  let card = CardService.newCardBuilder()
    .addSection(cardSection1)
    .build();
  return card;
}

function test() {
  let param1 = "doc";
  console.log(`${deploy_link}${param1}`);
}

function openLinkCallback(e) {
  let param = e.parameters.param;
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl(`${deploy_link}${param}`)
      .setOpenAs(CardService.OpenAs.OVERLAY)
      .setOnClose(CardService.OnClose.RELOAD))
    .build();
}

function onFormCard() {
  PropertiesService.getUserProperties().setProperties({ "status": "changed" });
  return onhomepage();
}

function goBack() {
  deleteProp();
  return onhomepage()
}

function deleteProp() {
  PropertiesService.getUserProperties().deleteProperty("status");
}