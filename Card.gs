function onhomepage(e) {
  let formInput = JSON.parse(PropertiesService.getUserProperties().getProperty("inputData"));

  let formInputLength;
  if (formInput) {
    formInputLength = Object.keys(formInput).length;
  }


  let cardHeader1 = CardService.newCardHeader()
    .setTitle('Generate Notice')
    .setImageStyle(CardService.ImageStyle.CIRCLE);

  let cardSection2TextInput1 = CardService.newTextInput();
  let cardSection2TextInput2 = CardService.newTextInput();
  let cardSection2TextInput3 = CardService.newTextInput();
  let cardSection2DatePTimePicker1 = CardService.newDateTimePicker()

  if (formInputLength > 1) {
    cardSection2TextInput1 = cardSection2TextInput1.setValue(formInput.name);
    cardSection2TextInput2 = cardSection2TextInput2.setValue(formInput.profiles)
    cardSection2TextInput3 = cardSection2TextInput3.setValue(formInput.package)
    cardSection2DatePTimePicker1 = cardSection2DatePTimePicker1.setValueInMsSinceEpoch(formInput.dateTime.msSinceEpoch);
  }

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

  let cardSection2ButtonList1Button1Action1 = CardService.newAction()
    .setFunctionName('saveData')
    .setParameters({ "e": JSON.stringify(e) })

  let cardSection2ButtonList1Button1 = CardService.newTextButton()
    .setText('Save Data')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(cardSection2ButtonList1Button1Action1);

  let cardSection2Divider1 = CardService.newDivider();

  let param1 = "doc";
  let param2 = "sheet";

  let cardSection2DecoratedText1Button1 = CardService.newImageButton().setOpenLink(CardService.newOpenLink()
    .setUrl(`https://script.google.com/a/macros/cloudfort.in/s/AKfycbwpnfPkC1DY4WEOC4HO9CCR95deYlW4wroZiyCJASE/dev?param=${param1}`)
    .setOpenAs(CardService.OpenAs.OVERLAY)
    .setOnClose(CardService.OnClose.RELOAD))
    .setIconUrl("https://i.ibb.co/56ykKjx/211608-folder-icon.png")
    .setAltText('Select Template');

  let cardSection2DecoratedText1 = CardService.newDecoratedText()
    .setText('Select the Template')
    .setBottomLabel('Choose a Doc file')
    .setButton(cardSection2DecoratedText1Button1);

  let filePick = JSON.parse(PropertiesService.getUserProperties().getProperty("filePick"));
  let prop = PropertiesService.getUserProperties().getProperty("files");

  if (prop == null || !filePick) {
    clearFilesFromPropServ();

  } else {
    let docs = JSON.parse(prop);
    console.log("docouter", docs[0])

    cardSection2DecoratedText1 = CardService.newDecoratedText().setText(docs[0].name)
      .setButton(cardSection2DecoratedText1Button1);
  }

  // PropertiesService.getUserProperties().setProperty("filePick", JSON.stringify(false));

  let cardSection2Divider2 = CardService.newDivider();

  let cardSection2DecoratedText2Button1 = CardService.newImageButton().setOpenLink(CardService.newOpenLink()
    .setUrl(`https://script.google.com/a/macros/cloudfort.in/s/AKfycbwpnfPkC1DY4WEOC4HO9CCR95deYlW4wroZiyCJASE/dev?param=${param2}`)
    .setOpenAs(CardService.OpenAs.OVERLAY)
    .setOnClose(CardService.OnClose.RELOAD))
    .setIconUrl("https://i.ibb.co/56ykKjx/211608-folder-icon.png")
    .setAltText('Select Sheet');

  let cardSection2DecoratedText2 = CardService.newDecoratedText()
    .setText('Select the Sheet with Email')
    .setButton(cardSection2DecoratedText2Button1);

  let sheetFilePick = JSON.parse(PropertiesService.getUserProperties().getProperty("sheetFilePick"));
  let sheetProp = PropertiesService.getUserProperties().getProperty("sheetFiles");

  if (sheetProp == null || !sheetFilePick) {
    clearSheetFilesFromPropServ();
  } else {
    let docs = JSON.parse(sheetProp);
    console.log("docouter", docs[0])
    let url = docs[0].url;
    console.log("url=", url);
    var col = fetchHead(url);
    cardSection2DecoratedText2 = CardService.newDecoratedText().setText(docs[0].name)
      .setButton(cardSection2DecoratedText2Button1);
  }

  // PropertiesService.getUserProperties().setProperty("sheetFilePick", JSON.stringify(false));
  let condition = false;

  let cardSection1SelectionInput1 = CardService.newSelectionInput()
    .setFieldName('Select ')
    .setTitle('Select Email Column')
    .setType(CardService.SelectionInputType.DROPDOWN);

  if (!col) {
    col = [""]
  }

  col.forEach((element) => {
    cardSection1SelectionInput1 = cardSection1SelectionInput1.addItem(element, element, condition);
  })

  let cardSection1ButtonList1Button1Action1 = CardService.newAction()
    .setFunctionName('createNotice')
    .setParameters({ "e": JSON.stringify(e) })

  let cardSection1ButtonList1Button1 = CardService.newTextButton()
    .setText('Save')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(cardSection1ButtonList1Button1Action1);

  let status = 1;

  let cardSection1ButtonList1Button2Action1 = CardService.newAction()
    .setFunctionName('clearCard')
    .setParameters({ "status": JSON.stringify(status) })

  let cardSection1ButtonList1Button2 = CardService.newTextButton()
    .setText('Clear')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(cardSection1ButtonList1Button2Action1);

  let cardSection1ButtonList1 = CardService.newButtonSet()
    .addButton(cardSection1ButtonList1Button1)
    .addButton(cardSection1ButtonList1Button2);

  let cardSection2 = CardService.newCardSection()
    .addWidget(cardSection2TextInput1)
    .addWidget(cardSection2TextInput2)
    .addWidget(cardSection2TextInput3)
    .addWidget(cardSection2DatePTimePicker1)
    .addWidget(cardSection2Divider1)
    .addWidget(cardSection2ButtonList1Button1)
    .addWidget(cardSection2Divider1)
    .addWidget(cardSection2DecoratedText1)
    .addWidget(cardSection2Divider2)
    .addWidget(cardSection2DecoratedText2)
    .addWidget(cardSection2Divider2)
    .addWidget(cardSection1SelectionInput1)
    .addWidget(cardSection2Divider2)
    .addWidget(cardSection1ButtonList1);

  let card = CardService.newCardBuilder()
    .setHeader(cardHeader1)
    .addSection(cardSection2)
    .build();
  return card;
}

function saveData(e) {
  PropertiesService.getUserProperties().setProperties({ "inputData": JSON.stringify(e.formInput) });

}

function clearCard(e) {
  PropertiesService.getUserProperties().deleteProperty("inputData");
  PropertiesService.getUserProperties().setProperty("filePick", JSON.stringify(false));
  PropertiesService.getUserProperties().setProperty("sheetFilePick", JSON.stringify(false));
  return CardService.newNavigation().updateCard(onhomepage(e));
}

