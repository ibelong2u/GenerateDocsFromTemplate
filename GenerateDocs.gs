function generateDocs() {

  /***********************************************
  Update template file and destination folder url:
  ************************************************/

  const GoogleDocsTemplate = 'https://docs.google.com/document/d/...';
  
  const GoogleDriveOutputFolder = 'https://drive.google.com/drive/folders/...';

  /***********************************************/


  let ui = SpreadsheetApp.getUi();
  let templateID, destinationID;

  try {
    templateID = DriveApp.getFileById(GoogleDocsTemplate.match(/[-\w]{25,}/));
  }
  catch(e) {
    ui.alert("Check GoogleDocsTemplate url", "The script is not able to access the file.", ui.ButtonSet.OK);
    return;
  }

  try {
    destinationID = DriveApp.getFolderById(GoogleDriveOutputFolder.match(/[\w-_]{15,}/));
  }
  catch(e) {
    ui.alert("Check GoogleDriveOutputFolder url", "The script is not able to access the folder.", ui.ButtonSet.OK);
    return;
  }

  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const sourceRows = sourceSheet.getDataRange().getValues();

  const headers = sourceRows[0];
  
  const indexFilename = headers.indexOf("Filename");

  if(indexFilename == -1) {
    ui.alert("Filename of generated Docs?", "Your Sheet must have a 'Filename' column.", ui.ButtonSet.OK);
    return;
  }

  const indexGeneratedGoogleDocs = headers.indexOf("Generated Google Docs");

  let docsGenerated = 0;

  try {

    for(var i = 1; i < sourceRows.length; i++) {

      if (sourceRows[i][indexFilename] == '')
        continue;

      if (indexGeneratedGoogleDocs != -1  && sourceRows[i][indexGeneratedGoogleDocs] != '') {
      
        let redo = ui.alert(sourceRows[i][indexFilename] + " doc url already exists!", "Do you want to create a new one? (Yes = Create new file, No = Skip this one, Cancel = Abort the script)", ui.ButtonSet.YES_NO_CANCEL);

        if (redo == ui.Button.CANCEL)
          return;

        if (redo == ui.Button.NO)
          continue;

      }

      const outputDoc = templateID.makeCopy(sourceRows[i][indexFilename], destinationID);
      const outputID = DocumentApp.openById(outputDoc.getId());
      const outputBody = outputID.getBody();
    
      headers.forEach(function(headername, headerindex) {
        outputBody.replaceText("{{" + headername + "}}", sourceRows[i][headerindex] );
      })

      docsGenerated++;

      outputID.saveAndClose();

      // console.log(outputDoc.getName() + " generated");

      if(indexGeneratedGoogleDocs == -1)
        continue;

      sourceSheet.getRange(i + 1, indexGeneratedGoogleDocs + 1).setValue(outputDoc.getUrl());

    }

    ui.alert("Done! 🎉", `Generated ${docsGenerated} docs.`, ui.ButtonSet.OK);

  }
  catch(e) {
     ui.alert("Sorry, something went wrong!", e, ui.ButtonSet.OK);
  }
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('Custom Scripts');
  menu.addItem('Generate docs from template', 'generateDocs');
  menu.addToUi();
}
