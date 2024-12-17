function onOpen() {
  ui
    .createMenu('Custom Menu')
      .addItem('Create documents', 'createDocsFromTemplate')
    .addToUi();
}

function createDocsFromTemplate() {
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
  const imageFolder = DriveApp.getFolderById(imageFolderId);

  let html;
  if (ui) {
    html = HtmlService.createHtmlOutput('');
  }

  for (const [customerName] of data.filter(item => item[1])) { // process only checked items
    let docsCopy, imageNotFound;
    const customerFolders = imageFolder.getFoldersByName(customerName);
    const customerFolderExist = customerFolders.hasNext();

    if (customerFolderExist) {
      docsCopy = DriveApp.getFileById(docsTemplateId).makeCopy(customerName, DriveApp.getFolderById(docsFolderId));
      const customerFolder = customerFolders.next();
      const bookmarks = DocumentApp.openById(docsCopy.getId())
        .getTabs()[0].asDocumentTab() // first tab
        .getBookmarks();

      imageNotFound = [];
      for (const bookmark of bookmarks.reverse()) {
        const position = bookmark.getPosition();
        const imageName = config.find(item => item.bookmarkId === bookmark.getId())?.imageName;
        const str = mimeTypes.map(type => `mimeType='${type}'`).join(' or ');
        const searchQuery = `title contains '${imageName}' and (${str})`;
        const images = customerFolder.searchFiles(searchQuery);

        if (images.hasNext()) {
          const file = images.next();
          position.getElement().asParagraph().clear(); // clear content
          const inlineImage = position.insertInlineImage(file.getBlob());
          const aspectRatio = inlineImage.getHeight() / inlineImage.getWidth();
          inlineImage
            .setWidth(200)
            .setHeight(200 * aspectRatio);

        } else {
          imageNotFound.push(imageName);
        }
      }
    }

    // output modal dialog or log the results
    if (ui) { // called from spreadsheet's ui
      if (customerFolderExist) {
        htmlTemplate.docsUrl = docsCopy.getUrl();
        htmlTemplate.customerName = customerName;
        html.append(htmlTemplate.evaluate().getContent());

        if (imageNotFound.length) { // images not found
          for (const imageName of imageNotFound.reverse()) {
            htmlTemplateNoImage.imageName = imageName;
            html.append(htmlTemplateNoImage.evaluate().getContent());
          }
        }

      } else { // Folder not found
        htmlTemplateNoFolder.customerName = customerName;
        html.append(htmlTemplateNoFolder.evaluate().getContent());
      }

      html.append('<div>&nbsp;</div>');

    } else { // called from script editor
      if (customerFolderExist) {
        Logger.log(`${customerName}: ${docsCopy.getUrl()}`);
        if (imageNotFound.length) {
          for (const imageName of imageNotFound.reverse()) {
            Logger.log(`${imageName} not found.`);
          }
        }

      } else {
        Logger.log(`${customerName}: Folder not found.`);
      }
    }
  }

  if (ui) {
    ui.showModalDialog(
      html
        .setWidth(300)
        .setHeight(700),
      'Documents created'
    );
  }

  // set all chekboxes to false
  sheet.getRange('B2:B' + sheet.getLastRow()).setValues(
    Array.from({ length: data.length }, () => [false])
  );

}
