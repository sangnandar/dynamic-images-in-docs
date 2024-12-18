const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheetName = 'Sheet1';

const docsTemplateId = '1CW9z7_LHVccwIwjnmoFLFaA8P1LyJ1dQponbraKzSCQ'; // docs template file
const config = [
  {
    imageName: 'Figure 1',
    bookmarkId: 'id.wglkb02zor96'
  },
  {
    imageName: 'Figure 2',
    bookmarkId: 'id.ifl1btgwoni'
  },
  {
    imageName: 'Figure 3',
    bookmarkId: 'kix.jithfpq0bbwm'
  }
];

// image file types
const mimeTypes = [
  MimeType.PNG,
  MimeType.JPEG,
  MimeType.BMP,
  MimeType.GIF
];

const imageFolderId = '1f7e-keb55S_OtqgszN8PRYKGZbYuEjmt'; // stored images for each customers
const docsFolderId = '1fiJhcZzpVu79kqFJjKtLk5pKmgS2aJGc'; // created docs goes here

// templates for modal dialog
const htmlTemplate = HtmlService.createTemplate(`
  <div><a href="<?!= docsUrl ?>" target="_blank"><?!= customerName ?></a></div>
`);
const htmlTemplateNoFolder = HtmlService.createTemplate(`
  <div><?!= customerName ?>: Folder not found.</div>
`);
const htmlTemplateNoImage = HtmlService.createTemplate(`
  <div style="padding-left: 10px;"><?!= imageName ?> not found.</div>
`);

var ui; // return null if called from script editor
try {
  ui = SpreadsheetApp.getUi();
} catch (e) {
  Logger.log('You are using script editor.');
}
