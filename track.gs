const SHEETS_NAME = "Sheet1";
const SPREADSHEET_ID = "1Nxxqz7qRCd1cqm3IGoZZNwIOgUxtKCpcZ3lMjJN0GLg";

function doGet(e) {
  var recipient = e.parameter.recipient;
  
  // Update status "Opened" di Google Sheets jika email penerima membuka email
  if (recipient) {
    updateStatusIfEmailOpened(recipient);
  }
  
  // Memberikan respons pixel (gambar 1x1 piksel transparan)
  var pixelUrl = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  return ContentService.createTextOutput(pixelUrl).setMimeType(ContentService.MimeType.GIF);
}

function updateStatusIfEmailOpened(recipient) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS_NAME);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // Get headers to find the column index of "Email"
  var headers = values[0];
  var emailColumnIndex = headers.indexOf(RECIPIENT_COL);

  // Cari baris dengan email penerima yang sesuai
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var emailValue = row[emailColumnIndex];
    if (emailValue && emailValue.toLowerCase() === recipient.toLowerCase()) {
      var columnIndex = getColumnIndexByName(sheet, STATUS_COL);
      if (columnIndex !== -1) {
        sheet.getRange(i + 1, columnIndex + 1).setValue("Opened");
        break;
      }
    }
  }
}

// Function to get column index by name
function getColumnIndexByName(sheet, columnName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnIndex = headers.indexOf(columnName);
  return columnIndex;
}
