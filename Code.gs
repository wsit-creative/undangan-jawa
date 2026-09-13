const SHEET_NAME = 'Sheet1';

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const body = JSON.parse(e.postData.contents || '{}');
  sheet.appendRow([new Date(), body.name || '', body.message || '']);
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const data = rows.slice(1).map(r => ({
    time: r[0], name: r[1], message: r[2]
  }));
  const callback = e && e.parameter && e.parameter.callback;
  const output = JSON.stringify({ok:true, data:data});
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
