function getSpreadsheetUrls() {
  var folderId = "1eP4sqLorjK49ovA9qUd1t49xk-6bT0zz"; // 取得したいフォルダのIDを入力(フォルダIDはURLの最後のみ)
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // 既存データをクリア
  sheet.appendRow(["ファイル名", "URL"]);
  
  while (files.hasNext()) {
    var file = files.next();
    sheet.appendRow([file.getName(), file.getUrl()]);
  }
}
