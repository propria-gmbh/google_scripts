function logToSheet(label, content) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Log");

  if (!logSheet) {
    logSheet = ss.insertSheet("Log");
    logSheet.appendRow(["Timestamp", "Label", "Content"]);
  }

  const timestamp = new Date();
  logSheet.appendRow([timestamp, label, content]);
}
