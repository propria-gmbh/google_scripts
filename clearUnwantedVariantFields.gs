function clearUnwantedVariantFields() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const fieldsToClear = ["Variant SKU", "Variant Taxable", "Variant Barcode"];
  const indexes = fieldsToClear.map(field => headers.indexOf(field));

  if (indexes.some(idx => idx === -1)) {
    logToSheet("⚠️ Missing columns", "One or more columns not found: " + fieldsToClear.join(", "));
    return;
  }

  for (let row = 1; row < data.length; row++) {
    indexes.forEach(idx => {
      if (data[row][idx]) {
        sheet.getRange(row + 1, idx + 1).setValue("");
      }
    });
  }

  logToSheet("🧹 Cleared fields", fieldsToClear.join(", "));
}
