function translateOptionsAndValues() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  let targetLang = props.getProperty("targetLanguage");
  if (!targetLang) {
    const response = ui.prompt(
      "🌍 Target language for option translation",
      "Enter one of: English, Italian, French, German, Swedish, Danish",
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    targetLang = response.getResponseText().trim();
    props.setProperty("targetLanguage", targetLang);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const translationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Translations");
  if (!translationSheet) {
    logToSheet("❌ Error", "Sheet 'Translations' not found");
    return;
  }

  const translationsData = translationSheet.getDataRange().getValues();
  const translationHeaders = translationsData[0];
  const targetIdx = translationHeaders.indexOf(targetLang);
  if (targetIdx === -1) {
    logToSheet("❌ Error", `Target language '${targetLang}' not found in Translations`);
    return;
  }

  const knownOptionNameCols = ["Option1 Name", "Option2 Name", "Option3 Name"];
  const knownOptionValueCols = ["Option1 Value", "Option2 Value", "Option3 Value"];
  const optionNameIndices = knownOptionNameCols.map(name => headers.indexOf(name)).filter(i => i !== -1);
  const optionValueIndices = knownOptionValueCols.map(name => headers.indexOf(name)).filter(i => i !== -1);

  const translatedNames = new Set();
  const translatedValues = new Set();
  const missingNames = new Set();
  const missingValues = new Set();

  for (let r = 1; r < data.length; r++) {
    for (const i of optionNameIndices) {
      const orig = data[r][i]?.toString().trim();
      if (!orig) continue;

      const translated = findTranslation(orig, translationsData, targetIdx);
      if (translated) {
        data[r][i] = translated;
        translatedNames.add(orig);
      } else {
        missingNames.add(orig);
      }
    }

    for (const i of optionValueIndices) {
      const orig = data[r][i]?.toString().trim();
      if (!orig) continue;

      const translated = findTranslation(orig, translationsData, targetIdx);
      if (translated) {
        data[r][i] = translated;
        translatedValues.add(orig);
      } else {
        missingValues.add(orig);
      }
    }
  }

  sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));

  logToSheet("🌍 Option names translated", translatedNames.size ? Array.from(translatedNames).join(", ") : "None");
  logToSheet("❌ Option names missing", missingNames.size ? Array.from(missingNames).join(", ") : "None");

  logToSheet("🌍 Option values translated", translatedValues.size ? Array.from(translatedValues).join(", ") : "None");
  logToSheet("❌ Option values missing", missingValues.size ? Array.from(missingValues).join(", ") : "None");
}

function findTranslation(original, translationsData, targetIdx) {
  for (let i = 1; i < translationsData.length; i++) {
    const row = translationsData[i];
    if (row.some(cell => typeof cell === "string" && cell.trim().toLowerCase() === original.toLowerCase())) {
      return row[targetIdx];
    }
  }
  return null;
}
