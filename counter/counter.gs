const SMIIRL_COUNTER_ID = 'e08e3c3acc55';
const SMIIRL_COUNTER_TOKEN = '481c8dd48765ee7fae41aa37032a67bc';

function updateSmiirlFromShops() {
  Logger.clear();
  Logger.log("=== Начало updateSmiirlFromShops ===");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Shops");
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log") ||
                   SpreadsheetApp.getActiveSpreadsheet().insertSheet("Log");

  const allRows = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();

  let totalCount = 0;
  const errors = [];
  let offset = 0;

  allRows.forEach(([shopUrl, token]) => {
    if (!shopUrl || !token) return;

    if (shopUrl.toLowerCase() === 'offset') {
      offset = parseInt(token) || 0;
      Logger.log("Offset: " + offset);
      return;
    }

    const apiUrl = `https://${shopUrl}/admin/api/2024-04/orders/count.json?status=any`;
    Logger.log("Запрос к: " + apiUrl);

    try {
      const response = UrlFetchApp.fetch(apiUrl, {
        method: 'get',
        headers: { 'X-Shopify-Access-Token': token },
        muteHttpExceptions: true
      });

      const raw = response.getContentText();
      Logger.log(`Ответ от ${shopUrl}: ${raw}`);

      let result;
      try {
        result = JSON.parse(raw);
      } catch (jsonErr) {
        errors.push(`${shopUrl}: ошибка парсинга JSON`);
        return;
      }

      if (result && result.count !== undefined) {
        Logger.log(`${shopUrl} → ${result.count}`);
        totalCount += result.count;
      } else {
        errors.push(`${shopUrl}: нет поля count`);
      }
    } catch (e) {
      errors.push(`${shopUrl}: ${e.toString()}`);
    }
  });

  totalCount += offset;
  Logger.log("Итоговое значение (с учётом offset): " + totalCount);

  const smiirlUrl = `http://api.smiirl.com/${SMIIRL_COUNTER_ID}/set-number/${SMIIRL_COUNTER_TOKEN}/${totalCount}`;
  Logger.log("Обновление Smiirl: " + smiirlUrl);

  try {
    const smiirlResponse = UrlFetchApp.fetch(smiirlUrl);
    Logger.log("Smiirl ответ: " + smiirlResponse.getContentText());
  } catch (e) {
    const errMsg = `Ошибка при обновлении Smiirl: ${e.toString()}`;
    Logger.log(errMsg);
    errors.push(errMsg);
  }

// Автоочистка старых строк в листе Log — сохраняем только последние 1000
const maxRows = 1000;
const lastRow = logSheet.getLastRow();
if (lastRow > maxRows) {
  const rowsToDelete = lastRow - maxRows;
  logSheet.deleteRows(2, rowsToDelete); // не трогаем заголовок
}


  logSheet.appendRow([
    new Date(),
    totalCount,
    errors.join(" | ") || "OK"
  ]);

  Logger.log("=== Конец updateSmiirlFromShops ===");
}
