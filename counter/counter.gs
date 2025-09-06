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

    const totalUrl = `https://${shopUrl}/admin/api/2024-04/orders/count.json?status=any`;
    const cancelledUrl = `https://${shopUrl}/admin/api/2024-04/orders/count.json?status=cancelled`;
    Logger.log("Запрос к (все): " + totalUrl);
    Logger.log("Запрос к (отмененные): " + cancelledUrl);

    try {
      const totalResponse = UrlFetchApp.fetch(totalUrl, {
        method: 'get',
        headers: { 'X-Shopify-Access-Token': token },
        muteHttpExceptions: true
      });
      const cancelledResponse = UrlFetchApp.fetch(cancelledUrl, {
        method: 'get',
        headers: { 'X-Shopify-Access-Token': token },
        muteHttpExceptions: true
      });

      const totalRaw = totalResponse.getContentText();
      const cancelledRaw = cancelledResponse.getContentText();
      Logger.log(`Ответ (все) от ${shopUrl}: ${totalRaw}`);
      Logger.log(`Ответ (отмененные) от ${shopUrl}: ${cancelledRaw}`);

      let totalResult;
      let cancelledResult;
      try {
        totalResult = JSON.parse(totalRaw);
        cancelledResult = JSON.parse(cancelledRaw);
      } catch (jsonErr) {
        errors.push(`${shopUrl}: ошибка парсинга JSON`);
        return;
      }

      if (totalResult && totalResult.count !== undefined && cancelledResult && cancelledResult.count !== undefined) {
        const nonCancelled = Math.max(0, Number(totalResult.count) - Number(cancelledResult.count));
        Logger.log(`${shopUrl} → всего: ${totalResult.count}, отмененные: ${cancelledResult.count}, неотмененные: ${nonCancelled}`);
        totalCount += nonCancelled;
      } else {
        errors.push(`${shopUrl}: нет поля count в одном из ответов`);
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
