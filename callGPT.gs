function callGPT(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: "gpt-4",
      messages: [{ role: "user", content: prompt }]
    }),
    headers: { Authorization: "Bearer " + apiKey },
    timeout: 60000
  });

  const ui = SpreadsheetApp.getUi();
  const rawText = response.getContentText();
  logToSheet("📦 RAW RESPONSE", rawText);

  try {
    const json = JSON.parse(rawText);

    // Если пришла ошибка от OpenAI
    if (json.error) {
      const message = json.error.message || "Unknown error";
      const code = json.error.code || "unknown";

      logToSheet("❌ GPT ERROR", `${code}: ${message}`);

      switch (code) {
        case "insufficient_quota":
          ui.alert("❌ Ошибка GPT: превышена квота.\n\nЗайди на https://platform.openai.com/account/usage и проверь баланс.");
          break;
        case "invalid_api_key":
          ui.alert("❌ Неверный API-ключ OpenAI. Проверь настройки в Script Properties.");
          break;
        case "rate_limit_exceeded":
          ui.alert("⚠️ Превышен лимит скорости запросов к OpenAI. Подожди 10–20 секунд и повтори.");
          break;
        default:
          ui.alert(`❌ Ошибка GPT [${code}]:\n${message}`);
      }

      return null;
    }

    const text = json.choices[0].message.content;
    logToSheet("✅ GPT RESPONSE", text);

    return parseGptResponse(text);
  } catch (e) {
    logToSheet("❌ PARSE ERROR", e.toString());
    ui.alert("❌ Ошибка при разборе ответа от GPT. См. вкладку Log.");
    return null;
  }
}
