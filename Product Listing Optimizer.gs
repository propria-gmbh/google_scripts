function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Product Listing")
    .addItem("Process First Product", "processFirstProductBlock")
    .addToUi();
}

function processFirstProductBlock() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();

  const currentTargetLang = props.getProperty("targetLanguage") || "Danish";
  const langResponse = ui.prompt("🎯 Confirm or change target language", `Current value: ${currentTargetLang}`, ui.ButtonSet.OK_CANCEL);
  if (langResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetLanguage = langResponse.getResponseText().trim() || currentTargetLang;
  props.setProperty("targetLanguage", targetLanguage);

  const currentOriginalLang = props.getProperty("originalLanguage") || "Swedish";
  const origLangResponse = ui.prompt("📝 Confirm or change original language", `Current value: ${currentOriginalLang}`, ui.ButtonSet.OK_CANCEL);
  if (origLangResponse.getSelectedButton() !== ui.Button.OK) return;
  const originalLanguage = origLangResponse.getResponseText().trim() || currentOriginalLang;
  props.setProperty("originalLanguage", originalLanguage);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const titleIdx = headers.indexOf("Title");
  const bodyIdx = headers.indexOf("Body (HTML)");
  const handleIdx = headers.indexOf("Handle");
  const imageIdx = headers.indexOf("Image Src");
  const tagsIdx = headers.indexOf("Tags");
  const typeIdx = headers.indexOf("Type");

  let startRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][titleIdx]) {
      startRow = i;
      break;
    }
  }

  if (startRow === -1) {
    ui.alert("❌ Не найдено ни одной строки с Title");
    return;
  }

  let endRow = startRow + 1;
  while (endRow < data.length && !data[endRow][titleIdx]) endRow++;

  const productRows = data.slice(startRow, endRow);
  const mainRow = productRows[0];
  const image1 = mainRow[imageIdx];
  const image2 = productRows[1] ? productRows[1][imageIdx] : "";
  const title = mainRow[titleIdx];
  const body = mainRow[bodyIdx];

  const prompt = buildUniversalPrompt(title, body, image1, image2, targetLanguage, originalLanguage);
  const result = callGPT(prompt);

  if (!result) {
    ui.alert("❌ GPT не вернул корректный результат");
    return;
  }

  // Генерация тегов и типа товара
  const genderTag = /dame|woman|femme|donna|damen/i.test(result.title + result.body) ? "damen" :
                    /herr|man|homme|uomo|herren/i.test(result.title + result.body) ? "herren" : "unisex";
  const typeMatch = result.title.toLowerCase().match(/blazer|shirt|dress|jumpsuit|jeans|pants|skirt|coat|sweater|jacket|tunic|kimono|top|polo|hoodie|t-shirt/);
  const typeTag = typeMatch ? typeMatch[0] : "apparel";
  const finalTags = genderTag === "unisex" ? `damen,herren,unisex,${typeTag}` : `${genderTag},${typeTag}`;

  // Запись результата в таблицу
  sheet.getRange(startRow + 1, titleIdx + 1).setValue(result.title);
  sheet.getRange(startRow + 1, bodyIdx + 1).setValue(result.body);
  sheet.getRange(startRow + 1, tagsIdx + 1).setValue(finalTags);
  sheet.getRange(startRow + 1, typeIdx + 1).setValue(capitalizeFirst(typeTag));

  for (let r = startRow; r < endRow; r++) {
    sheet.getRange(r + 1, handleIdx + 1).setValue(result.handle);
  }

  ui.alert("✅ Готово. Первый товар обновлён, включая теги и тип.");
}

function capitalizeFirst(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function buildUniversalPrompt(title, body, image1, image2, targetLanguage, originalLanguage) {
  return `You are a professional product listing copywriter, ranked among the top 5 in the world for e-commerce conversion copywriting.
You specialize in high-performing Shopify listings, especially for dropshipping products in competitive niches.

Your task is to transform a competitor’s product listing into a completely new, high-converting product title and HTML product description.
The result must be in the specified TARGET LANGUAGE and adapted for the TARGET MARKET and audience.

Your writing must sound 100% natural and native, perfectly matching the tone, expectations, and habits of local shoppers.
It must not resemble AI-generated text — it should feel human-written, brand-aligned, and conversion-focused.

ORIGINAL LANGUAGE: ${originalLanguage}
TARGET LANGUAGE: ${targetLanguage}
TARGET MARKET: auto-detect based on language

The goal is not a literal or word-for-word translation.
Instead, the text must be fully localized and rewritten to suit the culture, tone, and shopping behavior of the target market.

Use your expertise as a copywriter to adapt, rephrase, and reframe the content so it resonates deeply with the audience.
Rewriting is encouraged — only the product’s core idea should remain.

COMPETITOR LISTING:
Title: ${title}
Description: ${body}
Image 1: ${image1}
Image 2: ${image2}

You must not reuse or copy anything from this listing directly — rephrase and recreate everything.

🏷️ PRODUCT TITLE FORMAT:
[Name] | [Gender] [Feature] [Feature] [Product Type] | [Attribute]

• Remove brackets around features
• Remove ™, ®, ©
• Do NOT include banned brand names or materials
• Title must be emotional, localized, and sound native
• First name must always be newly invented, never copied from the input listing
• It must be appropriate for the target audience and market
• Reusing the original name is strictly forbidden
• Gender words must match the grammar of the target language

🔗 HANDLE FORMAT:
• All lowercase
• Remove all special characters
• Replace spaces with dashes
• Auto-generate it based on the final product title

📝 HTML DESCRIPTION FORMAT:
<p><strong>[HEADLINE – ALL CAPS]</strong></p>
<p>[Intro paragraph: what the product is, who it’s for, and what makes it great]</p>
<img src="${image1}" style="display: block; margin: 1em 0; text-align: left;" />
<p><strong>WHY CHOOSE THE [PRODUCT NAME + TYPE]?</strong></p>
<p>✓ [Benefit 1]</p>
<p>✓ [Benefit 2]</p>
<p>✓ [Benefit 3]</p>
<p>✓ [Benefit 4]</p>
<p>✓ [Benefit 5]</p>
<img src="${image2}" style="display: block; margin: 1em 0; text-align: left;" />
<p><strong>[FINAL CALL TO ACTION – ALL CAPS]</strong></p>
<p>[Short motivational sentence encouraging to buy now]</p>

📌 MANDATORY RULES:
• All <strong> headlines must be ALL CAPS
• Each ✓ benefit must be a separate <p> element — no <ul> or <li>
• Never use emojis
• Never use ALL CAPS in paragraphs (only in headlines)
• No brackets ([]) or parentheses in the output
• Product type must be auto-detected from context
• Do not proceed if one or both image URLs are missing
• If target language = original language, still rephrase and adapt — do not reuse the source verbatim

🚫 FORBIDDEN MATERIALS:
Linen, Cashmere, Cotton, Wool, Polyester, Spandex, Leather, Faux, Viscose, Silk, Denim, Fur,
Nylon, Acetate, EVA, Fleece, Tweed, Sherpa, Lace, Satin, Velvet, Rayon, Teddy

🚫 FORBIDDEN BRAND NAMES:
Coco, Chanel, Celine, Elara, Zara, Ami, Brioni, Chloé, Kenzo, Santoni, Tod, Vince, Zilli, Calvin

✅ FORMAT FOR COPY-FRIENDLY OUTPUT:
Wrap the final Title and Handle in code-ready blocks like this:

Title: \`[insert title]\`
Handle: \`[insert handle]\`
`;  
}


function parseGptResponse(text) {
  const titleMatch = text.match(/Title:\s*`([^`]+)`/);
  const handleMatch = text.match(/Handle:\s*`([^`]+)`/);
  const bodyStart = text.indexOf("<p><strong>");
  const body = bodyStart !== -1 ? text.slice(bodyStart).trim() : "";

  const title = titleMatch ? titleMatch[1] : null;
  const handle = handleMatch ? handleMatch[1] : null;

  if (title && handle && body) {
    return { title, handle, body };
  } else {
    return null;
  }
}
