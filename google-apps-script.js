/**
 * GOOGLE APPS SCRIPT ДЛЯ СБОРА ЛИДОВ С КВИЗА АЛЬФА СЕЛЛЕР
 * 
 * Инструкция по настройке за 2 минуты:
 * 1. Откройте Google Таблицу (или создайте новую на Google Диске: sheets.new).
 * 2. В меню выберите: «Расширения» (Extensions) -> «Apps Script».
 * 3. Удалите весь существующий там код и вставьте содержимое этого файла.
 * 4. Нажмите кнопку «Развернуть» (Deploy) -> «Новое развертывание» (New deployment).
 * 5. В типе выберите «Веб-приложение» (Web app).
 * 6. Настройки:
 *    - Описание: "Alfa Seller Quiz Webhook"
 *    - Запуск от имени: "Я" (Me)
 *    - У кого есть доступ (Who has access): "Все" (Anyone) — ЭТО ВАЖНО, чтобы квиз мог слать данные без авторизации.
 * 7. Нажмите «Развернуть» (Deploy), разрешите доступ (Authorize access).
 * 8. Скопируйте полученный «URL веб-приложения» (Web app URL) — вида https://script.google.com/macros/s/.../exec
 * 9. Вставьте этот URL в настройках квиза (иконка шестерёнки в шапке).
 */

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000); // Предотвращает конфликты при одновременной отправке с нескольких планшетов

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Если таблица пустая, автоматически создаем шапку со стилями
    if (sheet.getLastRow() === 0) {
      var headers = [
        "ID Лида",
        "Дата и время",
        "Имя клиента",
        "Номер телефона",
        "Telegram",
        "Логин в МАКС",
        "Компания / Бренд",
        "Сегмент (Лид)",
        "Основной продукт",
        "Все рекомендации",
        "Индустрия / Ниша",
        "Где продаёт",
        "Оборот в месяц",
        "Состояние сайта",
        "Главные боли",
        "Планы на полгода",
        "Бюджет на маркетинг",
        "Клиент банка"
      ];
      
      sheet.appendRow(headers);
      
      // Стилизуем шапку в цвета Альфа-Банка
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#1B1B21");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);
    }

    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var t = data.answersText || {};

    var row = [
      data.id || "",
      data.createdDate || new Date().toLocaleString("ru-RU", { timeZone: "Europe/Moscow" }),
      data.name || "",
      data.phone || "",
      data.telegram || "",
      data.maxLogin || "",
      data.company || "",
      data.temperatureLabel || data.temperature || "",
      data.mainProduct || "",
      (data.productNames || []).join(", "),
      data.industry || t.industry || "",
      t.channels || data.channels || "",
      t.turnover || data.turnover || "",
      t.site || data.siteState || "",
      t.pains || data.pains || "",
      t.plans || data.plans || "",
      t.budget || data.budget || "",
      t.bank || data.bankClient || ""
    ];

    sheet.appendRow(row);

    // Подсветка строки по температуре лида
    var lastRow = sheet.getLastRow();
    if (data.temperature === "hot") {
      sheet.getRange(lastRow, 8).setBackground("#FFD2CE").setFontColor("#900"); // Светло-красный для горячих
    } else if (data.temperature === "warm") {
      sheet.getRange(lastRow, 8).setBackground("#F4FFCC").setFontColor("#3E6B00"); // Светло-лаймовый для теплых
    }

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({ status: "success", row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    message: "Alfa Seller Quiz Webhook is active and waiting for POST requests."
  })).setMimeType(ContentService.MimeType.JSON);
}
