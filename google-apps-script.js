/**
 * GOOGLE APPS SCRIPT ДЛЯ СБОРА ЛИДОВ С КВИЗА АЛЬФА-СЕЛЛЕР
 * 
 * Инструкция по обновлению скрипта в вашей Google Таблице:
 * 1. Откройте Google Таблицу.
 * 2. В верхнем меню выберите: «Расширения» (Extensions) -> «Apps Script».
 * 3. Замените весь код на содержимое этого файла.
 * 4. Нажмите «Сохранить» (значок дискеты Ctrl+S / Cmd+S).
 * 5. Нажмите: «Развернуть» (Deploy) -> «Управление развертываниями» (Manage deployments) ->
 *    значок карандаша (Редактировать) -> в поле «Версия» выберите «Новая версия» (New version) -> «Развернуть» (Deploy).
 * 
 * Готово! Все поля (включая «Email», «Номер лотереи», новые вопросы) будут автоматически записываться в таблицу.
 */

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000); // Предотвращает конфликты при одновременной отправке с нескольких устройств

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var defaultHeaders = [
      "ID Лида",
      "Дата и время",
      "Имя клиента",
      "Номер телефона",
      "Email",
      "Telegram / МАКС",
      "Компания / Бренд",
      "Номер лотереи",
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
      "Кто ведёт сайт и рекламу",
      "Клиент банка"
    ];

    // Если таблица абсолютно пустая — создаём шапку
    if (lastRow === 0) {
      sheet.appendRow(defaultHeaders);
      var headerRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
      headerRange.setBackground("#121212"); // Фирменный глубокий графит
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);
      lastRow = 1;
      lastCol = defaultHeaders.length;
    }

    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var t = data.answersText || {};
    var msgr = data.messenger || data.telegram || data.maxLogin || "";
    var lottery = data.lottery || data.lotteryNumber || "";

    // Читаем текущие заголовки таблицы
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    
    // Проверяем наличие колонки Email
    var emailColIdx = -1;
    var lotteryColIdx = -1;
    for (var i = 0; i < currentHeaders.length; i++) {
      var h = String(currentHeaders[i]).toLowerCase();
      if (h.indexOf("email") !== -1 || h.indexOf("почта") !== -1) {
        emailColIdx = i;
      }
      if (h.indexOf("лотере") !== -1 || h.indexOf("билет") !== -1 || h.indexOf("купон") !== -1) {
        lotteryColIdx = i;
      }
    }

    // Если в таблице ещё нет колонки Email, аккуратно вставляем её после колонки «Телефон»
    if (emailColIdx === -1 && lastRow > 0) {
      var phoneColIdx = -1;
      for (var p = 0; p < currentHeaders.length; p++) {
        var ph = String(currentHeaders[p]).toLowerCase();
        if (ph.indexOf("телефон") !== -1 || ph.indexOf("phone") !== -1) {
          phoneColIdx = p + 1; // 1-based index
          break;
        }
      }
      if (phoneColIdx > 0) {
        sheet.insertColumnAfter(phoneColIdx);
        sheet.getRange(1, phoneColIdx + 1).setValue("Email").setBackground("#121212").setFontColor("#FFFFFF").setFontWeight("bold");
      } else {
        sheet.getRange(1, currentHeaders.length + 1).setValue("Email").setBackground("#121212").setFontColor("#FFFFFF").setFontWeight("bold");
      }
      lastCol = sheet.getLastColumn();
      currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // Если в таблице ещё нет колонки «Номер лотереи», добавляем её после «Компания / Бренд»
    if (lotteryColIdx === -1 && lastRow > 0) {
      var compColIdx = -1;
      for (var k = 0; k < currentHeaders.length; k++) {
        var kh = String(currentHeaders[k]).toLowerCase();
        if (kh.indexOf("компан") !== -1 || kh.indexOf("бренд") !== -1) {
          compColIdx = k + 1;
          break;
        }
      }
      if (compColIdx > 0) {
        sheet.insertColumnAfter(compColIdx);
        sheet.getRange(1, compColIdx + 1).setValue("Номер лотереи").setBackground("#121212").setFontColor("#FFFFFF").setFontWeight("bold");
      } else {
        sheet.getRange(1, currentHeaders.length + 1).setValue("Номер лотереи").setBackground("#121212").setFontColor("#FFFFFF").setFontWeight("bold");
      }
      lastCol = sheet.getLastColumn();
      currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // Формируем значения строки строго по заголовкам колонок
    var row = [];
    for (var c = 0; c < currentHeaders.length; c++) {
      var col = String(currentHeaders[c]).trim().toLowerCase();
      if (col.indexOf("id") !== -1) {
        row.push(data.id || "");
      } else if (col.indexOf("дата") !== -1 || col.indexOf("date") !== -1) {
        row.push(data.createdDate || new Date().toLocaleString("ru-RU", { timeZone: "Europe/Moscow" }));
      } else if (col.indexOf("имя") !== -1 || col.indexOf("name") !== -1) {
        row.push(data.name || "");
      } else if (col.indexOf("телефон") !== -1 || col.indexOf("phone") !== -1) {
        row.push(data.phone || "");
      } else if (col.indexOf("email") !== -1 || col.indexOf("почта") !== -1) {
        row.push(data.email || "");
      } else if (col.indexOf("telegram") !== -1 || col.indexOf("макс") !== -1 || col.indexOf("мессенджер") !== -1) {
        row.push(msgr);
      } else if (col.indexOf("логин") !== -1) {
        row.push(data.maxLogin || msgr);
      } else if (col.indexOf("компан") !== -1 || col.indexOf("бренд") !== -1) {
        row.push(data.company || "");
      } else if (col.indexOf("лотере") !== -1 || col.indexOf("билет") !== -1 || col.indexOf("купон") !== -1) {
        row.push(lottery);
      } else if (col.indexOf("сегмент") !== -1 || col.indexOf("температур") !== -1) {
        row.push(data.temperatureLabel || data.temperature || "");
      } else if (col.indexOf("основной продукт") !== -1) {
        row.push(data.mainProduct || "");
      } else if (col.indexOf("рекомендац") !== -1 || col.indexOf("все продукт") !== -1) {
        row.push((data.productNames || []).join(", "));
      } else if (col.indexOf("индустри") !== -1 || col.indexOf("ниша") !== -1 || col.indexOf("сфера") !== -1) {
        row.push(data.industry || t.industry || "");
      } else if (col.indexOf("где") !== -1 || col.indexOf("канал") !== -1) {
        row.push(t.channels || data.channels || "");
      } else if (col.indexOf("оборот") !== -1) {
        row.push(t.turnover || data.turnover || "");
      } else if (col.indexOf("сайт") !== -1 && col.indexOf("кто") === -1) {
        row.push(t.site || data.siteState || "");
      } else if (col.indexOf("боли") !== -1 || col.indexOf("рост") !== -1 || col.indexOf("сдержива") !== -1) {
        row.push(t.pains || data.pains || "");
      } else if (col.indexOf("план") !== -1) {
        row.push(t.plans || data.plans || "");
      } else if (col.indexOf("бюджет") !== -1) {
        row.push(t.budget || data.budget || "");
      } else if (col.indexOf("кто") !== -1 || col.indexOf("команд") !== -1 || col.indexOf("маркетолог") !== -1) {
        row.push(t.team || data.team || "");
      } else if (col.indexOf("банк") !== -1 || col.indexOf("клиент") !== -1) {
        row.push(t.bank || data.bankClient || "");
      } else {
        row.push("");
      }
    }

    sheet.appendRow(row);

    // Подсветка строки по температуре лида
    var newRow = sheet.getLastRow();
    for (var s = 0; s < currentHeaders.length; s++) {
      var sh = String(currentHeaders[s]).toLowerCase();
      if (sh.indexOf("сегмент") !== -1 || sh.indexOf("температур") !== -1) {
        if (data.temperature === "hot") {
          sheet.getRange(newRow, s + 1).setBackground("#FFD2CE").setFontColor("#900");
        } else if (data.temperature === "warm") {
          sheet.getRange(newRow, s + 1).setBackground("#F4FFCC").setFontColor("#3E6B00");
        }
        break;
      }
    }

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({ status: "success", row: newRow }))
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
