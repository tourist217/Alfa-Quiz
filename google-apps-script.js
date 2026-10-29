/**
 * GOOGLE APPS SCRIPT ДЛЯ СБОРА ЛИДОВ И ПРИВЯЗКИ ЛОТЕРЕИ АЛЬФА-СЕЛЛЕР
 * 
 * Инструкция по обновлению скрипта в вашей Google Таблице:
 * 1. Откройте Google Таблицу.
 * 2. В верхнем меню выберите: «Расширения» (Extensions) -> «Apps Script».
 * 3. Замените весь код на содержимое этого файла.
 * 4. Нажмите «Сохранить» (значок дискеты Ctrl+S / Cmd+S).
 * 5. Нажмите: «Развернуть» (Deploy) -> «Управление развертываниями» (Manage deployments) ->
 *    значок карандаша (Редактировать) -> в поле «Версия» выберите «Новая версия» (New version) -> «Развернуть» (Deploy).
 * 
 * Теперь:
 * 1. Лиды создаются со всеми ответами и контактами.
 * 2. Когда участник вводит номер лотереи на странице «Спасибо», скрипт находит созданную строку
 *    и моментально записывает номер лотереи прямо в эту же строку к клиенту!
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
    var lottery = String(data.lottery || data.lotteryNumber || "").trim();

    // Читаем текущие заголовки таблицы
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    
    // Проверяем наличие колонок Email и Номер лотереи
    var emailColIdx = -1;
    var lotteryColIdx = -1;
    var idColIdx = -1;
    var phoneColIdx = -1;

    for (var i = 0; i < currentHeaders.length; i++) {
      var h = String(currentHeaders[i]).toLowerCase();
      if (h.indexOf("id") !== -1 && idColIdx === -1) idColIdx = i + 1;
      if ((h.indexOf("телефон") !== -1 || h.indexOf("phone") !== -1) && phoneColIdx === -1) phoneColIdx = i + 1;
      if ((h.indexOf("email") !== -1 || h.indexOf("почта") !== -1) && emailColIdx === -1) emailColIdx = i + 1;
      if ((h.indexOf("лотере") !== -1 || h.indexOf("билет") !== -1 || h.indexOf("купон") !== -1) && lotteryColIdx === -1) lotteryColIdx = i + 1;
    }

    // Если в таблице ещё нет колонки Email, аккуратно вставляем её после колонки «Телефон»
    if (emailColIdx === -1 && lastRow > 0) {
      var pCol = phoneColIdx > 0 ? phoneColIdx : 4;
      sheet.insertColumnAfter(pCol);
      sheet.getRange(1, pCol + 1).setValue("Email").setBackground("#121212").setFontColor("#FFFFFF").setFontWeight("bold");
      lastCol = sheet.getLastColumn();
      currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      // Переопределяем индексы
      for (var i2 = 0; i2 < currentHeaders.length; i2++) {
        var h2 = String(currentHeaders[i2]).toLowerCase();
        if ((h2.indexOf("лотере") !== -1 || h2.indexOf("билет") !== -1 || h2.indexOf("купон") !== -1) && lotteryColIdx === -1) lotteryColIdx = i2 + 1;
      }
    }

    // Если в таблице ещё нет колонки «Номер лотереи», добавляем её
    if (lotteryColIdx === -1 && lastRow > 0) {
      var compColIdx = -1;
      for (var k = 0; k < currentHeaders.length; k++) {
        var kh = String(currentHeaders[k]).toLowerCase();
        if (kh.indexOf("компан") !== -1 || kh.indexOf("бренд") !== -1) {
          compColIdx = k + 1;
          break;
        }
      }
      var insertAfter = compColIdx > 0 ? compColIdx : Math.min(7, currentHeaders.length);
      sheet.insertColumnAfter(insertAfter);
      lotteryColIdx = insertAfter + 1;
      sheet.getRange(1, lotteryColIdx).setValue("Номер лотереи").setBackground("#121212").setFontColor("#FFFFFF").setFontWeight("bold");
      lastCol = sheet.getLastColumn();
      currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // ========================================================
    // РЕЖИМ 1: ПРИВЯЗКА / ОБНОВЛЕНИЕ НОМЕРА ЛОТЕРЕИ К СУЩЕСТВУЮЩЕМУ ЛИДУ
    // ========================================================
    if (data.action === "updateLottery" || (lottery && lastRow > 1 && !data.channels && !data.mainProduct)) {
      var targetRow = -1;
      var searchId = String(data.id || "").trim();
      var cleanPhone = String(data.phone || "").replace(/\D/g, "");

      var numRows = lastRow - 1;
      if (numRows > 0) {
        var idValues = idColIdx > 0 ? sheet.getRange(2, idColIdx, numRows, 1).getValues() : [];
        var phValues = phoneColIdx > 0 ? sheet.getRange(2, phoneColIdx, numRows, 1).getValues() : [];

        // Ищем строку снизу вверх (недавний созданный лид находится в самом конце таблицы)
        for (var r = numRows - 1; r >= 0; r--) {
          var rowId = idValues[r] ? String(idValues[r][0]).trim() : "";
          var rowPh = phValues[r] ? String(phValues[r][0]).replace(/\D/g, "") : "";
          
          if ((searchId && rowId === searchId) || (cleanPhone && cleanPhone.length >= 10 && rowPh.indexOf(cleanPhone.slice(-10)) !== -1)) {
            targetRow = r + 2; // 1-based, с учётом строки заголовка
            break;
          }
        }
      }

      if (targetRow > 0) {
        if (lotteryColIdx > 0) {
          var cell = sheet.getRange(targetRow, lotteryColIdx);
          cell.setValue(lottery);
          cell.setFontWeight("bold");
          cell.setBackground("#F4FFCC"); // Акцентный мягкий лайм
        }
        lock.releaseLock();
        return ContentService.createTextOutput(JSON.stringify({ status: "success", updated: true, row: targetRow, lottery: lottery }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ========================================================
    // РЕЖИМ 2: СОЗДАНИЕ НОВОГО ЛИДА
    // ========================================================
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
        var pVal = String(data.phone || "").trim();
        // Экранируем знак +, чтобы Google Таблицы не пытались вычислять телефон как формулу
        if (pVal && pVal.charAt(0) === "+") {
          pVal = "'" + pVal;
        }
        row.push(pVal);
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
    // Принудительно устанавливаем формат "Обычный текст" для колонки телефона, чтобы исключить #ERROR!
    if (phoneColIdx > 0) {
      sheet.getRange(sheet.getLastRow(), phoneColIdx).setNumberFormat("@");
    }

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
