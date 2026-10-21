/**
 * Hudor — збір відповідей опитування та дій відвідувачів у Google Таблицю.
 *
 * 1. Створіть нову Google Таблицю.
 * 2. Розширення → Apps Script. Вставте цей код замість вмісту Code.gs, збережіть.
 * 3. Розгорнути → Нове розгортання → Тип: Веб-застосунок.
 *    Виконувати як: Я. Хто має доступ: Будь-хто.
 * 4. Скопіюйте URL веб-застосунку (…/exec) і вставте його в index.html:
 *    const CONFIG = { endpoint: "https://script.google.com/macros/s/…/exec", … }
 *
 * Скрипт створить два аркуші:
 *   Leads  — заповнені опитування з контактами (одна заявка = один рядок);
 *   Events — усі дії: перегляди, кліки, кроки опитування, голоси за напрями.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents || "{}");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.type === "lead") writeLead_(ss, data);
    else writeEvent_(ss, data);
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err);
  } finally {
    lock.releaseLock();
  }
}

function sheet_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function writeLead_(ss, d) {
  var a = d.answers || {}, r = d.result || {}, c = d.contact || {}, k = d.consent || {};
  var sh = sheet_(ss, "Leads", [
    "Час", "Сесія", "Варіант", "UTM source", "UTM campaign",
    "Сфера", "Водокористування", "Стан документів", "Виклики", "Строки", "Формат (обраний)", "Цікаві напрями",
    "Рівень ризику", "Основна потреба", "Документи", "Рекомендований формат",
    "Ім’я", "Підприємство", "Email", "Телефон", "Коментар",
    "Згода на контакт", "Згода на аналітику"
  ]);
  var u = d.utm || {};
  sh.appendRow([
    d.ts, d.sid, d.variant, u.utm_source || "", u.utm_campaign || "",
    a.sector || "", (a.water || []).join(", "), a.status || "", (a.pains || []).join(", "),
    a.urgency || "", a.format || "", (a.extra || []).join(", "),
    r.level || "", r.need || "", (r.docs || []).join(", "), r.formatLabel || "",
    c.name || "", c.company || "", c.email || "", c.phone || "", c.note || "",
    k.contact ? "так" : "ні", k.analytics ? "так" : "ні"
  ]);
}

function writeEvent_(ss, d) {
  var sh = sheet_(ss, "Events", ["Час", "Сесія", "Варіант", "Подія", "Деталі", "UTM source", "Referrer"]);
  var u = d.utm || {};
  sh.appendRow([d.ts, d.sid, d.variant, d.event, JSON.stringify(d.props || {}), u.utm_source || "", d.ref || ""]);
}
