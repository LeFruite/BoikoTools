// ============================================================
//  «САМОСТІЙНІ» — шаблон. Нічого змінювати в коді не потрібно:
//  пін, групи та ключі задаються на сторінці підключення (start.html) і зберігаються у «Властивостях скрипту».
// ============================================================
// Адреса сторінки підключення на вашому сайті (її відкриває пункт меню «Самостійні → Встановлення»).
var START_PAGE = "https://ВАШ-САЙТ/start.html";

var PROPS_ = PropertiesService.getScriptProperties();
// Пін вчителя. Якщо ще не задано — вхід неможливий (випадкове значення, яке ніхто не введе).
var TEACHER_PIN = PROPS_.getProperty("TEACHER_PIN") || "\u0000not-configured";
// Ключ для скрипта-помічника (виписує оцінки з Extra). Порожній — функція вимкнена.
var MARKS_KEY = PROPS_.getProperty("MARKS_KEY") || "";
// Групи («7-А», «8-2»…) — задаються на сторінці підключення
var KNOWN_GROUPS = (function () { try { return JSON.parse(PROPS_.getProperty("GROUPS") || "[]"); } catch (e) { return []; } })();
// ============================================================


// Apps Script для Google Таблиці: приймає дані від тесту, створює аркуш для кожної роботи
// та автоматично оцінює відкриті питання за ключовими поняттями.
// Після вставки: Deploy → Manage deployments → Edit → New version.

var BASE_HEADERS = ["Час", "Клас", "Учень", "Автобал", "Статус", "Виходів зі сторінки",
                    "Бал відкритих (авто)", "Підсумок (авто)", "Перевірка відкритих"];

var OPEN_MAX = 2;    // макс. балів за ОДНЕ відкрите питання
var MIN_WORDS = 3;   // відповідь коротша за це число слів отримує 0

/* ===== КРИТЕРІЇ ВІДКРИТИХ ПИТАНЬ =====
   Ключ — код питання з пулу (o01…o10 — нервова система, k01…k10 — типи клітин, g01…g09 — політична карта світу, h01…h04 — клітина/тканини/органи/системи). Кожне «поняття» — це назва (для вашого звіту)
   і регулярний вираз (p), що шукається в тексті відповіді (малі літери, без апострофів).
   Бал = OPEN_MAX × (знайдені / усі поняття) ÷ FULL_AT, але не більше OPEN_MAX:
   максимум дається вже за FULL_AT = 75% понять. */
var RUBRIC = {};   // критерії для відкритих питань беруться з редактора

/* ===== Оцінювання відкритих питань ===== */
var FULL_AT = 0.75; // частка знайдених «понять», за якої відкрите питання отримує максимум (2 бали)
function norm_(s) {
  return String(s || "").toLowerCase().replace(/[’ʼ`']/g, "").replace(/\s+/g, " ").trim();
}

// rmap — критерії з тесту нової системи ({pid: {max, rub}}); якщо немає — беремо RUBRIC зі старих тестів
function gradeOpenAnswer_(pid, answer, rmap) {
  var rub = RUBRIC[pid], max = OPEN_MAX;
  if (rmap && rmap[pid]) { rub = rmap[pid].rub; max = rmap[pid].max || OPEN_MAX; }
  if (!rub || !rub.length) return { score: 0, note: pid + ": немає критеріїв — перевірте вручну" };
  var t = norm_(answer);
  var words = t.split(" ").filter(function (w) { return w.length > 1; }).length;
  if (words < MIN_WORDS) return { score: 0, note: pid + ": 0/" + max + " (відповідь порожня або занадто коротка)" };
  var hit = 0, miss = [];
  rub.forEach(function (c) {
    var ok;
    try { ok = c.ps ? c.ps.filter(function (q) { return new RegExp(q).test(t); }).length >= (c.min || 1) : new RegExp(c.p).test(t); }
    catch (e) { ok = false; }
    if (ok) hit++; else miss.push(c.n);
  });
  var score = Math.round(Math.min(max, max * (hit / rub.length) / FULL_AT) * 4) / 4;   // максимум — за FULL_AT (75%) понять
  return { score: score, note: pid + ": " + score + "/" + max + (miss.length ? " — не знайдено: " + miss.join("; ") : " — знайдено все") };
}

// answers: {q1: "...", q10: "[o06] Питання: ... — Відповідь: ...", ...}
function gradeOpenAll_(answers, rmap) {
  var sum = 0, notes = [];
  Object.keys(answers).sort(function (a, b) { return parseInt(a.slice(1), 10) - parseInt(b.slice(1), 10); })
    .forEach(function (key) {
      var m = String(answers[key]).match(/^\[([a-z]\d+)\]\s*Питання:[\s\S]*?—\s*Відповідь:\s*([\s\S]*)$/);
      if (!m) return;
      var g = gradeOpenAnswer_(m[1], m[2], rmap);
      sum += g.score;
      notes.push(key + " " + g.note);
    });
  return { sum: Math.round(sum * 100) / 100, notes: notes.join("\n") };
}

/* ===== Допоміжні функції ===== */
function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(BASE_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, BASE_HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function safe_(v) {
  v = (v === undefined || v === null) ? "" : String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

// Повертає номер колонки за назвою; якщо немає — створює в кінці
function col_(sheet, headers, name) {
  var c = headers.indexOf(name) + 1;
  if (c === 0) {
    c = headers.length + 1;
    sheet.getRange(1, c).setValue(name).setFontWeight("bold");
    headers.push(name);
  }
  return c;
}

function writeGrade_(sheet, row, headers, autoValue, g) {
  var autoNum = parseFloat(String(autoValue).replace(",", ".")) || 0; // працює і зі старим «6.25 / 7.5»
  var total = Math.round((autoNum + g.sum) * 100) / 100;
  sheet.getRange(row, col_(sheet, headers, "Бал відкритих (авто)")).setValue(g.sum);
  sheet.getRange(row, col_(sheet, headers, "Підсумок (авто)")).setValue(total);
  sheet.getRange(row, col_(sheet, headers, "Перевірка відкритих")).setValue(g.notes);
}

/* ===== Клас «8-2» Google Sheets може сам перетворити на дату — порівнюємо з урахуванням цього ===== */
function clsMatch_(v, cn) {
  cn = String(cn).trim().toLowerCase();
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(v, tz, "d-M") === cn || Utilities.formatDate(v, tz, "M-d") === cn;
  }
  return String(v).trim().toLowerCase() === cn;
}
// додає рядок, а колонку «Клас» записує саме як текст
function addRow_(sheet, arr) {
  var r = sheet.getLastRow() + 1;
  sheet.getRange(r, 2).setNumberFormat("@");
  sheet.getRange(r, 1, 1, arr.length).setValues([arr]);
  return r;
}

/* ===== Прийом даних від тесту ===== */
function doPost(e) {
  if (e && e.parameter && /^setup_/.test(e.parameter.action || "")) return setupApi_(e.parameter);   // перше підключення (без піна, лише поки пін не задано)
  if (e && e.parameter && /^t_/.test(e.parameter.action || "")) return teacherApi_(e.parameter);   // редактор вчителя
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return ContentService.createTextOutput("SERVER_BUSY");
  }

  try {
    var p = e.parameter;
    var sheetName = (p.sheetName || "Без назви").toString().substring(0, 100);
    var testId = (p.testId || "").trim(), test = null;
    if (testId) {   // тест нової системи: аркуш і керування — за ID тесту
      test = loadTest_(testId);
      if (!test) return ContentService.createTextOutput("NO_TEST");
      sheetName = test.sheetName;
    }
    var sheet = getSheet_(sheetName);
    var className = (p.className || "").trim();
    var studentName = (p.studentName || "").trim();
    var action = p.action;
    var now = new Date().toLocaleString('uk-UA');

    var ctrl = getCtrl_(testId);
    var round = (action === "submit" && parseInt(p.round, 10)) ? parseInt(p.round, 10) : ctrl.round;

    var data = sheet.getDataRange().getValues();
    var headers0 = data[0].slice();
    var rc = headers0.indexOf("Раунд");
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      var rr = rc >= 0 ? (parseInt(data[i][rc], 10) || 1) : 1;
      if (rr !== round) continue;
      if (clsMatch_(data[i][1], className) &&
          data[i][2].toString().trim().toLowerCase() === studentName.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (action === "start") {
      if (ctrl.state !== "open") return ContentService.createTextOutput("NOT_OPEN");
      if (ctrl.left !== null && ctrl.left <= 0) return ContentService.createTextOutput("TIME_UP");
      if (test && (test.groups || []).indexOf(className) < 0) return ContentService.createTextOutput("NOT_ASSIGNED");
      if (rowIndex !== -1) return ContentService.createTextOutput("ALREADY_STARTED");
      addRow_(sheet, [now, className, studentName, "", "В процесі складання...", 0]);
      sheet.getRange(sheet.getLastRow(), col_(sheet, headers0, "Раунд")).setValue(round);
      return ContentService.createTextOutput("OK_STARTED|" + round + "|" + ctrl.stopSeq);
    }

    if (action === "submit") {
      var answers = {};
      try { answers = JSON.parse(p.answers || "{}"); } catch (err) {}
      var partial = p.partial === "true";
      var status = partial ? (p.back === "true" ? "В процесі складання..." : "Вийшов зі сторінки (не завершив)")
        : (p.byTeacher === "true" ? "Перервано вчителем" : (p.isInterrupted === "true" ? "Перервано!" : "Завершено"));
      // запізніле часткове збереження не має перезаписувати вже завершену роботу
      if (partial && rowIndex !== -1) {
        var cur = String(data[rowIndex - 1][4]);
        if (cur.indexOf("Завершено") === 0 || cur.indexOf("Перервано") === 0) return ContentService.createTextOutput("OK_IGNORED");
      }

      if (rowIndex === -1) {
        rowIndex = addRow_(sheet, [now, className, studentName, "", "", 0]);
        sheet.getRange(rowIndex, col_(sheet, headers0, "Раунд")).setValue(round);
      }
      sheet.getRange(rowIndex, 1).setValue(now);
      sheet.getRange(rowIndex, 4).setValue(parseFloat(p.testScore) || 0);
      sheet.getRange(rowIndex, 5).setValue(status);
      sheet.getRange(rowIndex, 6).setValue(parseInt(p.leaves || "0", 10) || 0);

      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      sheet.getRange(rowIndex, col_(sheet, headers, "Тип виходів")).setValue(safe_(p.leaveInfo || ""));
      sheet.getRange(rowIndex, col_(sheet, headers, "Секунд поза сторінкою")).setValue(parseInt(p.awaySec || "0", 10) || 0);
      sheet.getRange(rowIndex, col_(sheet, headers, "Спроб вставки")).setValue(parseInt(p.pastes || "0", 10) || 0);
      sheet.getRange(rowIndex, col_(sheet, headers, "Журнал виходів")).setValue(safe_(String(p.leaveLog || "").substring(0, 2000)));
      sheet.getRange(rowIndex, col_(sheet, headers, "Час на питання")).setValue(safe_(String(p.qTime || "").substring(0, 2000)));
      Object.keys(answers).forEach(function (key) {
        sheet.getRange(rowIndex, col_(sheet, headers, key)).setValue(safe_(answers[key]));
      });

      writeGrade_(sheet, rowIndex, headers, p.testScore, gradeOpenAll_(answers, rubricMap_(test)));
      // в журнал оцінки потрапляють лише коли ви самі їх заносите (кнопка в редакторі або пункт меню)
      return ContentService.createTextOutput("OK_SUBMITTED");
    }

    return ContentService.createTextOutput("INVALID_ACTION");
  } finally {
    lock.releaseLock();
  }
}

/* ===== Керування вчителя: почати / перервати / дозволити всім ще раз ===== */

// id — ID тесту нової системи; порожній id — старі окремі HTML-файли (спільне керування)
function ck_(id, k) { return id ? k + "|" + id : k; }
function getCtrl_(id, fresh) {
  // стан читається з кешу (до 5 с): сотні опитувань від учнів не навантажують сховище; setCtrl_ скидає кеш одразу
  var cache = CacheService.getScriptCache(), key = "CTRLC|" + (id || "");
  if (!fresh) { var hit = cache.get(key); if (hit) return ctrlOut_(JSON.parse(hit)); }
  var pr = PropertiesService.getScriptProperties();
  var out = {
    state: pr.getProperty(ck_(id, "CTRL_STATE")) || "closed",           // open | paused | closed
    round: parseInt(pr.getProperty(ck_(id, "CTRL_ROUND")), 10) || 1,    // номер «спроби»
    stopSeq: parseInt(pr.getProperty(ck_(id, "CTRL_STOPSEQ")), 10) || 0, // збільшується при кожному «Перервати»
    endsAt: parseInt(pr.getProperty(ck_(id, "CTRL_ENDS")), 10) || 0,      // спільний кінець часу на роботу (мс), 0 = таймера немає
    pausedAt: parseInt(pr.getProperty(ck_(id, "CTRL_PAUSEDAT")), 10) || 0
  };
  cache.put(key, JSON.stringify(out), 5);
  return ctrlOut_(out);
}
// додає серверний час і скільки мс лишилось (null — спільного таймера немає); на паузі час стоїть
function ctrlOut_(o) {
  var n = Date.now();
  o.now = n;
  o.left = o.endsAt ? ((o.state === "paused" && o.pausedAt) ? o.endsAt - o.pausedAt : o.endsAt - n) : null;
  return o;
}
function workMinutes_(id) {
  var t = id ? loadTest_(id) : null;
  if (t && t.noTimer === true) return 0;      // таймер вимкнено для цієї роботи
  var m = t ? +t.totalTime : 0;
  return m > 0 ? m : 15;
}

function setCtrl_(kind, id, extra) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var pr = PropertiesService.getScriptProperties();
    var c = getCtrl_(id, true);
    var now = Date.now();
    if (kind === "open") {
      if (c.state === "closed") {          // «Відкрити доступ»: запускається спільний таймер на роботу
        var wm = workMinutes_(id);
        if (wm > 0) pr.setProperty(ck_(id, "CTRL_ENDS"), String(now + wm * 60000)); else pr.deleteProperty(ck_(id, "CTRL_ENDS"));
        pr.deleteProperty(ck_(id, "CTRL_PAUSEDAT"));
      } else if (c.state === "paused") {   // після паузи таймер продовжується з того місця, де зупинився
        if (c.endsAt && c.pausedAt) pr.setProperty(ck_(id, "CTRL_ENDS"), String(c.endsAt + (now - c.pausedAt)));
        pr.deleteProperty(ck_(id, "CTRL_PAUSEDAT"));
      }
      pr.setProperty(ck_(id, "CTRL_STATE"), "open");
    } else if (kind === "pause") {      // пауза: тести заморожуються, правила виходу не діють
      if (c.state === "open") pr.setProperty(ck_(id, "CTRL_PAUSEDAT"), String(now));
      pr.setProperty(ck_(id, "CTRL_STATE"), "paused");
    } else if (kind === "extend") {      // додати хвилин до спільного таймера
      var add = Math.max(1, Math.min(60, parseInt(extra, 10) || 5)) * 60000;
      if (c.endsAt) pr.setProperty(ck_(id, "CTRL_ENDS"), String(c.endsAt + add));
    } else if (kind === "close") {
      pr.setProperty(ck_(id, "CTRL_STATE"), "closed");
      pr.setProperty(ck_(id, "CTRL_STOPSEQ"), String(c.stopSeq + 1));
      pr.deleteProperty(ck_(id, "CTRL_ENDS")); pr.deleteProperty(ck_(id, "CTRL_PAUSEDAT"));
    } else if (kind === "reset") {      // нова спроба; робота закрита, поки не натиснете «Почати»
      pr.setProperty(ck_(id, "CTRL_STATE"), "closed");
      pr.setProperty(ck_(id, "CTRL_STOPSEQ"), String(c.stopSeq + 1));
      pr.setProperty(ck_(id, "CTRL_ROUND"), String(c.round + 1));
      pr.deleteProperty(ck_(id, "CTRL_ENDS")); pr.deleteProperty(ck_(id, "CTRL_PAUSEDAT"));
    }
  } finally {
    lock.releaseLock();
  }
  CacheService.getScriptCache().remove("CTRLC|" + (id || ""));
  return getCtrl_(id, true);
}

// Викликається зі сторінки вчителя
function teacherAction(pin, kind) {
  if (!kind) return getCtrl_();   // лише показати поточний статус
  if (String(pin) !== TEACHER_PIN) return { error: "Невірний пін" };
  if (kind === "open" || kind === "pause" || kind === "close" || kind === "reset") setCtrl_(kind);
  return getCtrl_();
}

// Статус для учнів (JSON) або сторінка вчителя (?panel=1)
function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.mat) return jsonOut_(matGet_(q.mat));              // власна HTML-сторінка (Матеріали)
  if (q.live_info) return jsonOut_(liveInfo_(q.live_info));      // живий урок: список класу
  if (q.live_vote) return jsonOut_(liveVoteStudent_(q));
  if (q.live) return jsonOut_(liveState_(q));
  if (q.review) return jsonOut_(reviewJson_(q));          // розбір для учня (якщо вчитель увімкнув)
  if (q.info) return jsonOut_(infoJson_(q.info));        // нова система: дані для сторінки учня
  if (q.body) return jsonOut_(bodyJson_(q.body));        // питання тесту (лише коли робота відкрита)
  if (q.ctrl) return jsonOut_(getCtrl_(q.ctrl));         // стан тесту (опитування кожні кілька секунд)
  if (e && e.parameter && e.parameter.marks) return marksJson_(e.parameter);
  if (e && e.parameter && e.parameter.panel) {
    return HtmlService.createHtmlOutput(PANEL_HTML_).setTitle("Керування самостійними")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }
  return ContentService.createTextOutput(JSON.stringify(getCtrl_())).setMimeType(ContentService.MimeType.JSON);
}

var PANEL_HTML_ =
  '<!doctype html><html><head><meta charset="utf-8"><style>' +
  'body{font-family:Arial,sans-serif;max-width:420px;margin:20px auto;padding:0 16px;text-align:center}' +
  'button{display:block;width:100%;padding:16px;margin:12px 0;font-size:18px;border:0;border-radius:8px;color:#fff;cursor:pointer}' +
  '#o{background:#27ae60}#p{background:#e67e22}#c{background:#c0392b}#r{background:#2980b9}input{padding:10px;font-size:16px;width:100%;box-sizing:border-box}' +
  '#st{font-size:20px;font-weight:bold;margin:16px 0}</style></head><body>' +
  '<h2>Керування самостійними</h2><input id="pin" type="password" placeholder="Пін вчителя">' +
  '<div id="st">…</div>' +
  '<button id="o" onclick="go(\'open\')">▶ Почати / продовжити написання</button>' +
  '<button id="p" onclick="go(\'pause\')">⏸ Пауза для всіх (тривога)</button>' +
  '<button id="c" onclick="go(\'close\')">⏹ Перервати всі роботи</button>' +
  '<button id="r" onclick="go(\'reset\')">🔄 Дозволити всім пройти ще раз</button>' +
  '<script>' +
  'function show(s){document.getElementById("st").innerText=s.error?s.error:({open:"Статус: ВІДКРИТО",paused:"Статус: ПАУЗА",closed:"Статус: ЗАКРИТО"}[s.state])+" · спроба №"+s.round;}' +
  'function go(k){var m={close:"Перервати ВСІ роботи просто зараз?",reset:"Дозволити всім пройти ще раз? Усі поточні роботи буде перервано, потім натисніть «Почати»."};' +
  'if(m[k]&&!confirm(m[k]))return;google.script.run.withSuccessHandler(show).teacherAction(document.getElementById("pin").value,k);}' +
  'google.script.run.withSuccessHandler(show).teacherAction("","");' +
  '</script></body></html>';

/* ===== Меню в таблиці ===== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Самостійні")
    .addItem("⚙ Встановлення та налаштування", "menuSetup_")
    .addSeparator()
    .addItem("▶ Почати / продовжити написання", "menuOpen_")
    .addItem("⏸ Пауза для всіх (тривога)", "menuPause_")
    .addItem("⏹ Перервати всі роботи", "menuClose_")
    .addItem("🔄 Дозволити всім пройти ще раз", "menuReset_")
    .addSeparator()
    .addItem("Перерахувати відкриті питання (поточний аркуш)", "regradeActiveSheet")
    .addItem("🧹 Прибрати зайві рядки (поточний аркуш)", "cleanupDuplicates")
    .addSeparator()
    .addItem("📒 Оцінки в журнал: лише порожні клітинки (поточний аркуш роботи)", "menuJournalFill_")
    .addItem("📒 Оцінки в журнал: перерахувати й перезаписати", "menuJournalOverwrite_")
    .addItem("📒 Створити аркуші журналу зі списків класів", "createJournals_")
    .addToUi();
}
// Якщо відкрито аркуш з результатами роботи з редактора — кнопки діють лише на цю роботу; інакше — на старі окремі файли.
function activeTestId_() {
  var t = findBySheet_(SpreadsheetApp.getActiveSheet().getName());
  return t ? t.id : "";
}
function menuOpen_()  { setCtrl_("open", activeTestId_());  toast_("Роботу відкрито — учні можуть починати"); }
function menuPause_() { setCtrl_("pause", activeTestId_()); toast_("Роботу поставлено на паузу"); }
function menuClose_() { setCtrl_("close", activeTestId_()); toast_("Роботу перервано"); }
function menuReset_() { var c = setCtrl_("reset", activeTestId_()); toast_("Нова спроба №" + c.round + ". Натисніть «Почати написання»"); }
function toast_(m) { SpreadsheetApp.getActiveSpreadsheet().toast(m, "Самостійні", 6); }

function regradeActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var tid = activeTestId_();
  var rmap = tid ? rubricMap_(loadTest_(tid)) : null;
  var data = sheet.getDataRange().getValues();
  var headers = data[0].slice();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][4]).indexOf("В процесі") === 0) continue;
    var answers = {};
    headers.forEach(function (h, c) {
      if (/^q\d+$/.test(h) && data[r][c]) answers[h] = String(data[r][c]);
    });
    writeGrade_(sheet, r + 1, headers, data[r][3], gradeOpenAll_(answers, rmap));
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Відкриті питання перераховано", "Готово", 5);
}


/* ===== Прибирання дублів: по одному рядку на учня за спробу (найповніший статус; за рівних — останній) ===== */
function cleanupDuplicates() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) { ui.alert("Тут нема що прибирати."); return; }
  var rc = data[0].indexOf("Раунд");
  var tz = Session.getScriptTimeZone();

  var known = {};  // тексти класів, що вже є в таблиці
  for (var i = 1; i < data.length; i++) if (!(data[i][1] instanceof Date)) known[String(data[i][1]).trim().toLowerCase()] = true;
  function clsKey(v) {
    if (!(v instanceof Date)) return String(v).trim().toLowerCase();
    var a = Utilities.formatDate(v, tz, "d-M"), b = Utilities.formatDate(v, tz, "M-d");
    return known[a] ? a : (known[b] ? b : a);
  }
  function rank(st) {
    st = String(st);
    if (st.indexOf("Завершено") === 0) return 4;
    if (st.indexOf("Перервано") === 0) return 3;
    if (st.indexOf("Вийшов") === 0) return 2;
    return 1;
  }
  var best = {};
  for (var r = 1; r < data.length; r++) {
    var key = clsKey(data[r][1]) + "|" + String(data[r][2]).trim().toLowerCase() + "|" + (rc >= 0 ? (parseInt(data[r][rc], 10) || 1) : 1);
    var rk = rank(data[r][4]);
    if (!best[key] || rk >= best[key].rk) best[key] = { row: r, rk: rk };
  }
  var del = [];
  for (var r2 = 1; r2 < data.length; r2++) {
    var k2 = clsKey(data[r2][1]) + "|" + String(data[r2][2]).trim().toLowerCase() + "|" + (rc >= 0 ? (parseInt(data[r2][rc], 10) || 1) : 1);
    if (best[k2].row !== r2) del.push(r2 + 1);
  }
  if (!del.length) { ui.alert("Зайвих рядків не знайдено."); return; }
  var ans = ui.alert("Буде видалено " + del.length + " зайвих рядків (лишається по одному на учня за спробу).\n" +
    "Перед цим автоматично створиться копія аркуша. Продовжити?", ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  sheet.copyTo(ss).setName((sheet.getName().substring(0, 70) + " (копія до очищення)").substring(0, 99));
  for (var d = del.length - 1; d >= 0; d--) sheet.deleteRow(del[d]);
  ss.toast("Видалено рядків: " + del.length, "Готово", 6);
}


/* ============================================================
   ЖУРНАЛ ОЦІНОК: по одному аркушу на клас («Журнал 6 клас», … «Журнал 9 клас»).
   У кожному — усі учні класу, біля прізвища підписана ваша група. Після самостійної
   з'являється нова колонка з оцінкою (12-бальна шкала) — окрема для кожної групи.
   ============================================================ */

// Учні й групи (за вашими списками з системи). Порожня група — допишіть вручну в таблиці.
// Біологія: 7, 8 класи; географія: 6, 9 класи. У 8 класі лише групи 8-2 і 8-4.
// Класи (паралелі) беруться з назв груп: «7-А» → 7. Учнів додають на сторінці «Списки груп».
var ROSTER = (function () {
  var o = {};
  KNOWN_GROUPS.forEach(function (g) { var m = String(g).match(/^\d+/); if (m && !o[m[0]]) o[m[0]] = []; });
  return o;
})();

// Назви груп перенесено на початок файлу (змінна KNOWN_GROUPS).

// Роботи: ключ — точна назва аркуша з результатами (SHEET_NAME у тесті).
// title — підпис колонки в журналі, max — максимум балів (автобал + 2 за кожне відкрите питання).
// Для нової роботи додайте рядок сюди.
var WORKS = {};   // старі окремі роботи (не потрібно в новій системі)

var GRADE_MAX = 12;  // шкала оцінювання
// Оцінка = частка набраних балів × 12, округлена; мінімум 1. Інша шкала — змініть цю функцію.
function gradeFromScore_(total, max) {
  if (!(max > 0) || isNaN(total)) return "";
  var g = Math.round(GRADE_MAX * total / max);
  return Math.max(1, Math.min(GRADE_MAX, g));
}

// Справжня назва групи (в аркуші «8-2» могло стати датою)
function canonClass_(v) {
  for (var i = 0; i < KNOWN_GROUPS.length; i++) if (clsMatch_(v, KNOWN_GROUPS[i])) return KNOWN_GROUPS[i];
  return v instanceof Date ? "" : String(v).trim();
}
function gradeOf_(cls) { var m = String(cls).match(/^\d+/); return m ? m[0] : ""; }

/* ----- порівняння прізвищ: «Іваненко Марія Петрівна» ~ «Іваненко Марія» ~ «Іваненко М. П.» ----- */
function normName_(s) { return norm_(s).replace(/\s+/g, ""); }

function parseName_(s) {
  var t = norm_(s).split(" ");
  var f = t[1] || "";
  var isIni = f.indexOf(".") >= 0 || f.length <= 1;
  return { sur: t[0] || "", full: isIni ? "" : f, ini: f.charAt(0) };
}
function lev_(a, b) {
  var m = a.length, n = b.length, d = [], i, j;
  for (i = 0; i <= m; i++) { d[i] = [i]; }
  for (j = 1; j <= n; j++) d[0][j] = j;
  for (i = 1; i <= m; i++) for (j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
  return d[m][n];
}
// 0 — однаково; більше — гірше; 99 — не підходить. Допускає дрібні розбіжності в написанні (Олена ~ Єлена, Олександр ~ Александр)
function nameScore_(a, b) {
  if (!a.sur || !b.sur) return 99;
  var s = lev_(a.sur, b.sur);
  if (s > 3) return 99;
  var bothFull = a.full && b.full;
  var f = bothFull ? lev_(a.full, b.full) : 0;
  if (a.ini !== b.ini && !(bothFull && f <= 2)) return 99;   // різні початкові літери — лише коли імена майже збігаються
  if (f > 3 || s + f > 4) return 99;
  return s + f;
}

function journalSheet_(grade) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = "Журнал " + grade + " клас";
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, 2).setValues([["Учень", "Група"]]).setFontWeight("bold");
    sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat("@");   // «8-2» не має ставати датою
    var list = ROSTER[grade] || [];
    if (list.length) sh.getRange(2, 1, list.length, 2).setValues(list);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(2);
    sh.setColumnWidth(1, 270);
    sh.setColumnWidth(2, 70);
  }
  return sh;
}

// колонка (робота + спроба + група): шукаємо за прихованим ключем у примітці; якщо немає — додаємо праворуч
function journalCol_(sh, work, round, cls, isoDate) {
  var key = work.title + "|" + round + "|" + cls;
  var lastCol = sh.getLastColumn();
  var notes = sh.getRange(1, 1, 1, lastCol).getNotes()[0];
  var earlier = 0;   // скільки колонок цієї роботи для цієї групи вже є (повторне складання)
  for (var c = 2; c < notes.length; c++) {
    var k = String(notes[c]).split("\n")[0];
    if (k === key) return c + 1;
    var p = k.split("|");
    if (p[0] === work.title && p[2] === cls) earlier++;
  }
  var col = lastCol + 1;
  var iso = isoDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var hdr = iso.slice(8, 10) + "." + iso.slice(5, 7) + " " + work.title + (earlier > 0 ? " (спроба " + (earlier + 1) + ")" : "") + " · " + cls;
  // у примітці: ключ і дата написання (за нею оцінки переносяться на сайт)
  sh.getRange(1, col).setValue(hdr).setNote(key + "\n" + iso).setFontWeight("bold").setWrap(true);
  sh.setColumnWidth(col, 120);
  return col;
}

// рядок учня: найкращий збіг; при рівності — той, чия група збігається; якщо учня немає — додаємо в кінець
function journalRow_(sh, studentName, cls, noCreate) {
  var last = sh.getLastRow();
  var vals = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  var ep = parseName_(studentName), best = 99, cand = [];
  vals.forEach(function (r, i) {
    var sc = nameScore_(ep, parseName_(r[0]));
    if (sc < best) { best = sc; cand = [i + 2]; } else if (sc === best && sc < 99) cand.push(i + 2);
  });
  if (best === 99) {
    if (noCreate) return -3;   // знімаємо позначку — нового рядка не створюємо
    sh.getRange(last + 1, 1, 1, 2).setValues([[studentName, cls]]);
    return last + 1;
  }
  if (cand.length > 1) {
    var byG = cand.filter(function (row) { return String(vals[row - 2][1]).trim() === cls; });
    if (byG.length === 1) return byG[0];
    return -1;
  }
  return cand[0];
}

// 1 — записано, 0 — клітинка вже заповнена, -1 — неоднозначний учень, -2 — не вдалося визначити клас/оцінку
function putGrade_(work, cls, studentName, round, total, overwrite, isoDate, gradeOverride) {
  var g = gradeOverride ? gradeOverride : gradeFromScore_(total, work.max), grade = gradeOf_(cls);
  if (g === "" || !grade || !ROSTER[grade]) return -2;
  var sh = journalSheet_(grade);
  var row = journalRow_(sh, studentName, cls);
  if (row < 0) return -1;
  var col = journalCol_(sh, work, round, cls, isoDate);
  var cell = sh.getRange(row, col);
  if (!overwrite && String(cell.getValue()).trim() !== "") return 0;
  cell.setValue(g).setHorizontalAlignment("center");
  return 1;
}

function isFinal_(status) { return /^(Завершено|Перервано)/.test(String(status)); }
function num_(v) { return parseFloat(String(v).replace(",", ".")); }

// оцінка одного рядка (викликається одразу після здачі роботи)
function journalFromRow_(sheet, rowIndex, overwrite) {
  var work = workFor_(sheet.getName());
  if (!work) return false;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var v = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  if (!isFinal_(v[4])) return false;
  var rc = headers.indexOf("Раунд"), tc = headers.indexOf("Підсумок (авто)");
  if (tc < 0) return false;
  return putGrade_(work, canonClass_(v[1]), String(v[2]).trim(), rc >= 0 ? (parseInt(v[rc], 10) || 1) : 1, num_(v[tc]), overwrite) === 1;
}

// Дата написання роботи з клітинки «Час» (рядок «16.10.2026, 11:12:56» або справжня дата)
function isoFromTime_(v) {
  var tz = Session.getScriptTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  var m = String(v).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  m = String(v).match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
}

var TEACHER_GRADE_COL = "Оцінка (вчитель)";

// Заносить у журнал оцінки з аркуша роботи. only — необов'язковий набір ключів «група|прізвище|спроба».
// Якщо ви виставили оцінку вручну — береться вона, інакше обчислена автоматично.
function journalPush_(sheet, work, overwrite, only) {
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var rc = h.indexOf("Раунд"), tc = h.indexOf("Підсумок (авто)"), oc = h.indexOf(TEACHER_GRADE_COL);
  if (tc < 0) return { error: "no_total" };
  var best = {};
  for (var r = 1; r < data.length; r++) {
    if (!isFinal_(data[r][4])) continue;
    var ov = oc >= 0 ? num_(data[r][oc]) : NaN;
    if (data[r][tc] === "" && !(ov > 0)) continue;
    var cls = canonClass_(data[r][1]);
    var rnd = rc >= 0 ? (parseInt(data[r][rc], 10) || 1) : 1;
    var key = cls + "|" + normName_(data[r][2]) + "|" + rnd;
    var rk = String(data[r][4]).indexOf("Завершено") === 0 ? 2 : 1;
    if (!best[key] || rk >= best[key].rk) best[key] = { r: r, rk: rk, cls: cls, rnd: rnd, key: key };
  }
  var out = { done: 0, kept: 0, bad: [], total: 0 };
  Object.keys(best).forEach(function (k) {
    var b = best[k];
    if (only && !only[k]) return;
    out.total++;
    var ov2 = oc >= 0 ? num_(data[b.r][oc]) : NaN;
    var res = putGrade_(work, b.cls, String(data[b.r][2]).trim(), b.rnd, num_(data[b.r][tc]), overwrite,
                        isoFromTime_(data[b.r][0]), ov2 > 0 ? Math.round(ov2) : 0);
    if (res === 1) out.done++; else if (res === 0) out.kept++; else out.bad.push(String(data[b.r][2]).trim() + " (" + b.cls + ")");
  });
  return out;
}

function syncJournal_(overwrite) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var work = workFor_(sheet.getName());
  if (!work) { ui.alert("Для аркуша «" + sheet.getName() + "» немає роботи (ні в редакторі, ні в WORKS у Code.gs).\nВідкрийте аркуш з результатами роботи."); return; }
  var r = journalPush_(sheet, work, overwrite, null);
  if (r.error) { ui.alert("На цьому аркуші немає колонки «Підсумок (авто)»."); return; }
  var msg = "Записано оцінок: " + r.done + " з " + r.total + (overwrite ? " (з перезаписом)" : "; уже були в журналі: " + r.kept);
  if (r.bad.length) msg += ". Не вдалося (неоднозначно або невідомий клас): " + r.bad.join(", ");
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Журнал", 15);
}
function menuJournalFill_() { syncJournal_(false); }
function menuJournalOverwrite_() { syncJournal_(true); }

function createJournals_() {
  Object.keys(ROSTER).forEach(function (g) { journalSheet_(g); });
  SpreadsheetApp.getActiveSpreadsheet().toast("Аркуші журналу створено", "Журнал", 5);
}

/* ============================================================
   ВИДАЧА ОЦІНОК ДЛЯ ПЕРЕНЕСЕННЯ НА САЙТ (читає скрипт-помічник у браузері вчителя)
   Захищено ключем MARKS_KEY (на початку файлу). Поки він порожній — вимкнено.
   ============================================================ */
// Ключ MARKS_KEY задається на самому початку файлу.

function marksJson_(p) {
  function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
  if (!MARKS_KEY || String(p.key) !== MARKS_KEY) return out({ error: "Ключ не заданий або невірний" });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), works = [], absent = [];
  Object.keys(ROSTER).forEach(function (grade) {
    if (p.grade && String(p.grade) !== grade) return;
    var sh = ss.getSheetByName("Журнал " + grade + " клас");
    if (!sh || sh.getLastColumn() < 3) return;
    var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
    var notes = sh.getRange(1, 1, 1, lastCol).getNotes()[0];
    var vals = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    for (var c = 2; c < lastCol; c++) {
      var parts = String(notes[c]).split("\n");
      var iso = parts[1] || "";
      if (parts[0] && parts[0].indexOf(ATT_TITLE + "|") === 0) {            // колонка відвідуваності → «Н»
        if (p.date && iso !== p.date) continue;
        var ak = parts[0].split("|");
        vals.forEach(function (r) { if (String(r[c]).trim() === "Н") absent.push([String(r[0]).trim(), iso, ak[2] || "", String(r[1]).trim()]); });
        continue;
      }
      if (!parts[0]) continue;
      if (p.date && iso !== p.date) continue;
      var k = parts[0].split("|");
      var rows = [];
      vals.forEach(function (r) {
        var n = num_(r[c]);
        if (String(r[c]).trim() !== "" && !isNaN(n)) rows.push([String(r[0]).trim(), n, String(r[1]).trim()]);
      });
      works.push({ grade: grade, group: k[2] || "", title: k[0], date: iso, rows: rows });
    }
  });
  return out({ works: works, absent: absent });
}


/* ============================================================
   НОВА СИСТЕМА: редактор тестів + один універсальний програвач (index.html на GitHub Pages)
   Тести зберігаються в аркуші «_Тести» цієї таблиці, малюнки — у папці Google Drive.
   Усі запити редактора захищені TEACHER_PIN.
   ============================================================ */
var TESTS_SHEET = "_Тести";
var IMG_FOLDER = "Самостійні — малюнки";
var CHUNK = 40000;   // у одну клітинку Таблиці вміщується до 50 000 символів
// Аркуш: A id | B назва | C оновлено | D групи | E аркуш результатів | F макс. балів | G частин | H… дані (JSON частинами)

function jsonOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function testsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TESTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TESTS_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([["ID", "Назва", "Оновлено", "Групи", "Аркуш результатів", "Макс. балів", "Частин"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function testRowOf_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function loadTest_(id) {
  if (!/^[A-Za-z0-9_\-]{3,40}$/.test(String(id))) return null;
  var cache = CacheService.getScriptCache();
  var hit = cache.get("T|" + id);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var sh = testsSheet_();
  var row = testRowOf_(sh, id);
  if (row < 0) return null;
  var n = parseInt(sh.getRange(row, 7).getValue(), 10) || 0;
  if (n < 1) return null;
  var parts = sh.getRange(row, 8, 1, n).getValues()[0];
  var t;
  try { t = JSON.parse(parts.join("")); } catch (e) { return null; }
  try { cache.put("T|" + id, JSON.stringify(t), 1800); } catch (e) {}   // кеш не вміщує дуже великі тести — тоді просто читаємо з аркуша
  return t;
}

// максимальний автобал + бали за відкриті (середнє, якщо в розділі різні)
function maxOf_(t) {
  var sum = 0;
  (t.sections || []).forEach(function (s) {
    var qs = s.questions || [];
    if (!qs.length) return;
    var n = s.type === "tf" ? (+s.n || 1) : Math.min(+s.n || qs.length, qs.length);
    if (s.type === "open") {
      var avg = qs.reduce(function (a, q) { return a + (+q.max || OPEN_MAX); }, 0) / qs.length;
      sum += n * avg;
    } else {
      sum += n * (+s.points || 0);
    }
  });
  return Math.round(sum * 100) / 100;
}

function cleanSheetName_(s) {
  return String(s || "Без назви").replace(/[\[\]\*\?:\/\\]/g, " ").replace(/\s+/g, " ").trim().substring(0, 90);
}

function saveTest_(id, t) {
  if (!/^[A-Za-z0-9_\-]{3,40}$/.test(String(id))) throw new Error("bad id");
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = testsSheet_();
    var row = testRowOf_(sh, id);
    t.id = id;
    // аркуш результатів: стабільна назва; унікальна серед тестів
    if (!t.sheetName) {
      var base = cleanSheetName_("Самостійна " + (t.title || id)), name = base, k = 2;
      var used = {};
      var last = sh.getLastRow();
      if (last >= 2) sh.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) { if (String(r[0]) !== String(id)) used[String(r[4])] = true; });
      while (used[name] || (WORKS[name] && true)) name = base.substring(0, 85) + " " + (k++);
      t.sheetName = name;
    }
    var json = JSON.stringify(t), parts = [];
    for (var i = 0; i < json.length; i += CHUNK) parts.push(json.substring(i, i + CHUNK));
    if (!parts.length) parts = ["{}"];
    var meta = [id, t.title || "", new Date().toISOString(), (t.groups || []).join(","), t.sheetName, maxOf_(t), parts.length];
    if (row < 0) row = Math.max(2, sh.getLastRow() + 1);
    sh.getRange(row, 1, 1, 7).setNumberFormat("@").setValues([meta]);
    sh.getRange(row, 8, 1, parts.length).setNumberFormat("@").setValues([parts]);
    bustTests_(id);
    return t;
  } finally {
    lock.releaseLock();
  }
}

// статична частина списку (без стану) кешується; стан беремо одним пакетом
function bustTests_(id) {
  var c = CacheService.getScriptCache();
  c.removeAll(["TLM", "T|" + id, "TB|" + id]);
}
function listTests_() {
  var cache = CacheService.getScriptCache(), meta = null, hit = cache.get("TLM");
  if (hit) { try { meta = JSON.parse(hit); } catch (e) {} }
  if (!meta) {
    meta = [];
    var sh = testsSheet_(), last = sh.getLastRow(), fold = tFolders_();
    if (last >= 2) sh.getRange(2, 1, last - 1, 7).getValues().forEach(function (r) {
      if (!r[0]) return;
      meta.push({ id: String(r[0]), title: String(r[1]), updated: String(r[2]), groups: String(r[3]).split(",").filter(Boolean),
                  sheetName: String(r[4]), max: +r[5] || 0, folder: fold[String(r[0])] || "" });
    });
    try { cache.put("TLM", JSON.stringify(meta), 60); } catch (e) {}
  }
  var cm = ctrlBulk_(meta.map(function (m) { return m.id; }));
  return meta.map(function (m) { var o = {}; for (var k in m) o[k] = m[k]; o.ctrl = ctrlOut_(cm[m.id]); return o; });
}
// стан усіх робіт: один getAll із кешу + один читання властивостей для відсутніх
function ctrlBulk_(ids) {
  var cache = CacheService.getScriptCache(), keys = ids.map(function (i) { return "CTRLC|" + i; });
  var got = keys.length ? cache.getAll(keys) : {}, out = {}, miss = [];
  ids.forEach(function (id) {
    var h = got["CTRLC|" + id];
    if (h) { try { out[id] = JSON.parse(h); return; } catch (e) {} }
    miss.push(id);
  });
  if (miss.length) {
    var all = PropertiesService.getScriptProperties().getProperties(), put = {};
    miss.forEach(function (id) {
      var o = {
        state: all[ck_(id, "CTRL_STATE")] || "closed",
        round: parseInt(all[ck_(id, "CTRL_ROUND")], 10) || 1,
        stopSeq: parseInt(all[ck_(id, "CTRL_STOPSEQ")], 10) || 0,
        endsAt: parseInt(all[ck_(id, "CTRL_ENDS")], 10) || 0,
        pausedAt: parseInt(all[ck_(id, "CTRL_PAUSEDAT")], 10) || 0
      };
      out[id] = o; put["CTRLC|" + id] = JSON.stringify(o);
    });
    try { cache.putAll(put, 5); } catch (e) {}
  }
  return out;
}

function findBySheet_(name) {
  if (!name || name === TESTS_SHEET) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < rows.length; i++) if (String(rows[i][4]) === String(name)) return { id: String(rows[i][0]), title: String(rows[i][1]), max: +rows[i][5] || 0 };
  return null;
}

// для журналу: спершу старі роботи (WORKS), потім роботи з редактора
function workFor_(sheetName) {
  if (WORKS[sheetName]) return WORKS[sheetName];
  var t = findBySheet_(sheetName);
  return t ? { title: t.title, max: t.max } : null;
}

/* ----- критерії відкритих питань з тесту ----- */
function escRe_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function conceptToRub_(c) {
  var name = c.n || c.k || "";
  if (c.ps && c.ps.length) return { n: name, ps: c.ps, min: c.min || 1 };
  if (c.p) return { n: name, p: c.p };
  var kws = String(c.k || "").split(/[,;\n]/).map(function (w) { return norm_(w); }).filter(Boolean).map(escRe_);
  return { n: name, p: kws.length ? kws.join("|") : "(?!)" };   // «(?!)» — ніколи не збігається
}
function rubricMap_(t) {
  if (!t) return null;
  var map = {};
  (t.sections || []).forEach(function (s) {
    if (s.type !== "open") return;
    (s.questions || []).forEach(function (q) {
      map[q.id] = { max: +q.max || OPEN_MAX, rub: (q.concepts || []).map(conceptToRub_) };
    });
  });
  return map;
}

/* ----- список учнів: береться з аркушів «Журнал N клас» (колонки «Учень» і «Група») ----- */
function rosterMap_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get("ROSTER");
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var out = {};
  Object.keys(ROSTER).forEach(function (grade) {
    var sh = journalSheet_(grade), last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      var name = String(r[0]).trim(), g = canonClass_(r[1]);
      if (name && g) (out[g] = out[g] || []).push(name);
    });
  });
  try { cache.put("ROSTER", JSON.stringify(out), 45); } catch (e) {}
  return out;
}

/* ----- для сторінки учня ----- */
function infoJson_(id) {
  var t = loadTest_(id);
  if (!t || !(t.groups || []).length) return { error: "not_found" };
  var ros = rosterMap_(), groups = {};
  t.groups.forEach(function (g) { groups[g] = (ros[g] || []).slice(); });
  return { ok: true, id: id, title: t.title, subject: t.subject || "", topic: t.topic || "", groups: groups,
           ctrl: getCtrl_(id), awayOne: +t.awayOne || 5, awayTotal: +t.awayTotal || 15, rules: rulesOf_(t) };
}
function bodyJson_(id) {
  var t = loadTest_(id);
  if (!t || !(t.groups || []).length) return { error: "not_found" };
  if (getCtrl_(id).state === "closed") return { error: "closed" };   // питання не віддаємо, поки вчитель не відкрив роботу
  var cc = CacheService.getScriptCache(), ph = cc.get("TB|" + id);
  if (ph) { try { return { test: JSON.parse(ph) }; } catch (e) {} }
  var pub = JSON.parse(JSON.stringify(t));
  delete pub.groups;
  (pub.sections || []).forEach(function (s) {   // критерії відкритих питань учням не віддаємо
    (s.questions || []).forEach(function (q) { delete q.ai; });
    if (s.type === "open") (s.questions || []).forEach(function (q) { delete q.concepts; });
  });
  try { cc.put("TB|" + id, JSON.stringify(pub), 1800); } catch (e) {}
  return { test: pub };
}

/* ===== Генерація питань із матеріалу за допомогою Gemini ===== */
// Ключ: Apps Script → Налаштування проєкту → Властивості скрипту → GEMINI_KEY (необов'язково GEMINI_MODEL, за замовчуванням gemini-2.5-flash)
var AI_TYPES_ = ["mcq", "multi", "match", "short", "open", "tf", "order", "cloze", "num", "group", "find", "label", "tap", "pmcq"];
function aiCounts_(c) {
  var out = { total: 0 }; c = c || {};
  AI_TYPES_.forEach(function (k) { var n = Math.max(0, Math.min(15, parseInt(c[k], 10) || 0)); out[k] = n; out.total += n; });
  return out;
}
function aiPrompt_(spec, c) {
  var lvl = { easy: "простіші, на розуміння і запам'ятовування", mid: "середньої складності", hard: "складніші, на аналіз і застосування знань" }[spec.level] || "середньої складності";
  var need = [];
  if (c.mcq) need.push("mcq: " + c.mcq);
  if (c.multi) need.push("multi: " + c.multi);
  if (c.match) need.push("match: " + c.match);
  if (c.short) need.push("short: " + c.short);
  if (c.open) need.push("open: " + c.open);
  if (c.tf) need.push("tf: " + c.tf);
  if (c.order) need.push("order: " + c.order);
  if (c.cloze) need.push("cloze: " + c.cloze);
  if (c.num) need.push("num: " + c.num);
  if (c.group) need.push("group: " + c.group);
  if (c.find) need.push("find: " + c.find);
  if (c.pmcq) need.push("pmcq: " + c.pmcq);
  if (c.label) need.push("label: " + c.label);
  if (c.tap) need.push("tap: " + c.tap);
  var s = "Ти асистент учителя" + (spec.subject ? " (" + String(spec.subject).substring(0, 40) + ")" : "") + ". На основі НАДАНОГО матеріалу (слайди/текст уроку) склади питання для самостійної роботи учнів " +
    (spec.grade ? String(spec.grade).substring(0, 20) + " класу" : "") + ". Рівень: " + lvl + ".\n" +
    "ПРАВИЛА:\n- Лише українською. Використовуй ТІЛЬКИ факти з матеріалу; нічого не вигадуй і не додавай із власних знань. Якщо в матеріалі чогось немає — не питай про це.\n" +
    "- Враховуй і текст, і таблиці та підписи на малюнках зі слайдів. Не став питань про зовнішній вигляд картинок.\n" +
    "- Питання мають бути різні, без повторів, чіткі, з однозначною відповіддю. Не використовуй «усе вищезазначене» і «жодне з наведених».\n" +
    "- ЗАБОРОНЕНО питати про структуру викладу: в якому порядку подавався матеріал, що було раніше чи пізніше в презентації/тексті, яке слово, приклад чи ім'я було названо або згадано в матеріалі, що «сказано/показано на слайді». Не пиши «згідно з текстом/слайдом/матеріалом/презентацією». Питання перевіряють розуміння ЗМІСТУ — факти, причини, властивості, поняття, закономірності — і мають бути зрозумілими самі по собі, без матеріалу перед очима.\n" +
    "- Тип order — лише для природних послідовностей у самому предметі (етапи процесу, ряд за величиною, хронологія подій), а не для порядку подання матеріалу. Тип find — для пошуку термінів певної категорії (наприклад, усіх назв планет), а не слів, які «були названі».\n" +
    (spec.hasImages ? "- Малюнки: якщо питання потребує малюнка, схеми чи графіка зі слайда (без нього відповісти неможливо), вкажи \"page\" — номер цього слайда/сторінки. Типи pmcq, label, tap — ЗАВЖДИ з \"page\" (слайд, де є змістовний малюнок чи схема; декоративні картинки ігноруй). Координати точок \"pt\" — [y, x] у діапазоні 0–1000 відносно ВСІЄЇ сторінки (0,0 — лівий верхній кут), точка має стояти на самому об'єкті.\n" : "- Малюнків немає: типи pmcq, label, tap не створюй.\n") +
    "- Для кожного питання вкажи \"src\": номер слайда/сторінки, звідки воно.\n" +
    "Потрібна кількість питань кожного типу: " + need.join(", ") + ".\n" +
    "ФОРМАТ ВІДПОВІДІ — лише JSON без пояснень, рівно такого вигляду (типи з нульовою кількістю пропусти або залиш порожнім масивом):\n" +
    "{\"mcq\":[{\"text\":\"питання\",\"correct\":\"правильна відповідь\",\"wrong\":[\"хибна 1\",\"хибна 2\",\"хибна 3\"],\"src\":\"5\"}],\n" +
    " \"multi\":[{\"text\":\"питання (оберіть усі правильні)\",\"options\":[\"варіант 1\",\"варіант 2\",\"варіант 3\",\"варіант 4\",\"варіант 5\"],\"correctIdx\":[0,2],\"src\":\"6\"}],\n" +
    " \"match\":[{\"text\":\"Установіть відповідність:\",\"pairs\":[[\"ліва частина 1\",\"права 1\"],[\"ліва 2\",\"права 2\"],[\"ліва 3\",\"права 3\"],[\"ліва 4\",\"права 4\"]],\"src\":\"10\"}],\n" +
    " \"short\":[{\"text\":\"питання з відповіддю в 1–3 слова\",\"accepted\":[\"варіант відповіді\",\"інша форма слова\"],\"src\":\"7\"}],\n" +
    " \"open\":[{\"text\":\"питання, що потребує пояснення в 1–3 реченнях\",\"concepts\":[{\"n\":\"що має бути названо\",\"k\":\"корінь1, корінь2, корінь3\"}],\"src\":\"9\"}],\n" +
    " \"tf\":[{\"t\":\"твердження\",\"c\":true,\"src\":\"3\"}],\n" +
    " \"order\":[{\"text\":\"Розташуйте ... у правильному порядку:\",\"items\":[\"1-й за порядком\",\"2-й\",\"3-й\",\"4-й\"],\"src\":\"8\"}],\n" +
    " \"cloze\":[{\"text\":\"Речення, де пропущені слова взято в подвійні квадратні дужки: [[слово]] ... [[слово]].\",\"extra\":[\"зайве слово 1\",\"зайве слово 2\"],\"src\":\"4\"}],\n" +
    " \"num\":[{\"text\":\"питання з числовою відповіддю\",\"ans\":150,\"tol\":5,\"unit\":\"млн км\",\"src\":\"11\"}],\n" +
    " \"group\":[{\"text\":\"Розподіліть за групами:\",\"cats\":[\"Група А\",\"Група Б\"],\"items\":[{\"t\":\"елемент\",\"c\":0},{\"t\":\"елемент\",\"c\":1}],\"src\":\"12\"}],\n" +
    " \"find\":[{\"text\":\"Торкніться в тексті ...:\",\"passage\":\"Текст із 2–4 реченнями, де шукані слова позначені зірочками: *слово*.\",\"src\":\"2\"}],\n" +
    " \"pmcq\":[{\"text\":\"питання про малюнок (наприклад, що позначено цифрою 2)\",\"correct\":\"правильна\",\"wrong\":[\"хибна 1\",\"хибна 2\",\"хибна 3\"],\"page\":5,\"src\":\"5\"}],\n" +
    " \"label\":[{\"text\":\"Підпишіть частини ... на малюнку:\",\"marks\":[{\"pt\":[250,430],\"a\":\"назва частини\"},{\"pt\":[600,300],\"a\":\"назва\"},{\"pt\":[480,720],\"a\":\"назва\"}],\"extra\":[\"зайвий підпис\"],\"page\":5,\"src\":\"5\"}],\n" +
    " \"tap\":[{\"text\":\"Торкніться на малюнку ...\",\"targets\":[{\"pt\":[300,500]}],\"page\":5,\"src\":\"5\"}]}\n" +
    "Уточнення: у mcq правильна відповідь одна, хибні — правдоподібні й схожі за довжиною. У multi 4–5 варіантів, правильних 2–3 (correctIdx — номери з нуля). У match 3–5 пар, короткі. " +
    "У short у accepted додай 1–4 прийнятні форми (відмінки, синоніми). У open 2–4 поняття, а в k — корені слів без закінчень через кому (щоб відповідь учня можна було знайти за ними). У tf перемішай правильні й хибні твердження. " +
    "У order 4–6 пунктів, у ПРАВИЛЬНОМУ порядку (сайт сам перемішає); питання про послідовність, етапи, ряд за зростанням/спаданням, яке є в матеріалі. " +
    "У cloze 1–4 пропуски в дужках [[ ]] лише на ключових термінах, extra — 2–3 правдоподібні зайві слова. " +
    "У num ans — число (крапка як десятковий роздільник), tol — допустима похибка (0, якщо відповідь точна), unit — одиниці або порожній рядок; лише якщо в матеріалі є числа. " +
    "У group 2–3 групи (cats) і 4–8 елементів, у c — номер групи з нуля; у кожній групі хоча б один елемент. " +
    "У pmcq/label/tap назви в підписах короткі (1–3 слова), у label 3–6 міток, що не накладаються. У find passage — 2–4 речення з матеріалу, де шукані слова або терміни (1–5) позначені *зірочками*, а в text сказано, що саме шукати.";
  if (spec.extra) s += "\nДодаткові побажання вчителя (лише щодо змісту й стилю питань): " + String(spec.extra).substring(0, 500);
  return s;
}
function aiStr_(x, max) { return String(x == null ? "" : x).replace(/\s+/g, " ").trim().substring(0, max || 300); }
function aiClean_(o, c) {
  var out = { mcq: [], multi: [], match: [], short: [], open: [], tf: [], order: [], cloze: [], num: [], group: [], find: [], label: [], tap: [], pmcq: [] };
  function pg(x) { var n = parseInt(x && x.page, 10); return n >= 1 && n <= 500 ? n : 0; }
  function pt01(p) { if (!Array.isArray(p) || p.length < 2) return null; var y = parseFloat(p[0]), x = parseFloat(p[1]); if (!isFinite(x) || !isFinite(y)) return null; return { x: Math.round(Math.min(1, Math.max(0, x / 1000)) * 10000) / 10000, y: Math.round(Math.min(1, Math.max(0, y / 1000)) * 10000) / 10000 }; }
  function arr(k) { return Array.isArray(o && o[k]) ? o[k] : []; }
  function src(x) { var s = aiStr_(x && x.src, 12); return s ? "сл. " + s.replace(/^(сл\.?|слайд|стор\.?)\s*/i, "") : ""; }
  arr("mcq").forEach(function (q) {
    var text = aiStr_(q.text, 400), cor = aiStr_(q.correct, 200), wr = (Array.isArray(q.wrong) ? q.wrong : []).map(function (w) { return aiStr_(w, 200); }).filter(function (w) { return w && w !== cor; });
    wr = wr.filter(function (w, i) { return wr.indexOf(w) === i; }).slice(0, 4);
    if (text && cor && wr.length >= 2) out.mcq.push({ text: text, o: [cor].concat(wr), src: src(q) });
  });
  arr("multi").forEach(function (q) {
    var text = aiStr_(q.text, 400), opts = (Array.isArray(q.options) ? q.options : []).map(function (w) { return aiStr_(w, 200); }), idx = [];
    (Array.isArray(q.correctIdx) ? q.correctIdx : []).forEach(function (i) { i = parseInt(i, 10); if (i >= 0 && i < opts.length && idx.indexOf(i) < 0 && opts[i]) idx.push(i); });
    if (text && opts.length >= 4 && opts.length <= 6 && opts.every(function (w) { return w; }) && idx.length >= 1 && idx.length < opts.length) out.multi.push({ text: text, o: opts, c: idx.sort(function (a, b) { return a - b; }), src: src(q) });
  });
  arr("match").forEach(function (q) {
    var pairs = (Array.isArray(q.pairs) ? q.pairs : []).map(function (p) { return Array.isArray(p) ? [aiStr_(p[0], 120), aiStr_(p[1], 120)] : null; }).filter(function (p) { return p && p[0] && p[1]; }).slice(0, 6);
    if (pairs.length >= 3) out.match.push({ text: aiStr_(q.text, 300) || "Установіть відповідність:", pairs: pairs, src: src(q) });
  });
  arr("short").forEach(function (q) {
    var acc = (Array.isArray(q.accepted) ? q.accepted : []).map(function (w) { return aiStr_(w, 80); }).filter(function (w, i, a) { return w && a.indexOf(w) === i; }).slice(0, 6);
    var text = aiStr_(q.text, 400);
    if (text && acc.length) out.short.push({ text: text, acc: acc, src: src(q) });
  });
  arr("open").forEach(function (q) {
    var cs = (Array.isArray(q.concepts) ? q.concepts : []).map(function (x) { return { n: aiStr_(x && x.n, 160), k: aiStr_(Array.isArray(x && x.k) ? x.k.join(", ") : (x && x.k), 200) }; }).filter(function (x) { return x.n && x.k; }).slice(0, 5);
    var text = aiStr_(q.text, 500);
    if (text && cs.length) out.open.push({ text: text, max: 2, concepts: cs, src: src(q) });
  });
  arr("tf").forEach(function (q) {
    var t = aiStr_(q.t, 300);
    if (t && (q.c === true || q.c === false)) out.tf.push({ t: t, c: q.c, src: src(q) });
  });
  arr("order").forEach(function (q) {
    var it = (Array.isArray(q.items) ? q.items : []).map(function (w) { return aiStr_(w, 160); }).filter(function (w) { return w; });
    var uniq = it.filter(function (w, i) { return it.indexOf(w) === i; }).slice(0, 7), text = aiStr_(q.text, 300);
    if (text && uniq.length >= 3 && uniq.length === it.length) out.order.push({ text: text, items: uniq, src: src(q) });
  });
  arr("cloze").forEach(function (q) {
    var text = aiStr_(q.text, 700), gaps = text.match(/\[\[[^\]]+\]\]/g) || [];
    var ex = (Array.isArray(q.extra) ? q.extra : []).map(function (w) { return aiStr_(w, 60); }).filter(function (w, i, a) { return w && a.indexOf(w) === i; }).slice(0, 4);
    if (text && gaps.length >= 1 && gaps.length <= 5) out.cloze.push({ text: text, extra: ex, src: src(q) });
  });
  arr("num").forEach(function (q) {
    var ans = parseFloat(String(q.ans).replace(",", ".")), tol = parseFloat(String(q.tol == null ? 0 : q.tol).replace(",", ".")), text = aiStr_(q.text, 400);
    if (text && isFinite(ans)) out.num.push({ text: text, ans: ans, tol: isFinite(tol) && tol >= 0 ? tol : 0, unit: aiStr_(q.unit, 20), src: src(q) });
  });
  arr("group").forEach(function (q) {
    var cats = (Array.isArray(q.cats) ? q.cats : []).map(function (w) { return aiStr_(w, 60); }).filter(function (w) { return w; }).slice(0, 4);
    var items = (Array.isArray(q.items) ? q.items : []).map(function (x) { return { t: aiStr_(x && x.t, 80), c: parseInt(x && x.c, 10) }; })
      .filter(function (x) { return x.t && x.c >= 0 && x.c < cats.length; }).slice(0, 10);
    var used = {}; items.forEach(function (x) { used[x.c] = 1; });
    if (cats.length >= 2 && items.length >= 4 && Object.keys(used).length === cats.length) out.group.push({ text: aiStr_(q.text, 300) || "Розподіліть за групами:", cats: cats, items: items, src: src(q) });
  });
  arr("find").forEach(function (q) {
    var pass = String(q.passage == null ? "" : q.passage).replace(/\s+/g, " ").trim().substring(0, 900), marks = pass.match(/\*[^*\s][^*]*\*/g) || [];
    if (aiStr_(q.text, 300) && marks.length >= 1 && marks.length <= 6) out.find.push({ text: aiStr_(q.text, 300), passage: pass, src: src(q) });
  });
  arr("pmcq").forEach(function (q) {
    var text = aiStr_(q.text, 400), cor = aiStr_(q.correct, 200), wr = (Array.isArray(q.wrong) ? q.wrong : []).map(function (w) { return aiStr_(w, 200); }).filter(function (w) { return w && w !== cor; });
    wr = wr.filter(function (w, i) { return wr.indexOf(w) === i; }).slice(0, 4);
    if (text && cor && wr.length >= 2 && pg(q)) out.pmcq.push({ text: text, o: [cor].concat(wr), page: pg(q), src: src(q) });
  });
  arr("label").forEach(function (q) {
    var marks = (Array.isArray(q.marks) ? q.marks : []).map(function (m) { var p = pt01(m && m.pt); return p && aiStr_(m.a, 60) ? { x: p.x, y: p.y, a: aiStr_(m.a, 60) } : null; }).filter(function (m) { return m; }).slice(0, 6);
    var ex = (Array.isArray(q.extra) ? q.extra : []).map(function (w) { return aiStr_(w, 60); }).filter(function (w, i, a) { return w && a.indexOf(w) === i; }).slice(0, 3);
    if (marks.length >= 2 && pg(q)) out.label.push({ text: aiStr_(q.text, 300) || "Підпишіть частини на малюнку:", marks: marks, extra: ex, page: pg(q), src: src(q) });
  });
  arr("tap").forEach(function (q) {
    var tg = (Array.isArray(q.targets) ? q.targets : []).map(function (m) { var p = pt01(m && m.pt); return p ? { x: p.x, y: p.y, r: 0.04 } : null; }).filter(function (m) { return m; }).slice(0, 3);
    if (aiStr_(q.text, 300) && tg.length && pg(q)) out.tap.push({ text: aiStr_(q.text, 300), targets: tg, page: pg(q), src: src(q) });
  });
  AI_TYPES_.forEach(function (k) { out[k] = out[k].slice(0, c[k] || 0); });
  return out;
}
function aiGenerate_(p) {
  var pr = PropertiesService.getScriptProperties();
  var key = pr.getProperty("GEMINI_KEY");
  if (!key) return { error: "Не задано ключ Gemini. Apps Script → Налаштування проєкту → Властивості скрипту → додайте GEMINI_KEY." };
  var model = pr.getProperty("GEMINI_MODEL") || "gemini-3.8-flash";
  var spec; try { spec = JSON.parse(p.spec || "{}"); } catch (e) { return { error: "Некоректні налаштування генерації" }; }
  var images = []; try { images = JSON.parse(p.images || "[]"); } catch (e) {}
  var text = String(p.text || "").substring(0, 120000);
  var c = aiCounts_(spec.counts);
  if (!c.total) return { error: "Вкажіть, скільки питань потрібно хоча б одного типу" };
  if (!images.length && text.replace(/\s+/g, "").length < 200) return { error: "У матеріалі замало тексту (потрібно хоча б кілька речень)" };
  var first = Math.max(1, parseInt(spec.firstPage, 10) || 1);
  var parts = [{ text: aiPrompt_(spec, c) }];
  if (text) parts.push({ text: "ТЕКСТ МАТЕРІАЛУ (витягнуто з файлу, може бути неповним):\n" + text });
  images.slice(0, 40).forEach(function (b64, i) {
    parts.push({ text: "Слайд/сторінка " + (first + i) + ":" });
    parts.push({ inline_data: { mime_type: "image/jpeg", data: String(b64) } });
  });
  var body = { contents: [{ role: "user", parts: parts }], generationConfig: { temperature: 0.4, responseMimeType: "application/json" } };
  var models = [model], fb = pr.getProperty("GEMINI_FALLBACK") || "gemini-2.5-flash";
  if (fb && fb !== model) models.push(fb);
  var res, code = 0, raw = "", used = model, tried = {};
  var call = function (m) {
    tried[m] = 1; used = m;
    for (var at = 0; at < 2; at++) {   // при перевантаженні (503/429/500) пробуємо ще раз
      try {
        res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(m) + ":generateContent",
          { method: "post", contentType: "application/json", headers: { "x-goog-api-key": key }, payload: JSON.stringify(body), muteHttpExceptions: true });
      } catch (e) { return { error: "Не вдалося звернутися до Gemini: " + String(e.message || e).replace(key, "***").substring(0, 200) }; }
      code = res.getResponseCode(); raw = res.getContentText();
      if (code === 503 || code === 429 || code === 500) { if (at === 0) Utilities.sleep(4000); continue; }
      break;
    }
    return null;
  };
  for (var mi = 0; mi < models.length && (mi === 0 || code === 503 || code === 429 || code === 500 || code === 404); mi++) {
    var fe = call(models[mi]); if (fe) return fe;
    if (code === 200) break;
  }
  if (code === 404 || code === 400 && /model/i.test(raw)) {   // моделі застаріли — беремо найновішу доступну «flash» зі списку Gemini
    try {
      var lr = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models?pageSize=200", { headers: { "x-goog-api-key": key }, muteHttpExceptions: true });
      var ms = (JSON.parse(lr.getContentText()).models || []).filter(function (m) {
        return /flash/i.test(m.name) && !/lite|image|tts|live|audio|thinking|exp/i.test(m.name) && (m.supportedGenerationMethods || []).indexOf("generateContent") >= 0;
      }).map(function (m) { return m.name.replace(/^models\//, ""); });
      var ver = function (n) { var x = n.match(/(\d+(?:\.\d+)?)/); return x ? parseFloat(x[1]) : 0; };
      ms.sort(function (a, b) { return ver(b) - ver(a) || (/preview/i.test(a) ? 1 : 0) - (/preview/i.test(b) ? 1 : 0); });
      for (var k = 0; k < ms.length && k < 3; k++) {
        if (tried[ms[k]]) continue;
        var fe2 = call(ms[k]); if (fe2) return fe2;
        if (code === 200) { pr.setProperty("GEMINI_MODEL", ms[k]); break; }   // запам'ятовуємо робочу модель
      }
    } catch (e) {}
  }
  var j; try { j = JSON.parse(raw); } catch (e) { j = null; }
  if (code !== 200) return { error: (code === 503 || code === 429 ? "Gemini зараз перевантажений. Спробуйте за кілька хвилин. " : "") + "Gemini відповів помилкою " + code + (j && j.error && j.error.message ? ": " + String(j.error.message).replace(key, "***").substring(0, 300) : "") };
  var out = j && j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts;
  var txt = out ? out.map(function (x) { return x.text || ""; }).join("") : "";
  if (!txt) return { error: "Gemini не повернув питань" + (j && j.promptFeedback && j.promptFeedback.blockReason ? " (заблоковано: " + j.promptFeedback.blockReason + ")" : "") + ". Спробуйте інший файл." };
  txt = txt.replace(/^\s*```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "");
  var obj; try { obj = JSON.parse(txt); } catch (e) { return { error: "Gemini повернув відповідь у неправильному форматі. Спробуйте ще раз." }; }
  var clean = aiClean_(obj, c), got = 0;
  AI_TYPES_.forEach(function (k) { got += clean[k].length; });
  if (!got) return { error: "Не вдалося отримати придатних питань. Спробуйте ще раз або змініть матеріал." };
  return { ok: true, data: clean, model: used };
}

/* ===== Списки груп: перегляд і редагування учнів (аркуші «Журнал N клас») ===== */
function stuGet_() {
  var out = {}, names = KNOWN_GROUPS.slice();
  Object.keys(ROSTER).forEach(function (grade) {
    var sh = journalSheet_(grade), last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r, i) {
      var name = String(r[0]).trim(), g = canonClass_(r[1]);
      if (!name || !g) return;
      (out[g] = out[g] || []).push({ name: name, row: i + 2, grade: grade });
      if (names.indexOf(g) < 0) names.push(g);
    });
  });
  return { ok: true, groups: names, students: out };
}
function stuSheet_(group) {
  var g = canonClass_(group), grade = gradeOf_(g);
  if (!g || !grade || !ROSTER[grade]) return null;
  return { grade: grade, group: g, sh: journalSheet_(grade) };
}
function stuMutate_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var kind = p.kind, name = String(p.name || "").replace(/\s+/g, " ").trim();
    if (kind === "add") {
      var names = name ? [name] : [];
      if (p.names) names = String(p.names).split(/\r?\n/).map(function (x) { return x.replace(/\s+/g, " ").trim(); }).filter(Boolean);
      if (!names.length) return { error: "Введіть ім'я учня" };
      var a = stuSheet_(p.group); if (!a) return { error: "Невідома група" };
      var last = a.sh.getLastRow(), have = {};
      if (last > 1) a.sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r, i) {
        var k = normName_(r[0]); if (k) have[k] = { row: i + 2, g: canonClass_(r[1]) };
      });
      var rows = [], skipped = [], revived = 0;
      names.forEach(function (n) {
        if (n.length < 3 || n.length > 80) { skipped.push(n); return; }
        var h = have[normName_(n)];
        if (h) {
          if (h.g) { skipped.push(n + " (уже є в групі " + h.g + ")"); return; }
          // учня раніше прибрали з групи (рядок і оцінки лишилися) — повертаємо в групу
          a.sh.getRange(h.row, 2).setNumberFormat("@").setValue(a.group); h.g = a.group; revived++; return;
        }
        have[normName_(n)] = { row: -1, g: a.group }; rows.push([n, a.group]);
      });
      if (rows.length) {
        a.sh.getRange(last + 1, 2, rows.length, 1).setNumberFormat("@");
        a.sh.getRange(last + 1, 1, rows.length, 2).setValues(rows);
      }
      CacheService.getScriptCache().remove("ROSTER");
      var res = stuGet_(); res.added = rows.length + revived; res.skipped = skipped; return res;
    }
    var grade = String(p.grade || ""), row = parseInt(p.row, 10);
    if (!ROSTER[grade] || !(row >= 2)) return { error: "Некоректний запис" };
    var sh = journalSheet_(grade);
    if (row > sh.getLastRow() || normName_(sh.getRange(row, 1).getValue()) !== normName_(p.old)) return { error: "Список змінився. Оновіть сторінку." };
    if (kind === "edit") {
      var g2 = canonClass_(p.group);
      if (name.length < 3 || name.length > 80) return { error: "Некоректне ім'я" };
      if (gradeOf_(g2) !== grade) return { error: "Перенести в інший клас тут не можна: додайте учня в нову групу, а в цій приберіть." };
      sh.getRange(row, 2).setNumberFormat("@");
      sh.getRange(row, 1, 1, 2).setValues([[name, g2]]);
    } else if (kind === "remove") {
      sh.getRange(row, 2).setNumberFormat("@").setValue("");   // з групи прибираємо, а рядок і оцінки в журналі лишаються
    } else return { error: "kind" };
    CacheService.getScriptCache().remove("ROSTER");
    return stuGet_();
  } finally { lock.releaseLock(); }
}

/* ===== Таблиця оцінок групи для сайту (лише для вчителя) ===== */
function jviewGet_(p) {
  var g = canonClass_(p.group), grade = gradeOf_(g);
  if (!g || !grade || !ROSTER[grade]) return { error: "Невідома група" };
  var sh = journalSheet_(grade), lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, group: g, cols: [], rows: [] };
  var notes = lastCol >= 1 ? sh.getRange(1, 1, 1, lastCol).getNotes()[0] : [];
  var vals = sh.getRange(2, 1, lastRow - 1, Math.max(2, lastCol)).getValues();
  var cols = [];
  for (var c = 2; c < lastCol; c++) {
    var parts = String(notes[c]).split("\n"), key = parts[0] || "";
    if (!key) continue;
    var k = key.split("|"), att = key.indexOf(ATT_TITLE + "|") === 0;
    if ((att ? k[2] : k[2]) !== g) continue;
    cols.push({ c: c, att: att, d: parts[1] || "", t: att ? "Відвідуваність" + (k[1] ? " (урок " + k[1] + ")" : "") : k[0] });
  }
  cols.sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : a.c - b.c; });
  var rows = [];
  vals.forEach(function (r) {
    var name = String(r[0]).trim();
    if (!name || canonClass_(r[1]) !== g) return;
    rows.push({ name: name, v: cols.map(function (x) { var m = r[x.c]; return m === "" || m == null ? "" : String(m).trim(); }) });
  });
  rows.sort(function (a, b) { return a.name.localeCompare(b.name, "uk"); });
  return { ok: true, group: g, cols: cols.map(function (x) { return { t: x.t, d: x.d, att: x.att }; }), rows: rows };
}

/* ----- запити редактора ----- */
function teacherApi_(p) {
  if (String(p.pin) !== TEACHER_PIN) return jsonOut_({ error: "pin" });
  try {
    switch (p.action) {
      case "t_auth": return jsonOut_({ ok: true });
      case "t_grade_set": return jsonOut_(gradeSet_(p));
      case "t_settings_get": return jsonOut_(settingsGet_());
      case "t_settings_set": return jsonOut_(settingsSet_(p));
      case "t_list": return jsonOut_({ tests: listTests_() });
      case "t_boot": return jsonOut_({ tests: listTests_(), mk: mkList_() });
      case "t_folder": return jsonOut_(setTestFolder_(p));
      case "t_mat_list": return jsonOut_(matList_());
      case "t_mat_save": return jsonOut_(matSave_(p));
      case "t_mat_meta": return jsonOut_(matMeta_(p));
      case "t_mat_del": return jsonOut_(matDel_(p));
      case "t_roster": return jsonOut_({ groups: rosterMap_() });
      case "t_stu_get": return jsonOut_(stuGet_());
      case "t_jview": return jsonOut_(jviewGet_(p));
      case "t_stu": return jsonOut_(stuMutate_(p));
      case "t_live_create": return jsonOut_(liveCreate_(p));
      case "t_live_state": return jsonOut_(liveTeacherState_(p));
      case "t_live_ctl": return jsonOut_(liveTeacherCtl_(p));
      case "t_live_vote": return jsonOut_(liveTeacherVote_(p));
      case "t_get": {
        var t = loadTest_(p.id);
        return jsonOut_(t ? { test: t } : { error: "not_found" });
      }
      case "t_save": {
        var obj = JSON.parse(p.json);
        // при повторному збереженні назва аркуша результатів не змінюється
        var row0 = testRowOf_(testsSheet_(), p.id);
        if (row0 > 0) { var old = loadTest_(p.id); if (old && old.sheetName) obj.sheetName = old.sheetName; }
        var saved = saveTest_(p.id, obj);
        return jsonOut_({ ok: true, sheetName: saved.sheetName, max: maxOf_(saved) });
      }
      case "t_delete": {
        var lock = LockService.getScriptLock(); lock.waitLock(20000);
        try {
          var sh = testsSheet_(), r = testRowOf_(sh, p.id);
          if (r > 0) sh.deleteRow(r);
          bustTests_(p.id);
        } finally { lock.releaseLock(); }
        return jsonOut_({ ok: true });
      }
      case "t_ctrl": {
        if (!loadTest_(p.id)) return jsonOut_({ error: "not_found" });
        if (["open", "pause", "close", "reset", "extend"].indexOf(p.kind) < 0) return jsonOut_({ error: "kind" });
        return jsonOut_({ ctrl: setCtrl_(p.kind, p.id, p.min) });
      }
      case "t_plan_get": return jsonOut_(planGet_());
      case "t_plan_set": return jsonOut_(planSet_(p));
      case "t_plan_del": return jsonOut_(planDel_(p));
      case "t_mk_list": return jsonOut_(mkList_());
      case "t_mk_info": return jsonOut_(mkInfo_(p));
      case "t_mk_create": return jsonOut_(mkCreate_(p));
      case "t_mk_done": return jsonOut_(mkDone_(p));
      case "t_mk_update": return jsonOut_(mkUpdate_(p));
      case "t_mk_cancel": return jsonOut_(mkCancel_(p));
      case "t_mk_ack": return jsonOut_(mkAck_(p));
      case "t_ai_prep": return jsonOut_(aiPrep_(p));
      case "t_ai_clean": return jsonOut_(aiClean2_(p));
      case "t_ai_model": return jsonOut_(aiSetModel_(p));
      case "t_ai_gen": return jsonOut_(aiGenerate_(p));
      case "t_monitor": return jsonOut_(monitorJson_(p.id));
      case "t_results": return jsonOut_(resultsJson_(p.id));
      case "t_grade": return jsonOut_(setTeacherGrade_(p));
      case "t_res_del": return jsonOut_(deleteResult_(p));
      case "t_comment": return jsonOut_(setComments_(p));
      case "t_seat_get": return jsonOut_(seatGet_(p));
      case "t_seat_save": return jsonOut_(seatSave_(p));
      case "t_att_get": return jsonOut_(attGet_(p));
      case "t_att_set": return jsonOut_(attSet_(p));
      case "t_journal": return jsonOut_(journalFromEditor_(p));
      case "t_img": return jsonOut_(saveImage_(p));
      case "t_bank_list": return jsonOut_(bankList_());
      case "t_bank_save": return jsonOut_(bankSave_(p));
      case "t_bank_del": return jsonOut_(bankDel_(p));
      case "t_bank_topic": return jsonOut_(bankTopic_(p));
      case "t_backup": return jsonOut_(backupJson_());
    }
    return jsonOut_({ error: "unknown_action" });
  } catch (err) {
    return jsonOut_({ error: String(err && err.message || err) });
  }
}

// хто з учнів у якому стані (для екрана «Видача» в редакторі)
function monitorJson_(id) {
  var t = loadTest_(id);
  if (!t) return { error: "not_found" };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
  var ctrl = getCtrl_(id), rows = [];
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues(), h = data[0];
    var ix = function (n) { return h.indexOf(n); };
    for (var r = 1; r < data.length; r++) {
      var rd = ix("Раунд") >= 0 ? (parseInt(data[r][ix("Раунд")], 10) || 1) : 1;
      if (rd !== ctrl.round) continue;
      rows.push({ cls: canonClass_(data[r][1]), name: String(data[r][2]), status: String(data[r][4]),
                  auto: data[r][3], total: ix("Підсумок (авто)") >= 0 ? data[r][ix("Підсумок (авто)")] : "",
                  leaves: data[r][5], away: ix("Секунд поза сторінкою") >= 0 ? data[r][ix("Секунд поза сторінкою")] : "" });
    }
  }
  return { ctrl: ctrl, max: maxOf_(t), rows: rows };
}

function saveImage_(p) {
  var mime = String(p.mime || "image/jpeg");
  if (!/^image\/(jpeg|png|webp|gif)$/.test(mime)) return { error: "mime" };
  var bytes = Utilities.base64Decode(String(p.data || ""));
  if (bytes.length > 8 * 1024 * 1024) return { error: "big" };
  var it = DriveApp.getFoldersByName(IMG_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(IMG_FOLDER);
  var ext = mime.split("/")[1].replace("jpeg", "jpg");
  var file = folder.createFile(Utilities.newBlob(bytes, mime, cleanSheetName_(p.name || "img").substring(0, 60) + "." + ext));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var id = file.getId();
  return { ok: true, id: id, url: "https://drive.google.com/thumbnail?id=" + id + "&sz=w1600" };
}


/* ===== Папки (класи) для робіт і Матеріали (власні HTML-сторінки на сайті) ===== */
var MAT_SHEET = "_Матеріали", MAT_FOLDER = "Самостійні — матеріали";
function tFolders_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("TFOLDERS") || "{}"); } catch (e) { return {}; } }
function setTestFolder_(p) {
  var id = String(p.id || ""); if (!/^[A-Za-z0-9_\-]{3,40}$/.test(id)) return { error: "bad id" };
  var f = String(p.folder || "").replace(/\s+/g, " ").trim().substring(0, 40);
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var m = tFolders_(); if (f) m[id] = f; else delete m[id];
    PropertiesService.getScriptProperties().setProperty("TFOLDERS", JSON.stringify(m));
    CacheService.getScriptCache().remove("TLM");
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function matSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(MAT_SHEET);
  if (!sh) { sh = ss.insertSheet(MAT_SHEET); sh.getRange(1, 1, 1, 6).setValues([["ID", "Назва", "Папка", "Файл", "Оновлено", "Розмір"]]).setFontWeight("bold"); sh.setFrozenRows(1); }
  return sh;
}
function matRow_(sh, id) {
  var last = sh.getLastRow(); if (last < 2) return -1;
  var v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 2;
  return -1;
}
function matList_() {
  var sh = matSheet_(), last = sh.getLastRow(), out = [];
  if (last >= 2) sh.getRange(2, 1, last - 1, 6).getValues().forEach(function (r) {
    if (r[0]) out.push({ id: String(r[0]), title: String(r[1]), folder: String(r[2]), updated: String(r[4]), size: +r[5] || 0 });
  });
  return { ok: true, items: out };
}
function matSave_(p) {
  var html = String(p.html || "");
  if (html.length < 20) return { error: "Порожній файл" };
  if (html.length > 6 * 1024 * 1024) return { error: "Файл завеликий (понад 6 МБ)" };
  var title = String(p.title || "").replace(/\s+/g, " ").trim().substring(0, 120) || "Без назви";
  var folder = String(p.folder || "").replace(/\s+/g, " ").trim().substring(0, 40);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = matSheet_(), id = String(p.id || ""), row = id ? matRow_(sh, id) : -1;
    if (id && row < 0) return { error: "Матеріал не знайдено" };
    var it = DriveApp.getFoldersByName(MAT_FOLDER), dir = it.hasNext() ? it.next() : DriveApp.createFolder(MAT_FOLDER);
    var file;
    if (row > 0) { file = DriveApp.getFileById(String(sh.getRange(row, 4).getValue())); file.setContent(html); }
    else {
      id = Utilities.getUuid().replace(/-/g, "").substring(0, 8);
      file = dir.createFile(Utilities.newBlob(html, "text/plain", id + ".html.txt"));
      row = Math.max(2, sh.getLastRow() + 1);
    }
    sh.getRange(row, 1, 1, 6).setNumberFormat("@").setValues([[id, title, folder, file.getId(), new Date().toISOString(), html.length]]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
function matMeta_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = matSheet_(), row = matRow_(sh, p.id); if (row < 0) return { error: "Матеріал не знайдено" };
    var title = String(p.title || "").replace(/\s+/g, " ").trim().substring(0, 120), folder = String(p.folder || "").replace(/\s+/g, " ").trim().substring(0, 40);
    if (title) sh.getRange(row, 2).setValue(title);
    sh.getRange(row, 3).setNumberFormat("@").setValue(folder);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function matDel_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = matSheet_(), row = matRow_(sh, p.id); if (row < 0) return { error: "Матеріал не знайдено" };
    try { DriveApp.getFileById(String(sh.getRange(row, 4).getValue())).setTrashed(true); } catch (e) {}
    sh.deleteRow(row); return { ok: true };
  } finally { lock.releaseLock(); }
}
// Публічно за посиланням (як і самі роботи): віддає сторінку учневі
function matGet_(id) {
  id = String(id || ""); if (!/^[a-f0-9]{8}$/.test(id)) return { error: "bad id" };
  var sh = matSheet_(), row = matRow_(sh, id); if (row < 0) return { error: "not_found" };
  var r = sh.getRange(row, 1, 1, 4).getValues()[0];
  try { return { ok: true, title: String(r[1]), html: DriveApp.getFileById(String(r[3])).getBlob().getDataAsString("UTF-8") }; }
  catch (e) { return { error: "Файл недоступний" }; }
}

/* ----- перегляд відповідей, ручна оцінка, занесення в журнал (з редактора) ----- */
function resultsJson_(id) {
  var t = loadTest_(id);
  if (!t) return { error: "not_found" };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
  var rows = [];
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues(), h = data[0];
    var ix = function (n) { return h.indexOf(n); };
    for (var r = 1; r < data.length; r++) {
      var ans = {};
      h.forEach(function (name, c) { if (/^q\d+$/.test(name) && data[r][c] !== "") ans[name] = String(data[r][c]); });
      var tm = data[r][0];
      rows.push({
        cls: canonClass_(data[r][1]), name: String(data[r][2]).trim(),
        round: ix("Раунд") >= 0 ? (parseInt(data[r][ix("Раунд")], 10) || 1) : 1,
        time: tm instanceof Date ? Utilities.formatDate(tm, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm") : String(tm),
        status: String(data[r][4]), auto: data[r][3],
        open: ix("Бал відкритих (авто)") >= 0 ? data[r][ix("Бал відкритих (авто)")] : "",
        total: ix("Підсумок (авто)") >= 0 ? data[r][ix("Підсумок (авто)")] : "",
        final: ix(TEACHER_GRADE_COL) >= 0 ? data[r][ix(TEACHER_GRADE_COL)] : "",
        leaves: data[r][5], away: ix("Секунд поза сторінкою") >= 0 ? data[r][ix("Секунд поза сторінкою")] : "",
        pastes: ix("Спроб вставки") >= 0 ? data[r][ix("Спроб вставки")] : "",
        leaveLog: ix("Журнал виходів") >= 0 ? String(data[r][ix("Журнал виходів")]) : "",
        qTime: ix("Час на питання") >= 0 ? String(data[r][ix("Час на питання")]) : "",
        notes: ix("Перевірка відкритих") >= 0 ? String(data[r][ix("Перевірка відкритих")]) : "",
        comments: ix(COMMENT_COL) >= 0 ? commentsOf_(data[r][ix(COMMENT_COL)]) : {},
        answers: ans
      });
    }
  }
  return { max: maxOf_(t), ctrl: getCtrl_(id), rows: rows };
}

// Повністю видалити роботу учня (рядок у результатах)
function deleteResult_(p) {
  var t = loadTest_(p.id);
  if (!t) return { error: "not_found" };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
    if (!sh) return { error: "no_sheet" };
    var data = sh.getDataRange().getValues(), rc = data[0].indexOf("Раунд"), want = parseInt(p.round, 10) || 1, row = -1;
    for (var r = 1; r < data.length; r++) {
      var rnd = rc >= 0 ? (parseInt(data[r][rc], 10) || 1) : 1;
      if (rnd === want && clsMatch_(data[r][1], p.cls) && String(data[r][2]).trim().toLowerCase() === String(p.name).trim().toLowerCase()) row = r + 1;
    }
    if (row < 0) return { error: "Роботу не знайдено" };
    sh.deleteRow(row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// Оцінка вчителя (1–12) записується в колонку «Оцінка (вчитель)»; порожнє значення повертає автоматичну
function setTeacherGrade_(p) {
  var t = loadTest_(p.id);
  if (!t) return { error: "not_found" };
  var g = String(p.grade || "").trim() === "" ? "" : Math.round(num_(p.grade));
  if (g !== "" && !(g >= 1 && g <= GRADE_MAX)) return { error: "grade_range" };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
    if (!sh) return { error: "no_sheet" };
    var data = sh.getDataRange().getValues(), headers = data[0].slice();
    var rc = headers.indexOf("Раунд"), want = parseInt(p.round, 10) || 1, row = -1;
    for (var r = 1; r < data.length; r++) {
      var rnd = rc >= 0 ? (parseInt(data[r][rc], 10) || 1) : 1;
      if (rnd === want && clsMatch_(data[r][1], p.cls) && String(data[r][2]).trim().toLowerCase() === String(p.name).trim().toLowerCase()) row = r + 1;
    }
    if (row < 0) return { error: "no_row" };
    sh.getRange(row, col_(sh, headers, TEACHER_GRADE_COL)).setValue(g);
    return { ok: true, grade: g };
  } finally { lock.releaseLock(); }
}

// Заносить у журнал: усі завершені роботи або лише вказані учні (p.only — JSON [{cls,name,round}])
function journalFromEditor_(p) {
  var t = loadTest_(p.id);
  if (!t) return { error: "not_found" };
  var work = workFor_(t.sheetName);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
  if (!work || !sh) return { error: "no_sheet" };
  var only = null;
  if (p.only) {
    only = {};
    JSON.parse(p.only).forEach(function (x) { only[canonClass_(x.cls) + "|" + normName_(x.name) + "|" + (parseInt(x.round, 10) || 1)] = true; });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return journalPush_(sh, work, p.overwrite === "true", only); } finally { lock.releaseLock(); }
}


/* ===== Правила гри для учня, розбір після здачі, коментарі вчителя ===== */
var COMMENT_COL = "Коментарі вчителя";
function round2_(x) { return Math.round(x * 100) / 100; }
function rulesOf_(t) {
  var secs = [], total = 0, secTime = 0, readT = (t.readTime == null || t.readTime === "") ? 8 : (+t.readTime || 0);
  (t.sections || []).forEach(function (s) {
    var qs = s.questions || []; if (!qs.length) return;
    var n = s.type === "tf" ? (+s.n || 1) : Math.min(+s.n || qs.length, qs.length);
    var per = s.type === "open" ? qs.reduce(function (a, q) { return a + (+q.max || OPEN_MAX); }, 0) / qs.length : (+s.points || 0);
    var time = +s.time || 60;
    secs.push({ type: s.type, n: n, count: s.type === "tf" ? Math.min(+s.count || 4, qs.length) : 0, time: time, per: round2_(per), sum: round2_(n * per) });
    total += n * per; secTime += n * (time + readT);
  });
  return { sections: secs, max: round2_(total), readTime: readT,
           totalMin: t.noTimer === true ? 0 : (+t.totalTime > 0 ? +t.totalTime : 15),
           showGrade: t.showGrade !== false, review: t.review === true };
}

function parseAns_(str) {
  str = String(str);
  var m = str.match(/^\[([^\]|]+?)\s*\|\s*\+([\d.,]+)\]\s*([\s\S]*)$/);
  if (m) return { pid: m[1].trim(), earned: parseFloat(m[2].replace(",", ".")), shown: m[3] };
  m = str.match(/^\[([^\]]+)\]\s*Питання:\s*([\s\S]*?)\s*—\s*Відповідь:\s*([\s\S]*)$/);
  if (m) return { pid: m[1], open: true, qtext: m[2], shown: m[3] };
  return { pid: "", shown: str };
}

function findRow_(data, want, cls, name, latest) {
  var h = data[0], rc = h.indexOf("Раунд"), row = -1, best = -1;
  for (var r = 1; r < data.length; r++) {
    var rnd = rc >= 0 ? (parseInt(data[r][rc], 10) || 1) : 1;
    if (!clsMatch_(data[r][1], cls) || String(data[r][2]).trim().toLowerCase() !== String(name).trim().toLowerCase()) continue;
    if (latest) { if (rnd >= best) { best = rnd; row = r + 1; } }
    else if (rnd === want) row = r + 1;
  }
  return row;
}

function commentsOf_(raw) {
  try { var o = JSON.parse(String(raw || "{}")); return (o && typeof o === "object") ? o : {}; } catch (e) { return {}; }
}

// розбір для учня: лише коли вчитель увімкнув «Розбір для учнів» і учень уже здав роботу
function reviewJson_(q) {
  var t = loadTest_(q.review);
  if (!t) return { error: "not_found" };
  if (t.review !== true) return { error: "off" };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
  if (!sh || sh.getLastRow() < 2) return { error: "no_row" };
  var data = sh.getDataRange().getValues(), h = data[0];
  var row = findRow_(data, 0, q.cls || "", q.name || "", true);
  if (row < 0) return { error: "no_row" };
  var d = data[row - 1], ix = function (n) { return h.indexOf(n); };
  if (!isFinal_(d[4])) return { error: "not_finished" };
  var byId = {};
  (t.sections || []).forEach(function (s) { (s.questions || []).forEach(function (qq) { if (qq.id) byId[qq.id] = { s: s, q: qq }; }); });
  var notes = ix("Перевірка відкритих") >= 0 ? String(d[ix("Перевірка відкритих")]).split("\n") : [];
  var cm = ix(COMMENT_COL) >= 0 ? commentsOf_(d[ix(COMMENT_COL)]) : {};
  var slots = h.filter(function (n) { return /^q\d+$/.test(n) && d[h.indexOf(n)] !== ""; })
    .sort(function (a, b) { return parseInt(a.slice(1), 10) - parseInt(b.slice(1), 10); });
  var items = slots.map(function (slot) {
    var a = parseAns_(d[h.indexOf(slot)]), it = { n: slot.slice(1), title: "", shown: a.shown, comment: cm[slot] || "" }, info = byId[a.pid];
    if (a.open) {
      it.title = a.qtext; it.open = true;
      var line = notes.filter(function (l) { return l.indexOf(slot + " ") === 0; })[0] || "";
      var mm = line.match(/:\s*([\d.,]+)\/([\d.,]+)/);
      if (mm) { it.earned = parseFloat(mm[1].replace(",", ".")); it.max = parseFloat(mm[2].replace(",", ".")); }
      if (info && info.q.concepts) it.good = "Очікувані поняття: " + info.q.concepts.map(function (c) { return c.n; }).filter(Boolean).join("; ");
    } else if (a.pid.indexOf("tf:") === 0) {
      var ids = a.pid.slice(3).split(",").map(function (x) { return parseInt(x, 10); });
      var sec = (t.sections || []).filter(function (s) { return s.type === "tf" && ids.every(function (i) { return i <= (s.questions || []).length; }); })[0];
      it.title = "Так / Ні"; it.max = sec ? +sec.points : null; it.earned = a.earned; it.shown = "";
      var got = {}; a.shown.replace(/(\d+)\)(Так|Ні|—)/g, function (_, i, v) { got[i] = v; return ""; });
      it.rows = ids.map(function (id, k) { var st = sec && sec.questions[id - 1]; if (!st) return null; var v = got[k + 1] || "—", right = st.c ? "Так" : "Ні"; return { t: st.t, v: v, right: right, ok: v === right }; }).filter(Boolean);
    } else {
      var qq = info && info.q, ss = info && info.s;
      it.title = qq ? qq.text : "(питання змінено вчителем)";
      if (ss && ss.type === "cloze" && qq) it.title = String(qq.text).replace(/\[\[[^\]]+\]\]/g, "_____"); it.max = ss ? +ss.points : null; it.earned = a.earned;
      if (qq && ss) {
        if (ss.type === "mcq") it.good = "Правильно: " + qq.o[0];
        else if (ss.type === "multi") it.good = "Правильно: " + qq.o.filter(function (_, i) { return (qq.c || []).indexOf(i) >= 0; }).join("; ");
        else if (ss.type === "short") it.good = "Прийнятні відповіді: " + (qq.acc || []).join(", ");
        else if (ss.type === "match") it.good = "Правильно: " + qq.pairs.map(function (p) { return p[0] + " → " + p[1]; }).join("; ");
        else if (ss.type === "label") it.good = "Правильно: " + qq.marks.map(function (m, i) { return (i + 1) + ") " + m.a; }).join("; ");
        else if (ss.type === "order") it.good = "Правильний порядок: " + (qq.items || []).join(" → ");
        else if (ss.type === "cloze") it.good = "Правильно: " + (String(qq.text).match(/\[\[[^\]]+\]\]/g) || []).map(function (g, i) { return (i + 1) + ") " + g.slice(2, -2); }).join("; ");
        else if (ss.type === "num") it.good = "Правильно: " + qq.ans + (+qq.tol ? " ± " + qq.tol : "") + (qq.unit ? " " + qq.unit : "");
        else if (ss.type === "group") it.good = "Правильно: " + (qq.items || []).map(function (x) { return x.t + " → " + (qq.cats || [])[x.c]; }).join("; ");
        else if (ss.type === "find") it.good = "Треба було виділити: " + (String(qq.passage).match(/\S*\*\S*/g) || []).map(function (x) { return x.replace(/\*/g, ""); }).join("; ");
      }
    }
    return it;
  });
  var max = maxOf_(t), total = num_(d[ix("Підсумок (авто)")]);
  if (isNaN(total)) total = num_(d[3]) || 0;
  var fin = ix(TEACHER_GRADE_COL) >= 0 ? d[ix(TEACHER_GRADE_COL)] : "";
  var out = { ok: true, title: t.title, status: String(d[4]), total: round2_(total), max: max, items: items, general: cm.all || "" };
  if (t.showGrade !== false) { out.grade = fin !== "" ? fin : gradeFromScore_(total, max); out.byTeacher = fin !== ""; }
  return out;
}

// коментарі вчителя: {q3: "...", all: "..."} у колонці «Коментарі вчителя»
function setComments_(p) {
  var t = loadTest_(p.id);
  if (!t) return { error: "not_found" };
  var o = commentsOf_(p.json), clean = {};
  Object.keys(o).forEach(function (k) { var v = String(o[k] || "").trim(); if (v && /^(all|q\d+)$/.test(k)) clean[k] = v.substring(0, 1500); });
  var s = JSON.stringify(clean);
  if (s.length > 40000) return { error: "too_long" };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.sheetName);
    if (!sh) return { error: "no_sheet" };
    var data = sh.getDataRange().getValues(), headers = data[0].slice();
    var row = findRow_(data, parseInt(p.round, 10) || 1, p.cls, p.name, false);
    if (row < 0) return { error: "no_row" };
    sh.getRange(row, col_(sh, headers, COMMENT_COL)).setValue(s === "{}" ? "" : s);
    return { ok: true };
  } finally { lock.releaseLock(); }
}


/* ===== Розсадка та відвідуваність: Н / З у «Журналі N клас» (колонка на дату) ===== */
var SEAT_SHEET = "_Розсадка";   // A група | B схема (JSON) | C оновлено
var ATT_TITLE = "Відвідуваність";
function seatSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SEAT_SHEET);
  if (!sh) { sh = ss.insertSheet(SEAT_SHEET); sh.getRange(1, 1, 1, 3).setValues([["Група", "Схема розсадки (JSON)", "Оновлено"]]).setFontWeight("bold"); sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat("@"); }
  return sh;
}
// Схема розсадки в кабінеті: основний кабінет зберігається під назвою групи (як раніше), інші — «група @ кабінет»
var SEAT_SEP = " @ ";
function seatKey_(p) {
  var cls = canonClass_(p.cls); if (!cls) return "";
  var room = String(p.room || "").replace(/\s+/g, " ").trim().substring(0, 40);
  return room && room !== "main" ? cls + SEAT_SEP + room : cls;
}
function seatRow_(sh, key) {
  var last = sh.getLastRow(); if (last < 2) return -1;
  var v = sh.getRange(2, 1, last - 1, 1).getValues(), room = String(key).indexOf(SEAT_SEP) >= 0;
  for (var i = 0; i < v.length; i++) {
    var c = v[i][0], isRoom = String(c).indexOf(SEAT_SEP) >= 0;
    if (room ? String(c) === String(key) : (!isRoom && clsMatch_(c, key))) return i + 2;
  }
  return -1;
}
function seatGet_(p) {
  var cls = seatKey_(p), sh = seatSheet_(), r = seatRow_(sh, cls);
  if (r < 0) return { layout: null };
  try { return { layout: JSON.parse(String(sh.getRange(r, 2).getValue())) }; } catch (e) { return { layout: null }; }
}
function seatSave_(p) {
  var cls = seatKey_(p); if (!cls) return { error: "cls" };
  var obj; try { obj = JSON.parse(p.json); } catch (e) { return { error: "json" }; }
  var s = JSON.stringify(obj); if (s.length > 40000) return { error: "too_long" };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = seatSheet_(), r = seatRow_(sh, cls);
    if (r < 0) { r = sh.getLastRow() + 1; sh.getRange(r, 1).setNumberFormat("@").setValue(cls); }
    sh.getRange(r, 2).setValue(s); sh.getRange(r, 3).setValue(new Date());
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function attIso_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d || "")) ? String(d) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
var GRD_TITLE = "Оцінка за урок";
function attKey_(cls, lesson, title) { return (title || ATT_TITLE) + "|" + (lesson || "") + "|" + cls; }
// колонка відвідуваності: шукаємо за ключем у примітці (+ дата); create — додати, якщо немає
function attCol_(sh, cls, iso, lesson, create, title) {
  var key = attKey_(cls, lesson, title), lastCol = sh.getLastColumn();
  var notes = lastCol >= 1 ? sh.getRange(1, 1, 1, lastCol).getNotes()[0] : [];
  for (var c = 2; c < notes.length; c++) {
    var parts = String(notes[c]).split("\n");
    if (parts[0] === key && parts[1] === iso) return c + 1;
  }
  if (!create) return -1;
  var col = lastCol + 1, hdr = iso.slice(8, 10) + "." + iso.slice(5, 7) + " " + (title || ATT_TITLE) + " · " + cls + (lesson ? " (урок " + lesson + ")" : "");
  sh.getRange(1, col).setValue(hdr).setNote(key + "\n" + iso).setFontWeight("bold").setWrap(true);
  sh.setColumnWidth(col, 120);
  return col;
}
function attGet_(p) {
  var cls = canonClass_(p.cls), grade = gradeOf_(cls);
  if (!grade || !ROSTER[grade]) return { error: "cls" };
  var sh = journalSheet_(grade), iso = attIso_(p.date), col = attCol_(sh, cls, iso, String(p.lesson || ""), false), marks = {};
  if (col > 0 && sh.getLastRow() >= 2) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, col).getValues();
    v.forEach(function (r) { var m = String(r[col - 1]).trim(); if (m === "Н" || m === "З") marks[String(r[0]).trim()] = m; });
  }
  var grades = {}, gc = attCol_(sh, cls, iso, String(p.lesson || ""), false, GRD_TITLE);
  if (gc > 0 && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, gc).getValues().forEach(function (r) { var m = String(r[gc - 1]).trim(); if (m !== "") grades[String(r[0]).trim()] = m; });
  }
  return { ok: true, date: iso, marks: marks, grades: grades };
}
function attSet_(p) {
  var cls = canonClass_(p.cls), grade = gradeOf_(cls), mark = String(p.mark || "");
  if (!grade || !ROSTER[grade]) return { error: "cls" };
  if (mark !== "" && mark !== "Н" && mark !== "З") return { error: "mark" };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = journalSheet_(grade), iso = attIso_(p.date), lesson = String(p.lesson || "");
    var col = attCol_(sh, cls, iso, lesson, mark !== "");
    if (col < 0) return { ok: true };                    // знімати нічого: колонки ще немає
    var row = journalRow_(sh, String(p.name || ""), cls, mark === "");
    if (row === -3) return { ok: true };
    if (row < 0) return { error: "ambiguous" };
    var cell = sh.getRange(row, col);
    if (mark === "") cell.clearContent(); else cell.setValue(mark).setHorizontalAlignment("center");
    return { ok: true };
  } finally { lock.releaseLock(); }
}
// оцінка за урок (1–12) з розсадки: окрема колонка «Оцінка за урок» на дату; порожня — прибрати
function gradeSet_(p) {
  var cls = canonClass_(p.cls), grade = gradeOf_(cls), g = String(p.grade == null ? "" : p.grade).trim();
  if (!grade || !ROSTER[grade]) return { error: "cls" };
  if (g !== "" && !(/^\d{1,2}$/.test(g) && +g >= 1 && +g <= GRADE_MAX)) return { error: "grade" };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = journalSheet_(grade), iso = attIso_(p.date), lesson = String(p.lesson || "");
    var col = attCol_(sh, cls, iso, lesson, g !== "", GRD_TITLE);
    if (col < 0) return { ok: true };
    var row = journalRow_(sh, String(p.name || ""), cls, g === "");
    if (row === -3) return { ok: true };
    if (row < 0) return { error: "ambiguous" };
    var cell = sh.getRange(row, col);
    if (g === "") cell.clearContent(); else cell.setValue(+g).setHorizontalAlignment("center");
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ===== Живий урок: симуляції в класі (імена, ролі та голоси на екрані вчителя) ===== */
// Метадані (крок, відкриті результати) — у PropertiesService + кеш; кожен екран сценарію — окрема властивість (ліміт 9 КБ на значення);
// голоси, ролі та «присутність» — лише в кеші (швидко, без блокувань).
var LIVE_TTL_ = 21600;
function liveC_() { return CacheService.getScriptCache(); }
function liveGet_(prefix, id) {
  var c = liveC_(), s = c.get(prefix + "|" + id);
  if (!s) { s = PropertiesService.getScriptProperties().getProperty("L" + prefix + "|" + id); if (s) c.put(prefix + "|" + id, s, LIVE_TTL_); }
  return s ? JSON.parse(s) : null;
}
function liveSet_(prefix, id, obj) {
  var s = JSON.stringify(obj);
  PropertiesService.getScriptProperties().setProperty("L" + prefix + "|" + id, s);
  liveC_().put(prefix + "|" + id, s, LIVE_TTL_);
}
function liveStr_(v, n) { return String(v == null ? "" : v).substring(0, n); }
function liveScreens_(m) {
  var c = liveC_(), s = c.get("SC|" + m.id), out = [];
  if (s) return JSON.parse(s);
  var pr = PropertiesService.getScriptProperties();
  for (var i = 0; i < m.n; i++) out.push(JSON.parse(pr.getProperty("LS|" + m.id + "|" + i)));
  c.put("SC|" + m.id, JSON.stringify(out), LIVE_TTL_);
  return out;
}
function liveLoad_(id) {
  if (!/^[a-z0-9]{4,12}$/.test(String(id))) return null;
  var m = liveGet_("M", id); if (!m) return null;
  m.screens = liveScreens_(m); return m;
}
function liveSaveMeta_(m) { var c = {}; for (var k in m) if (k !== "screens") c[k] = m[k]; liveSet_("M", m.id, c); }

function liveOpt_(x) {
  if (typeof x === "string") return { t: liveStr_(x, 140) };
  var o = { t: liveStr_(x && x.t, 140) };
  if (x && x.s) o.s = liveStr_(x.s, 40);
  if (x && typeof x.w === "number") o.w = x.w;
  return o;
}
function liveOpts_(a) {
  var r = (a || []).slice(0, 6).map(liveOpt_);
  if (r.length < 2) throw new Error("options");
  return r;
}
function liveClean_(sc) {
  if (!sc || !sc.screens || !sc.screens.length || sc.screens.length > 20) throw new Error("scenario");
  var ids = {}, roles = {}, out = { title: liveStr_(sc.title, 120), roles: {}, screens: [], noTeacher: !!sc.noTeacher };
  ["agg", "ark"].forEach(function (k) {
    var r = (sc.roles || {})[k] || {};
    out.roles[k] = { title: liveStr_(r.title || (k === "agg" ? "Союзник агресора" : "Союзник Аркадії"), 60), text: liveStr_(r.text, 300) };
  });
  sc.screens.forEach(function (s) {
    if (!/^[A-Za-z0-9_]{1,20}$/.test(String(s.id)) || ids[s.id]) throw new Error("screen id");
    if (["info", "press", "choice"].indexOf(s.kind) < 0) throw new Error("kind");
    ids[s.id] = 1;
    var o = { id: s.id, kind: s.kind, title: liveStr_(s.title, 160), text: liveStr_(s.text, 700), prompt: liveStr_(s.prompt, 300),
              button: liveStr_(s.button, 160), revealLabel: liveStr_(s.revealLabel, 100) };
    if (s.hintBy) o.hintBy = { agg: liveStr_(s.hintBy.agg, 300), ark: liveStr_(s.hintBy.ark, 300) };
    if (s.kind === "choice") {
      if (s.options) { o.common = true; o.optionsBy = { agg: liveOpts_(s.options), ark: liveOpts_(s.options) }; }
      else o.optionsBy = { agg: liveOpts_((s.optionsBy || {}).agg), ark: liveOpts_((s.optionsBy || {}).ark) };
      if (typeof s.correct === "number") o.correct = s.correct;
      if (s.meter) {
        var mt = s.meter;
        o.meter = { label: liveStr_(mt.label || "Скарбниця агресора", 60), unit: liveStr_(mt.unit || "", 20),
                    start: +mt.start || 100, floor: +mt.floor || 0, threshold: +mt.threshold || 25 };
      }
    }
    if (s.result) o.result = { tone: s.result.tone === "good" ? "good" : (s.result.tone === "bad" ? "bad" : "info"), title: liveStr_(s.result.title, 120), text: liveStr_(s.result.text, 500) };
    out.screens.push(o);
  });
  return out;
}
function liveCreate_(p) {
  var sc = liveClean_(JSON.parse(p.scenario)), groups = JSON.parse(p.groups || "[]");
  if (!groups.length) throw new Error("groups");
  var ros = rosterMap_(), r = {}, i;
  groups.forEach(function (g) { r[g] = (ros[g] || []).slice(); });
  var alphabet = "abcdefghjkmnpqrstuvwxyz23456789", id = "";
  for (i = 0; i < 6; i++) id += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  var m = { id: id, title: sc.title, roles: sc.roles, noTeacher: !!sc.noTeacher, n: sc.screens.length, groups: groups, step: 0, revealed: {}, created: Date.now() };
  var pr = PropertiesService.getScriptProperties();
  sc.screens.forEach(function (s, k) {
    var js = JSON.stringify(s); if (js.length > 8500) throw new Error("big");
    pr.setProperty("LS|" + id + "|" + k, js);
  });
  liveSet_("M", id, m); liveSet_("R", id, r);
  // прибираємо старі сесії (> 3 діб)
  var all = pr.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf("LM|") !== 0) return;
    try {
      var o = JSON.parse(all[k]);
      if (Date.now() - o.created > 3 * 86400000) {
        var sid = k.substring(3); pr.deleteProperty(k); pr.deleteProperty("LR|" + sid);
        for (var j = 0; j < (o.n || 0); j++) pr.deleteProperty("LS|" + sid + "|" + j);
      }
    } catch (e) {}
  });
  return { ok: true, id: id };
}
function liveKeyP_(id, cls, name) { return "LP|" + id + "|" + cls + "|" + name; }
function liveKeyV_(id, sid, cls, name) { return "LV|" + id + "|" + sid + "|" + cls + "|" + name; }
function liveKeyR_(id, cls, name) { return "LRO|" + id + "|" + cls + "|" + name; }
function liveMember_(id, cls, name) {
  var r = liveGet_("R", id); return !!(r && r[cls] && r[cls].indexOf(name) >= 0);
}
function liveGetAll_(keys) {
  var out = {}, c = liveC_();
  for (var i = 0; i < keys.length; i += 90) { var part = c.getAll(keys.slice(i, i + 90)); for (var k in part) out[k] = part[k]; }
  return out;
}
// Роль призначається при першому підключенні: кожен наступний учень отримує роль, якої зараз менше (поровну)
function liveRole_(m, cls, name) {
  var c = liveC_(), k = liveKeyR_(m.id, cls, name), r = c.get(k);
  if (r) return r;
  var lock = LockService.getScriptLock(), locked = false;
  try { lock.waitLock(5000); locked = true; } catch (e) {}
  try {
    r = c.get(k); if (r) return r;
    var ca = +(c.get("LCA|" + m.id) || 0), cb = +(c.get("LCB|" + m.id) || 0);
    r = ca < cb ? "agg" : (cb < ca ? "ark" : (Math.random() < 0.5 ? "agg" : "ark"));
    c.put(k, r, LIVE_TTL_);
    if (r === "agg") c.put("LCA|" + m.id, String(ca + 1), LIVE_TTL_); else c.put("LCB|" + m.id, String(cb + 1), LIVE_TTL_);
    return r;
  } finally { if (locked) lock.releaseLock(); }
}
function liveTexts_(arr) { return arr.map(function (o) { return o.t; }); }
function livePub_(sc, role) {
  var o = { id: sc.id, kind: sc.kind, title: sc.title, text: sc.text, prompt: sc.prompt, button: sc.button,
            hint: sc.hintBy ? sc.hintBy[role] || "" : "", options: sc.optionsBy ? liveTexts_(sc.optionsBy[role]) : null };
  if (sc.meter) o.meter = sc.meter;
  return o;
}
function liveInfo_(id) {
  var m = liveGet_("M", id); if (!m) return { error: "not_found" };
  return { ok: true, title: m.title, groups: liveGet_("R", id) || {} };
}
function liveState_(q) {
  var m = liveLoad_(q.live); if (!m) return { error: "not_found" };
  var cls = String(q.cls || ""), name = String(q.name || "");
  if (!liveMember_(m.id, cls, name)) return { error: "not_member" };
  var c = liveC_(); c.put(liveKeyP_(m.id, cls, name), String(Date.now()), 600);
  var role = liveRole_(m, cls, name), sc = m.screens[m.step], rev = !!m.revealed[sc.id], mine = c.get(liveKeyV_(m.id, sc.id, cls, name));
  var out = { ok: true, title: m.title, role: role, roleInfo: m.roles[role], step: m.step, total: m.screens.length, screen: livePub_(sc, role), revealed: rev, my: mine ? mine.split("|")[0] : null };
  if (sc.meter) { var mt = c.get("LMT|" + m.id + "|" + sc.id); out.meter = mt ? JSON.parse(mt) : liveMeter_(sc, []); }
  if (rev) { out.result = sc.result || null; if (typeof sc.correct === "number") out.correct = sc.correct; }
  return out;
}
function liveVote_(m, role, cls, name, sid, v) {
  var sc = m.screens[m.step];
  if (sc.id !== sid || m.revealed[sc.id]) return { error: "closed" };
  if (sc.kind === "press") v = "1";
  else if (sc.kind === "choice") { v = String(parseInt(v, 10)); if (!(+v >= 0 && +v < sc.optionsBy[role].length)) return { error: "value" }; }
  else return { error: "kind" };
  liveC_().put(liveKeyV_(m.id, sid, cls, name), v + "|" + Date.now(), LIVE_TTL_);
  return { ok: true, my: v };
}
function liveVoteStudent_(q) {
  var m = liveLoad_(q.live_vote); if (!m) return { error: "not_found" };
  var cls = String(q.cls || ""), name = String(q.name || "");
  if (!liveMember_(m.id, cls, name)) return { error: "not_member" };
  return liveVote_(m, liveRole_(m, cls, name), cls, name, String(q.screen || ""), q.v);
}
// Скарбниця агресора: кожен голос союзника Аркадії знижує її (за вагою варіанта), союзника агресора — частково повертає
function liveMeter_(sc, players) {
  var cfg = sc.meter, nArk = 0, wsum = 0, wn = 0, cap = 0;
  players.forEach(function (p) { if (!p.teacher && p.role) cap++; if (!p.teacher && p.role === "ark") nArk++; });
  sc.optionsBy.ark.forEach(function (o) { if ((o.w || 0) > 0) { wsum += o.w; wn++; } });
  var avg = wn ? wsum / wn : 1, unit = (cfg.start - cfg.floor) / (Math.max(1, nArk) * avg), ev = [];
  players.forEach(function (p) {
    if (p.teacher || p.vote === null || !p.role) return;
    var o = sc.optionsBy[p.role][+p.vote];
    if (o && o.w) ev.push({ ts: p.ts || 0, d: o.w * unit });
  });
  ev.sort(function (a, b) { return a.ts - b.ts; });
  var v = cfg.start, series = [v];
  ev.forEach(function (e) { v = Math.max(0, Math.min(cfg.start, v - e.d)); series.push(Math.round(v * 10) / 10); });
  var now = series[series.length - 1];
  return { label: cfg.label, unit: cfg.unit, start: cfg.start, floor: cfg.floor, threshold: cfg.threshold, series: series, now: now, cap: cap, reached: now <= cfg.threshold };
}
// Вчитель: стан для табло (імена, ролі, присутність, голоси поточного кроку, скарбниця)
function liveTeacherState_(p) {
  var m = liveLoad_(p.id); if (!m) return { error: "not_found" };
  var r = liveGet_("R", m.id) || {}, sc = m.screens[m.step], me = liveStr_(p.name, 40) || "Вчитель", c = liveC_();
  var list = [];
  Object.keys(r).forEach(function (g) { r[g].forEach(function (n) { list.push({ cls: g, name: n }); }); });
  if (!m.noTeacher) list.push({ cls: "T", name: me, teacher: true });
  if (p.name !== undefined) c.put(liveKeyP_(m.id, "T", me), String(Date.now()), 600);
  var keys = [];
  list.forEach(function (x) { keys.push(liveKeyP_(m.id, x.cls, x.name)); keys.push(liveKeyV_(m.id, sc.id, x.cls, x.name)); keys.push(liveKeyR_(m.id, x.cls, x.name)); });
  var got = liveGetAll_(keys), now = Date.now(), counts = { agg: 0, ark: 0 };
  var players = list.map(function (x) {
    var pr = got[liveKeyP_(m.id, x.cls, x.name)], v = got[liveKeyV_(m.id, sc.id, x.cls, x.name)];
    var role = x.teacher ? "agg" : (got[liveKeyR_(m.id, x.cls, x.name)] || null);
    if (role && !x.teacher) counts[role]++;
    var vp = v ? v.split("|") : null;
    return { cls: x.cls === "T" ? "" : x.cls, name: x.name, teacher: !!x.teacher, role: role, state: pr ? (now - (+pr) < 9000 ? "on" : "off") : "none",
             vote: vp ? vp[0] : null, ts: vp ? +vp[1] : 0 };
  });
  var out = { ok: true, id: m.id, title: m.title, roles: m.roles, noTeacher: !!m.noTeacher, step: m.step, total: m.screens.length,
    screen: { id: sc.id, kind: sc.kind, title: sc.title, text: sc.text, prompt: sc.prompt, button: sc.button, common: !!sc.common,
              optionsBy: sc.optionsBy ? { agg: sc.optionsBy.agg.map(function (o) { return { t: o.t, s: o.s || "" }; }), ark: sc.optionsBy.ark.map(function (o) { return { t: o.t, s: o.s || "" }; }) } : null,
              hintBy: sc.hintBy || null },
    revealed: !!m.revealed[sc.id], result: sc.result || null, correct: typeof sc.correct === "number" ? sc.correct : null,
    revealLabel: sc.revealLabel || "Показати результат", players: players, counts: counts,
    outline: m.screens.map(function (s) { return { id: s.id, title: s.title || s.prompt || s.id, kind: s.kind }; }) };
  if (sc.meter) { out.meter = liveMeter_(sc, players); c.put("LMT|" + m.id + "|" + sc.id, JSON.stringify(out.meter), 300); }
  return out;
}
function liveTeacherCtl_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var m = liveLoad_(p.id); if (!m) return { error: "not_found" };
    var sc = m.screens[m.step], c = liveC_();
    switch (p.op) {
      case "next": if (m.step < m.screens.length - 1) m.step++; break;
      case "prev": if (m.step > 0) m.step--; break;
      case "goto": m.step = Math.max(0, Math.min(m.screens.length - 1, parseInt(p.value, 10) || 0)); break;
      case "reveal": m.revealed[sc.id] = true; break;
      case "hide": delete m.revealed[sc.id]; break;
      case "clear": {
        var r = liveGet_("R", m.id) || {}, keys = [];
        Object.keys(r).forEach(function (g) { r[g].forEach(function (n) { keys.push(liveKeyV_(m.id, sc.id, g, n)); }); });
        keys.push(liveKeyV_(m.id, sc.id, "T", liveStr_(p.name, 40) || "Вчитель"));
        for (var i = 0; i < keys.length; i += 90) c.removeAll(keys.slice(i, i + 90));
        c.remove("LMT|" + m.id + "|" + sc.id);
        delete m.revealed[sc.id];
        break;
      }
      case "roles": {   // перерозподілити ролі заново
        var rr = liveGet_("R", m.id) || {}, rk = [];
        Object.keys(rr).forEach(function (g) { rr[g].forEach(function (n) { rk.push(liveKeyR_(m.id, g, n)); }); });
        for (var j = 0; j < rk.length; j += 90) c.removeAll(rk.slice(j, j + 90));
        c.put("LCA|" + m.id, "0", LIVE_TTL_); c.put("LCB|" + m.id, "0", LIVE_TTL_);
        return { ok: true };
      }
      case "delete": {
        var pr = PropertiesService.getScriptProperties(); pr.deleteProperty("LM|" + m.id); pr.deleteProperty("LR|" + m.id);
        for (var s = 0; s < m.n; s++) pr.deleteProperty("LS|" + m.id + "|" + s);
        c.remove("M|" + m.id); c.remove("R|" + m.id); c.remove("SC|" + m.id);
        return { ok: true, deleted: true };
      }
      default: return { error: "op" };
    }
    liveSaveMeta_(m);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function liveTeacherVote_(p) {
  var m = liveLoad_(p.id); if (!m) return { error: "not_found" };
  return liveVote_(m, "agg", "T", liveStr_(p.name, 40) || "Вчитель", String(p.screen || ""), p.v);
}


/* ============================================================
   БАНК ПИТАНЬ і РЕЗЕРВНА КОПІЯ
   ============================================================ */
var BANK_SHEET = "_Банк";
var BANK_TYPES_ = ["mcq", "multi", "match", "short", "open", "tf", "label", "tap", "order", "cloze", "num", "group", "find"];
function bankSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(BANK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BANK_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([["ID", "Тип", "Тема", "Клас", "Предмет", "Питання (JSON)", "Додано"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}
function bankList_() {
  var sh = bankSheet_(), last = sh.getLastRow(), out = [];
  if (last < 2) return { items: out };
  sh.getRange(2, 1, last - 1, 7).getValues().forEach(function (r) {
    if (!r[0]) return;
    var q; try { q = JSON.parse(String(r[5])); } catch (e) { return; }
    out.push({ id: String(r[0]), type: String(r[1]), topic: String(r[2]), grade: String(r[3]), subject: String(r[4]), q: q,
               added: r[6] instanceof Date ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[6]) });
  });
  return { items: out };
}
// p.items — JSON [{type, topic, grade, subject, q}]; однакові питання (за типом і текстом) не дублюються
function bankSave_(p) {
  var items; try { items = JSON.parse(p.items); } catch (e) { return { error: "json" }; }
  if (!Array.isArray(items) || !items.length) return { error: "Немає питань" };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = bankSheet_(), last = sh.getLastRow(), have = {};
    var keyOf = function (type, q) { return type + "|" + String((q && (q.text || q.t || q.passage || "")) || "").replace(/\s+/g, " ").trim().toLowerCase().substring(0, 200); };
    if (last > 1) sh.getRange(2, 2, last - 1, 5).getValues().forEach(function (r) { try { have[keyOf(String(r[0]), JSON.parse(String(r[4])))] = 1; } catch (e) {} });
    var rows = [], skipped = 0, now = new Date();
    items.forEach(function (it, i) {
      var type = String(it.type || ""), q = it.q;
      if (BANK_TYPES_.indexOf(type) < 0 || !q || typeof q !== "object") return;
      delete q.ai;                                  // позначка «ШІ» до банку не потрапляє
      var k = keyOf(type, q); if (have[k]) { skipped++; return; } have[k] = 1;
      var js = JSON.stringify(q); if (js.length > 45000) return;
      rows.push(["b" + now.getTime().toString(36) + i + Math.floor(Math.random() * 1296).toString(36), type,
                 String(it.topic || "").replace(/\s+/g, " ").trim().substring(0, 60), String(it.grade || "").substring(0, 12),
                 String(it.subject || "").substring(0, 40), js, now]);
    });
    if (rows.length) sh.getRange(last + 1, 1, rows.length, 7).setValues(rows);
    return { ok: true, added: rows.length, skipped: skipped };
  } finally { lock.releaseLock(); }
}
function bankDel_(p) {
  var ids; try { ids = JSON.parse(p.ids); } catch (e) { return { error: "json" }; }
  var want = {}; ids.forEach(function (x) { want[String(x)] = 1; });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = bankSheet_(), last = sh.getLastRow(), n = 0;
    if (last < 2) return { ok: true, deleted: 0 };
    var v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) if (want[String(v[i][0])]) { sh.deleteRow(i + 2); n++; }
    return { ok: true, deleted: n };
  } finally { lock.releaseLock(); }
}
function bankTopic_(p) {
  var ids; try { ids = JSON.parse(p.ids); } catch (e) { return { error: "json" }; }
  var want = {}; ids.forEach(function (x) { want[String(x)] = 1; });
  var topic = String(p.topic || "").replace(/\s+/g, " ").trim().substring(0, 60);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = bankSheet_(), last = sh.getLastRow(), n = 0;
    if (last < 2) return { ok: true, changed: 0 };
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (want[String(r[0])]) { sh.getRange(i + 2, 3).setValue(topic); n++; } });
    return { ok: true, changed: n };
  } finally { lock.releaseLock(); }
}

// Повна копія: усі роботи, папки й схеми розсадки (результати учнів, журнал і матеріали лишаються у Таблиці / на Диску)
function backupJson_() {
  var tests = [], skipped = [];
  listTests_().forEach(function (m) {
    var t = null; try { t = loadTest_(m.id); } catch (e) {}
    if (t) tests.push({ id: m.id, folder: m.folder || "", test: t }); else skipped.push(m.id);
  });
  var seats = {}, ssh = seatSheet_(), last = ssh.getLastRow();
  if (last > 1) ssh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    try { if (r[0]) seats[String(r[0])] = JSON.parse(String(r[1])); } catch (e) {}
  });
  var bank = bankList_().items;
  return { ok: true, kind: "samostiini-backup", version: 1, date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
           tests: tests, seats: seats, bank: bank, skipped: skipped };
}


/* ===== Gemini напряму з браузера вчителя (обходить ліміти Apps Script: 6 хв і розмір запиту) ===== */
function aiPrep_(p) {
  var pr = PropertiesService.getScriptProperties(), key = pr.getProperty("GEMINI_KEY");
  if (!key) return { error: "Не задано ключ Gemini. Apps Script → Налаштування проєкту → Властивості скрипту → додайте GEMINI_KEY." };
  var spec; try { spec = JSON.parse(p.spec || "{}"); } catch (e) { return { error: "Некоректні налаштування генерації" }; }
  var c = aiCounts_(spec.counts);
  if (!c.total) return { error: "Вкажіть, скільки питань потрібно хоча б одного типу" };
  return { ok: true, prompt: aiPrompt_(spec, c), key: key, model: pr.getProperty("GEMINI_MODEL") || "gemini-3.8-flash", fallback: pr.getProperty("GEMINI_FALLBACK") || "gemini-2.5-flash" };
}
function aiClean2_(p) {
  var c; try { c = aiCounts_(JSON.parse(p.counts || "{}")); } catch (e) { return { error: "counts" }; }
  var txt = String(p.raw || "").replace(/^\s*```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "");
  var obj; try { obj = JSON.parse(txt); } catch (e) { return { error: "Gemini повернув відповідь у неправильному форматі. Спробуйте ще раз." }; }
  var clean = aiClean_(obj, c), got = 0;
  AI_TYPES_.forEach(function (k) { got += clean[k].length; });
  if (!got) return { error: "Не вдалося отримати придатних питань. Спробуйте ще раз або змініть матеріал." };
  return { ok: true, data: clean };
}
function aiSetModel_(p) {
  var m = String(p.model || "").trim();
  if (/^[A-Za-z0-9._\-]{3,60}$/.test(m)) PropertiesService.getScriptProperties().setProperty("GEMINI_MODEL", m);
  return { ok: true };
}


/* ============================================================
   ВІДПРАЦЮВАННЯ ЗА «Н»: призначити дату, оцінка потрапляє в журнал на день здачі,
   а якщо учень не здав до терміну — автоматично ставиться 2 (на дату терміну) і з'являється сповіщення.
   Перевірка прострочених відбувається щоразу, коли відкрито сторінку «Відпрацювання» або список робіт.
   ============================================================ */
var MK_SHEET = "_Відпрацювання";
var MK_HEAD = ["ID", "Група", "Учень", "Пропуски (дати)", "Термін", "Тема / коментар", "Створено", "Статус", "Оцінка", "Дата здачі", "Сповіщення"];
function mkSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(MK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MK_SHEET);
    sh.getRange(1, 1, 1, MK_HEAD.length).setValues([MK_HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:K").setNumberFormat("@");
  }
  return sh;
}
function mkToday_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function mkIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var m = String(v).match(/\d{4}-\d{2}-\d{2}/); return m ? m[0] : "";
}
function mkAll_() {
  var sh = mkSheet_(), last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 11).getValues().forEach(function (r, i) {
    if (!r[0]) return;
    out.push({ row: i + 2, id: String(r[0]), group: String(r[1]), name: String(r[2]), abs: String(r[3]).split(",").filter(Boolean), due: mkIso_(r[4]),
               topic: String(r[5]), created: String(r[6]), status: String(r[7]), grade: String(r[8]) === "" ? "" : num_(r[8]), done: mkIso_(r[9]), notice: String(r[10]) === "1" });
  });
  return out;
}
// записує (або, якщо grade === "", стирає) оцінку за відпрацювання в журналі на вказану дату. Колонка — окрема на кожну дату й групу.
function mkJournalPut_(group, name, iso, grade) {
  var g = canonClass_(group), gr = gradeOf_(g);
  if (!g || !gr || !ROSTER[gr] || !iso) return -2;
  var sh = journalSheet_(gr), key = "Відпрацювання|" + iso + "|" + g, lastCol = sh.getLastColumn();
  var notes = sh.getRange(1, 1, 1, lastCol).getNotes()[0], col = -1;
  for (var c = 2; c < notes.length; c++) if (String(notes[c]).split("\n")[0] === key) { col = c + 1; break; }
  if (col < 0 && grade === "") return 0;
  var row = journalRow_(sh, name, g); if (row < 0) return -1;
  if (col < 0) {
    col = lastCol + 1;
    sh.getRange(1, col).setValue(iso.slice(8, 10) + "." + iso.slice(5, 7) + " Відпрацювання · " + g).setNote(key + "\n" + iso).setFontWeight("bold").setWrap(true);
    sh.setColumnWidth(col, 120);
  }
  var cell = sh.getRange(row, col);
  if (grade === "") cell.clearContent(); else cell.setValue(grade).setHorizontalAlignment("center");
  return 1;
}
// не здав до терміну (термін — включно) → 2 на дату терміну + сповіщення
function mkSweep_() {
  var today = mkToday_(), sh = mkSheet_(), changed = [];
  mkAll_().forEach(function (m) {
    if (m.status !== "pending" || !m.due || m.due >= today) return;
    mkJournalPut_(m.group, m.name, m.due, 2);
    sh.getRange(m.row, 8, 1, 4).setValues([["failed", "2", m.due, "1"]]);
    changed.push(m.id);
  });
  return changed;
}
function mkList_() {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var swept = mkSweep_(), all = mkAll_();
    var act = all.filter(function (m) { return m.status === "pending"; }).sort(function (a, b) { return a.due < b.due ? -1 : a.due > b.due ? 1 : 0; });
    var rest = all.filter(function (m) { return m.status !== "pending"; }).sort(function (a, b) { return (b.done || "") < (a.done || "") ? -1 : (b.done || "") > (a.done || "") ? 1 : 0; }).slice(0, 80);
    return { ok: true, today: mkToday_(), swept: swept.length, items: act.concat(rest) };
  } finally { lock.releaseLock(); }
}
// «Н» учня з журналу відвідуваності цього сайту + підказки (які роботи були цього дня в групі) + його відпрацювання
function mkInfo_(p) {
  var g = canonClass_(p.group), grade = gradeOf_(g);
  if (!g || !grade || !ROSTER[grade]) return { error: "Невідома група" };
  var sh = journalSheet_(grade), lastCol = sh.getLastColumn(), lastRow = sh.getLastRow(), abs = [], works = [], want = normName_(p.name);
  if (lastRow >= 2 && lastCol >= 3) {
    var notes = sh.getRange(1, 1, 1, lastCol).getNotes()[0], vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues(), ri = -1;
    vals.forEach(function (r, i) { if (ri < 0 && normName_(r[0]) === want && canonClass_(r[1]) === g) ri = i; });
    for (var c = 2; c < lastCol; c++) {
      var parts = String(notes[c]).split("\n"), key = parts[0] || ""; if (!key) continue;
      var k = key.split("|"); if (k[2] !== g) continue;
      var iso = parts[1] || "";
      if (key.indexOf(ATT_TITLE + "|") === 0) { if (ri >= 0 && String(vals[ri][c]).trim() === "Н") abs.push({ d: iso, l: k[1] || "" }); }
      else if (k[0] !== "Відпрацювання" && k[0] !== GRD_TITLE) works.push({ d: iso, t: k[0] });
    }
  }
  abs.sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : 0; });
  var ex = mkAll_().filter(function (m) { return canonClass_(m.group) === g && normName_(m.name) === want; });
  return { ok: true, group: g, name: p.name, abs: abs, works: works, existing: ex };
}
function mkCreate_(p) {
  var g = canonClass_(p.group), name = String(p.name || "").replace(/\s+/g, " ").trim(), gr = gradeOf_(g);
  if (!g || !gr || !ROSTER[gr] || !name) return { error: "Оберіть учня" };
  var due = String(p.due || ""); if (!/^\d{4}-\d{2}-\d{2}$/.test(due)) return { error: "Оберіть дату здачі" };
  if (due < mkToday_()) return { error: "Дата здачі вже минула" };
  var abs = []; try { abs = JSON.parse(p.abs || "[]").filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); }); } catch (e) {}
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = mkSheet_(), now = new Date(), id = "mk" + now.getTime().toString(36) + Math.floor(Math.random() * 1296).toString(36);
    sh.appendRow([id, g, name, abs.join(","), due, String(p.topic || "").substring(0, 400), Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"), "pending", "", "", ""]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
function mkFind_(id) { var a = mkAll_(); for (var i = 0; i < a.length; i++) if (a[i].id === String(id)) return a[i]; return null; }
// учень здав: оцінка йде в журнал на дату здачі (за замовчуванням — сьогодні)
function mkDone_(p) {
  var g = Math.round(num_(p.grade)); if (!(g >= 1 && g <= 12)) return { error: "Оцінка має бути від 1 до 12" };
  var date = /^\d{4}-\d{2}-\d{2}$/.test(String(p.date || "")) ? String(p.date) : mkToday_();
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var m = mkFind_(p.id); if (!m) return { error: "Не знайдено" };
    if (m.status !== "pending" && m.status !== "failed") return { error: "Це відпрацювання вже закрите" };
    if (m.status === "failed" && m.done && m.done !== date) mkJournalPut_(m.group, m.name, m.done, "");   // прибираємо автоматичну двійку
    var r = mkJournalPut_(m.group, m.name, date, g);
    if (r < 0) return { error: r === -1 ? "У журналі кілька учнів із такою назвою" : "Не вдалося записати в журнал" };
    mkSheet_().getRange(m.row, 8, 1, 4).setValues([["done", String(g), date, ""]]);
    return { ok: true, grade: g, date: date };
  } finally { lock.releaseLock(); }
}
function mkUpdate_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var m = mkFind_(p.id); if (!m || m.status !== "pending") return { error: "Не знайдено або вже закрите" };
    var sh = mkSheet_();
    if (p.due) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.due)) || String(p.due) < mkToday_()) return { error: "Некоректна дата" }; sh.getRange(m.row, 5).setValue(String(p.due)); }
    if (p.topic != null) sh.getRange(m.row, 6).setValue(String(p.topic).substring(0, 400));
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function mkCancel_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var m = mkFind_(p.id); if (!m) return { error: "Не знайдено" };
    if (m.status === "failed" && m.done) mkJournalPut_(m.group, m.name, m.done, "");
    mkSheet_().getRange(m.row, 8, 1, 4).setValues([["cancelled", "", "", ""]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function mkAck_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ids = null; try { ids = JSON.parse(p.ids || "null"); } catch (e) {}
    var sh = mkSheet_(), n = 0;
    mkAll_().forEach(function (m) { if (m.notice && (!ids || ids.indexOf(m.id) >= 0)) { sh.getRange(m.row, 11).setValue(""); n++; } });
    return { ok: true, acked: n };
  } finally { lock.releaseLock(); }
}


/* ============================================================
   ПЛАН ЗАНЯТЬ: дата → тема для класу («6») або конкретної групи («6-1»). Використовується на сторінці «Відпрацювання».
   ============================================================ */
var PLAN_SHEET = "_План";
function planSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(PLAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PLAN_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([["Клас / група", "Дата", "Тема", "Додано"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:D").setNumberFormat("@");
  }
  return sh;
}
function planRows_() {
  var sh = planSheet_(), last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
    var d = mkIso_(r[1]); if (!r[0] || !d) return;
    out.push({ k: String(r[0]).trim(), d: d, t: String(r[2]), a: String(r[3]) });
  });
  return out;
}
function planGet_() {
  var plans = {};
  planRows_().forEach(function (r) { (plans[r.k] = plans[r.k] || []).push({ d: r.d, t: r.t }); });
  Object.keys(plans).forEach(function (k) { plans[k].sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : 0; }); });
  return { ok: true, plans: plans };
}
// p.key — «6» або «6-1»; p.rows — JSON [{d,t}]; p.mode — «merge» (замінити лише дати з файлу) або «replace» (замінити весь план)
function planSet_(p) {
  var key = String(p.key || "").trim();
  if (!/^\d{1,2}(-\d{1,2})?$/.test(key)) return { error: "Оберіть клас або групу" };
  var rows; try { rows = JSON.parse(p.rows || "[]"); } catch (e) { return { error: "json" }; }
  var fresh = [], dates = {}, now = mkToday_();
  rows.forEach(function (x) {
    var d = String(x && x.d || ""), t = String(x && x.t || "").replace(/\s+/g, " ").trim().substring(0, 300);
    if (/^\d{4}-\d{2}-\d{2}$/.test(d) && t) { fresh.push([key, d, t, now]); dates[d] = 1; }
  });
  if (!fresh.length) return { error: "Немає жодного рядка з датою й темою" };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = planSheet_(), keep = [];
    planRows_().forEach(function (r) {
      if (r.k === key && (p.mode === "replace" || dates[r.d])) return;
      keep.push([r.k, r.d, r.t, r.a]);
    });
    var all = keep.concat(fresh), last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
    sh.getRange(2, 1, all.length, 4).setNumberFormat("@").setValues(all);
    return { ok: true, saved: fresh.length };
  } finally { lock.releaseLock(); }
}
function planDel_(p) {
  var key = String(p.key || "").trim();
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = planSheet_(), keep = [];
    planRows_().forEach(function (r) { if (r.k !== key) keep.push([r.k, r.d, r.t, r.a]); });
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 4).setNumberFormat("@").setValues(keep);
    return { ok: true };
  } finally { lock.releaseLock(); }
}


/* ============================================================
   ВСТАНОВЛЕННЯ: пін і групи задаються зі сторінки start.html
   ============================================================ */
function normGroups_(v) {
  var list = Array.isArray(v) ? v : String(v || "").split(/[\n,;]+/), seen = {}, out = [];
  list.forEach(function (x) {
    x = String(x).replace(/\s+/g, "").replace(/[–—]/g, "-");
    if (x && /^\d{1,2}/.test(x) && x.length <= 12 && !seen[x]) { seen[x] = 1; out.push(x); }
  });
  return out.slice(0, 60);
}
function settingsGet_() {
  return { ok: true, groups: KNOWN_GROUPS.slice(), hasMarksKey: !!MARKS_KEY };
}
// Перше підключення: працює лише доки пін не задано (адреса скрипту секретна, а той, хто її знає, і є власником)
function setupApi_(p) {
  var pr = PropertiesService.getScriptProperties();
  if (p.action === "setup_status") return jsonOut_({ ok: true, configured: !!pr.getProperty("TEACHER_PIN"), template: true });
  if (p.action !== "setup_init") return jsonOut_({ error: "action" });
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (pr.getProperty("TEACHER_PIN")) return jsonOut_({ error: "already" });
    var pin = String(p.pin || "");
    if (pin.length < 4 || pin.length > 40) return jsonOut_({ error: "Пін має містити від 4 до 40 символів" });
    var groups = normGroups_(p.groups);
    if (!groups.length) return jsonOut_({ error: "Вкажіть хоча б одну групу (наприклад 7-А)" });
    pr.setProperty("TEACHER_PIN", pin); pr.setProperty("GROUPS", JSON.stringify(groups));
    return jsonOut_({ ok: true, groups: groups });
  } finally { lock.releaseLock(); }
}
// Зміна груп і піна пізніше (потрібен чинний пін)
function settingsSet_(p) {
  var pr = PropertiesService.getScriptProperties(), lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (p.groups != null) {
      var groups = normGroups_(p.groups);
      if (!groups.length) return { error: "Вкажіть хоча б одну групу" };
      pr.setProperty("GROUPS", JSON.stringify(groups));
    }
    if (p.newPin) {
      var np = String(p.newPin);
      if (np.length < 4 || np.length > 40) return { error: "Пін має містити від 4 до 40 символів" };
      pr.setProperty("TEACHER_PIN", np);
    }
    CacheService.getScriptCache().remove("ROSTER");
    return { ok: true, groups: normGroups_(p.groups != null ? p.groups : JSON.parse(pr.getProperty("GROUPS") || "[]")) };
  } finally { lock.releaseLock(); }
}
function menuSetup_() {
  var url = ""; try { url = ScriptApp.getService().getUrl() || ""; } catch (e) {}
  var ok = /\/exec$/.test(url), link = START_PAGE + (ok ? "?api=" + encodeURIComponent(url) : "");
  var h = '<div style="font:14px/1.5 system-ui,sans-serif;padding:4px 2px">' +
    '<p><b>Крок 1.</b> Відкрийте <b>Розширення → Apps Script</b>. Угорі праворуч: <b>Розгорнути → Нове розгортання</b> → тип «Веб-застосунок».<br>' +
    'Виконувати від імені: <b>Я</b> · Хто має доступ: <b>Усі</b> → <b>Розгорнути</b>. Google попросить дозволи: «Додатково» → «Перейти…» → «Дозволити».</p>' +
    '<p><b>Крок 2.</b> Скопіюйте «URL веб-застосунку» і відкрийте сторінку підключення:</p>' +
    (ok ? '<p style="word-break:break-all;background:#f3f1ea;padding:8px;border-radius:8px">' + url + '</p>' : '<p style="color:#a55">Розгортання ще не виконано (або адресу не знайдено) — виконайте крок 1.</p>') +
    '<p><a target="_blank" href="' + link + '" style="display:inline-block;background:#1f7a5a;color:#fff;padding:10px 16px;border-radius:10px;text-decoration:none;font-weight:700">Відкрити сторінку підключення</a></p>' +
    '<p style="color:#667">Там ви оберете пін, вкажете групи — і кабінет відкриється.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h).setWidth(520).setHeight(420), "Встановлення «Самостійних»");
}
