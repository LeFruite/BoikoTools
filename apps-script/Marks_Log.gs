/**
 * Приймач для вивантаження оцінок з extra.boiko.com.ua.
 * Покладіть у НОВУ Google-таблицю (Розширення -> Apps Script), вкажіть LOG_KEY,
 * Розгорнути -> Нове розгортання -> Веб-застосунок (виконувати: від мого імені, доступ: усі).
 */
const LOG_KEY = 'ВАШ_КЛЮЧ';          // той самий, що KEY у користувацькому скрипті
const SHEET = 'Оцінки';
const HEAD = ['Дата', 'Клас', 'Категорія', 'Учень', 'Оцінка', 'Стиль мітки'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.grades) {                       // ?grades=1&key=LOG_KEY — оцінки для сторінки на вебсайті
    if (p.key !== LOG_KEY) return out_({ error: 'bad key' });
    return out_(gradesJson_());
  }
  return out_({ ok: true, hint: 'Marks log is alive' });
}

function gradesJson_() {
  const src = SpreadsheetApp.getActive().getSheetByName(SHEET);
  const res = { classes: {} };
  if (!src || src.getLastRow() < 2) return res;
  const by = {};
  src.getRange(2, 1, src.getLastRow() - 1, 5).getValues().forEach(r => {
    const d = String(r[0]).slice(0, 10), cls = String(r[1]), cat = String(r[2]), st = String(r[3]), m = String(r[4]);
    if (!cls || !st || !d) return;
    const c = by[cls] = by[cls] || { dates: {}, st: {} };
    c.dates[d] = 1;
    ((c.st[st] = c.st[st] || {})[d] = c.st[st][d] || []).push([m, cat]);
  });
  Object.keys(by).sort().forEach(cls => {
    const c = by[cls], dates = Object.keys(c.dates).sort();
    res.classes[cls] = {
      dates: dates,
      students: Object.keys(c.st).sort((a, b) => a.localeCompare(b, 'uk')).map(n => ({ n: n, c: dates.map(d => c.st[n][d] || []) }))
    };
  });
  return res;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== LOG_KEY) return out_({ error: 'bad key' });
    const sh = sheet_();
    if (d.reset) {
      sh.clearContents();
      sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
      sh.setFrozenRows(1);
      return out_({ ok: true, reset: true });
    }
    const rows = (d.rows || []).map(r => r.map(String));
    if (!rows.length) return out_({ ok: true, added: 0 });
    // без дублікатів при повторному запуску
    const last = sh.getLastRow();
    const seen = new Set();
    if (last > 1) sh.getRange(2, 1, last - 1, HEAD.length).getValues().forEach(r => seen.add(r.map(String).join('|')));
    const fresh = rows.filter(r => { const k = r.join('|'); if (seen.has(k)) return false; seen.add(k); return true; });
    if (fresh.length) sh.getRange(last + 1, 1, fresh.length, HEAD.length).setValues(fresh);
    return out_({ ok: true, added: fresh.length });
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.getRange('A:F').setNumberFormat('@');   // дати й оцінки як текст, щоб не перетворювались
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }


// ===================== ТАБЛИЦІ ПО КЛАСАХ =====================
// Для кожного класу — окремий аркуш «Журнал <клас>»: рядки = учні, стовпці = дати.
// У клітинці — оцінки за день через кому («9, 10»), «Н» — відсутність.
// Категорія кожної оцінки — у примітці клітинки (наведіть курсор). Справа — середній бал і кількість «Н».
// Запуск: меню «Оцінки» → «Оновити таблиці по класах» (або вручну функція buildGrids).

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Оцінки').addItem('Оновити таблиці по класах', 'buildGrids').addToUi();
}

function buildGrids() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SHEET);
  if (!src || src.getLastRow() < 2) { SpreadsheetApp.getUi().alert('Аркуш «' + SHEET + '» порожній'); return; }
  const data = src.getRange(2, 1, src.getLastRow() - 1, 5).getValues();   // Дата, Клас, Категорія, Учень, Оцінка
  const byClass = {};
  data.forEach(r => {
    const date = String(r[0]).slice(0, 10), cls = String(r[1]), cat = String(r[2]), st = String(r[3]), mark = String(r[4]);
    if (!cls || !st || !date) return;
    const c = byClass[cls] = byClass[cls] || { dates: {}, students: {} };
    c.dates[date] = 1;
    const s = c.students[st] = c.students[st] || {};
    (s[date] = s[date] || []).push({ mark, cat });
  });
  Object.keys(byClass).sort().forEach(cls => {
    const c = byClass[cls];
    const dates = Object.keys(c.dates).sort();
    const students = Object.keys(c.students).sort((a, b) => a.localeCompare(b, 'uk'));
    const name = 'Журнал ' + cls;
    let sh = ss.getSheetByName(name);
    if (sh) sh.clear(); else sh = ss.insertSheet(name);
    sh.clearNotes();
    const head = ['Учень'].concat(dates.map(d => d.slice(8, 10) + '.' + d.slice(5, 7)), ['Середній', 'К-сть «Н»']);
    const vals = [], notes = [];
    students.forEach(st => {
      const row = [st], nrow = [''];
      let sum = 0, n = 0, abs = 0;
      dates.forEach(d => {
        const list = c.students[st][d] || [];
        row.push(list.map(x => x.mark).join(', '));
        nrow.push(list.map(x => x.cat + ': ' + x.mark).join('\n'));
        list.forEach(x => {
          const v = parseFloat(x.mark.replace(',', '.'));
          if (x.mark === 'Н' || x.mark === 'H') abs++;
          else if (!isNaN(v)) { sum += v; n++; }
        });
      });
      row.push(n ? Math.round(sum / n * 100) / 100 : '', abs || '');
      nrow.push('', '');
      vals.push(row); notes.push(nrow);
    });
    sh.getRange(1, 1, 1, head.length).setNumberFormat('@').setValues([head]).setFontWeight('bold').setBackground('#e8eef7').setHorizontalAlignment('center');
    if (vals.length) {
      const rg = sh.getRange(2, 1, vals.length, head.length);
      rg.setNumberFormat('@').setValues(vals);
      rg.offset(0, 0, vals.length, head.length).setNotes(notes);
      sh.getRange(2, 2, vals.length, head.length - 1).setHorizontalAlignment('center');
      // «Н» сірим
      const area = sh.getRange(2, 2, vals.length, dates.length);
      sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Н').setBackground('#dddddd').setRanges([area]).build()]);
    }
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
    sh.setColumnWidth(1, 200);
    if (dates.length) sh.setColumnWidths(2, dates.length, 55);
  });
}
