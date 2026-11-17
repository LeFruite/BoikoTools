// ==UserScript==
// @name         Extra: заповнити оцінки з журналу
// @namespace    https://extra.boiko.com.ua/
// @version      1.1
// @description  На сторінці «Выставить оценки списком» підставляє в списки оцінки з вашого Google-журналу за обраною датою. Оцінки НЕ зберігаються автоматично — ви перевіряєте й самі натискаєте «Выставить».
// @match        https://extra.boiko.com.ua/marks/enter-group*
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// ==/UserScript==

(function () {
  'use strict';

  // ===== НАЛАШТУВАННЯ (вставте свої значення) =====
  const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx3Ov1bucIYd-kuNLa0gs-VxU3x9m6NEYn27RSizdQnZSmWBk_w_cgoqA1JSXCZlNE/exec";   // адреса вашого веб-застосунку Apps Script
  const KEY = "ТОЙ_САМИЙ_КЛЮЧ_ЩО_Й_У_MARKS_KEY";                          // значення MARKS_KEY з Code.gs
  // ================================================

  const norm = s => String(s || "").toLowerCase().replace(/[’ʼ`']/g, "").replace(/\s+/g, " ").trim();

  // «Сахно Варвара Даніїлівна» / «Непокупний А. М.» / «Папян К.В.» → прізвище + ім'я або ініціал
  function parseName(s) {
    const t = norm(s).split(" ");
    const sur = t[0] || "";
    const f = t[1] || "";
    const isInitial = f.indexOf(".") >= 0 || f.length <= 1;
    return { sur, full: isInitial ? "" : f, ini: f.charAt(0) };
  }
  function sameStudent(a, b) {            // a, b — результат parseName
    if (!a.sur || a.sur !== b.sur) return false;
    if (a.full && b.full) return a.full === b.full;
    return a.ini === b.ini;
  }

  function getJson(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: "GET", url,
        onload: r => { try { resolve(JSON.parse(r.responseText)); } catch (e) { reject(new Error("Сервер повернув не JSON (перевірте адресу й ключ)")); } },
        onerror: () => reject(new Error("Немає зв'язку з Google Apps Script")),
      });
    });
  }

  function setSelect(sel, value) {
    if (![...sel.options].some(o => o.value === value)) return false;
    sel.value = value;
    sel.dispatchEvent(new Event("input", { bubbles: true }));
    sel.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }

  async function run() {
    const dateEl = document.querySelector('input[type="date"]');
    const date = dateEl && dateEl.value;
    if (!date) { alert("Спершу оберіть на сторінці Клас, Предмет, Тип і ДАТУ (день, коли писали самостійну)."); return; }

    const clsSel = document.querySelector(".field select");
    const clsText = clsSel && clsSel.selectedOptions[0] ? clsSel.selectedOptions[0].textContent.trim() : "";
    const gm = clsText.match(/^\d+/);
    if (!gm) { alert("Оберіть клас на сторінці (зараз: «" + clsText + "»)."); return; }

    let res;
    try {
      res = await getJson(SCRIPT_URL + "?marks=1&key=" + encodeURIComponent(KEY) + "&grade=" + gm[0] + "&date=" + encodeURIComponent(date));
    } catch (e) { alert(e.message); return; }
    if (res.error) { alert("Журнал: " + res.error); return; }
    const absent = res.absent || [];       // [[учень, дата, група, ...]] — відсутні («Н»)
    if ((!res.works || !res.works.length) && !absent.length) { alert("У журналі немає ні оцінок, ні «Н» за " + date + " для " + gm[0] + " класу."); return; }
    res.works = res.works || [];

    // якщо за цей день було кілька робіт — обираємо
    const titles = [...new Set(res.works.map(w => w.title))];
    let title = titles[0] || "(лише «Н»)";
    if (titles.length > 1) {
      const n = parseInt(prompt("За цей день кілька робіт. Введіть номер:\n" + titles.map((t, i) => (i + 1) + ". " + t).join("\n")), 10);
      if (!n || n < 1 || n > titles.length) return;
      title = titles[n - 1];
    }
    const entries = [];
    res.works.filter(w => w.title === title).forEach(w => w.rows.forEach(r => entries.push({ name: r[0], grade: r[1], cls: w.cls })));

    const btn = [...document.querySelectorAll("button")].map(b => b.textContent.trim()).find(t => /^Выставить за/.test(t)) || "";
    if (!confirm("Робота: «" + title + "», дата " + date + ".\nУчнів з оцінками в журналі: " + entries.length + "\nВідсутніх («Н»): " + absent.length +
                 "\nНа сторінці зараз: " + btn.replace(/\s+/g, " ") + "\n\nПідставити оцінки в список? (Зберігатимете ви самі кнопкою внизу.)")) return;

    const rows = [...document.querySelectorAll("table tr")].map(tr => {
      const td = tr.querySelector("td");
      return { tr, sel: tr.querySelector("select"), p: parseName(td ? td.textContent : ""), existing: ((tr.querySelector(".tags .tag") || {}).textContent || "").trim() };
    }).filter(r => r.sel);

    const filled = [], notFound = [], ambiguous = [], hasMark = [], bad = [], nFilled = [];
    // «Н» (прогул) — найвищий пріоритет: така дитина не отримує оцінку за цей день
    const absNames = new Set();
    absent.forEach(a => {
      const ep = parseName(a[0]);
      const m = rows.filter(r => sameStudent(ep, r.p));
      if (m.length !== 1) { (m.length ? ambiguous : notFound).push(a[0] + " (Н)"); return; }
      if (setSelect(m[0].sel, "absent")) { m[0].tr.style.background = "#e6ecff"; nFilled.push(a[0]); absNames.add(a[0]); }
      else bad.push(a[0] + " (Н)");
    });
    entries.forEach(e => {
      if (absNames.has(e.name)) return;     // відсутній — оцінку не ставимо
      const ep = parseName(e.name);
      const m = rows.filter(r => sameStudent(ep, r.p));
      if (m.length === 0) { notFound.push(e.name); return; }
      if (m.length > 1) { ambiguous.push(e.name); return; }
      const r = m[0];
      // оцінка підставляється завжди — навіть якщо за цей день уже є інша (наприклад, за самостійну), тип обираєте ви зверху
      if (/^\d+$/.test(r.existing)) hasMark.push(e.name + " (там уже є " + r.existing + ")");
      if (setSelect(r.sel, String(Math.round(e.grade)))) { r.tr.style.background = "#e6ffe6"; filled.push(e.name); }
      else bad.push(e.name);
    });

    alert("Підставлено оцінок: " + filled.length + " з " + entries.length + ". Поставлено «Н»: " + nFilled.length + " з " + absent.length + "." +
      (hasMark.length ? "\n\nНа сторінці за цей день уже є інші оцінки (перевірте тип зверху, щоб не задублювати):\n" + hasMark.join("\n") : "") +
      (notFound.length ? "\n\nНе знайдено на сторінці:\n" + notFound.join("\n") : "") +
      (ambiguous.length ? "\n\nНеоднозначно (кілька збігів):\n" + ambiguous.join("\n") : "") +
      (bad.length ? "\n\nНе вдалося підставити:\n" + bad.join("\n") : "") +
      "\n\nПеревірте зелені (оцінки) та сині («Н») рядки й натисніть кнопку «Выставить…» внизу.");
  }

  const b = document.createElement("button");
  b.textContent = "📒 Заповнити з журналу";
  b.style.cssText = "position:fixed;right:16px;bottom:16px;z-index:99999;padding:12px 16px;font-size:15px;background:#27ae60;color:#fff;border:0;border-radius:8px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.3)";
  b.addEventListener("click", run);
  document.body.appendChild(b);
})();
