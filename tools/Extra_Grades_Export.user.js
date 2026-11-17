// ==UserScript==
// @name         Extra: вивантажити всі оцінки в Google-таблицю
// @namespace    https://extra.boiko.com.ua/
// @version      2.0
// @description  Перебирає класи × дати (категорію НЕ перемикає), читає виставлені оцінки та їх справжню категорію і записує в Google-таблицю
// @match        https://extra.boiko.com.ua/marks/enter-group*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// ==/UserScript==

(function () {
  'use strict';

  // ===== НАЛАШТУВАННЯ =====
  const SCRIPT_URL = "ВСТАВТЕ_АДРЕСУ_ВЕБ-ЗАСТОСУНКУ_З_НОВОЇ_ТАБЛИЦІ";  // .../exec
  const KEY = "ВАШ_КЛЮЧ";                                               // той самий, що LOG_KEY у Marks_Log.gs
  const CLASSES = ["6 Харків", "7 Харків", "8 Харків", "9 Харків"];     // які класи вивантажувати
  // Категорії, де ви ставите оцінки (регістр, крапки й скісні риски не мають значення): решту скрипт пропускає
  const CATEGORIES = ["урок", "Дом. завдання", "С\\Р", "Відпрацювання"];
  const DEFAULT_FROM = "2026-09-01";
  // Якщо сайт не віддає категорію оцінки в даних, її вгадують за кольором мітки.
  // Заповніть після кнопки «🔍 Діагностика» (кольори: link+dark, dark, info, warning, success, danger, primary, "" = без кольору)
  const STYLE_TO_CAT = {};   // напр.: { "info": "урок", "warning": "Дом. завдання", "": "С\\Р" }
  // ========================

  const W = typeof unsafeWindow !== "undefined" ? unsafeWindow : window;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => String(s || "").replace(/\s+/g, " ").trim();

  // --- відстеження запитів сторінки (щоб знати, коли оновлення завершилось) ---
  let pending = 0, seq = 0, tracking = false, captured = [];
  function grab(url, text) { if (!text || text[0] !== '{' && text[0] !== '[') return; try { captured.push({ url, json: JSON.parse(text) }); } catch (e) {} }
  try {
    const XP = W.XMLHttpRequest.prototype, oSend = XP.send, oOpen = XP.open;
    XP.open = function (m, u) { this.__xgUrl = String(u); return oOpen.apply(this, arguments); };
    XP.send = function () {
      pending++; seq++;
      this.addEventListener("loadend", () => {
        pending = Math.max(0, pending - 1);
        try { if (this.responseType === "" || this.responseType === "text") grab(this.__xgUrl, this.responseText); } catch (e) {}
      });
      return oSend.apply(this, arguments);
    };
    if (W.fetch) {
      const oFetch = W.fetch;
      W.fetch = function () {
        pending++; seq++;
        const p = oFetch.apply(this, arguments);
        const u = String(arguments[0] && arguments[0].url || arguments[0]);
        p.then(r => { try { r.clone().text().then(t => grab(u, t)).catch(() => {}); } catch (e) {} }, () => {});
        const done = () => { pending = Math.max(0, pending - 1); };
        p.then(done, done);
        return p;
      };
    }
    tracking = true;
  } catch (e) { console.warn("Не вдалося стежити за запитами, використаю фіксовані паузи", e); }

  async function settle(seq0) {
    const t0 = Date.now();
    if (!tracking) { await sleep(1500); return; }
    while (seq === seq0 && Date.now() - t0 < 1500) await sleep(50);   // чекаємо старту запиту
    for (let i = 0; i < 3; i++) {
      while (pending > 0 && Date.now() - t0 < 20000) await sleep(50);
      await sleep(250);                                               // час на перемальовування
      if (pending === 0) break;
    }
  }

  // --- робота з елементами сторінки ---
  const selects = () => [...document.querySelectorAll(".field .select select")];
  const classSel = () => selects()[0];
  const typeSel = () => selects()[2];
  const dateEl = () => document.querySelector('input[type="date"]');
  const refreshBtn = () => [...document.querySelectorAll("button")].find(b => /Обновить/.test(b.textContent));
  // «С\Р» може містити латинські C/P: зводимо схожі літери до кирилиці й прибираємо все, крім літер
  const LAT = { a: "а", c: "с", e: "е", i: "і", o: "о", p: "р", x: "х", y: "у", k: "к", h: "н", t: "т", m: "м", b: "в" };
  const catKey = s => String(s || "").toLowerCase().replace(/[a-z]/g, ch => LAT[ch] || ch).replace(/[^а-яіїєґ]/g, "");
  const WANT = new Set(CATEGORIES.map(catKey));
  const optTexts = sel => [...sel.options].filter(o => !o.disabled).map(o => norm(o.textContent));

  function pick(sel, text) {
    const i = [...sel.options].findIndex(o => !o.disabled && norm(o.textContent) === norm(text));
    if (i < 0) return false;
    sel.selectedIndex = i;
    sel.dispatchEvent(new Event("input", { bubbles: true }));
    sel.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }
  function setDate(v) {
    const el = dateEl();
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
    setter.call(el, v);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  // читає таблицю: [{student, mark, style}]
  function readRows() {
    const out = [];
    document.querySelectorAll("table tr").forEach(tr => {
      const tds = tr.querySelectorAll("td");
      if (tds.length < 2) return;
      const student = norm(tds[0].textContent);
      tds[1].querySelectorAll(".tag").forEach(t => {
        const mark = norm(t.textContent);
        if (!mark) return;
        const style = ["is-link", "is-dark", "is-info", "is-success", "is-warning", "is-danger", "is-primary"]
          .filter(c => t.classList.contains(c)).map(c => c.slice(3)).join("+");
        out.push({ student, mark, style });
      });
    });
    return out;
  }
  const tableCount = () => document.querySelectorAll("table tr td").length;

  // --- пошук у даних запиту: об'єкти, де є і оцінка, і категорія ---
  const nameOf = v => v == null ? "" : typeof v === "object" ? norm(v.name || v.title || v.label || v.text || "") : norm(v);
  function fromJson(list) {
    const out = [];
    const walk = o => {
      if (Array.isArray(o)) return o.forEach(walk);
      if (!o || typeof o !== "object") return;
      const keys = Object.keys(o);
      const kType = keys.find(k => /type|categor|kind|work|вид/i.test(k) && nameOf(o[k]));
      const kMark = keys.find(k => /^(mark|grade|value|score|rating|ball|point)s?$|ocen|оцін|оцен/i.test(k) && o[k] != null && typeof o[k] !== "object");
      if (kType && kMark) {
        const kStu = keys.find(k => /student|pupil|user|child|учен|name/i.test(k) && nameOf(o[k]));
        const kDate = keys.find(k => /date|day|дата|created|at$/i.test(k) && typeof o[k] === "string");
        out.push({ student: kStu ? nameOf(o[kStu]) : "", mark: String(o[kMark]), cat: nameOf(o[kType]), date: kDate ? String(o[kDate]).slice(0, 10) : "" });
      }
      keys.forEach(k => walk(o[k]));
    };
    list.forEach(c => walk(c.json));
    return out;
  }

  async function loadFor(cls, date) {
    const s0 = seq;
    captured = [];
    setDate(date);
    await sleep(60);
    refreshBtn().click();
    await settle(s0);
    await sleep(150);
    let rows = readRows();
    if (!tableCount()) { await sleep(1000); rows = readRows(); }   // таблиця ще порожня — один повтор
    // точні дані сайту: /api/marks/students-of-unit.json → учні → marks[] з mark_type.title
    const ids = new Set(), out = [];
    captured.filter(c => /students-of-unit/.test(c.url) && Array.isArray(c.json)).forEach(c => c.json.forEach(st => {
      (st.marks || []).forEach(m => {
        if (m.deleted_at || ids.has(m.id)) return;
        if (m.marked_at && String(m.marked_at).slice(0, 10) !== date) return;
        ids.add(m.id);
        out.push({ student: norm(st.name), mark: String(m.value), cat: norm(m.mark_type && m.mark_type.title), style: "", by: norm(m.created_by) });
      });
    }));
    if (!out.length && rows.some(r => r.mark !== "H" && r.mark !== "Н"))
      rows.filter(r => r.mark !== "H" && r.mark !== "Н").forEach(r => out.push(Object.assign({}, r, { cat: "? немає даних API", by: "" })));
    // відсутності («Н») беремо з таблиці
    rows.filter(r => r.mark === "H" || r.mark === "Н").forEach(r => out.push({ student: r.student, mark: "Н", cat: "Н", style: r.style, by: "" }));
    return out;
  }

  function post(payload) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: "POST", url: SCRIPT_URL,
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        data: JSON.stringify(Object.assign({ key: KEY }, payload)),
        onload: r => { try { const j = JSON.parse(r.responseText); j.error ? reject(new Error(j.error)) : resolve(j); } catch (e) { reject(new Error("Сервер повернув не JSON (перевірте адресу й ключ)")); } },
        onerror: () => reject(new Error("Немає зв'язку з Google Apps Script")),
      });
    });
  }

  function* workdays(from, to) {
    const d = new Date(from + "T00:00:00Z"), end = new Date(to + "T00:00:00Z");
    for (; d <= end; d.setUTCDate(d.getUTCDate() + 1)) {
      const w = d.getUTCDay();
      if (w === 0 || w === 6) continue;
      yield d.toISOString().slice(0, 10);
    }
  }

  // --- панель ---
  const panel = document.createElement("div");
  panel.style.cssText = "position:fixed;left:16px;bottom:16px;z-index:99999;font:14px sans-serif;background:#fff;border:1px solid #ccc;border-radius:8px;padding:10px 12px;box-shadow:0 2px 8px rgba(0,0,0,.3);max-width:340px";
  const btn = document.createElement("button");
  btn.textContent = "📊 Вивантажити оцінки в таблицю";
  btn.style.cssText = "padding:10px 14px;font-size:15px;background:#2980b9;color:#fff;border:0;border-radius:8px;cursor:pointer";
  const stop = document.createElement("button");
  stop.textContent = "⏹ Зупинити";
  stop.style.cssText = "display:none;margin-left:8px;padding:10px 12px;font-size:14px;border:1px solid #c0392b;color:#c0392b;background:#fff;border-radius:8px;cursor:pointer";
  const status = document.createElement("div");
  status.style.cssText = "margin-top:8px;color:#333;white-space:pre-line";
  const diag = document.createElement("button");
  diag.textContent = "🔍 Діагностика";
  diag.style.cssText = "margin-left:8px;padding:10px 12px;font-size:14px;border:1px solid #888;color:#333;background:#fff;border-radius:8px;cursor:pointer";
  diag.onclick = async () => {
    const d = dateEl().value;
    captured = []; const s0 = seq; refreshBtn().click(); await settle(s0); await sleep(300);
    const tags = [...document.querySelectorAll("table .tag")].slice(0, 15).map(t => ({ text: norm(t.textContent), cls: t.className, title: t.title || "", attrs: [...t.attributes].map(a => a.name + "=" + a.value).join(" ") }));
    const rep = JSON.stringify({ date: d, found: fromJson(captured).slice(0, 15), tags, requests: captured.map(c => ({ url: c.url, sample: JSON.stringify(c.json).slice(0, 1500) })).slice(0, 8) }, null, 1);
    try { await navigator.clipboard.writeText(rep); status.textContent = "Діагностику скопійовано в буфер — вставте її в чат."; }
    catch (e) { console.log(rep); prompt("Скопіюйте (Ctrl+C) і надішліть у чат:", rep.slice(0, 20000)); }
  };
  panel.append(btn, stop, diag, status);
  document.body.appendChild(panel);

  let stopFlag = false;
  stop.onclick = () => { stopFlag = true; status.textContent += "\nЗупиняю після поточного кроку…"; };

  btn.onclick = async () => {
    if (SCRIPT_URL.indexOf("ВСТАВТЕ") === 0) { alert("Спершу вставте SCRIPT_URL і KEY у верхній частині скрипта."); return; }
    const todayMax = dateEl().max || new Date().toISOString().slice(0, 10);
    const from = prompt("З якої дати? (РРРР-ММ-ДД)", DEFAULT_FROM);
    if (!from) return;
    const to = prompt("По яку дату? (не пізніше " + todayMax + ")", todayMax);
    if (!to) return;
    const classes = CLASSES.filter(c => optTexts(classSel()).includes(c));
    if (!classes.length) { alert("На сторінці не знайдено жодного з класів: " + CLASSES.join(", ")); return; }
    const days = [...workdays(from, to)];
    const total = classes.length;
    const runKey = "xg_done_" + from + "_" + to;
    let done = {};
    try { done = JSON.parse(localStorage.getItem(runKey) || "{}"); } catch (e) {}
    const resume = Object.keys(done).length > 0 && confirm("Знайдено незавершений прогін за цей період (" + Object.keys(done).length + " з " + total + " частин). Продовжити з місця зупинки?\n\nOK — продовжити, Скасувати — почати заново.");
    if (!resume) {
      done = {};
      if (confirm("Очистити аркуш «Оцінки» перед записом?\n(Рекомендовано при повному перезборі, щоб не лишилось застарілих оцінок.)"))
        try { await post({ reset: true }); } catch (e) { alert(e.message); return; }
    }
    if (!confirm("Буде опрацьовано: " + classes.length + " класи × " + days.length + " днів ≈ " + (classes.length * days.length) + " оновлень сторінки.\nЦе може тривати 20–40 хв. Не закривайте і не перемикайте цю вкладку на інший сайт.\n\nПочати?")) return;

    stopFlag = false; stop.style.display = ""; btn.disabled = true;
    // запам'ятаємо, що було вибрано, щоб повернути потім
    const keep = { cls: norm(classSel().selectedOptions[0].textContent), date: dateEl().value };
    let step = 0, sent = 0;
    try {
      outer:
      for (const cls of classes) {
        step++;
        const id = cls;
        if (done[id]) continue;
        if (stopFlag) break outer;
        pick(classSel(), cls); await sleep(500);
        const batch = [];
        for (let i = 0; i < days.length; i++) {
          if (stopFlag) break outer;
          status.textContent = "[" + step + "/" + total + "] " + cls + "\nдата " + days[i] + " (" + (i + 1) + "/" + days.length + ")\nзнайдено оцінок: " + (sent + batch.length);
          const rows = await loadFor(cls, days[i]);
          rows.forEach(r => { if (WANT.has(catKey(r.cat)) || /^\?|^Н$/.test(r.cat)) batch.push([days[i], cls, r.cat, r.student, r.mark, r.style]); });
        }
        if (batch.length) { await post({ rows: batch }); sent += batch.length; }
        done[id] = 1;
        localStorage.setItem(runKey, JSON.stringify(done));
      }
      status.textContent = stopFlag ? "Зупинено. Записано оцінок: " + sent + ". Можна продовжити пізніше." : "Готово ✅ Записано оцінок: " + sent + ".";
      if (!stopFlag) localStorage.removeItem(runKey);
    } catch (e) {
      status.textContent = "Помилка: " + e.message + "\nПрогрес збережено — запустіть ще раз і оберіть «продовжити».";
    } finally {
      pick(classSel(), keep.cls); setDate(keep.date);
      stop.style.display = "none"; btn.disabled = false;
    }
  };
})();
