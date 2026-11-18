// Адреса вашого Apps Script (Deploy → Manage deployments → Web app URL, закінчується на /exec).
// Вносити її сюди НЕ ОБОВ'ЯЗКОВО: відкрийте start.html на сайті — він перевірить адресу й запам'ятає її.
// Посилання для учнів, які створює кабінет вчителя, самі містять потрібну адресу.
(function () {
  var own = "";   // лише для особистого сайту одного вчителя; для спільного сайту ЗАЛИШАЙТЕ порожнім
  var host = "";  // домен особистого сайту, наприклад "ім'я.github.io" (необов'язково)
  var url = "", fromStorage = false;
  try { if (own && host && location.hostname.toLowerCase() === host.toLowerCase()) url = own; } catch (e) {}
  if (!url) {
    try {
      var q = new URLSearchParams(location.search).get("api");
      if (q) localStorage.setItem("script_api", q);
      var v = localStorage.getItem("script_api") || "";
      if (v) { url = /^https?:\/\//.test(v) ? v : "https://script.google.com/macros/s/" + v + "/exec"; fromStorage = true; window.SCRIPT_API_ID = v; }
    } catch (e) {}
  }
  window.SCRIPT_URL = url; window.SCRIPT_FROM_STORAGE = fromStorage;
})();
