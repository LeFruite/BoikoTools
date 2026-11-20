/* Єдине меню переходів для сторінок вчителя: кнопка ☰ унизу праворуч, звідти можна потрапити в кабінет і на будь-який інструмент. */
(function () {
  var LINKS = [
    ['🏠', 'Кабінет вчителя', 'editor.html'],
    ['🪑', 'Розсадка', 'seating.html'],
    ['📁', 'Матеріали', 'materials.html'],
    ['📋', 'Журнал і оцінки', 'journal.html'],
    ['🕒', 'Відпрацювання', 'makeup.html'],
    ['🗓️', 'План занять', 'plan.html'],
    ['👥', 'Списки груп', 'students.html'],
    ['📡', 'Живий урок', 'live-teacher.html'],
    ['📖', 'Довідка', 'docs.html'],
    ['⚙️', 'Налаштування', 'start.html#settings']
  ];
  var here = (location.pathname.split('/').pop() || 'index.html');
  if (here === 'index.html' || here === '' || document.getElementById('nav-fab')) return;
  function mount() {
    var st = document.createElement('style');
    st.textContent = '#nav-fab{position:fixed;right:14px;bottom:14px;z-index:900;width:48px;height:48px;border-radius:50%;border:0;background:#1f7a5a;color:#fff;font-size:22px;box-shadow:0 3px 12px rgba(0,0,0,.3);cursor:pointer}' +
      '#nav-pan{position:fixed;right:14px;bottom:70px;z-index:900;background:#fff;color:#1d2a2e;border:1px solid #dcd9ce;border-radius:14px;box-shadow:0 8px 30px rgba(0,0,0,.25);padding:6px;min-width:220px;max-height:75vh;overflow:auto;display:none;font:15px system-ui,sans-serif}' +
      '#nav-pan a{display:flex;gap:10px;align-items:center;padding:11px 12px;color:inherit;text-decoration:none;border-radius:9px;min-height:44px}#nav-pan a:hover,#nav-pan a.cur{background:#e3f2ec}#nav-pan a.cur{font-weight:800}';
    document.head.appendChild(st); document.body.style.paddingBottom = 'max(' + (parseInt(getComputedStyle(document.body).paddingBottom) || 0) + 'px, 72px)';
    var b = document.createElement('button'); b.id = 'nav-fab'; b.type = 'button'; b.textContent = '☰'; b.setAttribute('aria-label', 'Меню переходів');
    var p = document.createElement('nav'); p.id = 'nav-pan'; p.setAttribute('aria-label', 'Меню переходів');
    p.innerHTML = LINKS.map(function (l) { var cur = l[2].split('#')[0] === here; return '<a href="' + l[2] + '"' + (cur ? ' class="cur" aria-current="page"' : '') + '><span aria-hidden="true">' + l[0] + '</span>' + l[1] + '</a>'; }).join('');
    b.onclick = function (e) { e.stopPropagation(); p.style.display = p.style.display === 'block' ? 'none' : 'block'; };
    document.addEventListener('click', function (e) { if (!p.contains(e.target)) p.style.display = 'none'; });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') p.style.display = 'none'; });
    document.body.appendChild(p); document.body.appendChild(b);
  }
  if (document.body) mount(); else document.addEventListener('DOMContentLoaded', mount);
})();
