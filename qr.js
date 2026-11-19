/* Мінімальний генератор QR-кодів (байтовий режим, корекція M, версії 1–10). Використання: QR.svg("текст", пікселів) → рядок SVG або "" якщо текст задовгий. */
(function (g) {
  var DATA = [0, 16, 28, 44, 64, 86, 108, 124, 154, 182, 216], TOTAL = [0, 26, 44, 70, 100, 134, 172, 196, 242, 292, 346],
      ECB = [0, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26], NB = [0, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5],
      ALIGN = [[], [], [6, 18], [6, 22], [6, 26], [6, 30], [6, 34], [6, 22, 38], [6, 24, 42], [6, 26, 46], [6, 28, 50]];
  var EXP = [], LOG = [];
  (function () { var x = 1; for (var i = 0; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 256) x ^= 0x11D; } for (i = 255; i < 512; i++) EXP[i] = EXP[i - 255]; })();
  function mul(a, b) { return a && b ? EXP[LOG[a] + LOG[b]] : 0; }
  function rs(data, n) {
    var gen = [1];
    for (var i = 0; i < n; i++) { var nx = new Array(gen.length + 1).fill(0); for (var j = 0; j < gen.length; j++) { nx[j] ^= gen[j]; nx[j + 1] ^= mul(gen[j], EXP[i]); } gen = nx; }
    var res = new Array(n).fill(0);
    data.forEach(function (b) { var f = b ^ res.shift(); res.push(0); for (var k = 0; k < n; k++) res[k] ^= mul(gen[k + 1], f); });
    return res;
  }
  function utf8(s) { return Array.prototype.slice.call(new TextEncoder().encode(s)); }
  function build(text) {
    var bytes = utf8(text), ver = 0;
    for (var v = 1; v <= 10; v++) if (Math.ceil((12 + (v < 10 ? 8 : 16) + bytes.length * 8) / 8) <= DATA[v] && v) { ver = v; break; }
    if (!ver) return null;
    var bits = []; function put(val, len) { for (var i = len - 1; i >= 0; i--) bits.push((val >>> i) & 1); }
    put(4, 4); put(bytes.length, ver < 10 ? 8 : 16); bytes.forEach(function (b) { put(b, 8); });
    var cap = DATA[ver] * 8; put(0, Math.min(4, cap - bits.length)); while (bits.length % 8) bits.push(0);
    var cw = []; for (var i = 0; i < bits.length; i += 8) { var b = 0; for (var j = 0; j < 8; j++) b = (b << 1) | bits[i + j]; cw.push(b); }
    for (var p = 0xEC; cw.length < DATA[ver]; p ^= 0xEC ^ 0x11) cw.push(p);
    var nb = NB[ver], shortN = nb - TOTAL[ver] % nb, blockLen = Math.floor(TOTAL[ver] / nb), blocks = [], k = 0;
    for (i = 0; i < nb; i++) { var dl = blockLen - ECB[ver] + (i < shortN ? 0 : 1), d = cw.slice(k, k + dl); k += dl; blocks.push({ d: d, e: rs(d, ECB[ver]) }); }
    var out = [], maxD = blockLen - ECB[ver] + 1;
    for (i = 0; i < maxD; i++) blocks.forEach(function (bl) { if (i < bl.d.length) out.push(bl.d[i]); });
    for (i = 0; i < ECB[ver]; i++) blocks.forEach(function (bl) { out.push(bl.e[i]); });
    var size = ver * 4 + 17, M = [], F = [];
    for (i = 0; i < size; i++) { M.push(new Array(size).fill(false)); F.push(new Array(size).fill(false)); }
    function fn(x, y, dark) { if (x >= 0 && y >= 0 && x < size && y < size) { M[y][x] = !!dark; F[y][x] = true; } }
    for (i = 0; i < size; i++) { fn(6, i, i % 2 === 0); fn(i, 6, i % 2 === 0); }
    function finder(x, y) { for (var dy = -4; dy <= 4; dy++) for (var dx = -4; dx <= 4; dx++) { var d = Math.max(Math.abs(dx), Math.abs(dy)); fn(x + dx, y + dy, d !== 2 && d !== 4); } }
    finder(3, 3); finder(size - 4, 3); finder(3, size - 4);
    var ap = ALIGN[ver], n = ap.length;
    for (i = 0; i < n; i++) for (j = 0; j < n; j++) if (!((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)))
      for (var dy = -2; dy <= 2; dy++) for (var dx = -2; dx <= 2; dx++) fn(ap[i] + dx, ap[j] + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
    function fmt(mask) {
      var data = mask, rem = data; for (var q = 0; q < 10; q++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
      var bt = ((data << 10) | rem) ^ 0x5412; function bit(q) { return ((bt >>> q) & 1) !== 0; }
      for (q = 0; q <= 5; q++) fn(8, q, bit(q)); fn(8, 7, bit(6)); fn(8, 8, bit(7)); fn(7, 8, bit(8)); for (q = 9; q < 15; q++) fn(14 - q, 8, bit(q));
      for (q = 0; q < 8; q++) fn(size - 1 - q, 8, bit(q)); for (q = 8; q < 15; q++) fn(8, size - 15 + q, bit(q)); fn(8, size - 8, true);
    }
    fmt(0);
    if (ver >= 7) { var r = ver; for (i = 0; i < 12; i++) r = (r << 1) ^ ((r >>> 11) * 0x1F25); var vb = (ver << 12) | r; for (i = 0; i < 18; i++) { var c = ((vb >>> i) & 1) !== 0, a = size - 11 + i % 3, bb = Math.floor(i / 3); fn(a, bb, c); fn(bb, a, c); } }
    var idx = 0;
    for (var right = size - 1; right >= 1; right -= 2) {
      if (right === 6) right = 5;
      for (var vert = 0; vert < size; vert++) for (j = 0; j < 2; j++) {
        var x = right - j, up = ((right + 1) & 2) === 0, y = up ? size - 1 - vert : vert;
        if (!F[y][x] && idx < out.length * 8) { M[y][x] = ((out[idx >>> 3] >>> (7 - (idx & 7))) & 1) !== 0; idx++; }
      }
    }
    var MASKS = [function (x, y) { return (x + y) % 2 === 0; }, function (x, y) { return y % 2 === 0; }, function (x, y) { return x % 3 === 0; }, function (x, y) { return (x + y) % 3 === 0; },
      function (x, y) { return (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; }, function (x, y) { return x * y % 2 + x * y % 3 === 0; }, function (x, y) { return (x * y % 2 + x * y % 3) % 2 === 0; }, function (x, y) { return ((x + y) % 2 + x * y % 3) % 2 === 0; }];
    function applyMask(mk) { for (var y = 0; y < size; y++) for (var x = 0; x < size; x++) if (!F[y][x] && MASKS[mk](x, y)) M[y][x] = !M[y][x]; }
    function penalty() {
      var pen = 0, y, x, run, i2;
      function line(get) { run = 1; for (i2 = 1; i2 < size; i2++) { if (get(i2) === get(i2 - 1)) { run++; if (run === 5) pen += 3; else if (run > 5) pen++; } else run = 1; } }
      for (y = 0; y < size; y++) line(function (t) { return M[y][t]; });
      for (x = 0; x < size; x++) line(function (t) { return M[t][x]; });
      for (y = 0; y < size - 1; y++) for (x = 0; x < size - 1; x++) if (M[y][x] === M[y][x + 1] && M[y][x] === M[y + 1][x] && M[y][x] === M[y + 1][x + 1]) pen += 3;
      var pat = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0];
      function fnd(get) { for (var s = 0; s + 11 <= size; s++) { var a = true, b = true; for (var t = 0; t < 11; t++) { var v = get(s + t) ? 1 : 0; if (v !== pat[t]) a = false; if (v !== pat[10 - t]) b = false; } if (a || b) pen += 40; } }
      for (y = 0; y < size; y++) fnd(function (t) { return M[y][t]; });
      for (x = 0; x < size; x++) fnd(function (t) { return M[t][x]; });
      var dark = 0; for (y = 0; y < size; y++) for (x = 0; x < size; x++) if (M[y][x]) dark++;
      pen += Math.floor(Math.abs(dark * 20 - size * size * 10) / (size * size)) * 10;
      return pen;
    }
    var best = 0, bp = 1e9;
    for (var mk = 0; mk < 8; mk++) { applyMask(mk); fmt(mk); var pn = penalty(); if (pn < bp) { bp = pn; best = mk; } applyMask(mk); }
    applyMask(best); fmt(best);
    return M;
  }
  g.QR = {
    matrix: build,
    svg: function (text, px) {
      var m = build(text); if (!m) return ''; var n = m.length, q = 4, d = '';
      for (var y = 0; y < n; y++) for (var x = 0; x < n; x++) if (m[y][x]) d += 'M' + (x + q) + ' ' + (y + q) + 'h1v1h-1z';
      return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + (n + 2 * q) + ' ' + (n + 2 * q) + '" width="' + (px || 220) + '" height="' + (px || 220) + '" shape-rendering="crispEdges" role="img" aria-label="QR-код"><rect width="100%" height="100%" fill="#fff"/><path d="' + d + '" fill="#000"/></svg>';
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
