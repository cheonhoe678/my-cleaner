/* 그래프 — 외부 부품 없이 SVG 로 직접 그린다. 데이터는 textContent 로만 넣는다. */
(function (root) {
  'use strict';
  var NS = 'http://www.w3.org/2000/svg';

  function s(tag, attrs, txt) {
    var e = document.createElementNS(NS, tag);
    if (attrs) Object.keys(attrs).forEach(function (k) { e.setAttribute(k, attrs[k]); });
    if (txt != null) e.textContent = txt;
    return e;
  }
  function niceMax(v) {
    if (v <= 0) return 1;
    var p = Math.pow(10, Math.floor(Math.log(v) / Math.LN10)), f = v / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p;
  }
  function short(t, n) { return t.length > n ? t.slice(0, n - 1) + '…' : t; }
  function fmt(v) { return Math.round(v).toLocaleString('ko-KR'); }

  function frame(host, title, w, h, note) {
    while (host.firstChild) host.removeChild(host.firstChild);
    var t = document.createElement('h4'); t.className = 'ctitle'; t.textContent = title; host.appendChild(t);
    var svg = s('svg', { viewBox: '0 0 ' + w + ' ' + h, role: 'img', 'aria-label': title + (note ? ' — ' + note : ''), 'class': 'chart' });
    host.appendChild(svg);
    return svg;
  }
  function empty(host, title, msg) {
    while (host.firstChild) host.removeChild(host.firstChild);
    var t = document.createElement('h4'); t.className = 'ctitle'; t.textContent = title; host.appendChild(t);
    var p = document.createElement('p'); p.className = 'empty'; p.textContent = msg; host.appendChild(p);
  }

  /* items: [{name, count}] — 세로 막대 */
  function bars(host, title, items, opt) {
    opt = opt || {};
    var total = items.reduce(function (a, x) { return a + x.count; }, 0);
    if (!items.length || !total) return empty(host, title, '그릴 값이 없습니다.');
    var W = 560, H = 250, m = { l: 44, r: 10, t: 10, b: 28 };
    var svg = frame(host, title, W, H, opt.note);
    var mx = niceMax(Math.max.apply(null, items.map(function (x) { return x.count; })));
    var band = (W - m.l - m.r) / items.length, bw = band * 0.72;
    [0, 0.5, 1].forEach(function (f) {
      var y = H - m.b - (H - m.t - m.b) * f;
      svg.appendChild(s('line', { x1: m.l, x2: W - m.r, y1: y, y2: y, 'class': 'grid' }));
      svg.appendChild(s('text', { x: m.l - 6, y: y + 4, 'text-anchor': 'end' }, fmt(mx * f)));
    });
    items.forEach(function (x, i) {
      var bh = (H - m.t - m.b) * x.count / mx, bx = m.l + band * i + (band - bw) / 2, by = H - m.b - bh;
      var g = s('g');
      g.appendChild(s('rect', { x: bx, y: by, width: bw, height: Math.max(bh, 0), 'class': 'bar' }));
      g.appendChild(s('title', null, x.name + ': ' + fmt(x.count) + '건'));
      svg.appendChild(g);
      if (!opt.skip || i % opt.skip === 0) {
        svg.appendChild(s('text', { x: bx + bw / 2, y: H - m.b + 15, 'text-anchor': 'middle' }, short(x.name, 4)));
      }
    });
  }

  /* 가로 막대 — 큰 순서 그대로 위에서 아래로 */
  function hbars(host, title, items, opt) {
    opt = opt || {};
    if (!items.length) return empty(host, title, '그릴 값이 없습니다.');
    var rowH = 24, W = 560, m = { l: 112, r: 54, t: 6, b: 6 };
    var H = items.length * rowH + m.t + m.b;
    var svg = frame(host, title, W, H, opt.note);
    var mx = Math.max.apply(null, items.map(function (x) { return x.count; })) || 1;
    items.forEach(function (x, i) {
      var y = m.t + i * rowH, bw = (W - m.l - m.r) * x.count / mx;
      var g = s('g');
      g.appendChild(s('text', { x: m.l - 8, y: y + rowH / 2 + 4, 'text-anchor': 'end' }, short(x.name, 9)));
      g.appendChild(s('rect', { x: m.l, y: y + 3, width: Math.max(bw, 0), height: rowH - 8, 'class': 'bar' }));
      g.appendChild(s('text', { x: m.l + bw + 6, y: y + rowH / 2 + 4 }, fmt(x.count)));
      g.appendChild(s('title', null, x.name + ': ' + fmt(x.count) + '건' + (x.extra ? ' · ' + x.extra : '')));
      svg.appendChild(g);
    });
  }

  /* 선 그래프 */
  function line(host, title, items, opt) {
    opt = opt || {};
    var total = items.reduce(function (a, x) { return a + x.count; }, 0);
    if (items.length < 2 || !total) return empty(host, title, '그릴 값이 없습니다.');
    var W = opt.W || 560, H = opt.W ? 260 : 230, m = { l: 44, r: 12, t: 12, b: 28 };
    var svg = frame(host, title, W, H, opt.note);
    var mx = niceMax(Math.max.apply(null, items.map(function (x) { return x.count; })));
    [0, 0.5, 1].forEach(function (f) {
      var y = H - m.b - (H - m.t - m.b) * f;
      svg.appendChild(s('line', { x1: m.l, x2: W - m.r, y1: y, y2: y, 'class': 'grid' }));
      svg.appendChild(s('text', { x: m.l - 6, y: y + 4, 'text-anchor': 'end' }, fmt(mx * f)));
    });
    var px = function (i) { return m.l + (W - m.l - m.r) * i / (items.length - 1); };
    var py = function (v) { return H - m.b - (H - m.t - m.b) * v / mx; };
    svg.appendChild(s('polyline', { points: items.map(function (x, i) { return px(i) + ',' + py(x.count); }).join(' '), 'class': 'lineP' }));
    items.forEach(function (x, i) {
      var c = s('circle', { cx: px(i), cy: py(x.count), r: 2.5, 'class': 'dot' });
      c.appendChild(s('title', null, x.name + ': ' + fmt(x.count) + '건'));
      svg.appendChild(c);
    });
    [0, Math.floor((items.length - 1) / 2), items.length - 1].forEach(function (i) {
      svg.appendChild(s('text', { x: px(i), y: H - m.b + 15, 'text-anchor': i === 0 ? 'start' : i === items.length - 1 ? 'end' : 'middle' }, items[i].name.slice(5)));
    });
  }

  /* 열지도: 행 = 요일, 열 = 시 */
  function heatmap(host, title, rowNames, matrix, opt) {
    opt = opt || {};
    var mx = 0;
    matrix.forEach(function (r) { r.forEach(function (v) { if (v > mx) mx = v; }); });
    if (!mx) return empty(host, title, '그릴 값이 없습니다.');
    var W = opt.W || 560, m = { l: 30, r: 8, t: 8, b: 22 }, cw = (W - m.l - m.r) / 24, ch = opt.W ? 26 : 22;
    var H = m.t + ch * rowNames.length + m.b;
    var svg = frame(host, title, W, H, opt.note);
    rowNames.forEach(function (name, r) {
      svg.appendChild(s('text', { x: m.l - 6, y: m.t + r * ch + ch / 2 + 4, 'text-anchor': 'end' }, name));
      matrix[r].forEach(function (v, c) {
        var cell = s('rect', { x: m.l + c * cw + 1, y: m.t + r * ch + 1, width: cw - 2, height: ch - 2, rx: 2,
          'class': 'heat', 'fill-opacity': v ? (0.08 + 0.92 * v / mx).toFixed(3) : '0.04' });
        cell.appendChild(s('title', null, name + ' ' + (c < 10 ? '0' : '') + c + '시: ' + fmt(v) + '건'));
        svg.appendChild(cell);
      });
    });
    for (var c = 0; c < 24; c += 3) {
      svg.appendChild(s('text', { x: m.l + c * cw + cw / 2, y: H - 6, 'text-anchor': 'middle' }, (c < 10 ? '0' : '') + c));
    }
  }

  root.Charts = { bars: bars, hbars: hbars, line: line, heatmap: heatmap };
})(typeof window !== 'undefined' ? window : this);
