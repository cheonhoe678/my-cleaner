/* 화면 — 파일은 이 브라우저 안에서만 읽고 계산한다. 서버로 보내는 코드는 없다.
   통신은 같은 사이트의 rules/<파일 이름>.json 을 읽는 fetch 한 군데뿐이다. */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var state = { fileName: '', baseName: '', table: null, rules: null, result: null, twice: null };

  function el(tag, text, cls) {
    var e = document.createElement(tag);
    if (text != null) e.textContent = text;
    if (cls) e.className = cls;
    return e;
  }
  function show(id, on) { $(id).hidden = !on; }
  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }
  function fmt(v) {
    if (v == null) return '';
    if (Cleaner.isSerial(v)) return String(v.serial);
    return String(v);
  }

  /* ---------- 엑셀 · CSV 읽기 ---------- */
  function sheetToTable(ws) {
    if (!ws || !ws['!ref']) return { header: [], rows: [] };
    var r = XLSX.utils.decode_range(ws['!ref']);
    function cellValue(c) {
      if (!c || c.t === 'z') return null;
      if (c.t === 'n') {
        if (c.z && XLSX.SSF.is_date(c.z)) return { serial: c.v };
        return c.v;
      }
      if (c.t === 's') return c.v;
      if (c.t === 'b') return c.v ? 'TRUE' : 'FALSE';
      return c.w != null ? c.w : null;
    }
    var header = [], rows = [], x, y;
    for (x = r.s.c; x <= r.e.c; x++) {
      var h = ws[XLSX.utils.encode_cell({ r: r.s.r, c: x })];
      header.push(h && h.v != null ? String(h.v) : '');
    }
    for (y = r.s.r + 1; y <= r.e.r; y++) {
      var row = [];
      for (x = r.s.c; x <= r.e.c; x++) row.push(cellValue(ws[XLSX.utils.encode_cell({ r: y, c: x })]));
      rows.push(row);
    }
    return { header: header, rows: rows };
  }

  function readTable(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      var isCsv = /\.csv$/i.test(file.name);
      fr.onerror = function () { reject(new Error('파일을 읽지 못했습니다')); };
      fr.onload = function () {
        try {
          var wb = isCsv
            ? XLSX.read(fr.result, { type: 'string', raw: true })
            : XLSX.read(fr.result, { type: 'array', cellNF: true, cellDates: false });
          resolve({ table: sheetToTable(wb.Sheets[wb.SheetNames[0]]), sheets: wb.SheetNames.length });
        } catch (e) { reject(e); }
      };
      if (isCsv) fr.readAsText(file, 'utf-8'); else fr.readAsArrayBuffer(file);
    });
  }

  /* ---------- 1. 파일 넣기 ---------- */
  function onFile(file) {
    if (!file) return;
    state.fileName = file.name;
    state.baseName = file.name.replace(/\.[^.]+$/, '');
    state.result = null;
    show('s3', false); show('s4', false); show('s5', false); show('s6', false);
    readTable(file).then(function (res) {
      state.table = res.table;
      var info = state.fileName + ' — 줄 ' + res.table.rows.length + ' · 칸 ' + res.table.header.length +
        ' (' + res.table.header.join(' · ') + ')' + (res.sheets > 1 ? ' · 시트가 ' + res.sheets + '개라 첫 시트만 읽었습니다' : '');
      $('fileInfo').textContent = info; show('fileInfo', true);
      show('s2', true);
      loadRules();
    }).catch(function (e) {
      $('fileInfo').textContent = '읽지 못했습니다: ' + e.message; show('fileInfo', true);
    });
  }

  /* ---------- 2. 처방 읽기 ---------- */
  function loadRules() {
    show('noRules', false); show('rulesBox', false);
    fetch('rules/' + encodeURIComponent(state.baseName) + '.json', { cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error('rules/' + state.baseName + '.json 이(가) 없습니다');
      return r.json();
    }).then(applyRules).catch(function (e) {
      $('noRulesWhy').textContent = e.message + '. 처방 파일을 직접 골라 올릴 수 있습니다.';
      show('noRules', true);
    });
  }

  function applyRules(rules) {
    if (!rules || !Array.isArray(rules.rules)) {
      $('noRulesWhy').textContent = '처방 파일의 모양이 맞지 않습니다 (rules 목록이 없습니다).';
      show('noRules', true); show('rulesBox', false); return;
    }
    state.rules = rules;
    show('noRules', false);
    $('rulesMeta').textContent = '처방 ' + rules.rules.length + '개 · 정한 사람 ' + (rules.owner || '(적혀 있지 않음)') +
      (rules.file ? ' · 대상 ' + rules.file : '');
    var t = $('rulesTable'); clear(t);
    var head = el('tr');
    ['번호', '칸', '종류', '처방', '이유'].forEach(function (h) { head.appendChild(el('th', h)); });
    t.appendChild(head);
    rules.rules.forEach(function (r) {
      var tr = el('tr');
      tr.appendChild(el('td', r.id));
      tr.appendChild(el('td', Array.isArray(r.column) ? r.column.join(' · ') : r.column));
      tr.appendChild(el('td', r.action));
      tr.appendChild(el('td', r.do));
      tr.appendChild(el('td', r.why));
      t.appendChild(tr);
    });
    show('rulesBox', true);
  }

  /* ---------- 3. 정리 ---------- */
  function simpleTable(id, head, rows, emptyText) {
    var host = $(id).parentNode;
    clear(host);
    if (!rows.length) { host.appendChild(el('p', emptyText, 'empty')); return; }
    var t = el('table'); t.id = id;
    var hr = el('tr'); head.forEach(function (h) { hr.appendChild(el('th', h)); }); t.appendChild(hr);
    rows.forEach(function (r) {
      var tr = el('tr'); r.forEach(function (c) { tr.appendChild(el('td', c)); }); t.appendChild(tr);
    });
    host.appendChild(t);
  }

  function stat(k, v, s, cls) {
    var d = el('div', null, 'stat');
    d.appendChild(el('div', k, 'k'));
    d.appendChild(el('div', v, 'v' + (cls ? ' ' + cls : '')));
    if (s) d.appendChild(el('div', s, 's'));
    return d;
  }

  function run() {
    var res = Cleaner.cleanTable(state.table, state.rules);
    var twice = Cleaner.checkTwice(res, state.rules);
    state.result = res; state.twice = twice;
    var s = res.stats;
    var box = $('stats'); clear(box);
    box.appendChild(stat('행 전후', s.rowsBefore + ' → ' + s.rowsAfter, '줄 수'));
    box.appendChild(stat('지운 행', String(s.removed), '완전히 같은 줄'));
    box.appendChild(stat('고친 값', s.changedCells + '칸', s.changedRows + '줄'));
    box.appendChild(stat('표시한 행', s.markedRows + '줄', '표시 ' + s.markCount + '개'));
    box.appendChild(stat('두 번 돌려도 같은 결과', twice.same ? '예' : '아니오',
      twice.same ? '정리본에 다시 돌려도 바뀌는 것이 없음' : '다시 돌리니 고침 ' + twice.fixes + ' · 지움 ' + twice.removed,
      twice.same ? 'ok' : 'bad'));
    var w = $('warnings'); clear(w);
    if (res.warnings.length) {
      res.warnings.forEach(function (m) { w.appendChild(el('p', m)); }); show('warnings', true);
    } else show('warnings', false);

    var key = state.rules.key_column || '';
    simpleTable('removedTable', ['원본 행', key || '키', '같은 줄이 있는 행', '이유'],
      res.removed.map(function (r) { return [String(r.src), fmt(r.key), String(r.keptSrc), r.why]; }), '지운 행이 없습니다.');
    simpleTable('fixTable', ['원본 행', key || '키', '칸', '전', '후', '처방'],
      res.fixes.slice(0, 20).map(function (e) { return [String(e.src), fmt(e.key), e.col, e.before, e.after, e.rule]; }),
      '고친 값이 없습니다.');
    simpleTable('markTable', ['원본 행', key || '키', '칸', '까닭', '처방', '이유'],
      res.marks.slice(0, 20).map(function (e) { return [String(e.src), fmt(e.key), e.col, e.after, e.rule, e.why]; }),
      '표시한 행이 없습니다.');
    show('s3', true); show('s4', true); show('s5', true);
    setupAnalysis();
    $('compareOut').textContent = '';
  }

  /* ---------- 4. 내려받기 ---------- */
  function download(wb, name) {
    var data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    var blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
  }
  function cellOut(v) { return v == null ? '' : (Cleaner.isSerial(v) ? v.serial : v); }

  function downloadClean() {
    var r = state.result; if (!r) return;
    var aoa = [r.header.concat(['_표시'])];
    r.rows.forEach(function (row, i) { aoa.push(row.map(cellOut).concat([r.markText[i]])); });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '정리본');
    download(wb, state.baseName + '_정리본.xlsx');
  }
  function downloadLog() {
    var r = state.result; if (!r) return;
    var key = state.rules.key_column || '키';
    var aoa = [['구분', '처방', '원본 행', key, '칸', '전', '후 · 까닭', '이유']];
    r.fixes.forEach(function (e) { aoa.push(['고침', e.rule, e.src, cellOut(e.key), e.col, e.before, e.after, e.why]); });
    r.marks.forEach(function (e) { aoa.push(['표시', e.rule, e.src, cellOut(e.key), e.col, '', e.after, e.why]); });
    r.removed.forEach(function (e) {
      aoa.push(['삭제', e.rule, e.src, cellOut(e.key), '(줄 전체)', '', '같은 줄이 ' + e.keptSrc + '행에 있어 지움', e.why]);
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '바꾼 기록');
    download(wb, state.baseName + '_바꾼기록.xlsx');
  }

  /* ---------- 5. 정답과 대조 ---------- */
  function onAnswer(file) {
    if (!file || !state.result) return;
    var out = $('compareOut'); clear(out);
    readTable(file).then(function (res) {
      var key = state.rules.key_column;
      var c = Cleaner.compareWithAnswer(state.result, res.table, key);
      if (c.error) { out.appendChild(el('p', c.error, 'warn')); return; }
      var box = el('div', null, 'stats');
      box.appendChild(stat('정답 행', String(c.answerRows)));
      box.appendChild(stat('일치', c.matched + ' / ' + c.answerRows, '모든 칸이 같은 줄'));
      box.appendChild(stat('어긋난 칸', String(c.cells.length)));
      box.appendChild(stat('정리본에만 있는 줄', String(c.onlyMine.length), c.onlyMine.slice(0, 6).join(' · ')));
      box.appendChild(stat('정답에만 있는 줄', String(c.onlyAnswer.length), c.onlyAnswer.slice(0, 6).join(' · ')));
      out.appendChild(box);
      out.appendChild(el('h3', '어긋난 칸 (앞의 30개)'));
      var wrap = el('div', null, 'tablewrap');
      if (!c.cells.length) wrap.appendChild(el('p', '어긋난 칸이 없습니다.', 'empty'));
      else {
        var t = el('table'), hr = el('tr');
        [key, '칸', '정리본', '정답'].forEach(function (h) { hr.appendChild(el('th', h)); }); t.appendChild(hr);
        c.cells.slice(0, 30).forEach(function (x) {
          var tr = el('tr');
          [x.key, x.col, x.mine, x.ans].forEach(function (v) { tr.appendChild(el('td', v)); }); t.appendChild(tr);
        });
        wrap.appendChild(t);
      }
      out.appendChild(wrap);
    }).catch(function (e) { out.appendChild(el('p', '정답 파일을 읽지 못했습니다: ' + e.message, 'warn')); });
  }

  /* ---------- 6. 접속 분석 (그래프) ---------- */
  function rawTable() { return { header: state.table.header, rows: state.table.rows }; }

  function setupAnalysis() {
    var cfg = state.rules && state.rules.analysis;
    if (!cfg) { show('s6', false); return; }
    var sel = $('anaGroup'); clear(sel);
    var choices = (cfg.group_choices || [cfg.group_column]).filter(function (c) { return state.table.header.indexOf(c) >= 0; });
    if (!choices.length) choices = [cfg.group_column];
    choices.forEach(function (c) {
      var o = el('option', c); o.value = c; if (c === cfg.group_column) o.selected = true; sel.appendChild(o);
    });
    show('s6', true);
    renderAnalysis();
  }

  function renderAnalysis() {
    var cfg = state.rules && state.rules.analysis;
    if (!cfg || !state.result) return;
    var group = $('anaGroup').value || cfg.group_column;
    var rawA = Analyzer.analyze(rawTable(), cfg, group);
    var cleanA = Analyzer.analyze({ header: state.result.header, rows: state.result.rows }, cfg, group);
    var which = document.querySelector('input[name="anaTable"]:checked').value;
    var a = which === 'raw' ? rawA : cleanA;
    var nf = function (n) { return n.toLocaleString('ko-KR'); };

    var box = $('anaStats'); clear(box);
    box.appendChild(stat('그래프에 쓴 줄', nf(a.used) + ' / ' + nf(a.total)));
    box.appendChild(stat('서로 다른 유저', a.userKinds == null ? '—' : a.userKinds + '명'));
    box.appendChild(stat('평균 접속시간', a.avgMin == null ? '—' : a.avgMin.toFixed(1) + '분', '튀는 값 · 못 읽은 값은 뺌'));
    box.appendChild(stat('접속이 가장 많은 시간대', a.peakHour == null ? '—' : (a.peakHour < 10 ? '0' : '') + a.peakHour + '시'));
    box.appendChild(stat('가장 많은 ' + a.groupName, a.topGroup ? a.topGroup.name : '—', a.topGroup ? nf(a.topGroup.count) + '건' : ''));

    $('anaNote').textContent = a.groupName + ' 이름: 원본 ' + rawA.groupKinds + '가지 → 정리본 ' + cleanA.groupKinds +
      '가지 · 그래프에 쓴 줄: 원본 ' + nf(rawA.used) + ' → 정리본 ' + nf(cleanA.used);

    var ex = $('anaExcluded'); clear(ex);
    var keys = Object.keys(a.excluded);
    if (keys.length) {
      ex.appendChild(el('p', '그래프에서 뺀 줄 — ' + keys.map(function (k) { return k + ' ' + nf(a.excluded[k]) + '줄'; }).join(' · ')));
      show('anaExcluded', true);
    } else show('anaExcluded', false);

    var hours = a.hourly.map(function (c, h) { return { name: (h < 10 ? '0' : '') + h, count: c }; });
    Charts.bars($('chHour'), '시간대별 접속 건수 (시)', hours, { note: '접속 시각의 시 기준' });
    Charts.hbars($('chGroup'), a.groupName + '별 접속 건수' + (a.groups.length > 12 ? ' (상위 12개)' : ''), a.groups.slice(0, 12).map(function (g) {
      return { name: g.name, count: g.count, extra: g.avgMin == null ? '' : '평균 ' + g.avgMin.toFixed(1) + '분' };
    }), { note: '큰 순서' });
    Charts.bars($('chWeek'), '요일별 접속 건수', Analyzer.DAYS.map(function (d, i) { return { name: d, count: a.weekday[i] }; }));
    Charts.hbars($('chUser'), '접속이 많은 유저 10명', a.topUsers);
    Charts.line($('chDaily'), '일자별 접속 건수', a.daily, { W: 900 });
    Charts.heatmap($('chHeat'), '요일 × 시간대 접속 건수 (진할수록 많음)', Analyzer.DAYS, a.heat, { W: 900 });
  }

  /* ---------- 연결 ---------- */
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); });
  });
  drop.addEventListener('drop', function (e) { onFile(e.dataTransfer.files[0]); });
  $('file').addEventListener('change', function (e) { onFile(e.target.files[0]); });
  $('rulesFile').addEventListener('change', function (e) {
    var f = e.target.files[0]; if (!f) return;
    var fr = new FileReader();
    fr.onload = function () {
      try { applyRules(JSON.parse(fr.result)); }
      catch (err) { $('noRulesWhy').textContent = '처방 파일을 읽지 못했습니다: ' + err.message; show('noRules', true); }
    };
    fr.readAsText(f, 'utf-8');
  });
  $('runBtn').addEventListener('click', run);
  $('dlClean').addEventListener('click', downloadClean);
  $('dlLog').addEventListener('click', downloadLog);
  $('answerFile').addEventListener('change', function (e) { onAnswer(e.target.files[0]); });
  $('anaGroup').addEventListener('change', renderAnalysis);
  Array.prototype.forEach.call(document.querySelectorAll('input[name="anaTable"]'), function (r) { r.addEventListener('change', renderAnalysis); });
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) { e.preventDefault(); });

  // 시험용 손잡이 — 같은 계산을 화면 없이 부를 때만 쓴다
  window.__cleanerApp = { sheetToTable: sheetToTable, state: state };
})();
