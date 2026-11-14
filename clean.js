/* 정리 계산 — 처방 값(칸 이름 · 바꿀 말 · 경계)은 전부 rules/*.json 에서 읽는다.
   여기에는 «동작의 이름(type)»만 있다. 네트워크 코드는 없다. */
(function (root) {
  'use strict';

  var pad2 = function (n) { return (n < 10 ? '0' : '') + n; };

  function isBlank(v) { return v == null || (typeof v === 'string' && v.trim() === ''); }
  function isSerial(v) { return v !== null && typeof v === 'object' && typeof v.serial === 'number'; }

  // 엑셀 시리얼 → 글자 (시간대 계산 없이 정수로만 계산한다)
  function serialToTime(n) {
    var secs = Math.round((n - Math.floor(n)) * 86400);
    if (secs >= 86400) secs -= 86400;
    return pad2(Math.floor(secs / 3600)) + ':' + pad2(Math.floor(secs % 3600 / 60)) + ':' + pad2(secs % 60);
  }
  function serialToDate(n) {
    var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  function validYMD(y, m, d) {
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
  }
  function asList(c) { return Array.isArray(c) ? c : [c]; }

  function show(v) {
    if (v == null) return '(빈칸)';
    if (isSerial(v)) return String(v.serial);
    if (typeof v === 'string' && v.trim() === '') return '(공백)';
    return String(v);
  }

  /* table: { header: [칸 이름…], rows: [[값…]…] }
     값: null | 글자 | 숫자 | {serial: 엑셀 날짜·시간 숫자}
     rules: rules/*.json 을 읽은 것 */
  function cleanTable(table, rules) {
    var header = table.header.slice();
    var rows = table.rows.map(function (r, i) { return { src: i + 2, v: r.slice(), marks: [] }; });
    var log = [];       // {src, key, col, before, after, rule, kind}
    var warnings = [];
    var removed = [];
    var keyCol = rules.key_column ? header.indexOf(rules.key_column) : -1;

    function colIdx(name) {
      var i = header.indexOf(name);
      if (i < 0) warnings.push('칸 «' + name + '» 이(가) 이 파일에 없다');
      return i;
    }
    function keyOf(row) { return keyCol >= 0 ? row.v[keyCol] : null; }
    function change(row, ci, newV, rule, beforeText, afterText) {
      log.push({ src: row.src, key: keyOf(row), col: header[ci], before: beforeText, after: afterText,
        rule: rule.id, kind: '고침', why: rule.why });
      row.v[ci] = newV;
    }
    function mark(row, rule, cols) {
      row.marks.push({ id: rule.id, label: rule.label || rule.do, why: rule.why, cols: cols });
    }
    function parseTime(v, noSecRe) {
      if (typeof v !== 'string') return null;
      var m = /^\s*(\d{1,2}):(\d{2}):(\d{2})\s*$/.exec(v);
      if (m) return { sec: +m[1] * 3600 + +m[2] * 60 + +m[3], noSec: 0 };
      if (noSecRe && noSecRe.test(v)) {
        var m2 = /(\d{1,2})시 (\d{1,2})분/.exec(v);
        return { sec: +m2[1] * 3600 + +m2[2] * 60, noSec: 1 };
      }
      return null;
    }

    var ops = {
      strip_lower: function (rule) {
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci];
            if (typeof v !== 'string') return;
            var nv = v.trim().toLowerCase();
            if (nv !== v) change(row, ci, nv, rule, show(v), nv);
          });
        });
      },
      map_values: function (rule) {
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci];
            if (typeof v !== 'string') return;
            var t = v.trim();
            var nv = Object.prototype.hasOwnProperty.call(rule.map, t) ? rule.map[t] : t;
            if (nv !== v) change(row, ci, nv, rule, show(v), nv);
          });
        });
      },
      time_format: function (rule) {
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci], nv;
            if (isSerial(v)) { nv = serialToTime(v.serial); change(row, ci, nv, rule, nv + ' (엑셀 시간값)', nv); return; }
            if (typeof v !== 'string') return;
            var m = /^\s*(\d{1,2}):(\d{2}):(\d{2})\s*$/.exec(v);
            if (!m) return;
            nv = pad2(+m[1]) + ':' + m[2] + ':' + m[3];
            if (nv !== v) change(row, ci, nv, rule, show(v), nv);
          });
        });
      },
      date_format: function (rule) {
        var res = (rule.patterns || []).map(function (p) { return new RegExp(p); });
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci], nv;
            if (isSerial(v)) { nv = serialToDate(v.serial); change(row, ci, nv, rule, nv + ' (엑셀 날짜값)', nv); return; }
            if (typeof v !== 'string') return;
            for (var k = 0; k < res.length; k++) {
              var m = res[k].exec(v.trim());
              if (m && validYMD(+m[1], +m[2], +m[3])) {
                nv = m[1] + '-' + m[2] + '-' + m[3];
                if (nv !== v) change(row, ci, nv, rule, show(v), nv);
                return;
              }
            }
          });
        });
      },
      strip_unit: function (rule) {
        var re = new RegExp('^\\s*(-?\\d+(?:\\.\\d+)?)\\s*' + rule.unit.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*$');
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci];
            if (typeof v !== 'string') return;
            var m = re.exec(v);
            if (m) change(row, ci, Number(m[1]), rule, show(v), m[1]);
          });
        });
      },
      mark_pattern: function (rule) {
        var re = new RegExp(rule.pattern);
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci];
            if (typeof v === 'string' && re.test(v)) mark(row, rule, [cn]);
          });
        });
      },
      mark_blank: function (rule) {
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) { if (isBlank(row.v[ci])) mark(row, rule, [cn]); });
        });
      },
      mark_range: function (rule) {
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci];
            if (typeof v !== 'number') return;
            if ((rule.mark_if_at_most != null && v <= rule.mark_if_at_most) ||
                (rule.mark_if_at_least != null && v >= rule.mark_if_at_least)) mark(row, rule, [cn]);
          });
        });
      },
      mark_date_range: function (rule) {
        asList(rule.column).forEach(function (cn) {
          var ci = colIdx(cn); if (ci < 0) return;
          rows.forEach(function (row) {
            var v = row.v[ci];
            if (typeof v !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(v)) return;
            if (v < rule.from || v > rule.to) mark(row, rule, [cn]);
          });
        });
      },
      mark_duration_mismatch: function (rule) {
        var li = colIdx(rule.login), lo = colIdx(rule.logout), du = colIdx(rule.duration);
        if (li < 0 || lo < 0 || du < 0) return;
        var noSecRe = rule.no_seconds_pattern ? new RegExp(rule.no_seconds_pattern) : null;
        var tolPer = rule.tolerance_minutes_per_time_without_seconds || 0;
        rows.forEach(function (row) {
          var a = parseTime(row.v[li], noSecRe), b = parseTime(row.v[lo], noSecRe), mm = row.v[du];
          if (!a || !b || typeof mm !== 'number') return;
          var tol = tolPer * (a.noSec + b.noSec);
          var d = (b.sec - a.sec) / 60;
          if (Math.abs(d - mm) > tol && Math.abs(Math.round(d) - mm) > tol) mark(row, rule, [rule.login, rule.logout, rule.duration]);
        });
      },
      dedupe: function (rule) {
        var seen = {}, keep = [];
        rows.forEach(function (row) {
          var k = JSON.stringify(row.v.map(function (v) { return isBlank(v) ? null : v; }));
          if (Object.prototype.hasOwnProperty.call(seen, k)) {
            removed.push({ src: row.src, key: keyOf(row), keptSrc: seen[k], rule: rule.id, why: rule.why });
          } else { seen[k] = row.src; keep.push(row); }
        });
        rows = keep;
      }
    };

    rules.rules.forEach(function (rule) {
      var f = ops[rule.type];
      if (!f) { warnings.push('처방 ' + rule.id + ' 의 동작 «' + rule.type + '» 을(를) 이 정리기는 모른다'); return; }
      f(rule);
    });

    var alive = {};
    rows.forEach(function (r) { alive[r.src] = true; });
    var keptLog = log.filter(function (e) { return alive[e.src]; });
    var markLog = [];
    rows.forEach(function (row) {
      row.marks.forEach(function (m) {
        markLog.push({ src: row.src, key: keyOf(row), col: m.cols.join(' · '), before: '', after: m.label,
          rule: m.id, kind: '표시', why: m.why });
      });
    });

    var changedRowSet = {};
    keptLog.forEach(function (e) { changedRowSet[e.src] = true; });

    return {
      header: header,
      rows: rows.map(function (r) { return r.v; }),
      srcRows: rows.map(function (r) { return r.src; }),
      markText: rows.map(function (r) {
        return r.marks.map(function (m) { return m.id + ' ' + m.label; }).join(' / ');
      }),
      markCount: rows.map(function (r) { return r.marks.length; }),
      log: keptLog.concat(markLog),
      fixes: keptLog,
      marks: markLog,
      removed: removed,
      warnings: warnings,
      stats: {
        rowsBefore: table.rows.length,
        rowsAfter: rows.length,
        removed: removed.length,
        changedCells: keptLog.length,
        changedRows: Object.keys(changedRowSet).length,
        markedRows: rows.filter(function (r) { return r.marks.length > 0; }).length,
        markCount: markLog.length
      }
    };
  }

  // 정리본에 한 번 더 돌려서 바뀌는 것이 없는지 본다
  function checkTwice(first, rules) {
    var second = cleanTable({ header: first.header, rows: first.rows }, rules);
    var same = JSON.stringify(first.rows) === JSON.stringify(second.rows) &&
      JSON.stringify(first.markText) === JSON.stringify(second.markText) &&
      second.fixes.length === 0 && second.removed.length === 0;
    return { same: same, fixes: second.fixes.length, removed: second.removed.length };
  }

  /* 정답과 대조 — 정리본의 줄을 key 칸으로 정답 줄과 맞춘다 */
  function compareWithAnswer(first, answer, keyName) {
    var norm = function (v) {
      if (v == null) return '';
      if (typeof v === 'number') return String(v);
      return String(v);
    };
    var kc = first.header.indexOf(keyName), ak = answer.header.indexOf(keyName);
    if (kc < 0 || ak < 0) return { error: '«' + keyName + '» 칸을 정리본과 정답 양쪽에서 찾지 못했다' };
    var common = [];
    first.header.forEach(function (h, i) {
      var j = answer.header.indexOf(h);
      if (j >= 0) common.push({ name: h, mine: i, ans: j });
    });
    var ansMap = {};
    answer.rows.forEach(function (r) { ansMap[norm(r[ak]).trim()] = r; });
    var matched = 0, cells = [], onlyMine = [], seenKeys = {};
    first.rows.forEach(function (r) {
      var k = norm(r[kc]).trim();
      var a = ansMap[k];
      if (!a) { onlyMine.push(k); return; }
      seenKeys[k] = true;
      var ok = true;
      common.forEach(function (c) {
        var mine = norm(r[c.mine]), ans = norm(a[c.ans]);
        if (mine !== ans) { ok = false; cells.push({ key: k, col: c.name, mine: mine === '' ? '(빈칸)' : mine, ans: ans === '' ? '(빈칸)' : ans }); }
      });
      if (ok) matched++;
    });
    var onlyAnswer = Object.keys(ansMap).filter(function (k) { return !seenKeys[k]; });
    return { answerRows: answer.rows.length, matched: matched, cells: cells, onlyMine: onlyMine, onlyAnswer: onlyAnswer };
  }

  var api = { cleanTable: cleanTable, checkTwice: checkTwice, compareWithAnswer: compareWithAnswer,
    serialToTime: serialToTime, serialToDate: serialToDate, isBlank: isBlank, isSerial: isSerial };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Cleaner = api;
})(typeof window !== 'undefined' ? window : this);
