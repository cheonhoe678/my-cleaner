/* 접속 분석 계산 — 칸 이름과 기준값은 rules/*.json 의 "analysis" 에서 읽는다.
   읽지 못한 값은 지어내지 않고 «뺀 줄»로 센다. 네트워크 코드는 없다. */
(function (root) {
  'use strict';

  var DAYS = ['월', '화', '수', '목', '금', '토', '일'];
  var HOUR_RE = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$/;
  var HM_RE = /^\s*(\d{1,2})시\s*(\d{1,2})분\s*$/;
  var DATE_RE = /^\s*(\d{4})-(\d{2})-(\d{2})(?:[ T]00:00:00)?\s*$/;
  var NUM_RE = /^\s*[+-]?\d+(?:\.\d+)?\s*$/;

  // 셀 값을 글자로: 엑셀 날짜·시간 값은 엑셀 화면에 보이는 글자로 읽는다 (고치는 것이 아니라 읽는 것)
  function text(v, role) {
    if (v == null) return '';
    if (Cleaner.isSerial(v)) {
      if (role === 'time') return Cleaner.serialToTime(v.serial);
      if (role === 'date') return Cleaner.serialToDate(v.serial);
      return String(v.serial);
    }
    return String(v);
  }
  function hourOf(s) {
    var m = HOUR_RE.exec(s) || HM_RE.exec(s);
    if (!m) return null;
    var h = +m[1];
    return h >= 0 && h <= 23 ? h : null;
  }
  function dateOf(s) {
    var m = DATE_RE.exec(s);
    if (!m) return null;
    var y = +m[1], mo = +m[2], d = +m[3];
    var t = Date.UTC(y, mo - 1, d), dt = new Date(t);
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
    return { t: t, dow: (dt.getUTCDay() + 6) % 7, key: m[1] + '-' + m[2] + '-' + m[3] };
  }

  /* table: {header, rows}  cfg: rules.analysis  groupName: 묶을 칸 (없으면 cfg.group_column) */
  function analyze(table, cfg, groupName) {
    var h = table.header;
    var ix = function (n) { return n ? h.indexOf(n) : -1; };
    var ti = ix(cfg.time_column), di = ix(cfg.date_column), mi = ix(cfg.minutes_column), ui = ix(cfg.user_column);
    var gname = groupName || cfg.group_column, gi = ix(gname);
    var n = table.rows.length;
    var excluded = {};
    var add = function (k, c) { if (c) excluded[k] = c; };

    // 날짜
    var dates = new Array(n).fill(null), unreadDate = 0, farDate = 0;
    if (di >= 0) {
      var valid = [];
      table.rows.forEach(function (r, i) {
        var d = dateOf(text(r[di], 'date'));
        if (d) { dates[i] = d; valid.push(d.t); } else unreadDate++;
      });
      if (valid.length) {
        valid.sort(function (a, b) { return a - b; });
        var mid = valid[Math.floor(valid.length / 2)];
        var win = (cfg.date_window_days || 120) * 86400000;
        dates.forEach(function (d, i) {
          if (d && Math.abs(d.t - mid) > win) { dates[i] = null; farDate++; }
        });
      }
      add('날짜를 읽지 못함 (빈칸 · 다른 모양)', unreadDate);
      add('튀는 날짜 (가운데 날짜에서 ' + (cfg.date_window_days || 120) + '일 넘게 멂)', farDate);
    }

    // 시
    var hours = new Array(n).fill(null), unreadHour = 0;
    if (ti >= 0) {
      table.rows.forEach(function (r, i) {
        var v = hourOf(text(r[ti], 'time'));
        if (v == null) unreadHour++; else hours[i] = v;
      });
      add('시각을 읽지 못함 (빈칸 · 다른 모양)', unreadHour);
    }

    // 접속시간(분) — 평균에서만 뺀다
    var mins = new Array(n).fill(null), badMin = 0, unreadMin = 0;
    if (mi >= 0) {
      var lo = cfg.minutes_valid_above, hi = cfg.minutes_valid_below;
      table.rows.forEach(function (r, i) {
        var s = text(r[mi], 'num');
        if (!NUM_RE.test(s)) { unreadMin++; return; }
        var v = Number(s);
        if (v <= lo || v >= hi) { badMin++; return; }
        mins[i] = v;
      });
      add('튀는 접속시간 (' + lo + ' 이하 · ' + hi + ' 이상) — 평균에서만 뺌', badMin);
      add('접속시간을 읽지 못함 — 평균에서만 뺌', unreadMin);
    }

    // 집계
    var hourly = new Array(24).fill(0), weekday = new Array(7).fill(0), heat = [];
    var r7, c24;
    for (r7 = 0; r7 < 7; r7++) { heat.push(new Array(24).fill(0)); }
    var daily = {}, groups = {}, users = {}, userCount = {};
    var used = 0, sum = 0, cnt = 0;
    table.rows.forEach(function (r, i) {
      var hr = hours[i], d = dates[i];
      if (hr != null) { hourly[hr]++; }
      if (d) { weekday[d.dow]++; daily[d.key] = (daily[d.key] || 0) + 1; }
      if (hr != null && d) heat[d.dow][hr]++;
      var counted = ti >= 0 ? hr != null : (di >= 0 ? !!d : true);
      if (counted) used++;
      if (gi >= 0) {
        var g = text(r[gi], 'text').trim() || '(빈칸)';
        var o = groups[g] || (groups[g] = { name: g, count: 0, msum: 0, mcnt: 0 });
        o.count++;
        if (mins[i] != null) { o.msum += mins[i]; o.mcnt++; }
      }
      if (ui >= 0) {
        var u = text(r[ui], 'text').trim().toLowerCase();
        if (u) { users[u] = true; userCount[u] = (userCount[u] || 0) + 1; }
      }
      if (mins[i] != null) { sum += mins[i]; cnt++; }
    });

    var groupList = Object.keys(groups).map(function (k) {
      var o = groups[k];
      return { name: o.name, count: o.count, avgMin: o.mcnt ? o.msum / o.mcnt : null };
    }).sort(function (a, b) { return b.count - a.count || (a.name < b.name ? -1 : 1); });
    var top = groupList.filter(function (g) { return g.name !== '(빈칸)'; })[0] || null;
    var peak = null, mx = 0;
    hourly.forEach(function (c, hr) { if (c > mx) { mx = c; peak = hr; } });
    var topUsers = Object.keys(userCount).map(function (k) { return { name: k, count: userCount[k] }; })
      .sort(function (a, b) { return b.count - a.count || (a.name < b.name ? -1 : 1); }).slice(0, 10);
    var dailyList = Object.keys(daily).sort().map(function (k) { return { name: k, count: daily[k] }; });

    return {
      total: n, used: used, excluded: excluded, groupName: gname,
      hourly: hourly, weekday: weekday, heat: heat, daily: dailyList,
      groups: groupList, topGroup: top, topUsers: topUsers,
      userKinds: ui >= 0 ? Object.keys(users).length : null,
      groupKinds: gi >= 0 ? groupList.filter(function (g) { return g.name !== '(빈칸)'; }).length : null,
      avgMin: cnt ? sum / cnt : null, peakHour: peak,
      has: { time: ti >= 0, date: di >= 0, group: gi >= 0, minutes: mi >= 0, user: ui >= 0 }
    };
  }

  var api = { analyze: analyze, DAYS: DAYS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Analyzer = api;
})(typeof window !== 'undefined' ? window : this);
