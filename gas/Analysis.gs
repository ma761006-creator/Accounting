/**
 * 消費分析：總支出、日常平均每天、最大支出分類、和上一期比較。
 *
 * 金額都由程式從試算表計算（AI 只負責判斷要分析哪段期間）。
 * 固定支出（房租、水電等自動記帳）和日常花費分開看，
 * 不然房租這種大筆金額會讓「最大支出」永遠是同一項。
 */

var FIXED_SOURCE = '固定支出';

/**
 * 上一個比較期間：
 * - 從 1 號開始的期間（本月、上月）→ 上個月的同幾天，例如 10/1～10/9 對 9/1～9/9
 * - 其他期間 → 緊接在前、天數相同的期間，例如本週對上週
 */
function previousPeriod_(start, end) {
  var days = daysBetween_(start, end) + 1;
  if (start.slice(8) === '01') {
    var y = +start.slice(0, 4);
    var m = +start.slice(5, 7) - 1;
    if (m === 0) {
      y -= 1;
      m = 12;
    }
    var prevStart = ymd_(y, m, 1);
    var prevMonthEnd = ymd_(y, m, daysInMonth_(y, m));
    var prevEnd = addDays_(prevStart, days - 1);
    return [prevStart, prevEnd < prevMonthEnd ? prevEnd : prevMonthEnd];
  }
  var pEnd = addDays_(start, -1);
  return [addDays_(pEnd, -(days - 1)), pEnd];
}

function daysBetween_(a, b) {
  return Math.round((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 86400000);
}

/** 統計一段期間，把固定支出和日常花費分開。 */
function tally_(rows, start, end) {
  var t = { total: 0, daily: 0, fixed: 0, count: 0, byCategory: {}, byRecorder: {}, fixedItems: {}, byDate: {}, maxRow: null };
  rows.forEach(function (r) {
    var date = r[COL.date];
    if (date < start || date > end) return;
    var amount = Number(r[COL.amount]) || 0;
    t.total += amount;
    t.count += 1;
    if (r[COL.source] === FIXED_SOURCE) {
      t.fixed += amount;
      t.fixedItems[r[COL.item]] = (t.fixedItems[r[COL.item]] || 0) + amount;
      return;
    }
    t.daily += amount;
    t.byDate[date] = (t.byDate[date] || 0) + amount;
    if (!t.maxRow || amount > Number(t.maxRow[COL.amount])) t.maxRow = r;
    t.byCategory[r[COL.category]] = (t.byCategory[r[COL.category]] || 0) + amount;
    t.byRecorder[r[COL.recorder]] = (t.byRecorder[r[COL.recorder]] || 0) + amount;
  });
  return t;
}

/**
 * @param {string} start  YYYY-MM-DD
 * @param {string} end    YYYY-MM-DD（超過今天的部分不算天數）
 */
function analyze_(start, end, today) {
  if (end > today) end = today;
  if (start > end) start = end;
  var rows = readRows_();
  var prev = previousPeriod_(start, end);
  return {
    start: start,
    end: end,
    days: daysBetween_(start, end) + 1,
    prevStart: prev[0],
    prevEnd: prev[1],
    cur: tally_(rows, start, end),
    prev: tally_(rows, prev[0], prev[1])
  };
}

function formatAnalysis_(a) {
  var md = function (s) {
    return +s.slice(5, 7) + '/' + +s.slice(8);
  };
  var cur = a.cur;
  var lines = ['📊 消費分析（' + md(a.start) + '～' + md(a.end) + '）'];

  if (cur.count === 0) {
    lines.push('這段期間沒有紀錄。');
    return lines.join('\n');
  }

  // 總覽
  lines.push('');
  if (cur.fixed > 0) {
    lines.push('💵 總支出 $' + formatMoney_(cur.total) + '（日常 $' + formatMoney_(cur.daily) + '＋固定支出 $' + formatMoney_(cur.fixed) + '）');
  } else {
    lines.push('💵 總支出 $' + formatMoney_(cur.total));
  }
  lines.push('📅 日常花費平均每天 $' + formatMoney_(cur.daily / a.days) + '（' + a.days + ' 天）');

  // 和上一期比較（只比日常花費，固定支出每期差不多）
  var compareLabel = a.start.slice(8) === '01' ? '上月同期' : '前一段時間';
  var period = '（' + md(a.prevStart) + '～' + md(a.prevEnd) + '）';
  if (a.prev.daily > 0) {
    var diff = cur.daily - a.prev.daily;
    var pct = Math.round((Math.abs(diff) / a.prev.daily) * 100);
    if (Math.abs(diff) < 1 || pct === 0) {
      lines.push('📈 日常花費和' + compareLabel + period + '差不多');
    } else {
      lines.push((diff > 0 ? '📈' : '📉') + ' 日常花費比' + compareLabel + period +
        (diff > 0 ? '多' : '少') + ' $' + formatMoney_(Math.abs(diff)) + '（' + (diff > 0 ? '+' : '-') + pct + '%）');
    }
  } else {
    lines.push('📈 ' + compareLabel + period + '沒有日常花費紀錄，還無法比較');
  }

  // 最大支出分類與提醒
  var cats = Object.keys(cur.byCategory).sort(function (x, y) {
    return cur.byCategory[y] - cur.byCategory[x];
  });
  if (cats.length && cur.daily > 0) {
    var top = cats[0];
    var share = cur.byCategory[top] / cur.daily;
    lines.push('', '🔥 最大支出：' + top + ' $' + formatMoney_(cur.byCategory[top]) + '（' + Math.round(share * 100) + '%）');
    lines.push('💡 ' + insight_(top, share, cats.length));

    // 增加最多的分類
    var rising = cats.map(function (c) {
      return { c: c, d: cur.byCategory[c] - (a.prev.byCategory[c] || 0) };
    }).filter(function (x) {
      return a.prev.daily > 0 && x.d > 0;
    }).sort(function (x, y) {
      return y.d - x.d;
    })[0];
    // 最高單筆、最高單日（只看日常花費）
    if (cur.maxRow) {
      lines.push('🏆 最高單筆：' + cur.maxRow[COL.item] + ' $' + formatMoney_(cur.maxRow[COL.amount]) +
        '（' + md(String(cur.maxRow[COL.date])) + '，' + cur.maxRow[COL.recorder] + '）');
    }
    var days = Object.keys(cur.byDate);
    if (days.length > 1) {
      var topDay = days.sort(function (x, y) {
        return cur.byDate[y] - cur.byDate[x];
      })[0];
      lines.push('📆 最高單日：' + md(topDay) + ' $' + formatMoney_(cur.byDate[topDay]));
    }
    if (rising) {
      lines.push('📌 ' + rising.c + '比' + compareLabel + '多 $' + formatMoney_(rising.d) + '，是增加最多的項目');
    }

    lines.push('', '📂 日常花費分類');
    cats.forEach(function (c) {
      var line = '・' + c + ' $' + formatMoney_(cur.byCategory[c]) + '（' + Math.round((cur.byCategory[c] / cur.daily) * 100) + '%）';
      if (a.prev.daily > 0) {
        var d = cur.byCategory[c] - (a.prev.byCategory[c] || 0);
        if (Math.abs(d) >= 1) line += ' ' + (d > 0 ? '▲ +$' : '▼ -$') + formatMoney_(Math.abs(d));
      }
      lines.push(line);
    });
  }

  var fixedNames = Object.keys(cur.fixedItems);
  if (fixedNames.length) {
    lines.push('', '🔁 固定支出（全家共同）');
    fixedNames.forEach(function (n) {
      lines.push('・' + n + ' $' + formatMoney_(cur.fixedItems[n]));
    });
  }

  var people = Object.keys(cur.byRecorder);
  if (people.length > 1) {
    lines.push('', '👨‍👩‍👧 各家人（日常花費）');
    people.sort(function (x, y) {
      return cur.byRecorder[y] - cur.byRecorder[x];
    }).forEach(function (n) {
      lines.push('・' + n + ' $' + formatMoney_(cur.byRecorder[n]));
    });
  }
  return lines.join('\n');
}

function insight_(top, share, categoryCount) {
  if (categoryCount === 1) return '這段期間的日常花費都在「' + top + '」。';
  if (share >= 0.6) return top + '占了日常花費的' + Math.round(share * 10) + '成左右，是主要的花費來源。';
  if (share >= 0.4) return top + '是目前最大的花費，約占' + Math.round(share * 100) + '%。';
  return '各分類花費分布平均，沒有特別集中在哪一項。';
}

/**
 * 規則模式的分析指令：「分析」「本月分析」「上月分析」「幫我分析這個月的消費」「本週分析」。
 * 不是分析指令則回傳 null。
 */
function parseAnalysis_(text, today) {
  var t = normalizeText_(text);
  if (!/分析/.test(t)) return null;
  var s = t.replace(/幫我|請|分析|消費|花費|支出|一下|的|[?？!！。]/g, '').trim();
  var monthStart = today.slice(0, 8) + '01';
  if (!s) return { start_date: monthStart, end_date: today, category: '全部', keyword: '', detail: false };
  var q = parseQuery_(s, today);
  if (!q || q.keyword) return null;
  return { start_date: q.start_date, end_date: q.end_date, category: '全部', keyword: '', detail: false };
}
