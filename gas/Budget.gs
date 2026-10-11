/**
 * 每月預算：總預算與分類預算，記帳後接近或超過預算時提醒。
 *
 *   預算                 → 本月預算使用狀況（剩多少、每天可以花多少）
 *   預算 30000           → 設定每月總預算
 *   預算 餐飲 8000       → 設定分類預算（也可以說「餐飲預算 8000」）
 *   預算 交通 10%        → 依總預算的比例分配，改總預算時會跟著變
 *   預算 餐飲 刪除       → 刪除分類預算；「預算 刪除」清除全部
 *   預算建議／631        → 依收入用 631 法則建議，並列出上月各分類花費當參考
 *
 * 預算存在指令碼屬性 BUDGETS（JSON），全家共用。支出含固定支出、不含收入。
 */

var BUDGET_PROP = 'BUDGETS';
var BUDGET_WARN_RATIO = 0.8;

function parseBudgetCommand_(text) {
  var t = normalizeText_(text).replace(/\s+/g, ' ').trim();
  if (/^(631|631法則|預算(建議|分配|規劃))$/.test(t)) return { action: 'suggest' };
  var cats = CATEGORIES.join('|');
  var m = t.match(new RegExp('^(?:每月|本月|這個月)?(' + cats + '|總)?預算\\s*(.*)$'));
  if (!m) return null;
  var category = m[1] && m[1] !== '總' ? m[1] : '';
  var rest = m[2];
  var c = rest.match(new RegExp('^(' + cats + '|總|總共|全部)\\s*(.*)$'));
  if (c && !category) {
    category = CATEGORIES.indexOf(c[1]) >= 0 ? c[1] : '';
    rest = c[2];
  }
  rest = rest.replace(/^(設定|設為|改成|改為|是|為|[:：])\s*/, '').replace(/^(每月|一個月|每個月)\s*/, '').trim();
  if (!rest) return { action: 'status' };
  if (/^(刪除|清除|取消)$/.test(rest)) return { action: 'clear', category: category };
  var pct = rest.match(/^(\d+(?:\.\d+)?)\s*[%％]$/);
  if (pct) {
    var p = Number(pct[1]);
    return category && p > 0 && p <= 100 ? { action: 'set', category: category, percent: p } : { action: 'invalid' };
  }
  var amount = parseMoneyText_(rest);
  return amount > 0 ? { action: 'set', category: category, amount: amount } : { action: 'invalid' };
}

function getBudgets_() {
  var raw = getProp_(BUDGET_PROP, false);
  var data = {};
  try {
    data = raw ? JSON.parse(raw) : {};
  } catch (err) {
    data = {};
  }
  return { total: Number(data.total) > 0 ? Number(data.total) : 0, categories: data.categories || {} };
}

function saveBudgets_(b) {
  var props = PropertiesService.getScriptProperties();
  if (!b.total && !Object.keys(b.categories).length) {
    props.deleteProperty(BUDGET_PROP);
    return;
  }
  props.setProperty(BUDGET_PROP, JSON.stringify(b));
}

/** 分類預算的金額：比例的換算成總預算的幾 %；總預算沒設定時比例的不算。 */
function categoryBudget_(b, category) {
  var v = b.categories[category];
  if (v === undefined) return 0;
  if (typeof v === 'string' && /%$/.test(v)) return Math.round(b.total * parseFloat(v) / 100);
  return Number(v) || 0;
}

function applyBudgetCommand_(cmd, today) {
  if (cmd.action === 'suggest') return formatBudgetSuggestion_(today);
  if (cmd.action === 'invalid') {
    return '預算的用法：\n・預算 30000（每月總預算）\n・預算 餐飲 8000\n・預算 交通 10%（總預算的比例）\n・預算 餐飲 刪除\n看狀況：「預算」';
  }
  var b = getBudgets_();
  if (cmd.action === 'clear') {
    if (!cmd.category) {
      saveBudgets_({ total: 0, categories: {} });
      return '已清除所有預算。';
    }
    if (b.categories[cmd.category] === undefined) return '「' + cmd.category + '」沒有設定預算。';
    delete b.categories[cmd.category];
    saveBudgets_(b);
    return '已刪除「' + cmd.category + '」的預算。\n\n' + formatBudgetStatus_(today);
  }
  if (cmd.action === 'set') {
    if (cmd.percent && !b.total) return '要先設定總預算，才能用比例分配，例如「預算 30000」。';
    var label;
    if (!cmd.category) {
      b.total = cmd.amount;
      label = '每月總預算 $' + formatMoney_(cmd.amount);
    } else {
      b.categories[cmd.category] = cmd.percent ? cmd.percent + '%' : cmd.amount;
      label = cmd.category + '每月預算 $' + formatMoney_(categoryBudget_(b, cmd.category)) + (cmd.percent ? '（總預算的 ' + cmd.percent + '%）' : '');
    }
    saveBudgets_(b);
    return '📌 已設定' + label + '\n\n' + formatBudgetStatus_(today);
  }
  return formatBudgetStatus_(today);
}

/** 本月 1 號到今天的支出：總額、各分類、固定支出。 */
function monthSpending_(today) {
  var s = summarize_(today.slice(0, 8) + '01', today, '全部', '');
  var fixed = 0;
  s.rows.forEach(function (r) {
    if (r[COL.source] === '固定支出') fixed += Number(r[COL.amount]) || 0;
  });
  return { total: s.total, byCategory: s.byCategory, fixed: fixed };
}

function budgetBar_(ratio) {
  var filled = Math.min(10, Math.round(ratio * 10));
  return '▓'.repeat(filled) + '░'.repeat(10 - filled);
}

function budgetLine_(budget, spent, daysLeft) {
  var pct = Math.round(spent / budget * 100);
  var left = budget - spent;
  if (left < 0) return '已花 $' + formatMoney_(spent) + '（' + pct + '%）｜🚨 超支 $' + formatMoney_(-left);
  return '已花 $' + formatMoney_(spent) + '（' + pct + '%）｜剩 $' + formatMoney_(left) + '，每天 $' + formatMoney_(Math.floor(left / daysLeft)) +
    (spent >= budget * BUDGET_WARN_RATIO ? ' ⚠️' : '');
}

function formatBudgetStatus_(today) {
  var b = getBudgets_();
  var cats = CATEGORIES.filter(function (c) {
    return categoryBudget_(b, c) > 0;
  });
  if (!b.total && !cats.length) {
    return '還沒有設定預算。\n・預算 30000（每月總預算）\n・預算 餐飲 8000（分類預算）\n・預算建議（用 631 法則算給你看）';
  }
  var y = +today.slice(0, 4);
  var mo = +today.slice(5, 7);
  var daysLeft = daysInMonth_(y, mo) - +today.slice(8) + 1;
  var spend = monthSpending_(today);
  var lines = ['📌 ' + mo + ' 月預算（含今天還有 ' + daysLeft + ' 天）'];
  if (b.total) {
    lines.push('', '總預算 $' + formatMoney_(b.total) + '  ' + budgetBar_(spend.total / b.total));
    lines.push(budgetLine_(b.total, spend.total, daysLeft));
    if (spend.fixed > 0) lines.push('（含固定支出 $' + formatMoney_(spend.fixed) + '）');
  }
  if (cats.length) {
    lines.push('', '分類預算：');
    var sum = 0;
    cats.forEach(function (c) {
      var budget = categoryBudget_(b, c);
      sum += budget;
      lines.push('・' + c + ' $' + formatMoney_(budget) + '｜' + budgetLine_(budget, spend.byCategory[c] || 0, daysLeft));
    });
    if (b.total && sum > b.total) lines.push('⚠️ 分類預算加總 $' + formatMoney_(sum) + '，超過總預算');
  }
  return lines.join('\n');
}

/**
 * 記帳後檢查預算：這次記帳讓花費超過 80% 或超支時提醒；已經超支的分類每次記帳都提醒。
 * 只看本月的支出，沒設預算時回傳空字串。
 */
function budgetAlerts_(entries, today) {
  var b = getBudgets_();
  if (!b.total && !Object.keys(b.categories).length) return '';
  var month = today.slice(0, 7);
  var added = { total: 0 };
  entries.forEach(function (e) {
    if (e.category === INCOME_CATEGORY || String(e.date).slice(0, 7) !== month) return;
    added.total += e.amount;
    added[e.category] = (added[e.category] || 0) + e.amount;
  });
  if (!added.total) return '';
  var spend = monthSpending_(today);
  var alerts = [];
  var check = function (label, budget, after, add) {
    if (!(budget > 0)) return;
    var before = after - add;
    if (after > budget) {
      alerts.push('🚨 ' + label + '已超支 $' + formatMoney_(after - budget) + '（預算 $' + formatMoney_(budget) + '）');
    } else if (after >= budget * BUDGET_WARN_RATIO && before < budget * BUDGET_WARN_RATIO) {
      alerts.push('⚠️ ' + label + '預算只剩 $' + formatMoney_(budget - after) + '（已用 ' + Math.round(after / budget * 100) + '%）');
    }
  };
  check('本月', b.total, spend.total, added.total);
  CATEGORIES.forEach(function (c) {
    if (added[c]) check('本月' + c, categoryBudget_(b, c), spend.byCategory[c] || 0, added[c]);
  });
  return alerts.join('\n');
}

/** 每月收入：固定收入換算成每月；沒有固定收入就用上個月記到的收入。 */
function monthlyIncome_(today) {
  var income = 0;
  listRecurringItems_().forEach(function (i) {
    if (i.category === INCOME_CATEGORY && i.amount > 0) income += i.amount / Math.max(1, i.every);
  });
  if (income > 0) return { amount: income, from: '固定收入' };
  var last = lastMonthRange_(today);
  var s = summarize_(last.start, last.end, INCOME_CATEGORY, '');
  return { amount: s.total, from: '上月收入' };
}

function lastMonthRange_(today) {
  var y = +today.slice(0, 4);
  var m = +today.slice(5, 7) - 1;
  if (m === 0) {
    m = 12;
    y -= 1;
  }
  var mm = (m < 10 ? '0' : '') + m;
  return { start: y + '-' + mm + '-01', end: y + '-' + mm + '-' + daysInMonth_(y, m), month: m };
}

function formatBudgetSuggestion_(today) {
  var lines = ['💡 631 法則：收入的 60% 生活、30% 儲蓄投資、10% 保險'];
  var income = monthlyIncome_(today);
  if (income.amount > 0) {
    var a = Math.round(income.amount);
    lines.push('', '每月收入 $' + formatMoney_(a) + '（' + income.from + '）');
    lines.push('・60% 日常支出 $' + formatMoney_(a * 0.6) + ' → 傳「預算 ' + Math.round(a * 0.6) + '」');
    lines.push('・30% 儲蓄投資 $' + formatMoney_(a * 0.3));
    lines.push('・10% 保險防護 $' + formatMoney_(a * 0.1));
  } else {
    lines.push('', '還沒有收入紀錄，先設定「固定收入 薪水 85000 每月5號」或記「收入 薪水 85000」。');
  }
  var last = lastMonthRange_(today);
  var s = summarize_(last.start, last.end, '全部', '');
  if (s.total > 0) {
    lines.push('', '📊 ' + last.month + ' 月實際花費（可當分類預算的參考）');
    CATEGORIES.forEach(function (c) {
      if (s.byCategory[c]) {
        lines.push('・' + c + ' $' + formatMoney_(s.byCategory[c]) + '（' + Math.round(s.byCategory[c] / s.total * 100) + '%）');
      }
    });
    lines.push('合計 $' + formatMoney_(s.total));
  }
  lines.push('', '每個家庭狀況不同，有小孩、房貸的生活比例會比較高。先記帳一個月，再依實際花費調整，例如「預算 餐飲 8000」。');
  return lines.join('\n');
}
