/**
 * 年度收支計畫：依「個人收支計畫書」的格式，用固定收入、固定支出算出今年的計畫結餘，
 * 再和今年實際記到的收入、支出比較，並顯示離存錢目標還差多少。
 *
 *   年度計畫／收支計畫／年度總結 → 顯示計畫
 *   存錢目標 100萬               → 設定今年的存錢目標
 *   存錢目標                     → 顯示計畫（含目標）
 *   存錢目標 刪除                → 清除目標
 *
 * 存錢目標存在指令碼屬性 SAVING_GOAL，不需要另外的工作表。
 */

var SAVING_GOAL_PROP = 'SAVING_GOAL';

function parsePlanCommand_(text) {
  var t = normalizeText_(text).replace(/\s+/g, ' ').trim();
  if (/^(年度|收支|今年)(計畫|計劃|總結|收支|規劃)$/.test(t) || /^(收支計畫|收支計劃)書?$/.test(t)) {
    return { action: 'plan' };
  }
  var m = t.match(/^(?:今年)?(?:存錢|儲蓄|存款)目標\s*(.*)$/);
  if (!m) return null;
  var rest = m[1].replace(/^(設定|設為|改成|是|為|[:：])\s*/, '').trim();
  if (!rest) return { action: 'plan' };
  if (/^(刪除|清除|取消)$/.test(rest)) return { action: 'clearGoal' };
  var amount = parseMoneyText_(rest);
  return amount > 0 ? { action: 'setGoal', amount: amount } : { action: 'invalid' };
}

/** 「100萬」「1,000,000」「150萬元」「1.5萬」→ 數字；看不懂回傳 0。 */
function parseMoneyText_(s) {
  var m = String(s).replace(/元|塊/g, '').trim().match(/^(\d+(?:\.\d+)?)\s*(萬)?$/);
  if (!m) return 0;
  return Math.round(Number(m[1]) * (m[2] ? 10000 : 1));
}

function getSavingGoal_() {
  var raw = getProp_(SAVING_GOAL_PROP, false);
  var n = Math.round(Number(raw));
  return n > 0 ? n : 0;
}

function applyPlanCommand_(cmd, today) {
  var props = PropertiesService.getScriptProperties();
  if (cmd.action === 'setGoal') {
    props.setProperty(SAVING_GOAL_PROP, String(cmd.amount));
    return '🎯 已設定今年存錢目標 $' + formatMoney_(cmd.amount) + '\n\n' + formatPlan_(today);
  }
  if (cmd.action === 'clearGoal') {
    props.deleteProperty(SAVING_GOAL_PROP);
    return '已清除存錢目標。';
  }
  if (cmd.action === 'invalid') {
    return '存錢目標的用法：「存錢目標 100萬」「存錢目標 600000」\n看計畫：「年度計畫」';
  }
  return formatPlan_(today);
}

/** 一年發生幾次：每月 12 次、每 2 個月 6 次、每年 1 次。 */
function timesPerYear_(every) {
  return 12 / Math.max(1, every);
}

function formatPlan_(today) {
  var year = today.slice(0, 4);
  var items = listRecurringItems_();
  var goal = getSavingGoal_();
  var lines = ['📋 ' + year + ' 年度收支計畫'];

  var section = function (title, list, unit) {
    lines.push('', title);
    var sum = 0;
    if (!list.length) lines.push('・（還沒有設定）');
    list.forEach(function (item) {
      if (!(item.amount > 0)) {
        lines.push('・' + recurringLabel_(item) + '｜金額不固定（未計入）');
        return;
      }
      var value = unit === 'year' ? item.amount * timesPerYear_(item.every) : item.amount;
      sum += value;
      var extra = unit === 'year' ? '（' + cycleText_(item) + ' $' + formatMoney_(item.amount) + '）' : '';
      lines.push('・' + recurringLabel_(item) + ' $' + formatMoney_(value) + extra);
    });
    return sum;
  };

  // 貳、月淨收入：每月入帳的固定收入；每幾個月一次的收入平均到每月
  var incomes = items.filter(function (i) {
    return i.category === INCOME_CATEGORY;
  });
  var monthlyIncome = 0;
  lines.push('', '貳、月淨收入（固定收入）');
  if (!incomes.length) lines.push('・（還沒有設定，例如「固定收入 薪水 85000 每月5號」）');
  incomes.forEach(function (item) {
    if (!(item.amount > 0)) {
      lines.push('・' + recurringLabel_(item) + '｜金額不固定（未計入）');
      return;
    }
    var perMonth = item.amount / Math.max(1, item.every);
    monthlyIncome += perMonth;
    lines.push('・' + recurringLabel_(item) + ' $' + formatMoney_(perMonth) + (item.every > 1 ? '（' + cycleText_(item) + ' $' + formatMoney_(item.amount) + '，平均每月）' : ''));
  });
  lines.push('合計 $' + formatMoney_(monthlyIncome) + ' × 12 = $' + formatMoney_(monthlyIncome * 12) + ' / 年');

  // 參、月固定支出；肆、年固定支出（每幾個月一次的，換算成一年）
  var expenses = items.filter(function (i) {
    return i.category !== INCOME_CATEGORY;
  });
  var monthlyFixed = section('參、月固定支出', expenses.filter(function (i) {
    return i.every === 1;
  }), 'month');
  lines.push('合計 $' + formatMoney_(monthlyFixed) + ' × 12 = $' + formatMoney_(monthlyFixed * 12) + ' / 年');
  var yearlyFixed = section('肆、年固定支出（每幾個月或每年一次）', expenses.filter(function (i) {
    return i.every > 1;
  }), 'year');
  lines.push('合計 $' + formatMoney_(yearlyFixed) + ' / 年');

  // 伍、年度總結（計畫）
  var a = monthlyIncome * 12;
  var b = monthlyFixed * 12 + yearlyFixed;
  var net = a - b;
  lines.push('', '伍、年度總結（依固定收支）');
  lines.push('・年淨收入（A）$' + formatMoney_(a));
  lines.push('・年固定支出（B）$' + formatMoney_(b));
  lines.push('・年度淨損益（A−B）' + (net >= 0 ? '+$' : '−$') + formatMoney_(Math.abs(net)) + (net >= 0 ? '（結餘）' : '（赤字）'));
  lines.push('・平均每月 ' + (net >= 0 ? '+$' : '−$') + formatMoney_(Math.abs(net) / 12));
  if (goal > 0) {
    var gap = goal - net;
    // 計畫只算固定收支，日常花費要從結餘裡出
    lines.push('・存錢目標 $' + formatMoney_(goal) + '：' + (gap <= 0
      ? '日常花費每月控制在 $' + formatMoney_(-gap / 12) + ' 以內就能達成'
      : '固定收支就已經不夠，還差 $' + formatMoney_(gap) + '（每月要再多 $' + formatMoney_(gap / 12) + '）'));
  }

  // 今年實際：1/1 到今天記到的收入與支出（含日常花費）
  var start = year + '-01-01';
  var actualIncome = 0;
  var actualExpense = 0;
  readRows_().forEach(function (r) {
    if (r[COL.date] < start || r[COL.date] > today) return;
    var amount = Number(r[COL.amount]) || 0;
    if (r[COL.category] === INCOME_CATEGORY) actualIncome += amount;
    else actualExpense += amount;
  });
  var actualNet = actualIncome - actualExpense;
  var months = +today.slice(5, 7) - 1 + (+today.slice(8)) / daysInMonth_(+year, +today.slice(5, 7));
  lines.push('', '📊 今年實際（1/1～' + (+today.slice(5, 7)) + '/' + (+today.slice(8)) + '，含日常花費）');
  lines.push('・收入 $' + formatMoney_(actualIncome));
  lines.push('・支出 $' + formatMoney_(actualExpense));
  lines.push('・' + (actualNet >= 0 ? '結餘 $' : '⚠️ 赤字 $') + formatMoney_(Math.abs(actualNet)));
  if (goal > 0 && months > 0) {
    var pace = actualNet / months * 12;
    lines.push('・照目前速度，全年約可存 $' + formatMoney_(Math.max(0, pace)) + (pace >= goal ? '，可以達成目標 🎉' : '，離目標還差 $' + formatMoney_(goal - Math.max(0, pace))));
  }
  if (goal <= 0) lines.push('', '設定存錢目標：「存錢目標 100萬」');
  if (!incomes.length) lines.push('記收入：「收入 醫院 85000」，或設定「固定收入 薪水 85000 每月5號」');
  return lines.join('\n');
}
