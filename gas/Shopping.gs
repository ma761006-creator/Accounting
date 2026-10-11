/**
 * 購物規劃：全家共用的購物清單，以及想買的大東西的存錢計畫。
 *
 * 購物清單（工作表「購物清單」）：
 *   要買 衛生紙、牛奶     → 加進清單
 *   購物清單／要買什麼    → 列出還沒買的
 *   買好了 衛生紙         → 劃掉（記帳時品項對得上也會自動劃掉）
 *   購物清單 清空         → 全部清掉
 *
 * 存錢計畫（指令碼屬性 WISHLIST）：
 *   想買 iPhone 36000 3月前   → 加入；可以不寫期限
 *   購物計畫                 → 每項還要存多久、每月要存多少
 *   想買 刪除 iPhone          → 移除（買到了也用這個）
 */

var SHOPPING_SHEET_NAME = '購物清單';
var SHOPPING_HEADERS = ['品項', '加入者', '加入時間', '狀態'];
var WISHLIST_PROP = 'WISHLIST';

// ===== 購物清單 =====

function parseShoppingCommand_(text) {
  var t = normalizeText_(text).replace(/[ \t]+/g, ' ').trim();
  if (/^(購物清單|待買清單|採買清單|要買什麼|要買哪些|還要買什麼|買什麼)[?？]?$/.test(t)) return { action: 'list' };
  if (/^(購物清單|待買清單)\s*(清空|全部刪除|清除)$/.test(t)) return { action: 'clear' };
  var add = t.match(/^(?:要買|待買|記得買|購物清單\s*(?:加|新增|加入)?|加入購物清單|幫我記(?:得)?買)\s*[:：]?\s*([\s\S]+)$/);
  if (add) {
    // 「要買衛生紙 560」是記帳；「要買票嗎？」是聊天，都不是購物清單
    if (/(嗎|呢|吧|[?？])$/.test(add[1])) return null;
    if (/\d+\s*(元|塊)?\s*$/.test(add[1]) && !/[、，,\n]/.test(add[1])) return null;
    var items = splitShoppingItems_(add[1]);
    return items.length ? { action: 'add', items: items } : null;
  }
  var done = t.match(/^(?:買好了|買好|已買|買到了|劃掉)\s*[:：]?\s*([\s\S]+)$/);
  if (done && !/^\d/.test(done[1])) {
    var names = splitShoppingItems_(done[1]);
    return names.length ? { action: 'done', items: names } : null;
  }
  return null;
}

function splitShoppingItems_(s) {
  return String(s).split(/[、，,\n]+|\s+(?:和|跟|還有)\s+|\s+/).map(function (x) {
    return x.replace(/^(和|跟|還有)/, '').trim();
  }).filter(function (x) {
    return x && x.length <= 30;
  });
}

function getShoppingSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHOPPING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHOPPING_SHEET_NAME);
    sheet.appendRow(SHOPPING_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 還沒買的：[{ row, name, by }] */
function pendingShopping_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var items = [];
  sheet.getRange(2, 1, last - 1, SHOPPING_HEADERS.length).getValues().forEach(function (r, i) {
    if (String(r[0]).trim() && String(r[3]) !== '已買') items.push({ row: i + 2, name: String(r[0]).trim(), by: String(r[1] || '') });
  });
  return items;
}

function formatShoppingList_(sheet) {
  var items = pendingShopping_(sheet);
  if (!items.length) return '🛒 購物清單是空的。\n加東西：「要買 衛生紙、牛奶」';
  return '🛒 購物清單（' + items.length + ' 項）\n' + items.map(function (x) {
    return '・' + x.name + (x.by ? '（' + x.by + '）' : '');
  }).join('\n') + '\n\n買好了傳「買好了 ' + items[0].name + '」，或記帳時會自動劃掉。';
}

function applyShoppingCommand_(cmd, who) {
  var sheet = getShoppingSheet_();
  if (cmd.action === 'list') return formatShoppingList_(sheet);
  if (cmd.action === 'clear') {
    var pending = pendingShopping_(sheet);
    for (var i = pending.length - 1; i >= 0; i--) sheet.deleteRow(pending[i].row);
    return '🧹 已清空購物清單（' + pending.length + ' 項）。';
  }
  if (cmd.action === 'add') {
    var existing = pendingShopping_(sheet).map(function (x) {
      return x.name.toLowerCase();
    });
    var now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm');
    var added = [];
    cmd.items.forEach(function (name) {
      if (existing.indexOf(name.toLowerCase()) >= 0) return;
      sheet.appendRow([name, who || '', now, '未買']);
      existing.push(name.toLowerCase());
      added.push(name);
    });
    var head = added.length ? '🛒 已加入：' + added.join('、') : '這些都已經在清單上了。';
    return head + '\n\n' + formatShoppingList_(sheet);
  }
  if (cmd.action === 'done') {
    var crossed = crossOffShopping_(cmd.items);
    if (!crossed.length) return '購物清單上找不到「' + cmd.items.join('、') + '」。\n\n' + formatShoppingList_(sheet);
    return '✅ 已劃掉：' + crossed.join('、') + '\n\n' + formatShoppingList_(sheet);
  }
  return formatShoppingList_(sheet);
}

/** 把清單上名稱對得上的劃掉（任一方包含另一方），回傳劃掉的品項。 */
function crossOffShopping_(names) {
  var sheet = getSpreadsheet_().getSheetByName(SHOPPING_SHEET_NAME);
  if (!sheet) return [];
  var crossed = [];
  pendingShopping_(sheet).forEach(function (x) {
    var key = x.name.toLowerCase();
    var hit = names.some(function (n) {
      var k = String(n).toLowerCase();
      return k && (k.indexOf(key) >= 0 || key.indexOf(k) >= 0);
    });
    if (hit) {
      sheet.getRange(x.row, 4).setValue('已買');
      crossed.push(x.name);
    }
  });
  return crossed;
}

/** 記帳後：品項對得上購物清單的自動劃掉，回傳要附在回覆後面的文字。 */
function shoppingAfterRecord_(entries) {
  var names = entries.filter(function (e) {
    return e.category !== INCOME_CATEGORY;
  }).map(function (e) {
    return e.item;
  });
  if (!names.length) return '';
  var crossed = crossOffShopping_(names);
  return crossed.length ? '🛒 已從購物清單劃掉：' + crossed.join('、') : '';
}

/** 「購物規劃」：購物清單和存錢計畫一起看。 */
function formatShoppingPlan_(today) {
  return formatShoppingList_(getShoppingSheet_()) + '\n\n' + formatWishes_(today);
}

// ===== 想買的東西：存錢計畫 =====

function parseWishCommand_(text) {
  var t = normalizeText_(text).replace(/\s+/g, ' ').trim();
  if (/^(購物計畫|購物計劃|想買清單|想買什麼|存錢計畫|願望清單)$/.test(t)) return { action: 'list' };
  var m = t.match(/^(?:想買|想要買|計畫買|打算買|存錢買)\s*([\s\S]+)$/);
  if (!m) return null;
  var rest = m[1].trim();
  var del = rest.match(/^(?:刪除|移除|取消|買了|買到了)\s*(.+)$/);
  if (del) return { action: 'delete', name: del[1].trim() };
  var deadline = '';
  rest = rest.replace(/\s*((?:\d{4}\s*年\s*)?\d{1,2}\s*月(?:前|底|之前)?|年底(?:前)?|明年|\d+\s*個月(?:內|後|之內)?)\s*$/, function (all) {
    deadline = all.trim();
    return '';
  }).trim();
  var price = rest.match(/^(.+?)\s*(\d+(?:\.\d+)?)\s*(萬)?\s*(元|塊)?$/);
  if (!price || !price[1].trim()) return { action: 'invalid' };
  var amount = Math.round(Number(price[2]) * (price[3] ? 10000 : 1));
  if (!(amount > 0)) return { action: 'invalid' };
  return { action: 'add', name: price[1].trim(), price: amount, deadline: deadline };
}

/** 期限文字 → 'YYYY-MM'（那個月底前要買到）；看不懂或沒寫回傳 ''。 */
function wishDeadline_(text, today) {
  if (!text) return '';
  var y = +today.slice(0, 4);
  var mo = +today.slice(5, 7);
  var pad = function (yy, mm) {
    return yy + '-' + (mm < 10 ? '0' : '') + mm;
  };
  var n = text.match(/(\d+)\s*個月/);
  if (n) {
    // 「6個月內」：含這個月共 6 個月
    var total = mo + +n[1] - 1;
    return pad(y + Math.floor((total - 1) / 12), ((total - 1) % 12) + 1);
  }
  if (/年底/.test(text)) return pad(y, 12);
  if (/明年/.test(text)) return pad(y + 1, 12);
  var m = text.match(/(?:(\d{4})\s*年\s*)?(\d{1,2})\s*月/);
  if (!m || +m[2] < 1 || +m[2] > 12) return '';
  var yy = m[1] ? +m[1] : (+m[2] < mo ? y + 1 : y);
  return pad(yy, +m[2]);
}

function getWishes_() {
  try {
    var list = JSON.parse(getProp_(WISHLIST_PROP, false) || '[]');
    return Array.isArray(list) ? list : [];
  } catch (err) {
    return [];
  }
}

function saveWishes_(list) {
  var props = PropertiesService.getScriptProperties();
  if (list.length) props.setProperty(WISHLIST_PROP, JSON.stringify(list));
  else props.deleteProperty(WISHLIST_PROP);
}

function applyWishCommand_(cmd, today, who) {
  if (cmd.action === 'invalid') {
    return '存錢計畫的用法：\n・想買 iPhone 36000 3月前\n・想買 冷氣 4.5萬\n・購物計畫（看進度）\n・想買 刪除 iPhone';
  }
  var list = getWishes_();
  if (cmd.action === 'delete') {
    var key = cmd.name.toLowerCase();
    var left = list.filter(function (w) {
      return String(w.name).toLowerCase() !== key;
    });
    if (left.length === list.length) return '購物計畫裡沒有「' + cmd.name + '」。\n\n' + formatWishes_(today);
    saveWishes_(left);
    return '🗑️ 已移除「' + cmd.name + '」\n\n' + formatWishes_(today);
  }
  if (cmd.action === 'add') {
    var by = wishDeadline_(cmd.deadline, today);
    list = list.filter(function (w) {
      return String(w.name).toLowerCase() !== cmd.name.toLowerCase();
    });
    list.push({ name: cmd.name, price: cmd.price, by: by, who: who || '', added: today });
    saveWishes_(list);
    return '🎯 已加入購物計畫：' + cmd.name + ' $' + formatMoney_(cmd.price) + (by ? '（' + by.replace('-', '/') + ' 前）' : '') +
      '\n\n' + formatWishes_(today);
  }
  return formatWishes_(today);
}

/** 最近 3 個完整月份的平均每月結餘（收入 − 支出），沒有收入紀錄回傳 null。 */
function averageMonthlySaving_(today) {
  var y = +today.slice(0, 4);
  var m = +today.slice(5, 7);
  var months = [];
  for (var i = 1; i <= 3; i++) {
    var mm = m - i;
    var yy = y;
    while (mm <= 0) {
      mm += 12;
      yy -= 1;
    }
    months.push(yy + '-' + (mm < 10 ? '0' : '') + mm);
  }
  var income = 0;
  var expense = 0;
  var seen = {};
  readRows_().forEach(function (r) {
    var key = String(r[COL.date]).slice(0, 7);
    if (months.indexOf(key) < 0) return;
    var amount = Number(r[COL.amount]) || 0;
    if (r[COL.category] === INCOME_CATEGORY) {
      income += amount;
      seen[key] = true;
    } else {
      expense += amount;
    }
  });
  var n = Object.keys(seen).length;
  if (!n) return null;
  return Math.round((income - expense) / n);
}

function monthsUntil_(today, by) {
  var diff = (+by.slice(0, 4) - +today.slice(0, 4)) * 12 + (+by.slice(5, 7) - +today.slice(5, 7)) + 1;
  return Math.max(1, diff);
}

function formatWishes_(today) {
  var list = getWishes_();
  if (!list.length) return '🎯 購物計畫是空的。\n加入想買的東西：「想買 iPhone 36000 3月前」';
  var saving = averageMonthlySaving_(today);
  var lines = ['🎯 購物計畫'];
  var needPerMonth = 0;
  var total = 0;
  list.forEach(function (w) {
    total += w.price;
    var line = '・' + w.name + ' $' + formatMoney_(w.price);
    if (w.by) {
      var months = monthsUntil_(today, w.by);
      var per = Math.ceil(w.price / months);
      needPerMonth += per;
      line += '｜' + w.by.replace('-', '/') + ' 前，還有 ' + months + ' 個月，每月要存 $' + formatMoney_(per);
    } else if (saving > 0) {
      line += '｜照目前結餘約 ' + Math.max(1, Math.ceil(w.price / saving)) + ' 個月可以買';
    }
    lines.push(line);
  });
  lines.push('', '合計 $' + formatMoney_(total));
  if (saving === null) {
    lines.push('還沒有收入紀錄，算不出每月結餘。記收入：「收入 薪水 85000」');
  } else {
    lines.push('最近 3 個月平均每月結餘 $' + formatMoney_(saving));
    if (needPerMonth > 0) {
      lines.push(needPerMonth <= saving
        ? '✅ 有期限的項目每月要存 $' + formatMoney_(needPerMonth) + '，目前的結餘夠'
        : '⚠️ 有期限的項目每月要存 $' + formatMoney_(needPerMonth) + '，比結餘多 $' + formatMoney_(needPerMonth - saving) + '，可能要延後或調整預算');
    }
  }
  lines.push('', '買到了傳「想買 刪除 ' + list[0].name + '」');
  return lines.join('\n');
}
