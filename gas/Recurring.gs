/**
 * 固定支出：房租、水電這類定期的花費。
 *
 * 試算表「固定支出」工作表每列一項：
 *   名稱 | 金額 | 分類 | 扣款日 | 每幾個月 | 下次扣款日 | 啟用
 * - 金額有填：到了下次扣款日自動記帳。
 * - 金額空白（例如每期不同的水電費）：只提醒，繳費後再自己記。
 * - 扣款前一天會在群組提醒。
 *
 * dailyJob 由 setup 建立的每日觸發條件執行（每天早上 8 點左右）。
 */

var RECURRING_SHEET_NAME = '固定支出';
var RECURRING_HEADERS = ['名稱', '金額', '分類', '扣款日', '每幾個月', '下次扣款日', '啟用'];
var RCOL = { name: 0, amount: 1, category: 2, day: 3, every: 4, next: 5, enabled: 6 };

/** 每日執行：自動記帳到期的固定支出，並提醒明天要扣款的項目。 */
function dailyJob() {
  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var messages = withLock_(function () {
    return processRecurring_(today);
  });
  if (messages.length === 0) return;

  var target = getProp_('NOTIFY_TARGET_ID', false);
  if (!target) {
    console.log('尚未設定通知對象（把機器人加入群組後會自動設定）：\n' + messages.join('\n'));
    return;
  }
  pushText_(target, messages.join('\n\n'));
}

/** @return {Array<string>} 要通知的訊息 */
function processRecurring_(today) {
  var sheet = getSpreadsheet_().getSheetByName(RECURRING_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var tomorrow = addDays_(today, 1);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECURRING_HEADERS.length).getValues();
  var messages = [];

  values.forEach(function (r, i) {
    var item = readRecurringRow_(r);
    if (!item) return;
    var rowNumber = i + 2;

    if (!item.next) {
      item.next = firstDueDate_(today, item.day);
      sheet.getRange(rowNumber, RCOL.next + 1).setValue(item.next);
    }

    // 到期（含之前漏掉的）就記帳並排到下一期；最多補 12 期避免設定錯誤時跑太多次
    var guard = 0;
    while (item.next <= today && guard < 12) {
      guard++;
      if (item.amount > 0) {
        appendEntries_([{
          date: item.next,
          category: item.category,
          item: item.name,
          amount: item.amount,
          note: '固定支出自動記帳'
        }], {
          // 固定支出是全家的開銷，不算在任何一位家人身上
          recorder: FAMILY_RECORDER,
          userId: '',
          messageId: 'recurring:' + item.name + ':' + item.next,
          source: '固定支出'
        });
        messages.push('🔁 已自動記帳：' + item.name + ' $' + formatMoney_(item.amount) + '（' + item.next + '）');
      } else {
        messages.push('📅 今天是「' + item.name + '」繳費日，金額不固定。\n繳費後請傳「' + item.name + ' 金額」記帳。');
      }
      item.next = addMonths_(item.next, item.every, item.day);
      sheet.getRange(rowNumber, RCOL.next + 1).setValue(item.next);
    }

    if (item.next === tomorrow) {
      var amountText = item.amount > 0 ? ' $' + formatMoney_(item.amount) : '';
      messages.push('⏰ 提醒：明天（' + tomorrow + '）要繳「' + item.name + '」' + amountText);
    }
  });
  return messages;
}

/** 把一列轉成物件；未啟用或名稱空白回傳 null。 */
function readRecurringRow_(r) {
  var name = String(r[RCOL.name]).trim();
  var enabled = String(r[RCOL.enabled]).trim();
  if (!name || ['否', 'N', 'n', 'FALSE', 'false', '0'].indexOf(enabled) >= 0) return null;
  if (r[RCOL.enabled] === false) return null;

  var day = Math.round(Number(r[RCOL.day]));
  if (!(day >= 1 && day <= 31)) return null;
  var every = Math.round(Number(r[RCOL.every])) || 1;
  var category = String(r[RCOL.category]).trim();
  var next = r[RCOL.next];
  if (next instanceof Date) {
    next = Utilities.formatDate(next, TIMEZONE, 'yyyy-MM-dd');
  } else {
    next = /^\d{4}-\d{2}-\d{2}$/.test(String(next).trim()) ? String(next).trim() : '';
  }

  return {
    name: name,
    amount: Math.round(Number(r[RCOL.amount])) || 0,
    category: CATEGORIES.indexOf(category) >= 0 ? category : '其他',
    day: day,
    every: Math.max(1, every),
    next: next
  };
}

/** 從今天起算，第一個「每月 day 號」（今天就是扣款日則為今天）。 */
function firstDueDate_(today, day) {
  var y = +today.slice(0, 4);
  var m = +today.slice(5, 7);
  var thisMonth = ymd_(y, m, Math.min(day, daysInMonth_(y, m)));
  return thisMonth >= today ? thisMonth : addMonths_(thisMonth, 1, day);
}

/** 加 n 個月，日期固定為 day 號（該月沒有這天就用月底）。 */
function addMonths_(ymd, n, day) {
  var y = +ymd.slice(0, 4);
  var m = +ymd.slice(5, 7) + n;
  y += Math.floor((m - 1) / 12);
  m = ((m - 1) % 12) + 1;
  return ymd_(y, m, Math.min(day, daysInMonth_(y, m)));
}

function daysInMonth_(y, m) {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

/** 列出固定支出；filterName 有值時只列那一項。 */
function formatRecurringList_(filterName) {
  var items = listRecurringItems_().filter(function (item) {
    return !filterName || item.name === filterName;
  });
  if (items.length === 0) {
    return '目前沒有固定支出。\n' + recurringUsage_();
  }
  var lines = items.map(function (item) {
    var amount = item.amount > 0 ? '$' + formatMoney_(item.amount) : '金額不固定（只提醒）';
    var cycle = item.every === 1 ? '每月' : '每 ' + item.every + ' 個月';
    return '・' + item.name + '｜' + amount + '｜' + cycle + ' ' + item.day + ' 號｜下次 ' + (item.next || '未設定');
  });
  var footer = filterName ? '' : '\n\n新增或修改：「固定支出 Netflix 390 每月15號」\n刪除：「固定支出 刪除 Netflix」';
  return '🔁 固定支出\n' + lines.join('\n') + footer;
}

/** 建立「固定支出」工作表，預設放房租和水電（金額、扣款日請改成實際的）。 */
function ensureRecurringSheet_(today) {
  var ss = getSpreadsheet_();
  if (ss.getSheetByName(RECURRING_SHEET_NAME)) return;
  var sheet = ss.insertSheet(RECURRING_SHEET_NAME);
  sheet.appendRow(RECURRING_HEADERS);
  // 房租：金額請填上；水電：台灣多為兩個月一期且金額不固定，預設只提醒
  sheet.appendRow(['房租', '', '其他', 5, 1, firstDueDate_(today, 5), '是']);
  sheet.appendRow(['水電', '', '其他', 20, 2, firstDueDate_(today, 20), '是']);
  sheet.setFrozenRows(1);
}

/** 建立每天早上執行 dailyJob 的觸發條件（已存在就略過）。 */
function ensureDailyTrigger_() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'dailyJob';
  });
  if (exists) return;
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(8).inTimezone(TIMEZONE).create();
}

/**
 * 解析「固定支出」開頭的指令，不是這類指令回傳 null。
 *   固定支出                         → 列出
 *   固定支出 Netflix 390 每月15號     → 新增或更新
 *   固定支出 電費 不固定 每2個月 20號 → 金額不固定（只提醒）
 *   固定支出 房租 16000               → 只改金額
 *   固定支出 刪除 Netflix             → 刪除
 * 「訂閱」開頭也可以。
 */
function parseRecurringCommand_(text) {
  var t = normalizeText_(text)
    .replace(/每個月/g, '每月')
    .replace(/每(兩|二)個?月/g, '每2個月')
    .replace(/每三個?月/g, '每3個月');
  // 「修改固定支出房租…」「設定固定支出…」：開頭的動詞不影響意思
  var m = t.match(/^(?:修改|更改|更新|調整|設定|新增|改)?\s*(固定支出|訂閱)([\s\S]*)$/);
  if (!m) return null;
  var rest = m[2].trim();
  if (!rest) return { action: 'list' };

  // 一次設定多項：「固定支出」後面每一行一項
  var lines = rest.split(/\n+/).map(function (l) {
    return l.trim();
  }).filter(function (l) {
    return l;
  });
  if (lines.length > 1) {
    return {
      action: 'batch',
      cmds: lines.map(function (line) {
        var c = parseRecurringCommand_('固定支出 ' + line);
        return c && c.action !== 'list' ? c : { action: 'invalid', line: line };
      })
    };
  }
  // 「房租9900元每月1號」這種黏在一起的寫法，先拆成「房租 9900 每月 1號」
  rest = rest
    .replace(/每\s*(\d+)\s*個?月/g, ' 每$1個月 ')
    .replace(/每月/g, ' 每月 ')
    .replace(/(\d{1,2})\s*(號|日)/g, ' $1$2 ')
    .replace(/(\d+)\s*(元|塊錢|塊)/g, ' $1 ')
    .replace(/([^\d\s\-－每])(\d+)(?=\s|$)/g, '$1 $2')
    .trim();

  var del = rest.match(/^(刪除|移除)\s*(.+)$/);
  if (del) return { action: 'delete', name: del[2].trim() };

  var cmd = { action: 'upsert', name: '', amount: null, amountBlank: false, day: null, every: null, category: null };
  var nameParts = [];
  rest.split(/\s+/).forEach(function (token) {
    var cycleDay = token.match(/^每(月|(\d+)個?月)(?:(\d{1,2})(?:號|日)?)?$/);
    var day = token.match(/^(\d{1,2})(號|日)$/);
    var amount = token.match(/^(\d+)(元)?$/);
    if (cycleDay) {
      cmd.every = cycleDay[2] ? +cycleDay[2] : 1;
      if (cycleDay[3]) cmd.day = +cycleDay[3];
    } else if (day) {
      cmd.day = +day[1];
    } else if (amount && nameParts.length > 0) {
      cmd.amount = +amount[1];
    } else if (token === '不固定' || token === '金額不固定') {
      cmd.amountBlank = true;
    } else if (CATEGORIES.indexOf(token) >= 0) {
      cmd.category = token;
    } else {
      nameParts.push(token);
    }
  });
  cmd.name = nameParts.join(' ');
  if (!cmd.name) return { action: 'invalid' };
  // 沒有金額、日期、週期的不算指令（例如「固定支出好多」「固定支出有哪些」），交給問句判斷
  if (cmd.amount === null && !cmd.amountBlank && cmd.day === null && cmd.every === null && !cmd.category) return null;
  if (cmd.day !== null && !(cmd.day >= 1 && cmd.day <= 31)) return { action: 'invalid' };
  if (cmd.every !== null && !(cmd.every >= 1 && cmd.every <= 12)) return { action: 'invalid' };
  return cmd;
}

/** 「我有哪些固定支出」「房租什麼時候繳」這類問句：回傳要篩選的名稱（'' 表示全部），不是則回傳 null。 */
function matchRecurringQuestion_(text) {
  var t = normalizeText_(text);
  var asking = /什麼時候|幾號|哪天|何時|多少|哪些|清單|列表|有沒有|查|[?？]/.test(t);
  if (/固定支出|訂閱/.test(t) && asking) return '';
  if (!asking) return null;
  var names = listRecurringItems_().map(function (item) {
    return item.name;
  });
  for (var i = 0; i < names.length; i++) {
    if (t.toLowerCase().indexOf(names[i].toLowerCase()) >= 0) return names[i];
  }
  return null;
}

function listRecurringItems_() {
  var sheet = getSpreadsheet_().getSheetByName(RECURRING_SHEET_NAME);
  var items = [];
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, RECURRING_HEADERS.length).getValues().forEach(function (r) {
      var item = readRecurringRow_(r);
      if (item) items.push(item);
    });
  }
  return items;
}

/** 執行固定支出指令，回傳要回覆的文字。 */
function applyRecurringCommand_(cmd, today) {
  if (cmd.action === 'batch') {
    return cmd.cmds.map(function (c) {
      if (c.action === 'invalid') return '⚠️ 看不懂「' + (c.line || '') + '」，請寫成「名稱 金額 每月幾號」';
      return applyRecurringCommand_(c, today);
    }).join('\n\n');
  }
  if (cmd.action === 'list') return formatRecurringList_();
  if (cmd.action === 'invalid') return recurringUsage_();

  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(RECURRING_SHEET_NAME);
  if (!sheet) {
    ensureRecurringSheet_(today);
    sheet = ss.getSheetByName(RECURRING_SHEET_NAME);
  }
  var rowNumber = findRecurringRow_(sheet, cmd.name);

  if (cmd.action === 'delete') {
    if (!rowNumber) return '找不到固定支出「' + cmd.name + '」。傳「固定支出」可以看目前的項目。';
    sheet.deleteRow(rowNumber);
    return '🗑️ 已刪除固定支出「' + cmd.name + '」';
  }

  var item;
  if (rowNumber) {
    var r = sheet.getRange(rowNumber, 1, 1, RECURRING_HEADERS.length).getValues()[0];
    var name = String(r[RCOL.name]).trim();
    var day = cmd.day || Math.round(Number(r[RCOL.day])) || +today.slice(8);
    var every = cmd.every || Math.round(Number(r[RCOL.every])) || 1;
    var amount = cmd.amountBlank ? '' : (cmd.amount !== null ? cmd.amount : r[RCOL.amount]);
    var category = cmd.category || (CATEGORIES.indexOf(String(r[RCOL.category])) >= 0 ? String(r[RCOL.category]) : '其他');
    var current = readRecurringRow_(r);
    // 改了扣款日或週期，或原本沒有下次扣款日，就重新計算
    var next = (cmd.day || cmd.every || !current || !current.next) ? firstDueDate_(today, day) : current.next;
    item = { name: name, amount: amount, category: category, day: day, every: every, next: next };
    sheet.getRange(rowNumber, 1, 1, RECURRING_HEADERS.length).setValues([[
      name, amount, category, day, every, next, '是'
    ]]);
  } else {
    var newDay = cmd.day || +today.slice(8);
    var newEvery = cmd.every || 1;
    // 沒說幾號就用今天的日期，從下一期開始（今天這期多半已經付了）
    var newNext = cmd.day ? firstDueDate_(today, newDay) : addMonths_(today, newEvery, newDay);
    item = {
      name: cmd.name,
      amount: cmd.amountBlank || cmd.amount === null ? '' : cmd.amount,
      category: cmd.category || guessCategory_(cmd.name, loadKeywords_()),
      day: newDay,
      every: newEvery,
      next: newNext
    };
    sheet.appendRow([item.name, item.amount, item.category, item.day, item.every, item.next, '是']);
  }

  var amountText = item.amount !== '' && Number(item.amount) > 0 ? '$' + formatMoney_(item.amount) : '金額不固定（只提醒）';
  var cycle = item.every === 1 ? '每月' : '每 ' + item.every + ' 個月';
  return (rowNumber ? '✏️ 已更新' : '✅ 已新增') + '固定支出\n' +
    '・' + item.name + '｜' + amountText + '｜' + cycle + ' ' + item.day + ' 號\n' +
    '・下次扣款：' + item.next + '（前一天會在群組提醒）';
}

function findRecurringRow_(sheet, name) {
  if (sheet.getLastRow() < 2) return 0;
  var names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var target = String(name).trim().toLowerCase();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return 0;
}

function recurringUsage_() {
  return [
    '固定支出的用法：',
    '・固定支出 → 列出全部',
    '・固定支出 Netflix 390 每月15號 → 新增或更新',
    '・固定支出 電費 不固定 每2個月 20號 → 金額不固定，只提醒',
    '・固定支出 房租 16000 → 只改金額',
    '・固定支出 刪除 Netflix → 刪除'
  ].join('\n');
}

/** 「近期扣款」「最近要繳什麼」「這週要扣什麼」這類問句。 */
function matchDueSoonQuestion_(text) {
  var t = normalizeText_(text);
  return /(近期|最近|即將|快要|這週|本週|下週|接下來).*(扣款|繳|付|扣)/.test(t) ||
    /^(扣款|繳費)(提醒|清單|查詢)?$/.test(t);
}

/** 列出接下來 days 天內要扣款的固定支出。 */
function formatDueSoon_(today, days) {
  var limit = addDays_(today, days - 1);
  var items = listRecurringItems_().filter(function (item) {
    return item.next && item.next >= today && item.next <= limit;
  }).sort(function (a, b) {
    return a.next < b.next ? -1 : 1;
  });
  if (items.length === 0) {
    return '⏰ 接下來 ' + days + ' 天沒有要扣款的固定支出。\n傳「固定支出」可以看全部項目。';
  }
  var total = 0;
  var lines = items.map(function (item) {
    var left = daysBetween_(today, item.next);
    var when = left === 0 ? '今天' : left === 1 ? '明天' : left + ' 天後';
    if (item.amount > 0) total += item.amount;
    return '・' + (+item.next.slice(5, 7)) + '/' + (+item.next.slice(8)) + '（' + when + '）' + item.name + '｜' +
      (item.amount > 0 ? '$' + formatMoney_(item.amount) : '金額不固定');
  });
  return '⏰ 接下來 ' + days + ' 天要扣款\n' + lines.join('\n') +
    (total > 0 ? '\n\n已知金額合計 $' + formatMoney_(total) : '');
}
