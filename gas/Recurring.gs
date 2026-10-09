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
          recorder: '🔁 固定支出',
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

/** 回覆「固定支出」指令：列出目前的固定支出。 */
function formatRecurringList_() {
  var sheet = getSpreadsheet_().getSheetByName(RECURRING_SHEET_NAME);
  var items = [];
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, RECURRING_HEADERS.length).getValues().forEach(function (r) {
      var item = readRecurringRow_(r);
      if (item) items.push(item);
    });
  }
  if (items.length === 0) {
    return '目前沒有固定支出。\n可以在試算表的「固定支出」工作表新增。';
  }
  var lines = items.map(function (item) {
    var amount = item.amount > 0 ? '$' + formatMoney_(item.amount) : '金額不固定（只提醒）';
    var cycle = item.every === 1 ? '每月' : '每 ' + item.every + ' 個月';
    return '・' + item.name + '｜' + amount + '｜' + cycle + ' ' + item.day + ' 號｜下次 ' + (item.next || '未設定');
  });
  return '🔁 固定支出\n' + lines.join('\n') + '\n\n要修改請到試算表的「固定支出」工作表。';
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
