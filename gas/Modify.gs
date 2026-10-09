/**
 * 修改與刪除記帳。只能改、刪自己記的帳。
 *
 * 修改：
 *   修改 150            → 最近一筆的金額改成 150
 *   修改 交通            → 最近一筆的分類改成交通
 *   修改 品項 早餐        → 最近一筆的品項改成早餐
 *   修改 午餐 150        → 最近一筆「午餐」的金額改成 150
 *   午餐改成150、鯖魚改成 餐飲
 * 刪除：
 *   刪除                → 最近一次記的帳（同一則訊息的多筆一起刪）
 *   刪除 午餐            → 最近一筆「午餐」
 *   刪除 昨天 停車 60     → 昨天、品項含「停車」、金額 60 的那筆
 *
 * 回傳的指令物件：
 *   { action: 'edit', keyword, date, amount, newAmount, newCategory, newItem }
 *   { action: 'delete', keyword, date, amount }
 *   { action: 'invalid', usage }
 */

function parseModifyCommand_(text, today) {
  var t = normalizeText_(text);

  // 「重複記帳了，刪除」「刪除重複」：刪掉自己最近重複的那筆
  if (/重複|重覆|記兩次|記了兩次/.test(t) && /刪|移除|取消|撤銷/.test(t)) return { action: 'dedupe' };

  // 「取消」只能單獨使用；指定刪除哪筆要用「刪除」，避免「取消聚餐」這種聊天誤刪
  if (t === '取消') return { action: 'delete', keyword: '', date: '', amount: null };
  var del = t.match(/^刪除\s*(.*)$/);
  if (del) {
    var target = parseTarget_(del[1].trim(), today);
    if (!target) return { action: 'invalid', usage: modifyUsage_() };
    target.action = 'delete';
    return target;
  }

  // 「午餐改成150」「鯖魚改成 餐飲」
  var inline = t.match(/^(.+?)\s*改成\s*(.+)$/);
  var edit = t.match(/^(修改|更正|更改|改成)\s*(.*)$/);
  if (!inline && !edit) return null;

  var cmd = { action: 'edit', keyword: '', date: '', amount: null, newAmount: null, newCategory: '', newItem: '' };
  var keywordPart;
  var changePart;
  if (inline && !edit) {
    keywordPart = inline[1];
    changePart = inline[2];
  } else {
    // 「修改 午餐 品項 早餐」「修改 150」「修改 午餐 金額 150」
    var rest = edit[2].trim();
    var itemMatch = rest.match(/^(.*?)\s*(?:品項|改名|名稱)\s*[:：]?\s*(.+)$/);
    if (itemMatch) {
      keywordPart = itemMatch[1];
      cmd.newItem = itemMatch[2].trim();
      changePart = '';
    } else {
      var tokens = rest.split(/\s+/).filter(function (x) {
        return x;
      });
      var last = tokens.pop() || '';
      keywordPart = tokens.join(' ');
      changePart = last;
    }
  }

  var change = String(changePart || '').replace(/^(金額|分類)\s*[:：]?\s*/, '').trim();
  if (change) {
    var num = change.match(/^(\d+(?:\.\d+)?)\s*(元|塊錢|塊)?$/);
    if (num) {
      cmd.newAmount = Math.round(Number(num[1]));
    } else if (CATEGORIES.indexOf(change) >= 0) {
      cmd.newCategory = change;
    } else if (inline && !edit) {
      // 「計畫改成明天」這種聊天不是修改指令
      return null;
    } else {
      cmd.newItem = change;
    }
  }
  if (!(cmd.newAmount > 0) && !cmd.newCategory && !cmd.newItem) {
    return { action: 'invalid', usage: modifyUsage_() };
  }

  var target2 = parseTarget_(String(keywordPart || '').replace(/(金額|分類)$/, '').trim(), today);
  if (!target2) return { action: 'invalid', usage: modifyUsage_() };
  cmd.keyword = target2.keyword;
  cmd.date = target2.date;
  cmd.amount = target2.amount;
  return cmd;
}

/** 「昨天 停車 60」→ { date, keyword, amount }，空字串代表「最近一筆」。 */
function parseTarget_(s, today) {
  var target = { keyword: '', date: '', amount: null };
  s = s.replace(/^(剛剛|剛才|最近一筆|上一筆|那筆|的)\s*/, '').replace(/(那筆|的)$/, '').trim();
  if (!s) return target;
  var d = matchDatePrefix_(s, today);
  if (d) {
    target.date = d.date;
    s = s.slice(d.length).trim();
  }
  var m = s.match(/^(.*?)\s*(\d+(?:\.\d+)?)\s*(元|塊錢|塊)?$/);
  if (m && m[1].trim()) {
    target.amount = Math.round(Number(m[2]));
    s = m[1].trim();
  }
  target.keyword = s.replace(/^的|的$/g, '').trim();
  return target;
}

/**
 * 找出要修改或刪除的列。
 * @return {{ rows: Array<{row: number, values: Array}>, ambiguous: boolean }}
 *   沒有條件時回傳最近一則訊息的所有列；有條件時回傳最近一筆符合的列。
 */
function findTargetRows_(userId, target) {
  var sheet = getLedgerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [] };
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var dateOf = function (v) {
    return v instanceof Date ? Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd') : String(v);
  };

  var hasCriteria = target.keyword || target.date || target.amount;
  for (var i = values.length - 1; i >= 0; i--) {
    var r = values[i];
    if (r[COL.userId] !== userId) continue;
    if (!hasCriteria) {
      var messageId = r[COL.messageId];
      var group = [];
      for (var j = i; j >= 0; j--) {
        if (values[j][COL.userId] === userId && values[j][COL.messageId] === messageId) {
          group.unshift({ row: j + 2, values: values[j] });
        }
      }
      return { rows: group };
    }
    var text = (String(r[COL.item]) + ' ' + String(r[COL.note])).toLowerCase();
    if (target.keyword && text.indexOf(target.keyword.toLowerCase()) < 0) continue;
    if (target.date && dateOf(r[COL.date]) !== target.date) continue;
    if (target.amount && Number(r[COL.amount]) !== target.amount) continue;
    return { rows: [{ row: i + 2, values: r }] };
  }
  return { rows: [] };
}

function describeRow_(r) {
  var d = r[COL.date] instanceof Date ? Utilities.formatDate(r[COL.date], TIMEZONE, 'yyyy-MM-dd') : String(r[COL.date]);
  return d + ' ' + r[COL.category] + ' ' + r[COL.item] + ' $' + formatMoney_(r[COL.amount]);
}

function notFound_(target) {
  if (!target.keyword && !target.date && !target.amount) return '找不到你可以修改或刪除的紀錄。';
  var desc = [target.date, target.keyword, target.amount ? '$' + target.amount : ''].filter(function (x) {
    return x;
  }).join(' ');
  return '找不到你記的「' + desc + '」。\n只能修改或刪除自己記的帳；其他人記的可以到試算表直接改。';
}

/**
 * 刪除自己最近重複的帳：在自己最近 20 筆裡，日期、品項、金額都相同的只留第一筆。
 * 找不到完全相同的，再看同一天、同金額的（例如照片辨識出的品項名稱略有不同）。
 */
function deleteDuplicates_(userId) {
  var sheet = getLedgerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var dateOf = function (v) {
    return v instanceof Date ? Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd') : String(v);
  };
  var mine = [];
  for (var i = values.length - 1; i >= 0 && mine.length < 20; i--) {
    if (values[i][COL.userId] === userId) mine.unshift({ row: i + 2, values: values[i] });
  }
  var findDups = function (keyOf) {
    var seen = {};
    return mine.filter(function (x) {
      var key = keyOf(x.values);
      if (seen[key]) return true;
      seen[key] = true;
      return false;
    });
  };
  var dups = findDups(function (r) {
    return dateOf(r[COL.date]) + '|' + r[COL.item] + '|' + Number(r[COL.amount]);
  });
  if (!dups.length) {
    dups = findDups(function (r) {
      return dateOf(r[COL.date]) + '|' + Number(r[COL.amount]);
    });
  }
  // 由下往上刪，列號才不會跑掉
  for (var j = dups.length - 1; j >= 0; j--) sheet.deleteRow(dups[j].row);
  return dups;
}

/** 執行修改或刪除，回傳回覆文字。 */
function applyModifyCommand_(cmd, userId) {
  if (cmd.action === 'invalid') return cmd.usage;
  if (cmd.action === 'dedupe') {
    var removed = deleteDuplicates_(userId);
    if (!removed.length) return '你最近記的帳沒有重複的。\n要刪最近一筆請傳「刪除」，指定哪一筆例如「刪除 健身房」。';
    return '🗑️ 已刪除重複的 ' + removed.length + ' 筆（保留第一筆）：\n' + removed.map(function (x) {
      return '・' + describeRow_(x.values);
    }).join('\n');
  }
  var found = findTargetRows_(userId, cmd);
  if (found.rows.length === 0) {
    var isRecurring = cmd.action === 'delete' && cmd.keyword && listRecurringItems_().some(function (item) {
      return item.name.toLowerCase() === cmd.keyword.toLowerCase();
    });
    if (isRecurring) return '「' + cmd.keyword + '」是固定支出，要刪除請傳「固定支出 刪除 ' + cmd.keyword + '」。';
    return notFound_(cmd);
  }
  var sheet = getLedgerSheet_();

  if (cmd.action === 'delete') {
    // 由下往上刪，列號才不會跑掉
    for (var i = found.rows.length - 1; i >= 0; i--) {
      sheet.deleteRow(found.rows[i].row);
    }
    return '🗑️ 已刪除：\n' + found.rows.map(function (x) {
      return '・' + describeRow_(x.values);
    }).join('\n');
  }

  if (found.rows.length > 1) {
    return '你最近一次記了 ' + found.rows.length + ' 筆，請指定要改哪一筆，例如：\n' +
      found.rows.map(function (x) {
        return '・修改 ' + x.values[COL.item] + ' 150';
      }).slice(0, 3).join('\n');
  }

  var target = found.rows[0];
  var before = describeRow_(target.values);
  var after = target.values.slice();
  if (cmd.newAmount > 0) {
    sheet.getRange(target.row, COL.amount + 1).setValue(cmd.newAmount);
    after[COL.amount] = cmd.newAmount;
  }
  if (cmd.newCategory) {
    sheet.getRange(target.row, COL.category + 1).setValue(cmd.newCategory);
    after[COL.category] = cmd.newCategory;
  }
  if (cmd.newItem) {
    sheet.getRange(target.row, COL.item + 1).setValue(cmd.newItem);
    after[COL.item] = cmd.newItem;
  }
  return '✏️ 已修改\n・原本：' + before + '\n・改成：' + describeRow_(after);
}

function modifyUsage_() {
  return [
    '修改的用法（只能改自己記的帳）：',
    '・修改 150 → 最近一筆的金額',
    '・修改 交通 → 最近一筆的分類',
    '・修改 品項 早餐 → 最近一筆的品項',
    '・午餐改成150 → 最近一筆「午餐」',
    '',
    '刪除的用法：',
    '・刪除 → 最近一次記的帳',
    '・刪除 午餐 → 最近一筆「午餐」',
    '・刪除 昨天 停車 60 → 指定日期、品項、金額',
    '・刪除重複 → 重複記到的帳只留一筆'
  ].join('\n');
}
