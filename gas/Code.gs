/**
 * LINE 家庭記帳機器人：Webhook 入口。
 *
 * 部署成「網頁應用程式」後，把網址填到 LINE Developers 的 Webhook URL。
 */

// 用函式而不是全域變數：Code.gs 會比 Config.gs 先載入，此時 CATEGORIES 還沒定義
function helpText_() {
  if (getProvider_() === 'rules') {
    return [
      '📒 家庭記帳機器人',
      '',
      '記帳：品項 金額，例如',
      '・午餐 120',
      '・昨天 全聯 560',
      '・10/8 中油 1200',
      '・掛號 150 醫療（最後加分類可指定分類）',
      '・加油 1200、停車 60（多筆用「、」隔開）',
      '',
      '查詢：今天、本週、本月、上月、今年',
      '・可加分類或店名：「本月 餐飲」「全聯花多少」',
      '・加「明細」列出每一筆：「本月 明細」',
      '',
      '分析：「分析」「上月分析」「本週分析」',
      '・平均每天花多少、最大支出、和上個月同期比較',
      '',
      '固定支出：傳「固定支出」查看房租、水電等',
      '・新增或修改：固定支出 Netflix 390 每月15號',
      '',
      '刪除：傳「刪除」會刪掉你最近一次記的帳',
      '',
      '分類：' + CATEGORIES.join('、')
    ].join('\n');
  }
  return [
    '📒 家庭記帳機器人',
    '',
    '記帳：直接傳訊息，例如',
    '・午餐 120',
    '・昨天全聯 560 衛生紙',
    '・加油 1200、停車 60',
    '・或直接拍收據 / 發票照片',
    '',
    '查詢：例如',
    '・這個月花多少？',
    '・上個月餐飲多少',
    '・這個月全聯花多少',
    '・本月明細',
    '',
    '分析：例如',
    '・幫我分析這個月的消費',
    '・上個月花得比較多嗎',
    '',
    '固定支出：傳「固定支出」查看房租、水電等',
    '・新增：我每個月訂 Netflix 390，15 號扣款',
    '',
    '刪除：傳「刪除」會刪掉你最近一次記的帳',
    '',
    '分類：' + CATEGORIES.join('、')
  ].join('\n');
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  // 選填的簡易檢查：只處理送給自己這個機器人的事件。
  // Apps Script 讀不到 X-Line-Signature 標頭，無法驗證簽章，請勿公開 Webhook 網址。
  var botUserId = getProp_('LINE_BOT_USER_ID', false);
  if (botUserId && body.destination !== botUserId) {
    return ok_();
  }

  (body.events || []).forEach(function (event) {
    try {
      handleEvent_(event);
    } catch (err) {
      console.error(err && err.stack ? err.stack : err);
      if (event.replyToken) {
        replyText_(event.replyToken, friendlyError_(err, event));
      }
    }
  });
  return ok_();
}

function friendlyError_(err, event) {
  var msg = String((err && err.message) || err);
  var isImage = event.message && event.message.type === 'image';
  if (/忙線|額度/.test(msg)) {
    return '⏳ ' + msg + '，' + (isImage ? '照片暫時無法辨識，請稍後再傳一次，或先用文字記帳，例如「全聯 560」。' : '請稍後再試一次。');
  }
  return '⚠️ 處理失敗，請稍後再試一次。\n（' + msg.slice(0, 200) + '）';
}

function ok_() {
  return ContentService.createTextOutput('OK');
}

function handleEvent_(event) {
  // LINE 重送的事件不重複記帳
  if (event.webhookEventId && isDuplicate_(event.webhookEventId)) return;

  // 記住群組，固定支出的提醒會推播到這裡
  if (event.source && event.source.type !== 'user') {
    var groupId = event.source.groupId || event.source.roomId;
    if (event.type === 'join' || !getProp_('NOTIFY_TARGET_ID', false)) {
      PropertiesService.getScriptProperties().setProperty('NOTIFY_TARGET_ID', groupId);
    }
  }

  if (event.type === 'join' || event.type === 'follow') {
    replyText_(event.replyToken, helpText_());
    return;
  }
  if (event.type !== 'message') return;

  var message = event.message;
  var isGroup = event.source.type !== 'user';
  var input;

  if (message.type === 'text') {
    var text = message.text.trim();
    if (text === '說明' || text === '幫助' || text.toLowerCase() === 'help') {
      replyText_(event.replyToken, helpText_());
      return;
    }
    var today0 = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    var recurringCmd = parseRecurringCommand_(text);
    if (recurringCmd) {
      replyText_(event.replyToken, withLock_(function () {
        return applyRecurringCommand_(recurringCmd, today0);
      }));
      return;
    }
    var recurringName = matchRecurringQuestion_(text);
    if (recurringName !== null) {
      replyText_(event.replyToken, formatRecurringList_(recurringName));
      return;
    }
    if (text === '刪除' || text === '取消') {
      replyText_(event.replyToken, handleDelete_(event.source.userId));
      return;
    }
    input = { text: text };
  } else if (message.type === 'image') {
    // 照片只能靠 AI 辨識；AI 忙線時請家人稍後再傳
    if (getProvider_() === 'rules') {
      // 沒有 AI 無法讀收據；群組裡家人分享照片很常見，只在私訊提示
      if (!isGroup) {
        replyText_(event.replyToken, '目前沒有開啟 AI，無法辨識收據照片。\n請用文字記帳，例如「全聯 560」。');
      }
      return;
    }
    input = getImageContent_(message.id);
  } else {
    return;
  }


  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var parsed = parseMessage_(input, today);
  parsed.entries = parsed.entries.filter(function (e) {
    return e.amount > 0;
  });

  if (parsed.intent === 'record' && parsed.entries.length > 0) {
    var recorder = getDisplayName_(event.source);
    withLock_(function () {
      appendEntries_(parsed.entries, {
        recorder: recorder,
        userId: event.source.userId,
        messageId: message.id,
        source: message.type === 'image' ? '收據照片' : '文字'
      });
    });
    replyText_(event.replyToken, formatRecorded_(parsed.entries, recorder, parsed.statedTotal));
    return;
  }

  if (parsed.intent === 'recurring') {
    replyText_(event.replyToken, withLock_(function () {
      return applyRecurringCommand_(parsed.recurring, today);
    }));
    return;
  }

  if (parsed.intent === 'analysis') {
    replyText_(event.replyToken, formatAnalysis_(analyze_(parsed.query.start_date, parsed.query.end_date, today)));
    return;
  }

  if (parsed.intent === 'query') {
    var q = parsed.query;
    replyText_(event.replyToken, formatSummary_(q, summarize_(q.start_date, q.end_date, q.category, q.keyword)));
    return;
  }

  // 群組裡的閒聊不回應，避免洗版；私訊則提示用法
  if (!isGroup) {
    var hint = getProvider_() === 'rules' || parsed.aiError ? '\n記帳請用「品項 金額」，例如「午餐 120」。' : '';
    if (parsed.aiError) {
      replyText_(event.replyToken, '⏳ AI 暫時忙線，這則看不出要記帳還是查詢。' + hint + '\n也可以稍後再傳一次。');
      return;
    }
    replyText_(event.replyToken, '看不出要記帳還是查詢 🤔' + hint + '\n傳「說明」可以看使用方式。');
  }
}

function handleDelete_(userId) {
  var deleted = withLock_(function () {
    return deleteLastEntry_(userId);
  });
  if (deleted.length === 0) {
    return '找不到你可以刪除的紀錄。';
  }
  var lines = deleted.map(function (r) {
    return '・' + r[COL.date] + ' ' + r[COL.category] + ' ' + r[COL.item] + ' $' + formatMoney_(r[COL.amount]);
  });
  return '🗑️ 已刪除：\n' + lines.join('\n');
}

function formatRecorded_(entries, recorder, statedTotal) {
  var total = 0;
  var lines = entries.map(function (e) {
    total += e.amount;
    var line = '・' + e.date + '｜' + e.category + '｜' + e.item + '｜$' + formatMoney_(e.amount);
    if (e.note) line += '（' + e.note + '）';
    return line;
  });
  var text = '✅ 已記帳（' + recorder + '）\n' + lines.join('\n');
  if (entries.length > 1) {
    text += '\n合計 $' + formatMoney_(total);
  }
  if (statedTotal) {
    text += statedTotal === total
      ? '（和你寫的總計相符）'
      : '\n⚠️ 你寫的總計是 $' + formatMoney_(statedTotal) + '，和明細加總 $' + formatMoney_(total) + ' 不同，請確認';
  }
  return text + '\n\n記錯了？傳「刪除」即可撤銷。';
}

function formatSummary_(q, s) {
  var title = '📊 ' + q.start_date + ' ～ ' + q.end_date;
  var filters = [];
  if (q.category !== '全部') filters.push(q.category);
  if (q.keyword) filters.push('「' + q.keyword + '」');
  if (filters.length) title += '（' + filters.join('・') + '）';

  if (s.count === 0) {
    return title + '\n這段期間沒有紀錄。';
  }

  var lines = [title, '總計 $' + formatMoney_(s.total) + '（' + s.count + ' 筆）'];

  if (q.category === '全部' && Object.keys(s.byCategory).length > 1) {
    lines.push('', '依分類：');
    CATEGORIES.forEach(function (c) {
      if (s.byCategory[c]) {
        var pct = Math.round((s.byCategory[c] / s.total) * 100);
        lines.push('・' + c + ' $' + formatMoney_(s.byCategory[c]) + '（' + pct + '%）');
      }
    });
  }

  var recorders = Object.keys(s.byRecorder);
  if (recorders.length > 1) {
    lines.push('', '依記錄人：');
    recorders.forEach(function (name) {
      lines.push('・' + name + ' $' + formatMoney_(s.byRecorder[name]));
    });
  }

  if (q.detail) {
    var MAX_DETAIL = 40;
    lines.push('', '明細：');
    s.rows.slice(-MAX_DETAIL).forEach(function (r) {
      lines.push('・' + String(r[COL.date]).slice(5) + ' ' + r[COL.item] + ' $' + formatMoney_(r[COL.amount]) + '（' + r[COL.recorder] + '）');
    });
    if (s.rows.length > MAX_DETAIL) {
      lines.push('（只列出最近 ' + MAX_DETAIL + ' 筆，共 ' + s.rows.length + ' 筆，完整明細請看試算表）');
    }
  } else if (s.count > 1) {
    lines.push('', '想看每一筆？在查詢後面加「明細」。');
  }
  return lines.join('\n');
}

function formatMoney_(n) {
  return Math.round(Number(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function isDuplicate_(eventId) {
  var cache = CacheService.getScriptCache();
  var key = 'evt:' + eventId;
  if (cache.get(key)) return true;
  cache.put(key, '1', 6 * 60 * 60);
  return false;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** 在編輯器手動執行一次，用來建立帳本工作表並觸發授權。 */
function setup() {
  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  getLedgerSheet_();
  ensureKeywordSheet_();
  ensureRecurringSheet_(today);
  ensureDailyTrigger_();
  console.log('帳本、關鍵字、固定支出工作表與每日排程已就緒');
}
