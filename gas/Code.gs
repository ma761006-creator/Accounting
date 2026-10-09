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
      '・可加分類，例如「本月 餐飲」',
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
        replyText_(event.replyToken, '⚠️ 處理失敗，請稍後再試一次。\n（' + String(err.message || err).slice(0, 200) + '）');
      }
    }
  });
  return ok_();
}

function ok_() {
  return ContentService.createTextOutput('OK');
}

function handleEvent_(event) {
  // LINE 重送的事件不重複記帳
  if (event.webhookEventId && isDuplicate_(event.webhookEventId)) return;

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
    if (text === '刪除' || text === '取消') {
      replyText_(event.replyToken, handleDelete_(event.source.userId));
      return;
    }
    input = { text: text };
  } else if (message.type === 'image') {
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
    replyText_(event.replyToken, formatRecorded_(parsed.entries, recorder));
    return;
  }

  if (parsed.intent === 'query') {
    var q = parsed.query;
    replyText_(event.replyToken, formatSummary_(q, summarize_(q.start_date, q.end_date, q.category)));
    return;
  }

  // 群組裡的閒聊不回應，避免洗版；私訊則提示用法
  if (!isGroup) {
    var hint = getProvider_() === 'rules' ? '\n記帳請用「品項 金額」，例如「午餐 120」。' : '';
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

function formatRecorded_(entries, recorder) {
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
  return text + '\n\n記錯了？傳「刪除」即可撤銷。';
}

function formatSummary_(q, s) {
  var title = '📊 ' + q.start_date + ' ～ ' + q.end_date;
  if (q.category !== '全部') title += '（' + q.category + '）';

  if (s.count === 0) {
    return title + '\n這段期間沒有紀錄。';
  }

  var lines = [title, '總計 $' + formatMoney_(s.total) + '（' + s.count + ' 筆）'];

  if (q.category === '全部') {
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
  getLedgerSheet_();
  ensureKeywordSheet_();
  console.log('帳本、關鍵字工作表已就緒');
}
