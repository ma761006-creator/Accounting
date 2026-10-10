/**
 * LINE 家庭記帳機器人：Webhook 入口。
 *
 * 部署成「網頁應用程式」後，把網址填到 LINE Developers 的 Webhook URL。
 */

// 用函式而不是全域變數：Code.gs 會比 Config.gs 先載入，此時 CATEGORIES 還沒定義
function helpText_() {
  var ai = getProvider_() !== 'rules';
  return [
    '📒 家庭記帳機器人',
    '',
    '💰 記帳：' + (ai ? '直接說，例如「昨天全聯買衛生紙 560」' : '品項 金額，例如「午餐 120」「昨天 中油 1200」'),
    '・多筆：加油 1200、停車 60',
    ai ? '📷 發票辨識：直接傳發票或收據照片' : '📷 發票辨識：設定 Gemini 金鑰後可以傳照片',
    '',
    '🔎 查詢：本月、上月、本週、今天',
    '・本月 餐飲、全聯花多少、本月 明細',
    '📊 分析：分析、本週分析、上月分析',
    '🔥 洞察：分析裡會列出最大支出類別、最高單筆、最高單日',
    '',
    '✏️ 修改：修改 150、修改 交通、午餐改成150',
    '🗑️ 刪除：刪除（最近一筆）、刪除 午餐、刪除重複、恢復（復原上一次）',
    '',
    '🔔 訂閱／固定支出：固定支出、固定支出 Netflix 390 每月15號',
    '🏷️ 分類關鍵字：關鍵字 健身房 其他、關鍵字（看清單）',
    '⏰ 扣款提醒：近期扣款（另外每天早上會自動提醒）',
    '',
    '分類：' + CATEGORIES.join('、'),
    '',
    '👇 也可以直接點下面的按鈕'
  ].join('\n');
}

/**
 * 快速選單：每個按鈕送出的文字都是既有的固定指令，不需要另外的處理邏輯，也不會用到 AI 額度。
 */
function mainMenu_() {
  return [
    { label: '🔎 本月花費', text: '本月' },
    { label: '📋 本月明細', text: '本月 明細' },
    { label: '📊 本月分析', text: '分析' },
    { label: '🔁 固定支出', text: '固定支出' },
    { label: '⏰ 扣款提醒', text: '近期扣款' },
    { label: '💰 怎麼記帳', text: '記帳' },
    { label: '✏️ 修改刪除', text: '修改' }
  ];
}

function queryMenu_() {
  return [
    { label: '今天', text: '今天' },
    { label: '本週', text: '本週' },
    { label: '本月', text: '本月' },
    { label: '上月', text: '上月' },
    { label: '📋 本月明細', text: '本月 明細' },
    { label: '📊 分析', text: '分析' }
  ];
}

function analysisMenu_() {
  return [
    { label: '📆 本週分析', text: '本週分析' },
    { label: '🗓️ 本月分析', text: '分析' },
    { label: '⏪ 上月分析', text: '上月分析' },
    { label: '📋 本月明細', text: '本月 明細' }
  ];
}

function recordUsage_() {
  var ai = getProvider_() !== 'rules';
  var lines = [
    '💰 記帳方式',
    ai ? '直接用說的就可以，例如：' : '輸入「品項 金額」，例如：',
    '・午餐 120',
    '・昨天 中油 1200',
    '・加油 1200、停車 60（一次記多筆）'
  ];
  if (ai) lines.push('・也可以直接傳發票或收據照片');
  lines.push('', '記錯了傳「刪除」就能撤銷。');
  return lines.join('\n');
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
  if (/照片下載失敗/.test(msg)) {
    return '⚠️ 這張照片下載失敗，請重新拍照或從相簿重新傳一次（不要用轉傳的）。\n也可以先用文字記帳，例如「全聯 560」。\n\n' +
      '（' + msg.replace(/^.*照片下載失敗（|）$/g, '') + '）';
  }
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
    replyText_(event.replyToken, helpText_(), mainMenu_());
    return;
  }
  if (event.type !== 'message') return;

  var message = event.message;
  var isGroup = event.source.type !== 'user';
  var input;

  if (message.type === 'text') {
    var text = message.text.trim();
    if (/^(說明|幫助|help|選單|menu|功能)$/i.test(text)) {
      replyText_(event.replyToken, helpText_(), mainMenu_());
      return;
    }
    if (text === '記帳') {
      replyText_(event.replyToken, recordUsage_(), mainMenu_());
      return;
    }
    if (text === '查詢') {
      replyText_(event.replyToken, '🔎 要查哪段期間？點下面的按鈕，或直接輸入，例如「全聯花多少」。', queryMenu_());
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
    if (matchDueSoonQuestion_(text)) {
      replyText_(event.replyToken, formatDueSoon_(today0, 14), [{ label: '🔁 全部固定支出', text: '固定支出' }]);
      return;
    }
    var recurringName = matchRecurringQuestion_(text);
    if (recurringName !== null) {
      replyText_(event.replyToken, formatRecurringList_(recurringName));
      return;
    }
    var keywordCmd = parseKeywordCommand_(text);
    if (keywordCmd) {
      replyText_(event.replyToken, withLock_(function () {
        return applyKeywordCommand_(keywordCmd);
      }));
      return;
    }
    var modifyCmd = parseModifyCommand_(text, today0);
    if (modifyCmd) {
      replyText_(event.replyToken, withLock_(function () {
        return applyModifyCommand_(modifyCmd, event.source.userId);
      }));
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
    input = getImageContent_(message);
  } else {
    return;
  }


  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var parsed = parseMessage_(input, today);
  parsed.entries = parsed.entries.filter(function (e) {
    return e.amount > 0;
  });

  // 照片：同一張發票不會有兩筆一模一樣的帳，也不會另外多一筆「總計」
  if (message.type === 'image') parsed.entries = dedupeReceiptEntries_(parsed.entries);

  // AI 判斷的分類以家人設定的關鍵字為準（規則辨識本來就會用）
  if (getProvider_() !== 'rules' && !parsed.aiError) parsed.entries = applyCustomKeywords_(parsed.entries);

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

  if (parsed.intent === 'modify') {
    replyText_(event.replyToken, withLock_(function () {
      return applyModifyCommand_(parsed.modify, event.source.userId);
    }));
    return;
  }

  if (parsed.intent === 'analysis') {
    replyText_(event.replyToken, formatAnalysis_(analyze_(parsed.query.start_date, parsed.query.end_date, today)), analysisMenu_());
    return;
  }

  if (parsed.intent === 'query') {
    var q = parsed.query;
    replyText_(event.replyToken, formatSummary_(q, summarize_(q.start_date, q.end_date, q.category, q.keyword)), queryMenu_());
    return;
  }

  // 「你是誰」「你會什麼」：簡短自我介紹
  if (message.type === 'text' && /你是誰|你叫什麼|你會什麼|你可以做什麼|你能做什麼|自我介紹/.test(message.text)) {
    replyText_(event.replyToken, '我是家庭記帳機器人 📒\n' +
      '幫全家記帳、查詢與分析花費，也會管理房租、水電這類固定支出，扣款前一天提醒。\n\n' +
      '試試傳「午餐 120」，或點下面的按鈕。', mainMenu_());
    return;
  }

  // 群組裡的閒聊不回應，避免洗版；私訊則提示用法
  if (!isGroup) {
    var hint = getProvider_() === 'rules' || parsed.aiError ? '\n記帳請用「品項 金額」，例如「午餐 120」。' : '';
    if (parsed.aiError) {
      replyText_(event.replyToken, '⏳ AI 暫時忙線，這則看不出要記帳還是查詢。' + hint + '\n也可以稍後再傳一次。', mainMenu_());
      return;
    }
    replyText_(event.replyToken, '看不出要記帳還是查詢 🤔' + hint + '\n點下面的按鈕，或傳「說明」看使用方式。', mainMenu_());
  }
}

function dedupeReceiptEntries_(entries) {
  var seen = {};
  var unique = entries.filter(function (e) {
    var key = e.item + '|' + e.amount;
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
  if (unique.length < 3) {
    // 兩筆金額相同（例如「健身房月費 1500」和「總計 1500」）→ 只留第一筆
    if (unique.length === 2 && unique[0].amount === unique[1].amount) return [unique[0]];
    return unique;
  }
  // 有一筆剛好等於其他筆的加總 → 那是總計，不另外記
  var total = unique.reduce(function (sum, e) {
    return sum + e.amount;
  }, 0);
  var totalLine = unique.filter(function (e) {
    return e.amount * 2 === total;
  })[0];
  return totalLine ? unique.filter(function (e) {
    return e !== totalLine;
  }) : unique;
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
  var unspecified = entries.some(function (e) {
    return e.item === UNSPECIFIED_ITEM;
  });
  if (unspecified) {
    return text + '\n\n👉 這筆用在哪裡？傳「修改 品項 午餐」補上，或「修改 餐飲」改分類。';
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
  if (s.count > 1) {
    var top = s.rows.reduce(function (a, r) {
      return !a || Number(r[COL.amount]) > Number(a[COL.amount]) ? r : a;
    }, null);
    lines.push('🏆 最高單筆：' + top[COL.item] + ' $' + formatMoney_(top[COL.amount]) + '（' + String(top[COL.date]).slice(5) + '）');
  }

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

  // 筆數少就直接列出明細，不用另外打「明細」
  var AUTO_DETAIL = 10;
  if (q.detail || s.count <= AUTO_DETAIL) {
    var MAX_DETAIL = 40;
    lines.push('', '明細：');
    s.rows.slice(-MAX_DETAIL).forEach(function (r) {
      lines.push('・' + String(r[COL.date]).slice(5) + ' ' + r[COL.item] + ' $' + formatMoney_(r[COL.amount]) + '（' + r[COL.recorder] + '）');
    });
    if (s.rows.length > MAX_DETAIL) {
      lines.push('（只列出最近 ' + MAX_DETAIL + ' 筆，共 ' + s.rows.length + ' 筆，完整明細請看試算表）');
    }
  } else {
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
