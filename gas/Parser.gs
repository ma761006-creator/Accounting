/**
 * 用 AI 把一則訊息（文字或收據照片）解析成結構化資料。
 *
 * 使用哪種解析方式由指令碼屬性決定：
 *   AI_PROVIDER = 'rules'、'gemini' 或 'claude'。
 *   沒設定時：有 GEMINI_API_KEY 用 Gemini，有 ANTHROPIC_API_KEY 用 Claude，都沒有就用免費的規則辨識（Rules.gs）。
 *
 * 回傳格式：
 *   {
 *     intent: 'record' | 'query' | 'recurring' | 'other',
 *     entries: [{ date, category, item, amount, note }],   // intent = record
 *     query: { start_date, end_date, category, keyword, detail }  // intent = query
 *     recurring: { action, name, amount, amountBlank, day, every, category }  // intent = recurring
 *   }
 */

/**
 * @param {Object} input  { text: string } 或 { imageBase64: string, mediaType: string }
 * @param {string} today  YYYY-MM-DD
 */
function parseMessage_(input, today) {
  var provider = getProvider_();
  var parsed;
  if (provider === 'gemini') {
    parsed = parseWithGemini_(input, today);
  } else if (provider === 'claude') {
    parsed = parseWithClaude_(input, today);
  } else {
    parsed = parseWithRules_(input.text, today);
  }
  return validateParsed_(parsed, today);
}

/**
 * AI 回傳的資料寫入試算表前再檢查一次：金額為正數、日期格式正確且不在未來、分類合法。
 * 不合格的記帳筆數直接丟掉，查詢欄位補上預設值。
 */
function validateParsed_(parsed, today) {
  parsed = parsed || {};
  var intent = ['record', 'query', 'recurring', 'other'].indexOf(parsed.intent) >= 0 ? parsed.intent : 'other';
  var isDate = function (s) {
    return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s + 'T00:00:00Z'));
  };

  var entries = (parsed.entries || []).map(function (e) {
    var amount = Math.round(Number(e && e.amount));
    var item = String((e && e.item) || '').trim();
    if (!(amount > 0) || !isFinite(amount) || !item) return null;
    return {
      date: isDate(e.date) && e.date <= today ? e.date : today,
      category: CATEGORIES.indexOf(e.category) >= 0 ? e.category : '其他',
      item: item.slice(0, 100),
      amount: amount,
      note: String(e.note || '').trim().slice(0, 200)
    };
  }).filter(function (e) {
    return e;
  });

  var q = parsed.query || {};
  var monthStart = today.slice(0, 8) + '01';
  var start = isDate(q.start_date) ? q.start_date : monthStart;
  var end = isDate(q.end_date) ? q.end_date : today;
  if (start > end) {
    var tmp = start;
    start = end;
    end = tmp;
  }
  var query = {
    start_date: start,
    end_date: end,
    category: CATEGORIES.indexOf(q.category) >= 0 ? q.category : '全部',
    keyword: String(q.keyword || '').trim().slice(0, 50),
    detail: q.detail === true
  };

  var r = parsed.recurring || {};
  var int = function (v, min, max) {
    var n = Math.round(Number(v));
    return n >= min && n <= max ? n : null;
  };
  var recurring = {
    action: ['upsert', 'delete', 'list'].indexOf(r.action) >= 0 ? r.action : 'list',
    name: String(r.name || '').trim().slice(0, 50),
    amount: int(r.amount, 1, 10000000),
    amountBlank: false,
    day: int(r.day, 1, 31),
    every: int(r.every, 1, 12),
    category: null
  };

  if (intent === 'record' && entries.length === 0) intent = 'other';
  if (intent === 'recurring' && recurring.action !== 'list' && !recurring.name) intent = 'other';
  return { intent: intent, entries: entries, query: query, recurring: recurring };
}

/**
 * 呼叫外部 API，遇到暫時性錯誤（429、5xx、連線失敗）時等一下再試一次。
 * LINE 的回覆權杖有時效，所以只重試一次。
 */
function fetchWithRetry_(url, options) {
  var res;
  try {
    res = UrlFetchApp.fetch(url, options);
  } catch (err) {
    Utilities.sleep(1500);
    return UrlFetchApp.fetch(url, options);
  }
  var code = res.getResponseCode();
  if (code === 429 || code >= 500) {
    Utilities.sleep(1500);
    return UrlFetchApp.fetch(url, options);
  }
  return res;
}

/** @return {'rules'|'gemini'|'claude'} */
function getProvider_() {
  var provider = getProp_('AI_PROVIDER', false);
  if (provider) return provider;
  if (getProp_('GEMINI_API_KEY', false)) return 'gemini';
  if (getProp_('ANTHROPIC_API_KEY', false)) return 'claude';
  return 'rules';
}

// 以函式回傳而不是全域變數：Apps Script 依檔案順序載入，避免 CATEGORIES 尚未定義
function getParseSchema_() {
  return {
    type: 'object',
    properties: {
      intent: { type: 'string', enum: ['record', 'query', 'recurring', 'other'] },
      entries: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            date: { type: 'string', description: 'YYYY-MM-DD' },
            category: { type: 'string', enum: CATEGORIES },
            item: { type: 'string' },
            amount: { type: 'number' },
            note: { type: 'string' }
          },
          required: ['date', 'category', 'item', 'amount', 'note'],
          additionalProperties: false
        }
      },
      recurring: {
        type: 'object',
        properties: {
          action: { type: 'string', enum: ['upsert', 'delete', 'list'] },
          name: { type: 'string' },
          amount: { type: 'number', description: '每期金額；沒提到或金額不固定填 0' },
          day: { type: 'number', description: '每月幾號扣款；沒提到填 0' },
          every: { type: 'number', description: '每幾個月一期；沒提到填 0' }
        },
        required: ['action', 'name', 'amount', 'day', 'every'],
        additionalProperties: false
      },
      query: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'YYYY-MM-DD' },
          end_date: { type: 'string', description: 'YYYY-MM-DD' },
          category: { type: 'string', enum: ['全部'].concat(CATEGORIES) },
          keyword: { type: 'string', description: '品項或店名關鍵字，沒有就空字串' },
          detail: { type: 'boolean', description: '是否要列出每一筆明細' }
        },
        required: ['start_date', 'end_date', 'category', 'keyword', 'detail'],
        additionalProperties: false
      }
    },
    required: ['intent', 'entries', 'query', 'recurring'],
    additionalProperties: false
  };
}

function buildSystemPrompt_(today) {
  return [
    '你是家庭記帳助理，負責把家人在 LINE 傳來的訊息轉成記帳資料。',
    '今天是 ' + today + '（台灣時間）。「昨天」「上週五」等相對日期請依此換算成 YYYY-MM-DD。',
    '',
    '分類只能是：' + CATEGORIES.join('、') + '。',
    '- 餐飲：三餐、飲料、零食、外送、買菜',
    '- 交通：油錢、停車、捷運、公車、高鐵、計程車、過路費、車輛保養',
    '- 日用品：清潔用品、衛生紙、盥洗用品、家用小物',
    '- 醫療：看診、掛號、藥品、保健食品、牙醫',
    '- 其他：不屬於以上分類的消費',
    '',
    '判斷 intent：',
    '- record：訊息在記錄花費（例如「午餐 120」「全聯 560 衛生紙」或收據照片）。',
    '  一則訊息可能有多筆，請逐筆列在 entries。金額一律為新台幣正整數。',
    '  沒提到日期就用今天。item 寫簡短品項或店名，note 放其他補充（沒有就空字串）。',
    '  收據照片：以實付總金額為準，一張收據通常記成一筆；若品項明顯分屬不同分類，可依分類拆成多筆，金額加總需等於實付金額。',
    '- query：訊息在問花費統計（例如「這個月花多少」「上個月交通費」）。',
    '  請填 query 的日期區間（含頭尾）與分類，沒指定分類就用「全部」。沒指定期間就用本月 1 日到今天。',
    '  問特定店家或品項（例如「全聯花多少」「這個月 7-11」）時，把店名或品項填在 keyword，分類用「全部」；否則 keyword 為空字串。',
    '  要求列出明細、清單、每一筆時 detail 為 true，否則為 false。',
    '- recurring：新增、修改、刪除或詢問固定支出／訂閱（例如「我每個月訂 Netflix 390，15 號扣款」「房租改成 16000」「取消 Netflix」「我有哪些訂閱」）。',
    '  action：新增或修改用 upsert，刪除或取消用 delete，詢問用 list。name 為項目名稱（例如 Netflix、房租）。',
    '  amount、day、every 沒提到就填 0；「兩個月一期」every 為 2。單次的消費請用 record，不是 recurring。',
    '- other：閒聊或與記帳無關的訊息。',
    '',
    '不適用的欄位：entries 填空陣列；query 填今天日期、「全部」、空字串 keyword 與 false；recurring 填 list、空字串與 0。'
  ].join('\n');
}

var RECEIPT_PROMPT = '這是一張收據或發票照片，請記帳。';
