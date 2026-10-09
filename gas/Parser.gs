/**
 * 用 AI 把一則訊息（文字或收據照片）解析成結構化資料。
 *
 * 使用哪種解析方式由指令碼屬性決定：
 *   AI_PROVIDER = 'rules'、'gemini' 或 'claude'。
 *   沒設定時：有 GEMINI_API_KEY 用 Gemini，有 ANTHROPIC_API_KEY 用 Claude，都沒有就用免費的規則辨識（Rules.gs）。
 *
 * 回傳格式：
 *   {
 *     intent: 'record' | 'query' | 'analysis' | 'recurring' | 'modify' | 'other',
 *     entries: [{ date, category, item, amount, note }],   // intent = record
 *     query: { start_date, end_date, category, keyword, detail }  // intent = query 或 analysis
 *     recurring: { action, name, amount, amountBlank, day, every, category }  // intent = recurring
 *     modify: { action, keyword, date, amount, newAmount, newCategory, newItem }  // intent = modify
 *   }
 */

/**
 * @param {Object} input  { text: string } 或 { imageBase64: string, mediaType: string }
 * @param {string} today  YYYY-MM-DD
 */
function parseMessage_(input, today) {
  var provider = getProvider_();
  var parsed;
  var aiError = '';
  // 按鈕和固定指令（本月、本週分析、本月 明細…）不需要 AI：省額度，AI 忙線時也照常運作
  if (input.text && isFixedCommand_(input.text)) provider = 'rules';
  try {
    if (provider === 'gemini') {
      parsed = parseWithGemini_(input, today);
    } else if (provider === 'claude') {
      parsed = parseWithClaude_(input, today);
    } else {
      parsed = parseWithRules_(input.text, today);
    }
  } catch (err) {
    // AI 忙線或出錯時，文字訊息改用免費的規則辨識，「午餐 120」這類格式照樣能記帳
    if (provider === 'rules' || !input.text) throw err;
    console.warn('AI 解析失敗，改用規則辨識：' + (err.message || err));
    aiError = String(err.message || err);
    parsed = parseWithRules_(input.text, today);
  }
  var result = validateParsed_(parsed, today);
  result.aiError = aiError;
  return result;
}

function isFixedCommand_(text) {
  var t = normalizeText_(text).replace(/\s+/g, ' ');
  var period = '(今天|昨天|本週|這週|上週|本月|這個月|上月|上個月|今年)';
  return new RegExp('^' + period + '( ?明細)?$').test(t) ||
    new RegExp('^' + period + '? ?分析$').test(t);
}

/**
 * AI 回傳的資料寫入試算表前再檢查一次：金額為正數、日期格式正確且不在未來、分類合法。
 * 不合格的記帳筆數直接丟掉，查詢欄位補上預設值。
 */
function validateParsed_(parsed, today) {
  parsed = parsed || {};
  var intent = ['record', 'query', 'analysis', 'recurring', 'modify', 'other'].indexOf(parsed.intent) >= 0 ? parsed.intent : 'other';
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

  var mo = parsed.modify || {};
  var modify = {
    action: mo.action === 'delete' ? 'delete' : 'edit',
    keyword: String(mo.keyword || '').trim().slice(0, 50),
    date: isDate(mo.date) ? mo.date : '',
    amount: int(mo.amount, 1, 10000000),
    newAmount: int(mo.new_amount, 1, 10000000),
    newCategory: CATEGORIES.indexOf(mo.new_category) >= 0 ? mo.new_category : '',
    newItem: String(mo.new_item || '').trim().slice(0, 100)
  };
  if (intent === 'modify' && modify.action === 'edit' && !modify.newAmount && !modify.newCategory && !modify.newItem) {
    intent = 'other';
  }

  if (intent === 'record' && entries.length === 0) intent = 'other';
  if (intent === 'recurring' && recurring.action !== 'list' && !recurring.name) intent = 'other';
  var statedTotal = Math.round(Number(parsed.statedTotal));
  return {
    intent: intent, entries: entries, query: query, recurring: recurring, modify: modify,
    statedTotal: statedTotal > 0 ? statedTotal : null
  };
}

/**
 * 呼叫外部 API，遇到暫時性錯誤（429、5xx、連線失敗）時等一下再試，最多重試兩次。
 * LINE 的回覆權杖有時效，所以總等待時間控制在幾秒內。
 */
function fetchWithRetry_(url, options) {
  var waits = [1000, 3000];
  for (var attempt = 0; ; attempt++) {
    var res = null;
    try {
      res = UrlFetchApp.fetch(url, options);
    } catch (err) {
      if (attempt >= waits.length) throw err;
    }
    if (res) {
      var code = res.getResponseCode();
      if ((code !== 429 && code < 500) || attempt >= waits.length) return res;
    }
    Utilities.sleep(waits[attempt]);
  }
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
      intent: { type: 'string', enum: ['record', 'query', 'analysis', 'recurring', 'modify', 'other'] },
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
      modify: {
        type: 'object',
        properties: {
          action: { type: 'string', enum: ['edit', 'delete'] },
          keyword: { type: 'string', description: '要找的品項關鍵字；指最近一筆就空字串' },
          date: { type: 'string', description: 'YYYY-MM-DD；沒指定日期就空字串' },
          amount: { type: 'number', description: '原本的金額，用來找出那一筆；沒提到填 0' },
          new_amount: { type: 'number', description: '改成的金額；不改填 0' },
          new_category: { type: 'string', enum: [''].concat(CATEGORIES) },
          new_item: { type: 'string', description: '改成的品項名稱；不改就空字串' }
        },
        required: ['action', 'keyword', 'date', 'amount', 'new_amount', 'new_category', 'new_item'],
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
    required: ['intent', 'entries', 'query', 'recurring', 'modify'],
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
    '- 育兒：尿布、奶粉、副食品、嬰幼兒用品、童裝、玩具、托嬰、幼兒園、安親班、才藝課、小孩的學費與疫苗以外的花費',
    '- 其他：不屬於以上分類的消費',
    '',
    '判斷 intent：',
    '- record：訊息在記錄花費（例如「午餐 120」「全聯 560 衛生紙」或收據照片）。',
    '  一則訊息可能有多筆，請逐筆列在 entries。金額一律為新台幣正整數。',
    '  沒提到日期就用今天。item 寫簡短品項或店名，note 放其他補充（沒有就空字串）。',
    '  item、note 一律用繁體中文（例如 Lunch 寫「午餐」、简体字轉成繁體）；品牌與店名照原樣（例如 Netflix、7-11、全聯）。',
    '  訊息沒說用在哪裡（例如「我今天花了120元」）時不要猜，item 填「未說明」、category 填「其他」。',
    '  「金額 1125」「合計 1125」「總共 1125」這類總計行不是另一筆消費，不要記成 entries。',
    '  收據照片：以實付總金額為準，一張收據通常記成一筆；若品項明顯分屬不同分類，可依分類拆成多筆，金額加總需等於實付金額。',
    '  發票上的品項明細、小計、合計、應付、實付、信用卡簽單常出現同一個金額，只能記一次，不要重複記成多筆。',
    '- query：訊息在問花費統計（例如「這個月花多少」「上個月交通費」）。',
    '  請填 query 的日期區間（含頭尾）與分類，沒指定分類就用「全部」。沒指定期間就用本月 1 日到今天。',
    '  問特定店家或品項（例如「全聯花多少」「這個月 7-11」）時，把店名或品項填在 keyword，分類用「全部」；否則 keyword 為空字串。',
    '  要求列出明細、清單、每一筆時 detail 為 true，否則為 false。',
    '- analysis：要求分析消費、看花費趨勢或和上個月比較（例如「幫我分析這個月的消費」「上個月花得比較多嗎」「本週分析」）。',
    '  把要分析的期間填在 query 的 start_date、end_date（沒指定就用本月 1 日到今天），其他 query 欄位填預設值。',
    '- recurring：新增、修改、刪除或詢問固定支出／訂閱（例如「我每個月訂 Netflix 390，15 號扣款」「房租改成 16000」「取消 Netflix」「我有哪些訂閱」）。',
    '  action：新增或修改用 upsert，刪除或取消用 delete，詢問用 list。name 為項目名稱（例如 Netflix、房租）。',
    '  amount、day、every 沒提到就填 0；「兩個月一期」every 為 2。單次的消費請用 record，不是 recurring。',
    '- modify：修改或刪除已經記過的帳（例如「剛剛的午餐其實是150」「把昨天的停車費刪掉」「鯖魚那筆改成餐飲」）。',
    '  action 為 edit 或 delete；keyword、date、amount 用來找出那一筆（指最近一筆就留空），new_* 填要改成的值，不改的留空或 0。',
    '- other：閒聊或與記帳無關的訊息。',
    '',
    '不適用的欄位：entries 填空陣列；query 填今天日期、「全部」、空字串 keyword 與 false；recurring 填 list、空字串與 0；modify 填 edit、空字串與 0。'
  ].join('\n');
}

var RECEIPT_PROMPT = '這是一張收據或發票照片，請記帳。';
