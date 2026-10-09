// 自動產生，請勿直接修改。原始檔在 gas/，修改後執行 node scripts/bundle.js

// ===== Config.gs =====

/**
 * 設定檔：分類、模型、試算表欄位。
 *
 * 金鑰不要寫在這裡，請放在「專案設定 → 指令碼屬性」：
 *   LINE_CHANNEL_ACCESS_TOKEN  LINE Messaging API 的 Channel access token
 *   GEMINI_API_KEY             （選填）Gemini API 金鑰，設定後改用 AI 解析
 *   ANTHROPIC_API_KEY          （選填）Claude API 金鑰，設定後改用 AI 解析
 *   （兩個都沒設定時，使用免費的規則辨識，見 Rules.gs）
 *   SPREADSHEET_ID             （選填）帳本試算表 ID；若程式是從試算表「擴充功能」建立的可省略
 *   LINE_BOT_USER_ID           （選填）機器人的 userId（U 開頭），設定後只接受送給這個機器人的事件
 */

// 記帳分類。「其他」用來接住不屬於前四類的消費，不需要可以刪掉。
var CATEGORIES = ['餐飲', '交通', '日用品', '醫療', '其他'];

var CLAUDE_MODEL = 'claude-haiku-5-5';

var SHEET_NAME = '帳本';

var HEADERS = ['記錄時間', '消費日期', '分類', '品項', '金額', '記錄人', '備註', '來源', 'LINE userId', '訊息ID'];

// HEADERS 中各欄的位置（從 0 開始）
var COL = {
  createdAt: 0,
  date: 1,
  category: 2,
  item: 3,
  amount: 4,
  recorder: 5,
  note: 6,
  source: 7,
  userId: 8,
  messageId: 9
};

var TIMEZONE = 'Asia/Taipei';

function getProp_(key, required) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (required && !value) {
    throw new Error('缺少指令碼屬性：' + key);
  }
  return value;
}

// ===== Sheet.gs =====

/**
 * Google 試算表讀寫。
 */

function getSpreadsheet_() {
  var id = getProp_('SPREADSHEET_ID', false);
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('找不到試算表，請設定指令碼屬性 SPREADSHEET_ID');
  }
  return ss;
}

function getLedgerSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * @param {Array} entries  Claude 解析出的 entries
 * @param {Object} meta    { recorder, userId, messageId, source }
 */
function appendEntries(entries, meta) {
  var sheet = getLedgerSheet_();
  var now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  entries.forEach(function (e) {
    var row = [];
    row[COL.createdAt] = now;
    row[COL.date] = e.date;
    row[COL.category] = e.category;
    row[COL.item] = e.item;
    row[COL.amount] = e.amount;
    row[COL.recorder] = meta.recorder;
    row[COL.note] = e.note;
    row[COL.source] = meta.source;
    row[COL.userId] = meta.userId;
    // 加 ' 讓試算表當文字存，避免 18 位數的訊息 ID 被轉成數字而失去精度
    row[COL.messageId] = "'" + meta.messageId;
    sheet.appendRow(row);
  });
}

/** 讀出所有資料列（不含標題），日期統一轉成 YYYY-MM-DD 字串。 */
function readRows_() {
  var sheet = getLedgerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values.map(function (r) {
    var d = r[COL.date];
    if (d instanceof Date) {
      r[COL.date] = Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd');
    } else {
      r[COL.date] = String(d);
    }
    return r;
  });
}

/**
 * 統計區間內的花費。
 * @return {{ total: number, count: number, byCategory: Object, byRecorder: Object }}
 */
function summarize(startDate, endDate, category) {
  var result = { total: 0, count: 0, byCategory: {}, byRecorder: {} };
  readRows_().forEach(function (r) {
    var date = r[COL.date];
    if (date < startDate || date > endDate) return;
    if (category !== '全部' && r[COL.category] !== category) return;
    var amount = Number(r[COL.amount]) || 0;
    result.total += amount;
    result.count += 1;
    result.byCategory[r[COL.category]] = (result.byCategory[r[COL.category]] || 0) + amount;
    result.byRecorder[r[COL.recorder]] = (result.byRecorder[r[COL.recorder]] || 0) + amount;
  });
  return result;
}

/**
 * 刪除這位使用者最近一次記帳（同一則訊息記下的多筆會一起刪除）。
 * @return {Array} 被刪除的資料列；沒有可刪的則回傳空陣列
 */
function deleteLastEntry(userId) {
  var sheet = getLedgerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  var targetMessageId = null;
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][COL.userId] === userId) {
      targetMessageId = values[i][COL.messageId];
      break;
    }
  }
  if (targetMessageId === null) return [];

  var deleted = [];
  // 由下往上刪，列號才不會跑掉
  for (var j = values.length - 1; j >= 0; j--) {
    if (values[j][COL.userId] === userId && values[j][COL.messageId] === targetMessageId) {
      deleted.unshift(values[j]);
      sheet.deleteRow(j + 2);
    }
  }
  return deleted;
}

// ===== Line.gs =====

/**
 * LINE Messaging API 相關函式。
 */

function lineFetch_(url, options) {
  options = options || {};
  options.headers = options.headers || {};
  options.headers.Authorization = 'Bearer ' + getProp_('LINE_CHANNEL_ACCESS_TOKEN', true);
  options.muteHttpExceptions = true;
  return UrlFetchApp.fetch(url, options);
}

function replyText(replyToken, text) {
  var res = lineFetch_('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      replyToken: replyToken,
      // LINE 單則文字上限 5000 字
      messages: [{ type: 'text', text: text.slice(0, 5000) }]
    })
  });
  if (res.getResponseCode() !== 200) {
    console.error('LINE 回覆失敗 ' + res.getResponseCode() + '：' + res.getContentText());
  }
}

/** 取得使用者顯示名稱，快取 6 小時。 */
function getDisplayName(source) {
  var userId = source.userId;
  if (!userId) return '（未知）';

  var cache = CacheService.getScriptCache();
  var cacheKey = 'name:' + userId;
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var url;
  if (source.type === 'group') {
    url = 'https://api.line.me/v2/bot/group/' + source.groupId + '/member/' + userId;
  } else if (source.type === 'room') {
    url = 'https://api.line.me/v2/bot/room/' + source.roomId + '/member/' + userId;
  } else {
    url = 'https://api.line.me/v2/bot/profile/' + userId;
  }

  var res = lineFetch_(url, { method: 'get' });
  if (res.getResponseCode() !== 200) {
    return '（未知）';
  }
  var name = JSON.parse(res.getContentText()).displayName || '（未知）';
  cache.put(cacheKey, name, 6 * 60 * 60);
  return name;
}

/** 下載使用者傳來的圖片，回傳 { imageBase64, mediaType }。 */
function getImageContent(messageId) {
  var res = lineFetch_('https://api-data.line.me/v2/bot/message/' + messageId + '/content', {
    method: 'get'
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('下載圖片失敗 ' + res.getResponseCode());
  }
  var blob = res.getBlob();
  return {
    imageBase64: Utilities.base64Encode(blob.getBytes()),
    mediaType: blob.getContentType() || 'image/jpeg'
  };
}

// ===== Rules.gs =====

/**
 * 不用 AI 的規則辨識：沒有設定任何 AI 金鑰時使用，完全免費。
 *
 * 記帳格式：[日期] 品項 金額 [分類]，多筆用「、」「，」或換行隔開
 *   午餐 120 / 昨天 加油 1200 / 10/8 全聯 560 / 掛號 150 醫療 / 加油 1200、停車 60
 * 查詢格式：今天、本週、本月、上月、今年，可加分類，例如「本月 餐飲」
 *
 * 分類依關鍵字判斷，可在試算表的「關鍵字」工作表自行新增（優先於內建關鍵字）。
 */

var KEYWORD_SHEET_NAME = '關鍵字';

// 以函式回傳而不是全域變數：Apps Script 依檔案順序載入
function getDefaultKeywords_() {
  return {
    '餐飲': [
      '早餐', '午餐', '晚餐', '宵夜', '早午餐', '點心', '零食', '飲料', '咖啡', '手搖', '珍奶', '茶',
      '便當', '麵', '飯', '火鍋', '水果', '買菜', '菜市場', '外送', 'ubereats', 'foodpanda',
      '麥當勞', '肯德基', '摩斯', '星巴克',
      '7-11', '711', '7－11', '小七', '統一超商', '全家', '萊爾富'
    ],
    '交通': [
      '中油', '台塑石油', '加油', '油錢', '停車', '捷運', '公車', '客運', '高鐵', '台鐵', '火車',
      '計程車', '小黃', 'uber', '過路費', 'etag', '悠遊卡', '一卡通', '保養', '洗車'
    ],
    '日用品': [
      '全聯', '家樂福', '好市多', 'costco', '屈臣氏', '康是美', '寶雅', '大創',
      '衛生紙', '洗衣精', '洗碗精', '牙膏', '牙刷', '洗髮精', '沐浴乳', '垃圾袋', '清潔', '電池'
    ],
    '醫療': [
      '掛號', '看診', '診所', '醫院', '藥局', '藥', '牙醫', '牙科', '保健', '維他命', '眼科', '復健', '疫苗'
    ]
  };
}

function parseWithRules(text, today) {
  var t = normalizeText_(text);
  var query = parseQuery_(t, today);
  if (query) {
    return { intent: 'query', entries: [], query: query };
  }

  var keywords = loadKeywords_();
  var segments = t.split(/[、，,；;\n]+/).map(function (s) {
    return s.trim();
  }).filter(function (s) {
    return s;
  });

  var entries = [];
  for (var i = 0; i < segments.length; i++) {
    var entry = parseEntry_(segments[i], today, keywords);
    // 有任何一段看不懂就當作不是記帳，避免把聊天內容誤記
    if (!entry) {
      return { intent: 'other', entries: [], query: emptyQuery_(today) };
    }
    entries.push(entry);
  }
  if (entries.length === 0) {
    return { intent: 'other', entries: [], query: emptyQuery_(today) };
  }
  return { intent: 'record', entries: entries, query: emptyQuery_(today) };
}

function emptyQuery_(today) {
  return { start_date: today, end_date: today, category: '全部' };
}

/** 全形數字與符號轉半形、去掉千分位與貨幣符號。 */
function normalizeText_(text) {
  return String(text)
    .replace(/[０-９]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[／]/g, '/')
    .replace(/[－]/g, '-')
    .replace(/[．]/g, '.')
    .replace(/　/g, ' ')
    .replace(/(\d),(?=\d{3}(\D|$))/g, '$1')
    .replace(/NT\$|\$/gi, ' ')
    .trim();
}

/** 解析一筆「[日期] 品項 金額 [分類]」，不符合格式回傳 null。 */
function parseEntry_(segment, today, keywords) {
  var s = segment;

  var date = today;
  var dateMatch = matchDatePrefix_(s, today);
  if (dateMatch) {
    date = dateMatch.date;
    s = s.slice(dateMatch.length).trim();
  }

  var category = null;
  var tokens = s.split(/\s+/);
  if (tokens.length > 1 && CATEGORIES.indexOf(tokens[tokens.length - 1]) >= 0) {
    category = tokens.pop();
  } else if (tokens.length > 1 && CATEGORIES.indexOf(tokens[0]) >= 0) {
    category = tokens.shift();
  }
  s = tokens.join(' ');

  // 金額必須在最後，例如「午餐 120」「午餐120元」
  var m = s.match(/^(.*?)\s*(\d+(?:\.\d+)?)\s*(元|塊錢|塊)?$/);
  if (!m) return null;
  var item = m[1].trim();
  var amount = Math.round(Number(m[2]));
  if (!item || !(amount > 0)) return null;

  return {
    date: date,
    category: category || guessCategory_(item, keywords),
    item: item,
    amount: amount,
    note: ''
  };
}

/** 開頭的日期：今天、昨天、前天、10/8、10月8日、2026-10-08。 */
function matchDatePrefix_(s, today) {
  var relative = { '今天': 0, '昨天': -1, '前天': -2 };
  for (var word in relative) {
    if (s.indexOf(word) === 0) {
      return { date: addDays_(today, relative[word]), length: word.length };
    }
  }

  var full = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?=\s|[^\d]|$)/);
  if (full) {
    return { date: ymd_(+full[1], +full[2], +full[3]), length: full[0].length };
  }

  var md = s.match(/^(\d{1,2})(?:\/|月)(\d{1,2})日?(?=\s|[^\d]|$)/);
  if (md) {
    var year = +today.slice(0, 4);
    var date = ymd_(year, +md[1], +md[2]);
    // 比今天晚的日期當作去年
    if (date > today) date = ymd_(year - 1, +md[1], +md[2]);
    return { date: date, length: md[0].length };
  }
  return null;
}

/** 查詢：今天、昨天、本週、本月、上月、今年、查詢，可加分類，例如「本月 餐飲」「餐飲花多少」。 */
function parseQuery_(t, today) {
  var plain = t.replace(/[?？!！。]/g, '').trim();
  var s = plain.replace(/(花了?多少錢?|多少錢?|花費|統計|支出)$/, '').trim();
  var askedHowMuch = s !== plain;
  if (/\d/.test(s)) return null;

  var category = '全部';
  var rest = s.split(/\s+/).filter(function (p) {
    if (CATEGORIES.indexOf(p) >= 0) {
      category = p;
      return false;
    }
    return true;
  }).join('');
  // 也接受「本月餐飲」這種沒有空白的寫法
  CATEGORIES.forEach(function (c) {
    if (category === '全部' && rest.length > c.length && rest.slice(-c.length) === c) {
      category = c;
      rest = rest.slice(0, -c.length);
    }
  });

  var y = +today.slice(0, 4);
  var mo = +today.slice(5, 7);
  var monthStart = ymd_(y, mo, 1);
  var lastMonth = [ymd_(mo === 1 ? y - 1 : y, mo === 1 ? 12 : mo - 1, 1), addDays_(monthStart, -1)];
  var periods = {
    '今天': [today, today],
    '昨天': [addDays_(today, -1), addDays_(today, -1)],
    '本週': [startOfWeek_(today), today],
    '這週': [startOfWeek_(today), today],
    '本月': [monthStart, today],
    '這個月': [monthStart, today],
    '上月': lastMonth,
    '上個月': lastMonth,
    '今年': [ymd_(y, 1, 1), today],
    '查詢': [monthStart, today]
  };

  var range;
  if (periods.hasOwnProperty(rest)) {
    range = periods[rest];
  } else if (rest === '' && (askedHowMuch || category !== '全部')) {
    // 「花多少」「餐飲」「醫療花多少」：預設查本月
    range = [monthStart, today];
  } else {
    return null;
  }
  return { start_date: range[0], end_date: range[1], category: category };
}

function guessCategory_(item, keywords) {
  var lower = item.toLowerCase();
  for (var i = 0; i < keywords.length; i++) {
    if (lower.indexOf(keywords[i].word) >= 0) return keywords[i].category;
  }
  return '其他';
}

/** 合併試算表「關鍵字」工作表與內建關鍵字，自訂的優先，同來源中長的關鍵字優先。 */
function loadKeywords_() {
  var custom = readCustomKeywords_();
  var builtIn = [];
  var defaults = getDefaultKeywords_();
  Object.keys(defaults).forEach(function (category) {
    defaults[category].forEach(function (word) {
      builtIn.push({ word: word.toLowerCase(), category: category });
    });
  });
  var byLength = function (a, b) {
    return b.word.length - a.word.length;
  };
  return custom.sort(byLength).concat(builtIn.sort(byLength));
}

function readCustomKeywords_() {
  var sheet = getSpreadsheet_().getSheetByName(KEYWORD_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .filter(function (r) {
      return String(r[0]).trim() && CATEGORIES.indexOf(String(r[1]).trim()) >= 0;
    })
    .map(function (r) {
      return { word: String(r[0]).trim().toLowerCase(), category: String(r[1]).trim() };
    });
}

/** 建立「關鍵字」工作表（setup 時呼叫）。 */
function ensureKeywordSheet_() {
  var ss = getSpreadsheet_();
  if (ss.getSheetByName(KEYWORD_SHEET_NAME)) return;
  var sheet = ss.insertSheet(KEYWORD_SHEET_NAME);
  sheet.appendRow(['關鍵字', '分類']);
  sheet.appendRow(['（例）菜市場', '餐飲']);
  sheet.setFrozenRows(1);
}

function ymd_(y, m, d) {
  var date = new Date(Date.UTC(y, m - 1, d));
  return date.toISOString().slice(0, 10);
}

function addDays_(ymd, n) {
  var date = new Date(ymd + 'T00:00:00Z');
  date.setUTCDate(date.getUTCDate() + n);
  return date.toISOString().slice(0, 10);
}

/** 本週一。 */
function startOfWeek_(ymd) {
  var day = new Date(ymd + 'T00:00:00Z').getUTCDay(); // 0 = 週日
  return addDays_(ymd, day === 0 ? -6 : 1 - day);
}

// ===== Parser.gs =====

/**
 * 用 AI 把一則訊息（文字或收據照片）解析成結構化資料。
 *
 * 使用哪種解析方式由指令碼屬性決定：
 *   AI_PROVIDER = 'rules'、'gemini' 或 'claude'。
 *   沒設定時：有 GEMINI_API_KEY 用 Gemini，有 ANTHROPIC_API_KEY 用 Claude，都沒有就用免費的規則辨識（Rules.gs）。
 *
 * 回傳格式：
 *   {
 *     intent: 'record' | 'query' | 'other',
 *     entries: [{ date, category, item, amount, note }],   // intent = record
 *     query: { start_date, end_date, category }            // intent = query
 *   }
 */

/**
 * @param {Object} input  { text: string } 或 { imageBase64: string, mediaType: string }
 * @param {string} today  YYYY-MM-DD
 */
function parseMessage(input, today) {
  var provider = getProvider_();
  if (provider === 'gemini') {
    return parseWithGemini(input, today);
  }
  if (provider === 'claude') {
    return parseWithClaude(input, today);
  }
  return parseWithRules(input.text, today);
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
      intent: { type: 'string', enum: ['record', 'query', 'other'] },
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
      query: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'YYYY-MM-DD' },
          end_date: { type: 'string', description: 'YYYY-MM-DD' },
          category: { type: 'string', enum: ['全部'].concat(CATEGORIES) }
        },
        required: ['start_date', 'end_date', 'category'],
        additionalProperties: false
      }
    },
    required: ['intent', 'entries', 'query'],
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
    '- other：閒聊或與記帳無關的訊息。',
    '',
    '不適用的欄位：entries 填空陣列；query 填今天日期與「全部」。'
  ].join('\n');
}

var RECEIPT_PROMPT = '這是一張收據或發票照片，請記帳。';

// ===== Gemini.gs =====

/**
 * 用 Gemini API 解析訊息。需要指令碼屬性 GEMINI_API_KEY（Google AI Studio 可免費申請）。
 * 可用 GEMINI_MODEL 指定模型，預設使用最新的 Flash 模型。
 */

var GEMINI_DEFAULT_MODEL = 'gemini-flash-latest';

/**
 * Gemini 的 responseSchema 使用 OpenAPI 子集：型別要大寫、不支援 additionalProperties。
 */
function toGeminiSchema_(schema) {
  var out = { type: schema.type.toUpperCase() };
  if (schema.description) out.description = schema.description;
  if (schema.enum) out.enum = schema.enum;
  if (schema.required) out.required = schema.required;
  if (schema.items) out.items = toGeminiSchema_(schema.items);
  if (schema.properties) {
    out.properties = {};
    out.propertyOrdering = [];
    Object.keys(schema.properties).forEach(function (key) {
      out.properties[key] = toGeminiSchema_(schema.properties[key]);
      out.propertyOrdering.push(key);
    });
  }
  return out;
}

function parseWithGemini(input, today) {
  var parts = [];
  if (input.imageBase64) {
    parts.push({ inlineData: { mimeType: input.mediaType, data: input.imageBase64 } });
    parts.push({ text: RECEIPT_PROMPT });
  } else {
    parts.push({ text: input.text });
  }

  var body = {
    systemInstruction: { parts: [{ text: buildSystemPrompt_(today) }] },
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: toGeminiSchema_(getParseSchema_())
    }
  };

  var model = getProp_('GEMINI_MODEL', false) || GEMINI_DEFAULT_MODEL;
  var res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': getProp_('GEMINI_API_KEY', true) },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );

  var status = res.getResponseCode();
  if (status === 429) {
    throw new Error('Gemini 免費額度已用完，請稍後再試');
  }
  if (status !== 200) {
    throw new Error('Gemini API 錯誤 ' + status + '：' + res.getContentText().slice(0, 500));
  }

  var data = JSON.parse(res.getContentText());
  var candidate = data.candidates && data.candidates[0];
  if (!candidate) {
    var reason = data.promptFeedback && data.promptFeedback.blockReason;
    throw new Error('Gemini 沒有回覆' + (reason ? '（' + reason + '）' : ''));
  }
  if (candidate.finishReason && candidate.finishReason !== 'STOP') {
    throw new Error('Gemini 未完成回覆（' + candidate.finishReason + '）');
  }

  var text = '';
  (candidate.content.parts || []).forEach(function (p) {
    if (p.text && !p.thought) text += p.text;
  });
  if (!text) {
    throw new Error('Gemini 回覆中沒有文字內容');
  }
  return JSON.parse(text);
}

// ===== Claude.gs =====

/**
 * 用 Claude API 解析訊息。需要指令碼屬性 ANTHROPIC_API_KEY。
 */

var CLAUDE_MODEL = 'claude-haiku-5-5';

function parseWithClaude(input, today) {
  var content = [];
  if (input.imageBase64) {
    content.push({
      type: 'image',
      source: { type: 'base64', media_type: input.mediaType, data: input.imageBase64 }
    });
    content.push({ type: 'text', text: RECEIPT_PROMPT });
  } else {
    content.push({ type: 'text', text: input.text });
  }

  var body = {
    model: CLAUDE_MODEL,
    max_tokens: 4000,
    system: buildSystemPrompt_(today),
    output_config: {
      effort: 'low',
      format: { type: 'json_schema', schema: getParseSchema_() }
    },
    messages: [{ role: 'user', content: content }]
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': getProp_('ANTHROPIC_API_KEY', true),
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var status = res.getResponseCode();
  if (status !== 200) {
    throw new Error('Claude API 錯誤 ' + status + '：' + res.getContentText().slice(0, 500));
  }

  var data = JSON.parse(res.getContentText());
  if (data.stop_reason === 'refusal' || data.stop_reason === 'max_tokens') {
    throw new Error('Claude 未完成回覆（' + data.stop_reason + '）');
  }

  var textBlock = null;
  for (var i = 0; i < data.content.length; i++) {
    if (data.content[i].type === 'text') {
      textBlock = data.content[i];
      break;
    }
  }
  if (!textBlock) {
    throw new Error('Claude 回覆中沒有文字內容');
  }
  return JSON.parse(textBlock.text);
}

// ===== Code.gs =====

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
        replyText(event.replyToken, '⚠️ 處理失敗，請稍後再試一次。\n（' + String(err.message || err).slice(0, 200) + '）');
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
    replyText(event.replyToken, helpText_());
    return;
  }
  if (event.type !== 'message') return;

  var message = event.message;
  var isGroup = event.source.type !== 'user';
  var input;

  if (message.type === 'text') {
    var text = message.text.trim();
    if (text === '說明' || text === '幫助' || text.toLowerCase() === 'help') {
      replyText(event.replyToken, helpText_());
      return;
    }
    if (text === '刪除' || text === '取消') {
      replyText(event.replyToken, handleDelete_(event.source.userId));
      return;
    }
    input = { text: text };
  } else if (message.type === 'image') {
    if (getProvider_() === 'rules') {
      // 沒有 AI 無法讀收據；群組裡家人分享照片很常見，只在私訊提示
      if (!isGroup) {
        replyText(event.replyToken, '目前沒有開啟 AI，無法辨識收據照片。\n請用文字記帳，例如「全聯 560」。');
      }
      return;
    }
    input = getImageContent(message.id);
  } else {
    return;
  }

  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var parsed = parseMessage(input, today);
  parsed.entries = parsed.entries.filter(function (e) {
    return e.amount > 0;
  });

  if (parsed.intent === 'record' && parsed.entries.length > 0) {
    var recorder = getDisplayName(event.source);
    withLock_(function () {
      appendEntries(parsed.entries, {
        recorder: recorder,
        userId: event.source.userId,
        messageId: message.id,
        source: message.type === 'image' ? '收據照片' : '文字'
      });
    });
    replyText(event.replyToken, formatRecorded_(parsed.entries, recorder));
    return;
  }

  if (parsed.intent === 'query') {
    var q = parsed.query;
    replyText(event.replyToken, formatSummary_(q, summarize(q.start_date, q.end_date, q.category)));
    return;
  }

  // 群組裡的閒聊不回應，避免洗版；私訊則提示用法
  if (!isGroup) {
    var hint = getProvider_() === 'rules' ? '\n記帳請用「品項 金額」，例如「午餐 120」。' : '';
    replyText(event.replyToken, '看不出要記帳還是查詢 🤔' + hint + '\n傳「說明」可以看使用方式。');
  }
}

function handleDelete_(userId) {
  var deleted = withLock_(function () {
    return deleteLastEntry(userId);
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
