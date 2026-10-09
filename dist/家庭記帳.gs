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
function appendEntries_(entries, meta) {
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
 * 統計區間內的花費。keyword 有值時只算品項或備註包含關鍵字的紀錄。
 * @return {{ total: number, count: number, byCategory: Object, byRecorder: Object, rows: Array }}
 *   rows 依日期排序，供列出明細
 */
function summarize_(startDate, endDate, category, keyword) {
  var result = { total: 0, count: 0, byCategory: {}, byRecorder: {}, rows: [] };
  var kw = String(keyword || '').toLowerCase();
  readRows_().forEach(function (r) {
    var date = r[COL.date];
    if (date < startDate || date > endDate) return;
    if (category !== '全部' && r[COL.category] !== category) return;
    if (kw && (String(r[COL.item]) + ' ' + String(r[COL.note])).toLowerCase().indexOf(kw) < 0) return;
    var amount = Number(r[COL.amount]) || 0;
    result.rows.push(r);
    result.total += amount;
    result.count += 1;
    result.byCategory[r[COL.category]] = (result.byCategory[r[COL.category]] || 0) + amount;
    result.byRecorder[r[COL.recorder]] = (result.byRecorder[r[COL.recorder]] || 0) + amount;
  });
  result.rows.sort(function (a, b) {
    return a[COL.date] < b[COL.date] ? -1 : a[COL.date] > b[COL.date] ? 1 : 0;
  });
  return result;
}

/**
 * 刪除這位使用者最近一次記帳（同一則訊息記下的多筆會一起刪除）。
 * @return {Array} 被刪除的資料列；沒有可刪的則回傳空陣列
 */
function deleteLastEntry_(userId) {
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

function replyText_(replyToken, text) {
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

/** 主動推播（會用掉官方帳號每月的免費訊息則數，只用在每日提醒）。 */
function pushText_(to, text) {
  var res = lineFetch_('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text.slice(0, 5000) }] })
  });
  if (res.getResponseCode() !== 200) {
    console.error('LINE 推播失敗 ' + res.getResponseCode() + '：' + res.getContentText());
  }
}

/** 取得使用者顯示名稱，快取 6 小時。 */
function getDisplayName_(source) {
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
function getImageContent_(messageId) {
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
 * 查詢格式：今天、本週、本月、上月、今年，可加分類或關鍵字（「本月 餐飲」「全聯花多少」），加「明細」列出每一筆
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
      '食材', '生鮮', '魚', '肉', '蛋', '蔬菜', '青菜', '豆腐', '牛奶',
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

function parseWithRules_(text, today) {
  var t = normalizeText_(text);
  var analysis = parseAnalysis_(t, today);
  if (analysis) {
    return { intent: 'analysis', entries: [], query: analysis };
  }
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
  var statedTotal = null;
  for (var i = 0; i < segments.length; i++) {
    // 「金額1125元」「合計 1125」這類總計行不是另一筆消費，只拿來核對
    var totalLine = segments[i].match(/^(金額|合計|總計|總共|共計|小計|總額|共|total)\s*[:：]?\s*(\d+(?:\.\d+)?)\s*(元|塊錢|塊)?$/i);
    if (totalLine && segments.length > 1) {
      statedTotal = Math.round(Number(totalLine[2]));
      continue;
    }
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
  return { intent: 'record', entries: entries, query: emptyQuery_(today), statedTotal: statedTotal };
}

function emptyQuery_(today) {
  return { start_date: today, end_date: today, category: '全部', keyword: '', detail: false };
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
var UNSPECIFIED_ITEM = '未說明';

function parseEntry_(segment, today, keywords) {
  // 「我今天花了120元」：去掉開頭的「我」，讓後面的日期能被認出來
  var s = segment.replace(/^我\s*/, '');

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
  // 「午餐花了120」→ 品項「午餐」；「花了120」沒說用途 → 品項「未說明」，回覆時請記帳的人補上
  var spentWord = /(花了|花掉|用了|付了|買了|共花|總共)$/;
  if (spentWord.test(item)) {
    item = item.replace(spentWord, '').trim() || UNSPECIFIED_ITEM;
  }
  // 「買咖啡」→「咖啡」；只剩一個字時保留（「買菜」）
  item = item.replace(/^(買了?|去|在)\s*(?=\S{2,})/, '');
  // 「本月 7-11」這類被數字切開的品項不算記帳
  if (!item || /[-~～]$/.test(item) || !(amount > 0)) return null;

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

/**
 * 查詢：今天、昨天、本週、本月、上月、今年、查詢，可加分類或關鍵字，加「明細」會列出每一筆。
 *   本月 / 本月 餐飲 / 上個月餐飲多少 / 全聯花多少 / 本月 全聯 / 本月 明細 / 7-11 明細
 */
function parseQuery_(t, today) {
  var plain = t.replace(/[?？!！。]/g, '').trim();
  if (!plain) return null;
  var detail = /明細|清單/.test(plain);
  var s = plain.replace(/明細|清單/g, ' ').trim();
  var withoutSuffix = s.replace(/(花了?多少錢?|多少錢?|花費|統計|支出)$/, '').trim();
  var askedHowMuch = withoutSuffix !== s;
  s = withoutSuffix;
  // 有「花多少」「明細」這類字眼，才確定是查詢
  var explicit = askedHowMuch || detail;
  if (explicit) {
    // 「我這個月花多少錢」「我們家今天總共花多少」：去掉主詞和「總共」，不然會被當成搜尋字
    s = s.replace(/^(我們家|我們|我家|全家|家裡|大家|我)\s*/, '').replace(/\s*(總共|一共|全部|共)$/, '').trim();
    if (s === '全部' || s === '總共') s = '';
  }

  var category = '全部';
  var tokens = s.split(/\s+/).filter(function (p) {
    if (!p) return false;
    if (CATEGORIES.indexOf(p) >= 0) {
      category = p;
      return false;
    }
    return true;
  });
  // 也接受「上個月餐飲」這種沒有空白的寫法
  if (category === '全部' && tokens.length) {
    var last = tokens[tokens.length - 1];
    CATEGORIES.forEach(function (c) {
      if (category === '全部' && last.length > c.length && last.slice(-c.length) === c) {
        category = c;
        tokens[tokens.length - 1] = last.slice(0, -c.length);
      }
    });
  }

  var y = +today.slice(0, 4);
  var mo = +today.slice(5, 7);
  var monthStart = ymd_(y, mo, 1);
  var lastMonth = [ymd_(mo === 1 ? y - 1 : y, mo === 1 ? 12 : mo - 1, 1), addDays_(monthStart, -1)];
  var periods = {
    '今天': [today, today],
    '昨天': [addDays_(today, -1), addDays_(today, -1)],
    '本週': [startOfWeek_(today), today],
    '這週': [startOfWeek_(today), today],
    '這禮拜': [startOfWeek_(today), today],
    '上週': [addDays_(startOfWeek_(today), -7), addDays_(startOfWeek_(today), -1)],
    '上禮拜': [addDays_(startOfWeek_(today), -7), addDays_(startOfWeek_(today), -1)],
    '本月': [monthStart, today],
    '這個月': [monthStart, today],
    '上月': lastMonth,
    '上個月': lastMonth,
    '今年': [ymd_(y, 1, 1), today],
    '查詢': [monthStart, today]
  };
  var periodWords = Object.keys(periods).sort(function (a, b) {
    return b.length - a.length;
  });

  var range = null;
  var keyword = '';
  if (tokens.length && periods.hasOwnProperty(tokens[0])) {
    // 「本月」「本月 全聯」
    range = periods[tokens[0]];
    keyword = tokens.slice(1).join(' ');
  } else if (explicit) {
    // 「這個月花多少」「全聯花多少」「7-11 明細」
    var joined = tokens.join(' ');
    for (var i = 0; i < periodWords.length; i++) {
      if (joined.indexOf(periodWords[i]) === 0) {
        range = periods[periodWords[i]];
        keyword = joined.slice(periodWords[i].length).trim();
        break;
      }
    }
    if (!range) {
      range = [monthStart, today];
      keyword = joined;
    }
  } else if (tokens.length === 0 && category !== '全部') {
    // 只傳分類名稱「醫療」
    range = [monthStart, today];
  } else {
    return null;
  }

  // 「昨天 加油 1200」是記帳不是查詢
  if (!explicit && /\d+(\.\d+)?\s*(元|塊錢|塊)?$/.test(keyword)) return null;

  return { start_date: range[0], end_date: range[1], category: category, keyword: keyword, detail: detail };
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

// ===== Recurring.gs =====

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
  var m = t.match(/^(固定支出|訂閱)(.*)$/);
  if (!m) return null;
  var rest = m[2].trim();
  if (!rest) return { action: 'list' };

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

// ===== Analysis.gs =====

/**
 * 消費分析：總支出、日常平均每天、最大支出分類、和上一期比較。
 *
 * 金額都由程式從試算表計算（AI 只負責判斷要分析哪段期間）。
 * 固定支出（房租、水電等自動記帳）和日常花費分開看，
 * 不然房租這種大筆金額會讓「最大支出」永遠是同一項。
 */

var FIXED_SOURCE = '固定支出';

/**
 * 上一個比較期間：
 * - 從 1 號開始的期間（本月、上月）→ 上個月的同幾天，例如 10/1～10/9 對 9/1～9/9
 * - 其他期間 → 緊接在前、天數相同的期間，例如本週對上週
 */
function previousPeriod_(start, end) {
  var days = daysBetween_(start, end) + 1;
  if (start.slice(8) === '01') {
    var y = +start.slice(0, 4);
    var m = +start.slice(5, 7) - 1;
    if (m === 0) {
      y -= 1;
      m = 12;
    }
    var prevStart = ymd_(y, m, 1);
    var prevMonthEnd = ymd_(y, m, daysInMonth_(y, m));
    var prevEnd = addDays_(prevStart, days - 1);
    return [prevStart, prevEnd < prevMonthEnd ? prevEnd : prevMonthEnd];
  }
  var pEnd = addDays_(start, -1);
  return [addDays_(pEnd, -(days - 1)), pEnd];
}

function daysBetween_(a, b) {
  return Math.round((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 86400000);
}

/** 統計一段期間，把固定支出和日常花費分開。 */
function tally_(rows, start, end) {
  var t = { total: 0, daily: 0, fixed: 0, count: 0, byCategory: {}, byRecorder: {}, fixedItems: {}, byDate: {}, maxRow: null };
  rows.forEach(function (r) {
    var date = r[COL.date];
    if (date < start || date > end) return;
    var amount = Number(r[COL.amount]) || 0;
    t.total += amount;
    t.count += 1;
    if (r[COL.source] === FIXED_SOURCE) {
      t.fixed += amount;
      t.fixedItems[r[COL.item]] = (t.fixedItems[r[COL.item]] || 0) + amount;
      return;
    }
    t.daily += amount;
    t.byDate[date] = (t.byDate[date] || 0) + amount;
    if (!t.maxRow || amount > Number(t.maxRow[COL.amount])) t.maxRow = r;
    t.byCategory[r[COL.category]] = (t.byCategory[r[COL.category]] || 0) + amount;
    t.byRecorder[r[COL.recorder]] = (t.byRecorder[r[COL.recorder]] || 0) + amount;
  });
  return t;
}

/**
 * @param {string} start  YYYY-MM-DD
 * @param {string} end    YYYY-MM-DD（超過今天的部分不算天數）
 */
function analyze_(start, end, today) {
  if (end > today) end = today;
  if (start > end) start = end;
  var rows = readRows_();
  var prev = previousPeriod_(start, end);
  return {
    start: start,
    end: end,
    days: daysBetween_(start, end) + 1,
    prevStart: prev[0],
    prevEnd: prev[1],
    cur: tally_(rows, start, end),
    prev: tally_(rows, prev[0], prev[1])
  };
}

function formatAnalysis_(a) {
  var md = function (s) {
    return +s.slice(5, 7) + '/' + +s.slice(8);
  };
  var cur = a.cur;
  var lines = ['📊 消費分析（' + md(a.start) + '～' + md(a.end) + '）'];

  if (cur.count === 0) {
    lines.push('這段期間沒有紀錄。');
    return lines.join('\n');
  }

  // 總覽
  lines.push('');
  if (cur.fixed > 0) {
    lines.push('💵 總支出 $' + formatMoney_(cur.total) + '（日常 $' + formatMoney_(cur.daily) + '＋固定支出 $' + formatMoney_(cur.fixed) + '）');
  } else {
    lines.push('💵 總支出 $' + formatMoney_(cur.total));
  }
  lines.push('📅 日常花費平均每天 $' + formatMoney_(cur.daily / a.days) + '（' + a.days + ' 天）');

  // 和上一期比較（只比日常花費，固定支出每期差不多）
  var compareLabel = a.start.slice(8) === '01' ? '上月同期' : '前一段時間';
  var period = '（' + md(a.prevStart) + '～' + md(a.prevEnd) + '）';
  if (a.prev.daily > 0) {
    var diff = cur.daily - a.prev.daily;
    var pct = Math.round((Math.abs(diff) / a.prev.daily) * 100);
    if (Math.abs(diff) < 1 || pct === 0) {
      lines.push('📈 日常花費和' + compareLabel + period + '差不多');
    } else {
      lines.push((diff > 0 ? '📈' : '📉') + ' 日常花費比' + compareLabel + period +
        (diff > 0 ? '多' : '少') + ' $' + formatMoney_(Math.abs(diff)) + '（' + (diff > 0 ? '+' : '-') + pct + '%）');
    }
  } else {
    lines.push('📈 ' + compareLabel + period + '沒有日常花費紀錄，還無法比較');
  }

  // 最大支出分類與提醒
  var cats = Object.keys(cur.byCategory).sort(function (x, y) {
    return cur.byCategory[y] - cur.byCategory[x];
  });
  if (cats.length && cur.daily > 0) {
    var top = cats[0];
    var share = cur.byCategory[top] / cur.daily;
    lines.push('', '🔥 最大支出：' + top + ' $' + formatMoney_(cur.byCategory[top]) + '（' + Math.round(share * 100) + '%）');
    lines.push('💡 ' + insight_(top, share, cats.length));

    // 增加最多的分類
    var rising = cats.map(function (c) {
      return { c: c, d: cur.byCategory[c] - (a.prev.byCategory[c] || 0) };
    }).filter(function (x) {
      return a.prev.daily > 0 && x.d > 0;
    }).sort(function (x, y) {
      return y.d - x.d;
    })[0];
    // 最高單筆、最高單日（只看日常花費）
    if (cur.maxRow) {
      lines.push('🏆 最高單筆：' + cur.maxRow[COL.item] + ' $' + formatMoney_(cur.maxRow[COL.amount]) +
        '（' + md(String(cur.maxRow[COL.date])) + '，' + cur.maxRow[COL.recorder] + '）');
    }
    var days = Object.keys(cur.byDate);
    if (days.length > 1) {
      var topDay = days.sort(function (x, y) {
        return cur.byDate[y] - cur.byDate[x];
      })[0];
      lines.push('📆 最高單日：' + md(topDay) + ' $' + formatMoney_(cur.byDate[topDay]));
    }
    if (rising) {
      lines.push('📌 ' + rising.c + '比' + compareLabel + '多 $' + formatMoney_(rising.d) + '，是增加最多的項目');
    }

    lines.push('', '📂 日常花費分類');
    cats.forEach(function (c) {
      var line = '・' + c + ' $' + formatMoney_(cur.byCategory[c]) + '（' + Math.round((cur.byCategory[c] / cur.daily) * 100) + '%）';
      if (a.prev.daily > 0) {
        var d = cur.byCategory[c] - (a.prev.byCategory[c] || 0);
        if (Math.abs(d) >= 1) line += ' ' + (d > 0 ? '▲ +$' : '▼ -$') + formatMoney_(Math.abs(d));
      }
      lines.push(line);
    });
  }

  var fixedNames = Object.keys(cur.fixedItems);
  if (fixedNames.length) {
    lines.push('', '🔁 固定支出');
    fixedNames.forEach(function (n) {
      lines.push('・' + n + ' $' + formatMoney_(cur.fixedItems[n]));
    });
  }

  var people = Object.keys(cur.byRecorder);
  if (people.length > 1) {
    lines.push('', '👨‍👩‍👧 各家人（日常花費）');
    people.sort(function (x, y) {
      return cur.byRecorder[y] - cur.byRecorder[x];
    }).forEach(function (n) {
      lines.push('・' + n + ' $' + formatMoney_(cur.byRecorder[n]));
    });
  }
  return lines.join('\n');
}

function insight_(top, share, categoryCount) {
  if (categoryCount === 1) return '這段期間的日常花費都在「' + top + '」。';
  if (share >= 0.6) return top + '占了日常花費的' + Math.round(share * 10) + '成左右，是主要的花費來源。';
  if (share >= 0.4) return top + '是目前最大的花費，約占' + Math.round(share * 100) + '%。';
  return '各分類花費分布平均，沒有特別集中在哪一項。';
}

/**
 * 規則模式的分析指令：「分析」「本月分析」「上月分析」「幫我分析這個月的消費」「本週分析」。
 * 不是分析指令則回傳 null。
 */
function parseAnalysis_(text, today) {
  var t = normalizeText_(text);
  if (!/分析/.test(t)) return null;
  var s = t.replace(/幫我|請|分析|消費|花費|支出|一下|的|[?？!！。]/g, '').trim();
  var monthStart = today.slice(0, 8) + '01';
  if (!s) return { start_date: monthStart, end_date: today, category: '全部', keyword: '', detail: false };
  var q = parseQuery_(s, today);
  if (!q || q.keyword) return null;
  return { start_date: q.start_date, end_date: q.end_date, category: '全部', keyword: '', detail: false };
}

// ===== Modify.gs =====

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

/** 執行修改或刪除，回傳回覆文字。 */
function applyModifyCommand_(cmd, userId) {
  if (cmd.action === 'invalid') return cmd.usage;
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
    '・刪除 昨天 停車 60 → 指定日期、品項、金額'
  ].join('\n');
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

function parseWithGemini_(input, today) {
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
  var res = fetchWithRetry_(
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
  if (status !== 200) {
    console.error('Gemini API 錯誤 ' + status + '：' + res.getContentText().slice(0, 1000));
    if (status === 429) throw new Error('Gemini 免費額度已用完');
    if (status >= 500) throw new Error('Gemini 暫時忙線');
    throw new Error('Gemini API 錯誤 ' + status + '（請檢查 GEMINI_API_KEY 或 GEMINI_MODEL）');
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

function parseWithClaude_(input, today) {
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

  var res = fetchWithRetry_('https://api.anthropic.com/v1/messages', {
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
    console.error('Claude API 錯誤 ' + status + '：' + res.getContentText().slice(0, 1000));
    if (status === 429 || status >= 500) throw new Error('Claude 暫時忙線');
    throw new Error('Claude API 錯誤 ' + status + '（請檢查 ANTHROPIC_API_KEY）');
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
    '🗑️ 刪除：刪除（最近一筆）、刪除 午餐、刪除 昨天 停車 60',
    '',
    '🔔 訂閱／固定支出：固定支出、固定支出 Netflix 390 每月15號',
    '⏰ 扣款提醒：近期扣款（另外每天早上會自動提醒）',
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
    if (matchDueSoonQuestion_(text)) {
      replyText_(event.replyToken, formatDueSoon_(today0, 14));
      return;
    }
    var recurringName = matchRecurringQuestion_(text);
    if (recurringName !== null) {
      replyText_(event.replyToken, formatRecurringList_(recurringName));
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

  if (parsed.intent === 'modify') {
    replyText_(event.replyToken, withLock_(function () {
      return applyModifyCommand_(parsed.modify, event.source.userId);
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

  // 「你是誰」「你會什麼」：簡短自我介紹
  if (message.type === 'text' && /你是誰|你叫什麼|你會什麼|你可以做什麼|你能做什麼|自我介紹/.test(message.text)) {
    replyText_(event.replyToken, '我是家庭記帳機器人 📒\n' +
      '幫全家記帳、查詢與分析花費，也會管理房租、水電這類固定支出，扣款前一天提醒。\n\n' +
      '試試傳「午餐 120」，或傳「說明」看完整用法。');
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
