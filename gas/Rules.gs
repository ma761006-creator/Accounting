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
