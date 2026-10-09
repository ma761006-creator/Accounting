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
