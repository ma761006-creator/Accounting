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
var CATEGORIES = ['餐飲', '交通', '日用品', '醫療', '育兒', '旅遊', '娛樂', '寵物', '其他'];

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
