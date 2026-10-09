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
 * 統計區間內的花費。
 * @return {{ total: number, count: number, byCategory: Object, byRecorder: Object }}
 */
function summarize_(startDate, endDate, category) {
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
