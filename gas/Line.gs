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

/**
 * @param {Array<{label: string, text: string}>=} quickItems
 *   選填的快速選單按鈕：顯示在輸入框上方，點了就會送出 text，和自己打字一樣。
 */
function replyText_(replyToken, text, quickItems) {
  // LINE 單則文字上限 5000 字
  var message = { type: 'text', text: text.slice(0, 5000) };
  if (quickItems && quickItems.length) {
    // LINE 限制：最多 13 個按鈕，每個標籤最多 20 字
    message.quickReply = {
      items: quickItems.slice(0, 13).map(function (q) {
        return { type: 'action', action: { type: 'message', label: q.label.slice(0, 20), text: q.text } };
      })
    };
  }
  var res = lineFetch_('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ replyToken: replyToken, messages: [message] })
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
