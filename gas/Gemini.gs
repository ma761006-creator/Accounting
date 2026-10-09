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
  // Gemini 不接受空字串的選項；有空字串的就改成一般文字欄位（寫入前 validateParsed_ 會再檢查）
  if (schema.enum && schema.enum.indexOf('') < 0) out.enum = schema.enum;
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
  var cache = CacheService.getScriptCache();
  var noThinkKey = 'gemini-thinking-required:' + model;
  // 記帳是簡單的格式轉換，不需要模型先「思考」；關掉思考通常能快好幾秒
  if (!cache.get(noThinkKey)) body.generationConfig.thinkingConfig = { thinkingBudget: 0 };

  var send = function () {
    return fetchWithRetry_(
      'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': getProp_('GEMINI_API_KEY', true) },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      }
    );
  };
  var res = send();
  var status = res.getResponseCode();
  // 有些模型不能關掉思考，錯誤訊息的寫法也不一定：只要是 400 就拿掉這個設定重送一次，
  // 重送成功代表就是這個原因，記住 6 小時
  if (status === 400 && body.generationConfig.thinkingConfig) {
    console.warn('Gemini 400，改用預設思考設定重送：' + res.getContentText().slice(0, 300));
    delete body.generationConfig.thinkingConfig;
    res = send();
    status = res.getResponseCode();
    if (status === 200) cache.put(noThinkKey, '1', 6 * 60 * 60);
  }
  if (status !== 200) {
    console.error('Gemini API 錯誤 ' + status + '：' + res.getContentText().slice(0, 1000));
    if (status === 429) throw new Error('Gemini 免費額度已用完');
    if (status >= 500) throw new Error('Gemini 暫時忙線');
    // 附上 Gemini 的錯誤說明（不含金鑰），方便排查
    var detail = '';
    try {
      detail = JSON.parse(res.getContentText()).error.message || '';
    } catch (e) {
      detail = '';
    }
    throw new Error('Gemini API 錯誤 ' + status + (detail ? '：' + detail.slice(0, 150) : '（請檢查 GEMINI_API_KEY 或 GEMINI_MODEL）'));
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
