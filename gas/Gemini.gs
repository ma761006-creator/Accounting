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
