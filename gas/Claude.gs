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
