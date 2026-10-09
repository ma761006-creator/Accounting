// 在本機用模擬的 Apps Script 環境跑 gas/*.gs，檢查記帳、查詢、刪除流程。
// 執行：node test/run.js
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

function createEnv(claudeReplies, extraProps, sourceFiles) {
  const props = Object.assign({
    LINE_CHANNEL_ACCESS_TOKEN: 'line-token',
    ANTHROPIC_API_KEY: 'sk-test',
    LINE_BOT_USER_ID: 'Ubot'
  }, extraProps);
  const geminiRequests = [];
  const replies = [];
  const claudeRequests = [];
  const cacheStore = new Map();

  // 模擬試算表：appendRow 時把 ' 開頭轉成文字、YYYY-MM-DD 轉成 Date，和真的 Google 試算表一樣
  const rows = [];
  const toCell = (v) => {
    if (typeof v === 'string' && v.startsWith("'")) return v.slice(1);
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return new (vm.runInContext('Date', context))(v + 'T00:00:00+08:00');
    return v === undefined ? '' : v;
  };
  const makeSheet = (data) => ({
    data,
    getLastRow: () => data.length,
    appendRow: (r) => data.push(r.map(toCell)),
    setFrozenRows: () => {},
    deleteRow: (n) => data.splice(n - 1, 1),
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => data.slice(row - 1, row - 1 + numRows).map((r) => {
        const copy = r.slice(col - 1, col - 1 + numCols);
        while (copy.length < numCols) copy.push('');
        return copy;
      })
    })
  });
  const sheets = {};
  const ss = {
    getSheetByName: (name) => sheets[name] || null,
    // 帳本工作表沿用 rows，方便測試檢查
    insertSheet: (name) => (sheets[name] = makeSheet(name === '帳本' ? rows : []))
  };

  const response = (code, body, blob) => ({
    getResponseCode: () => code,
    getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)),
    getBlob: () => blob
  });

  const context = {
    console,
    JSON,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => props[k] || null
      })
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => cacheStore.get(k) || null,
        put: (k, v) => cacheStore.set(k, v)
      })
    },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    ContentService: { createTextOutput: (t) => ({ text: t }) },
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const parts = Object.fromEntries(
          new Intl.DateTimeFormat('en-CA', {
            timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
            hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
          }).formatToParts(d).map((p) => [p.type, p.value])
        );
        return fmt
          .replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day)
          .replace('HH', parts.hour).replace('mm', parts.minute).replace('ss', parts.second);
      },
      base64Encode: (bytes) => Buffer.from(bytes).toString('base64')
    },
    UrlFetchApp: {
      fetch: (url, opts) => {
        if (url === 'https://api.anthropic.com/v1/messages') {
          const req = JSON.parse(opts.payload);
          claudeRequests.push(req);
          const reply = claudeReplies.shift();
          return response(200, {
            stop_reason: 'end_turn',
            content: [{ type: 'thinking', thinking: '' }, { type: 'text', text: JSON.stringify(reply) }]
          });
        }
        if (url.startsWith('https://generativelanguage.googleapis.com/')) {
          geminiRequests.push({ url, headers: opts.headers, body: JSON.parse(opts.payload) });
          const reply = claudeReplies.shift();
          return response(200, {
            candidates: [{
              finishReason: 'STOP',
              content: { parts: [{ text: 'thinking...', thought: true }, { text: JSON.stringify(reply) }] }
            }]
          });
        }
        if (url === 'https://api.line.me/v2/bot/message/reply') {
          replies.push(JSON.parse(opts.payload).messages[0].text);
          return response(200, {});
        }
        if (url.startsWith('https://api.line.me/v2/bot/group/')) {
          return response(200, { displayName: url.includes('Umom') ? '媽媽' : '爸爸' });
        }
        if (url.startsWith('https://api.line.me/v2/bot/profile/')) {
          return response(200, { displayName: '爸爸' });
        }
        if (url.startsWith('https://api-data.line.me/v2/bot/message/')) {
          return response(200, '', { getBytes: () => [1, 2, 3], getContentType: () => 'image/jpeg' });
        }
        throw new Error('unexpected fetch ' + url);
      }
    }
  };
  vm.createContext(context);

  const dir = path.join(__dirname, '..', 'gas');
  // 依 README 教學的建立順序載入（Code.gs 最先），確認全域變數不會依賴尚未載入的檔案
  const files = sourceFiles || ['Code.gs', 'Config.gs', 'Claude.gs', 'Line.gs', 'Sheet.gs', 'Parser.gs', 'Gemini.gs', 'Rules.gs']
    .map((f) => path.join(dir, f));
  files.forEach((f) => {
    vm.runInContext(fs.readFileSync(f, 'utf8'), context, { filename: path.basename(f) });
  });

  let eventSeq = 0;
  const post = (message, source, extra) => {
    const event = Object.assign({
      type: 'message',
      webhookEventId: 'evt' + (++eventSeq),
      replyToken: 'rt',
      source: source || { type: 'group', groupId: 'G1', userId: 'Udad' },
      message
    }, extra);
    context.doPost({ postData: { contents: JSON.stringify({ destination: 'Ubot', events: [event] }) } });
    return event;
  };

  return { context, rows, sheets, replies, claudeRequests, geminiRequests, post };
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

const noQuery = { start_date: '2026-10-09', end_date: '2026-10-09', category: '全部' };

test('文字記帳：寫入試算表並回覆', () => {
  const env = createEnv([{
    intent: 'record',
    entries: [
      { date: '2026-10-08', category: '交通', item: '加油', amount: 1200, note: '' },
      { date: '2026-10-08', category: '交通', item: '停車', amount: 60, note: '' }
    ],
    query: noQuery
  }]);
  env.post({ type: 'text', id: '468789577898262530', text: '昨天加油 1200、停車 60' });

  assert.strictEqual(env.rows.length, 3); // 標題 + 2 筆
  assert.deepStrictEqual(env.rows[0], env.context.HEADERS);
  assert.strictEqual(env.rows[1][4], 1200);
  assert.strictEqual(env.rows[1][5], '爸爸');
  assert.strictEqual(env.rows[1][9], '468789577898262530');
  assert.match(env.replies[0], /已記帳（爸爸）/);
  assert.match(env.replies[0], /合計 \$1,260/);

  const req = env.claudeRequests[0];
  assert.strictEqual(req.model, 'claude-haiku-5-5');
  assert.strictEqual(req.output_config.format.type, 'json_schema');
  assert.deepStrictEqual(req.messages[0].content, [{ type: 'text', text: '昨天加油 1200、停車 60' }]);
});

test('收據照片：送出圖片給 Claude', () => {
  const env = createEnv([{
    intent: 'record',
    entries: [{ date: '2026-10-09', category: '日用品', item: '全聯', amount: 560, note: '衛生紙' }],
    query: noQuery
  }]);
  env.post({ type: 'image', id: '111' });
  const content = env.claudeRequests[0].messages[0].content;
  assert.strictEqual(content[0].type, 'image');
  assert.strictEqual(content[0].source.media_type, 'image/jpeg');
  assert.strictEqual(env.rows[1][7], '收據照片');
});

test('查詢：依分類與記錄人統計，含日期區間篩選', () => {
  const env = createEnv([
    { intent: 'record', entries: [{ date: '2026-10-01', category: '餐飲', item: '午餐', amount: 120, note: '' }], query: noQuery },
    { intent: 'record', entries: [{ date: '2026-10-05', category: '醫療', item: '掛號', amount: 150, note: '' }], query: noQuery },
    { intent: 'record', entries: [{ date: '2026-09-30', category: '餐飲', item: '晚餐', amount: 999, note: '' }], query: noQuery },
    { intent: 'query', entries: [], query: { start_date: '2026-10-01', end_date: '2026-10-09', category: '全部' } }
  ]);
  env.post({ type: 'text', id: '1', text: '午餐 120' });
  env.post({ type: 'text', id: '2', text: '掛號 150' }, { type: 'group', groupId: 'G1', userId: 'Umom' });
  env.post({ type: 'text', id: '3', text: '9/30 晚餐 999' });
  env.post({ type: 'text', id: '4', text: '這個月花多少' });

  const summary = env.replies[3];
  assert.match(summary, /總計 \$270（2 筆）/);
  assert.match(summary, /餐飲 \$120（44%）/);
  assert.match(summary, /醫療 \$150（56%）/);
  assert.match(summary, /媽媽 \$150/);
});

test('刪除：只刪自己最近一次的紀錄', () => {
  const env = createEnv([
    { intent: 'record', entries: [{ date: '2026-10-09', category: '餐飲', item: '早餐', amount: 80, note: '' }], query: noQuery },
    {
      intent: 'record',
      entries: [
        { date: '2026-10-09', category: '餐飲', item: '午餐', amount: 100, note: '' },
        { date: '2026-10-09', category: '餐飲', item: '飲料', amount: 50, note: '' }
      ],
      query: noQuery
    },
    { intent: 'record', entries: [{ date: '2026-10-09', category: '交通', item: '公車', amount: 15, note: '' }], query: noQuery }
  ]);
  env.post({ type: 'text', id: '1', text: '早餐 80' });
  env.post({ type: 'text', id: '2', text: '午餐 100 飲料 50' });
  env.post({ type: 'text', id: '3', text: '公車 15' }, { type: 'group', groupId: 'G1', userId: 'Umom' });
  env.post({ type: 'text', id: '4', text: '刪除' });

  const items = env.rows.slice(1).map((r) => r[3]);
  assert.deepStrictEqual(items, ['早餐', '公車']);
  assert.match(env.replies[3], /已刪除/);
  assert.match(env.replies[3], /午餐/);
  assert.match(env.replies[3], /飲料/);
});

test('群組閒聊不回覆、私訊會提示', () => {
  const other = { intent: 'other', entries: [], query: noQuery };
  const env = createEnv([other, other]);
  env.post({ type: 'text', id: '1', text: '晚上吃什麼' });
  assert.strictEqual(env.replies.length, 0);
  env.post({ type: 'text', id: '2', text: '你好' }, { type: 'user', userId: 'Udad' });
  assert.match(env.replies[0], /說明/);
});

test('重送的事件不重複記帳', () => {
  const rec = { intent: 'record', entries: [{ date: '2026-10-09', category: '餐飲', item: '午餐', amount: 120, note: '' }], query: noQuery };
  const env = createEnv([rec, rec]);
  const event = env.post({ type: 'text', id: '1', text: '午餐 120' });
  env.post(event.message, event.source, { webhookEventId: event.webhookEventId });
  assert.strictEqual(env.rows.length, 2);
  assert.strictEqual(env.claudeRequests.length, 1);
});

test('不是送給這個機器人的事件會被忽略', () => {
  const env = createEnv([]);
  env.context.doPost({ postData: { contents: JSON.stringify({ destination: 'Uother', events: [{ type: 'message' }] }) } });
  assert.strictEqual(env.replies.length, 0);
});

test('說明指令不呼叫 Claude', () => {
  const env = createEnv([]);
  env.post({ type: 'text', id: '1', text: '說明' });
  assert.strictEqual(env.claudeRequests.length, 0);
  assert.match(env.replies[0], /餐飲、交通、日用品、醫療/);
});

test('設定 GEMINI_API_KEY 時改用 Gemini', () => {
  const env = createEnv([
    { intent: 'record', entries: [{ date: '2026-10-09', category: '餐飲', item: '午餐', amount: 120, note: '' }], query: noQuery },
    { intent: 'record', entries: [{ date: '2026-10-09', category: '日用品', item: '全聯', amount: 560, note: '' }], query: noQuery }
  ], { GEMINI_API_KEY: 'gm-test' });
  env.post({ type: 'text', id: '1', text: '午餐 120' });
  env.post({ type: 'image', id: '2' });

  assert.strictEqual(env.claudeRequests.length, 0);
  assert.strictEqual(env.geminiRequests.length, 2);
  const req = env.geminiRequests[0];
  assert.match(req.url, /models\/gemini-flash-latest:generateContent$/);
  assert.strictEqual(req.headers['x-goog-api-key'], 'gm-test');
  const schema = req.body.generationConfig.responseSchema;
  assert.strictEqual(schema.type, 'OBJECT');
  assert.strictEqual(schema.properties.entries.items.properties.amount.type, 'NUMBER');
  assert.ok(!JSON.stringify(schema).includes('additionalProperties'));
  assert.deepStrictEqual(req.body.contents[0].parts, [{ text: '午餐 120' }]);
  assert.strictEqual(env.geminiRequests[1].body.contents[0].parts[0].inlineData.mimeType, 'image/jpeg');
  assert.deepStrictEqual(env.rows.slice(1).map((r) => r[3]), ['午餐', '全聯']);
});

test('AI_PROVIDER=claude 時即使有 Gemini 金鑰也用 Claude', () => {
  const env = createEnv([{ intent: 'other', entries: [], query: noQuery }], { GEMINI_API_KEY: 'gm', AI_PROVIDER: 'claude' });
  env.post({ type: 'text', id: '1', text: '你好' });
  assert.strictEqual(env.claudeRequests.length, 1);
  assert.strictEqual(env.geminiRequests.length, 0);
});

const RULES = { ANTHROPIC_API_KEY: '' };

test('規則辨識：記帳格式、日期、關鍵字分類', () => {
  const env = createEnv([], RULES);
  const parse = (text, today) => env.context.parseWithRules(text, today || '2026-10-09');
  const one = (text, today) => {
    const r = parse(text, today);
    assert.strictEqual(r.intent, 'record', text);
    assert.strictEqual(r.entries.length, 1, text);
    const e = r.entries[0];
    return [e.date, e.category, e.item, e.amount].join('|');
  };
  assert.strictEqual(one('午餐 120'), '2026-10-09|餐飲|午餐|120');
  assert.strictEqual(one('午餐120元'), '2026-10-09|餐飲|午餐|120');
  assert.strictEqual(one('全聯 560'), '2026-10-09|日用品|全聯|560');
  assert.strictEqual(one('7-11 85'), '2026-10-09|餐飲|7-11|85');
  assert.strictEqual(one('７－１１　８５'), '2026-10-09|餐飲|7-11|85');
  assert.strictEqual(one('中油 1,200'), '2026-10-09|交通|中油|1200');
  assert.strictEqual(one('昨天 加油 1200'), '2026-10-08|交通|加油|1200');
  assert.strictEqual(one('昨天加油1200'), '2026-10-08|交通|加油|1200');
  assert.strictEqual(one('10/8 全聯 560'), '2026-10-08|日用品|全聯|560');
  assert.strictEqual(one('10月8日 全聯 560'), '2026-10-08|日用品|全聯|560');
  assert.strictEqual(one('12/25 禮物 500'), '2025-12-25|其他|禮物|500');
  assert.strictEqual(one('掛號 150 醫療'), '2026-10-09|醫療|掛號|150');
  assert.strictEqual(one('醫療 口罩 99'), '2026-10-09|醫療|口罩|99');
  assert.strictEqual(one('全聯 300 餐飲'), '2026-10-09|餐飲|全聯|300');
  assert.strictEqual(one('ubereats 300'), '2026-10-09|餐飲|ubereats|300');
  assert.strictEqual(one('Uber 250'), '2026-10-09|交通|Uber|250');

  const multi = parse('加油 1200、停車 60');
  assert.strictEqual(multi.intent, 'record');
  assert.strictEqual(multi.entries.map((e) => e.amount).join(','), '1200,60');
});

test('規則辨識：聊天內容不會被誤記', () => {
  const env = createEnv([], RULES);
  ['我 3 點到', '晚上吃什麼', '午餐 120、明天見', '好', '收到！', '120'].forEach((text) => {
    assert.strictEqual(env.context.parseWithRules(text, '2026-10-09').intent, 'other', text);
  });
});

test('規則辨識：查詢期間與分類', () => {
  const env = createEnv([], RULES);
  const q = (text, today) => {
    const r = env.context.parseWithRules(text, today || '2026-10-09');
    assert.strictEqual(r.intent, 'query', text);
    return [r.query.start_date, r.query.end_date, r.query.category].join('|');
  };
  assert.strictEqual(q('本月'), '2026-10-01|2026-10-09|全部');
  assert.strictEqual(q('這個月花多少？'), '2026-10-01|2026-10-09|全部');
  assert.strictEqual(q('本月 餐飲'), '2026-10-01|2026-10-09|餐飲');
  assert.strictEqual(q('上個月餐飲多少'), '2026-09-01|2026-09-30|餐飲');
  assert.strictEqual(q('上月', '2026-01-15'), '2025-12-01|2025-12-31|全部');
  assert.strictEqual(q('本週'), '2026-10-05|2026-10-09|全部');
  assert.strictEqual(q('本週', '2026-10-11'), '2026-10-05|2026-10-11|全部');
  assert.strictEqual(q('今天'), '2026-10-09|2026-10-09|全部');
  assert.strictEqual(q('今年 交通'), '2026-01-01|2026-10-09|交通');
  assert.strictEqual(q('醫療'), '2026-10-01|2026-10-09|醫療');
  assert.strictEqual(q('花多少'), '2026-10-01|2026-10-09|全部');
});

test('規則模式：記帳不呼叫任何 AI，照片只在私訊提示', () => {
  const env = createEnv([], RULES);
  env.post({ type: 'text', id: '1', text: '中油 1200' });
  assert.strictEqual(env.claudeRequests.length + env.geminiRequests.length, 0);
  assert.strictEqual(env.rows[1][2], '交通');
  assert.match(env.replies[0], /已記帳/);

  env.post({ type: 'image', id: '2' });
  assert.strictEqual(env.replies.length, 1); // 群組不回
  env.post({ type: 'image', id: '3' }, { type: 'user', userId: 'Udad' });
  assert.match(env.replies[1], /無法辨識收據照片/);

  env.post({ type: 'text', id: '4', text: '說明' });
  assert.match(env.replies[2], /品項 金額/);
  assert.ok(!/收據/.test(env.replies[2]));
});

test('規則模式：試算表「關鍵字」工作表優先於內建關鍵字', () => {
  const env = createEnv([], RULES);
  env.context.setup();
  assert.ok(env.sheets['關鍵字']);
  env.sheets['關鍵字'].appendRow(['全聯', '餐飲']);
  env.sheets['關鍵字'].appendRow(['寵物', '不存在的分類']);
  env.post({ type: 'text', id: '1', text: '全聯 450' });
  env.post({ type: 'text', id: '2', text: '寵物飼料 300' });
  assert.deepStrictEqual(env.rows.slice(1).map((r) => r[2]), ['餐飲', '其他']);
});

test('dist/家庭記帳.gs 與 gas/ 一致，且單一檔案可以正常運作', () => {
  const bundle = require('../scripts/bundle.js');
  assert.strictEqual(fs.readFileSync(bundle.OUT, 'utf8'), bundle.build(), '請執行 node scripts/bundle.js');
  const env = createEnv([], RULES, [bundle.OUT]);
  env.post({ type: 'text', id: '1', text: '7-11 85' });
  assert.strictEqual(env.rows[1][2], '餐飲');
});

let failed = 0;
for (const t of tests) {
  try {
    t.fn();
    console.log('✓ ' + t.name);
  } catch (err) {
    failed++;
    console.log('✗ ' + t.name + '\n  ' + (err.stack || err));
  }
}
console.log(`\n${tests.length - failed}/${tests.length} 通過`);
process.exit(failed ? 1 : 0);
