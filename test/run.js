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
  const pushes = [];
  const triggers = [];
  const geminiStatuses = [];
  const imageStatus = { code: 200 };
  const replies = [];
  const quickReplies = [];
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
      setValue: (v) => {
        while (data[row - 1].length < col) data[row - 1].push('');
        data[row - 1][col - 1] = toCell(v);
      },
      setValues: (vals) => vals.forEach((vr, i) => vr.forEach((v, j) => {
        while (data[row - 1 + i].length < col + j) data[row - 1 + i].push('');
        data[row - 1 + i][col - 1 + j] = toCell(v);
      })),
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
        getProperty: (k) => props[k] || null,
        setProperty: (k, v) => { props[k] = v; }
      })
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss },
    CacheService: {
      getScriptCache: () => ({
        get: (k) => cacheStore.get(k) || null,
        put: (k, v) => cacheStore.set(k, v)
      })
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.map((name) => ({ getHandlerFunction: () => name })),
      newTrigger: (name) => {
        const chain = {
          timeBased: () => chain, everyDays: () => chain, atHour: () => chain, inTimezone: () => chain,
          create: () => triggers.push(name)
        };
        return chain;
      }
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
      base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
      sleep: () => {}
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
          const status = geminiStatuses.shift();
          if (status) return response(status, { error: { message: 'busy' } });
          const reply = claudeReplies.shift();
          return response(200, {
            candidates: [{
              finishReason: 'STOP',
              content: { parts: [{ text: 'thinking...', thought: true }, { text: JSON.stringify(reply) }] }
            }]
          });
        }
        if (url === 'https://api.line.me/v2/bot/message/push') {
          pushes.push(JSON.parse(opts.payload));
          return response(200, {});
        }
        if (url === 'https://api.line.me/v2/bot/message/reply') {
          const m = JSON.parse(opts.payload).messages[0];
          replies.push(m.text);
          quickReplies.push(m.quickReply ? m.quickReply.items.map((i) => i.action.label + '=' + i.action.text) : []);
          return response(200, {});
        }
        if (url.startsWith('https://api.line.me/v2/bot/group/')) {
          return response(200, { displayName: url.includes('Umom') ? '媽媽' : '爸爸' });
        }
        if (url.startsWith('https://api.line.me/v2/bot/profile/')) {
          return response(200, { displayName: '爸爸' });
        }
        if (url.startsWith('https://api-data.line.me/v2/bot/message/')) {
          if (imageStatus.code !== 200) return response(imageStatus.code, { message: 'Authentication failed' });
          return response(200, '', { getBytes: () => [1, 2, 3], getContentType: () => 'image/jpeg' });
        }
        throw new Error('unexpected fetch ' + url);
      }
    }
  };
  vm.createContext(context);

  const dir = path.join(__dirname, '..', 'gas');
  // 依 README 教學的建立順序載入（Code.gs 最先），確認全域變數不會依賴尚未載入的檔案
  const files = sourceFiles || ['Code.gs', 'Config.gs', 'Claude.gs', 'Line.gs', 'Sheet.gs', 'Parser.gs', 'Gemini.gs', 'Rules.gs', 'Recurring.gs', 'Analysis.gs', 'Modify.gs']
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

  return { context, props, rows, sheets, replies, quickReplies, imageStatus, pushes, triggers, geminiStatuses, claudeRequests, geminiRequests, post };
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
  const parse = (text, today) => env.context.parseWithRules_(text, today || '2026-10-09');
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
    assert.strictEqual(env.context.parseWithRules_(text, '2026-10-09').intent, 'other', text);
  });
});

test('規則辨識：查詢期間與分類', () => {
  const env = createEnv([], RULES);
  const q = (text, today) => {
    const r = env.context.parseWithRules_(text, today || '2026-10-09');
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
  const rec = env.context.parseWithRules_('我今天買咖啡花了80元', '2026-10-09');
  assert.strictEqual([rec.intent, rec.entries[0].item, rec.entries[0].category, rec.entries[0].amount].join('|'), 'record|咖啡|餐飲|80');
  assert.strictEqual(env.context.parseWithRules_('買菜 300', '2026-10-09').entries[0].item, '買菜');
  assert.strictEqual(q('上週花多少'), '2026-09-28|2026-10-04|全部');
  // 「我」「我們家」「總共」不能變成搜尋字，不然會查不到任何紀錄
  const kw = (text) => env.context.parseWithRules_(text, '2026-10-09').query;
  assert.strictEqual(q('我這個月花多少錢？'), '2026-10-01|2026-10-09|全部');
  assert.strictEqual(kw('我這個月花多少錢？').keyword, '');
  assert.strictEqual(q('我今天花多少錢？'), '2026-10-09|2026-10-09|全部');
  assert.strictEqual(kw('我們家今天總共花多少').keyword, '');
  assert.strictEqual(q('我上個月餐飲花多少'), '2026-09-01|2026-09-30|餐飲');
  assert.strictEqual(kw('我全聯花多少').keyword, '全聯');
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
  assert.strictEqual(fs.readFileSync(bundle.MANIFEST_OUT, 'utf8'), fs.readFileSync(bundle.MANIFEST_SRC, 'utf8'), '請執行 node scripts/bundle.js');
  const env = createEnv([], RULES, [bundle.OUT]);
  env.post({ type: 'text', id: '1', text: '7-11 85' });
  assert.strictEqual(env.rows[1][2], '餐飲');
});

test('執行選單只會出現 setup、doPost、dailyJob（其他函式都以 _ 結尾隱藏）', () => {
  const bundle = require('../scripts/bundle.js');
  const names = [...bundle.build().matchAll(/^function ([A-Za-z0-9_]+)\(/gm)].map((m) => m[1]);
  assert.deepStrictEqual(names.filter((n) => !n.endsWith('_')).sort(), ['dailyJob', 'doPost', 'setup']);
});

test('規則辨識：關鍵字與明細查詢', () => {
  const env = createEnv([], RULES);
  const q = (text) => {
    const r = env.context.parseWithRules_(text, '2026-10-09');
    assert.strictEqual(r.intent, 'query', text);
    return [r.query.start_date, r.query.end_date, r.query.category, r.query.keyword, r.query.detail].join('|');
  };
  assert.strictEqual(q('全聯花多少'), '2026-10-01|2026-10-09|全部|全聯|false');
  assert.strictEqual(q('本月 全聯'), '2026-10-01|2026-10-09|全部|全聯|false');
  assert.strictEqual(q('上個月全聯花多少'), '2026-09-01|2026-09-30|全部|全聯|false');
  assert.strictEqual(q('7-11花多少'), '2026-10-01|2026-10-09|全部|7-11|false');
  assert.strictEqual(q('7-11 明細'), '2026-10-01|2026-10-09|全部|7-11|true');
  assert.strictEqual(q('本月明細'), '2026-10-01|2026-10-09|全部||true');
  assert.strictEqual(q('上月 餐飲 明細'), '2026-09-01|2026-09-30|餐飲||true');
  assert.strictEqual(q('本月'), '2026-10-01|2026-10-09|全部||false');
  ['今天好熱', '本月 7-11', '全聯'].forEach((text) => {
    assert.strictEqual(env.context.parseWithRules_(text, '2026-10-09').intent, 'other', text);
  });
  assert.strictEqual(env.context.parseWithRules_('昨天 加油 1200', '2026-10-09').intent, 'record');
});

test('查詢回覆：關鍵字篩選與明細', () => {
  const env = createEnv([], RULES);
  const today = env.context.Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  env.post({ type: 'text', id: '1', text: '全聯 560' });
  env.post({ type: 'text', id: '2', text: '全聯 120 餐飲' }, { type: 'group', groupId: 'G1', userId: 'Umom' });
  env.post({ type: 'text', id: '3', text: '中油 1000' });
  env.post({ type: 'text', id: '4', text: '全聯 明細' });
  const reply = env.replies[3];
  assert.match(reply, /「全聯」/);
  assert.match(reply, /總計 \$680（2 筆）/);
  assert.match(reply, /明細：/);
  assert.match(reply, new RegExp(today.slice(5) + ' 全聯 \\$560（爸爸）'));
  assert.match(reply, /全聯 \$120（媽媽）/);
  assert.ok(!/中油/.test(reply));
});

test('快速選單：按鈕送出的指令都能直接處理，開了 AI 也不會呼叫 AI', () => {
  const env = createEnv([]);
  env.post({ type: 'text', id: '1', text: '說明' });
  const buttons = env.quickReplies[0];
  assert.ok(buttons.length >= 5 && buttons.length <= 13);
  buttons.forEach((b) => assert.ok(b.split('=')[0].length <= 20, b));
  // 逐一按下主選單的每個按鈕：都要有回覆，且不呼叫 AI
  buttons.forEach((b, i) => env.post({ type: 'text', id: 'b' + i, text: b.split('=')[1] }));
  assert.strictEqual(env.claudeRequests.length, 0);
  assert.strictEqual(env.replies.length, 1 + buttons.length);
  env.replies.slice(1).forEach((r) => assert.ok(!/看不出/.test(r), r));
  assert.match(env.replies[1], /📊 2026|📊 \d{4}-/);
  // 查詢、分析的回覆帶著下一步的按鈕
  assert.ok(env.quickReplies[1].some((b) => b === '上月=上月'));
  assert.ok(env.quickReplies[3].some((b) => b.endsWith('=上月分析')));
  // 其他按鈕：查詢子選單、分析期間
  env.post({ type: 'text', id: 'q', text: '查詢' });
  assert.ok(env.quickReplies[env.quickReplies.length - 1].includes('本週=本週'));
  ['本週分析', '上月分析', '今天', '上月'].forEach((t, i) => env.post({ type: 'text', id: 'x' + i, text: t }));
  assert.strictEqual(env.claudeRequests.length, 0);
  // 不是固定指令的才交給 AI
  assert.strictEqual(env.context.isFixedCommand_('本月 明細'), true);
  assert.strictEqual(env.context.isFixedCommand_('幫我分析這個月的消費'), false);
  assert.strictEqual(env.context.isFixedCommand_('本月全聯'), false);
});

test('照片下載失敗時回覆好懂的訊息，不顯示錯誤代碼', () => {
  const env = createEnv([], { GEMINI_API_KEY: 'g' });
  env.imageStatus.code = 401;
  env.post({ type: 'image', id: 'img1' });
  assert.match(env.replies[0], /照片下載失敗，請重新拍照/);
  assert.match(env.replies[0], /（401：Authentication failed）/);
  assert.strictEqual(env.rows.length, 0);
});

test('查詢回覆：10 筆以內直接列出明細', () => {
  const env = createEnv([], RULES);
  env.post({ type: 'text', id: '1', text: '午餐 120' });
  env.post({ type: 'text', id: '2', text: '咖啡 80' });
  env.post({ type: 'text', id: '3', text: '我今天花了多少錢' });
  const reply = env.replies[2];
  assert.match(reply, /總計 \$200（2 筆）/);
  assert.match(reply, /明細：/);
  assert.match(reply, /午餐 \$120/);
  assert.match(reply, /咖啡 \$80/);
});

test('AI 回傳的資料會先驗證再寫入', () => {
  const env = createEnv([{
    intent: 'record',
    entries: [
      { date: '2099-01-01', category: '娛樂', item: '電影', amount: 300.4, note: '' },
      { date: '2026-10-01', category: '餐飲', item: '午餐', amount: -5, note: '' },
      { date: 'yesterday', category: '餐飲', item: '', amount: 100, note: '' }
    ],
    query: noQuery
  }]);
  env.post({ type: 'text', id: '1', text: '看電影 300' });
  const today = env.context.Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  assert.strictEqual(env.rows.length, 2);
  assert.strictEqual(env.rows[1][2], '其他'); // 不合法的分類
  assert.strictEqual(env.rows[1][4], 300); // 金額取整數
  assert.strictEqual(env.context.Utilities.formatDate(env.rows[1][1], 'Asia/Taipei', 'yyyy-MM-dd'), today); // 未來日期改成今天
});

test('Gemini 暫時錯誤（503）會自動重試一次', () => {
  const env = createEnv([{ intent: 'record', entries: [{ date: '2026-10-09', category: '餐飲', item: '咖啡', amount: 80, note: '' }], query: noQuery }],
    { GEMINI_API_KEY: 'gm' });
  env.geminiStatuses.push(503);
  env.post({ type: 'text', id: '1', text: '咖啡 80' });
  assert.strictEqual(env.geminiRequests.length, 2);
  assert.strictEqual(env.rows.length, 2);
  assert.match(env.replies[0], /已記帳/);
});

test('固定支出：setup 建立工作表與每日排程，「固定支出」指令列出項目', () => {
  const env = createEnv([], RULES);
  env.context.setup();
  env.context.setup(); // 重複執行不會重複建立
  assert.deepStrictEqual(env.triggers, ['dailyJob']);
  const sheet = env.sheets['固定支出'];
  assert.deepStrictEqual(sheet.data.slice(1).map((r) => r[0]), ['房租', '水電']);
  env.post({ type: 'text', id: '1', text: '固定支出' });
  assert.match(env.replies[0], /房租｜金額不固定（只提醒）｜每月 5 號/);
  assert.match(env.replies[0], /水電｜金額不固定（只提醒）｜每 2 個月 20 號/);
});

test('固定支出：到期自動記帳、前一天提醒、金額空白只提醒', () => {
  const env = createEnv([], RULES);
  env.context.setup();
  const sheet = env.sheets['固定支出'];
  const D = (s) => new (vm.runInContext('Date', env.context))(s + 'T00:00:00+08:00');
  sheet.data[1] = ['房租', 15000, '其他', 5, 1, D('2026-11-05'), '是'];
  sheet.data[2] = ['水電', '', '其他', 31, 2, D('2026-11-30'), '是'];
  sheet.data.push(['停用的', 100, '其他', 1, 1, D('2026-11-01'), '否']);

  let msgs = env.context.processRecurring_('2026-11-04');
  assert.strictEqual(msgs.join('\n'), '⏰ 提醒：明天（2026-11-05）要繳「房租」 $15,000');
  assert.strictEqual(env.rows.length, 1); // 只有標題列

  msgs = env.context.processRecurring_('2026-11-05');
  assert.match(msgs.join('\n'), /已自動記帳：房租 \$15,000/);
  const ledger = env.rows.slice(1);
  assert.strictEqual(ledger.length, 1);
  assert.strictEqual(ledger[0][3], '房租');
  assert.strictEqual(ledger[0][5], '🔁 固定支出');
  const fmt = (d) => env.context.Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
  assert.strictEqual(fmt(sheet.data[1][5]), '2026-12-05');

  // 同一天再跑一次不會重複記帳
  env.context.processRecurring_('2026-11-05');
  assert.strictEqual(env.rows.length - 1, 1);

  msgs = env.context.processRecurring_('2026-11-30');
  assert.match(msgs.join('\n'), /今天是「水電」繳費日，金額不固定/);
  assert.strictEqual(env.rows.length - 1, 1);
  assert.strictEqual(fmt(sheet.data[2][5]), '2027-01-31'); // 每 2 個月，31 號
});

test('固定支出：提醒推播到機器人加入的群組', () => {
  const env = createEnv([], RULES);
  env.context.setup();
  env.post({ type: 'text', id: '1', text: '說明' }, { type: 'group', groupId: 'Gfamily', userId: 'Udad' });
  assert.strictEqual(env.props.NOTIFY_TARGET_ID, 'Gfamily');
  const sheet = env.sheets['固定支出'];
  const today = env.context.Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  sheet.data[1] = ['房租', 15000, '其他', +today.slice(8), 1, today, '是'];
  sheet.data.length = 2;
  env.context.dailyJob();
  assert.strictEqual(env.pushes.length, 1);
  assert.strictEqual(env.pushes[0].to, 'Gfamily');
  assert.match(env.pushes[0].messages[0].text, /已自動記帳：房租/);

  env.context.dailyJob(); // 沒有新事項就不推播
  assert.strictEqual(env.pushes.length, 1);
});

test('固定支出指令解析', () => {
  const env = createEnv([], RULES);
  const p = (t) => JSON.stringify(env.context.parseRecurringCommand_(t));
  assert.strictEqual(p('固定支出'), '{"action":"list"}');
  assert.strictEqual(p('固定支出 Netflix 390 每月15號'),
    '{"action":"upsert","name":"Netflix","amount":390,"amountBlank":false,"day":15,"every":1,"category":null}');
  assert.strictEqual(p('訂閱 YouTube Premium 199 每個月 3號'),
    '{"action":"upsert","name":"YouTube Premium","amount":199,"amountBlank":false,"day":3,"every":1,"category":null}');
  assert.strictEqual(p('固定支出 電費 不固定 每兩個月 20號'),
    '{"action":"upsert","name":"電費","amount":null,"amountBlank":true,"day":20,"every":2,"category":null}');
  assert.strictEqual(p('固定支出 房租 16000'),
    '{"action":"upsert","name":"房租","amount":16000,"amountBlank":false,"day":null,"every":null,"category":null}');
  assert.strictEqual(p('固定支出 刪除 Netflix'), '{"action":"delete","name":"Netflix"}');
  assert.strictEqual(p('固定支出 Netflix 390 40號'), '{"action":"invalid"}');
  assert.strictEqual(p('固定支出好多'), 'null');
  assert.strictEqual(p('午餐 120'), 'null');
});

test('固定支出：用 LINE 新增、修改、刪除與詢問', () => {
  const env = createEnv([], RULES);
  env.context.setup();
  const today = env.context.Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const sheet = env.sheets['固定支出'];
  const names = () => sheet.data.slice(1).map((r) => r[0]).join(',');
  const fmt = (d) => (typeof d === 'string' ? d : env.context.Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd'));

  env.post({ type: 'text', id: '1', text: '固定支出 Netflix 390 每月15號' });
  assert.strictEqual(names(), '房租,水電,Netflix');
  const netflix = sheet.data[3];
  assert.strictEqual(netflix[1], 390);
  assert.strictEqual(fmt(netflix[5]), env.context.firstDueDate_(today, 15));
  assert.match(env.replies[0], /✅ 已新增固定支出\n・Netflix｜\$390｜每月 15 號/);

  const rentNextBefore = fmt(sheet.data[1][5]);
  env.post({ type: 'text', id: '2', text: '固定支出 房租 16000' });
  assert.strictEqual(sheet.data[1][1], 16000);
  assert.strictEqual(fmt(sheet.data[1][5]), rentNextBefore); // 只改金額，下次扣款日不變
  assert.match(env.replies[1], /✏️ 已更新固定支出\n・房租｜\$16,000｜每月 5 號/);

  env.post({ type: 'text', id: '3', text: '固定支出 電費 不固定 每2個月 20號' });
  assert.match(env.replies[2], /電費｜金額不固定（只提醒）｜每 2 個月 20 號/);

  env.post({ type: 'text', id: '4', text: '固定支出 刪除 netflix' });
  assert.strictEqual(names(), '房租,水電,電費');
  assert.match(env.replies[3], /已刪除固定支出「netflix」/);

  env.post({ type: 'text', id: '5', text: '我有哪些固定支出？' });
  assert.match(env.replies[4], /房租.*\n.*水電.*\n.*電費/);
  env.post({ type: 'text', id: '6', text: '房租什麼時候繳' });
  assert.match(env.replies[5], /^🔁 固定支出\n・房租｜\$16,000/);
  assert.ok(!/水電/.test(env.replies[5]));

  // 群組閒聊不會誤建項目；「Netflix 訂閱 390」仍是一般記帳
  env.post({ type: 'text', id: '7', text: '固定支出好多' });
  assert.strictEqual(env.replies.length, 6);
  env.post({ type: 'text', id: '8', text: 'Netflix 訂閱 390' });
  assert.match(env.replies[6], /已記帳/);
  assert.strictEqual(names(), '房租,水電,電費');
});

test('固定支出：AI 自然語句新增', () => {
  const env = createEnv([{
    intent: 'recurring', entries: [], query: noQuery,
    recurring: { action: 'upsert', name: 'Spotify', amount: 199, day: 0, every: 0 }
  }], { GEMINI_API_KEY: 'gm' });
  env.context.setup();
  env.post({ type: 'text', id: '1', text: '我每個月訂 Spotify，月費 199' });
  const today = env.context.Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const row = env.sheets['固定支出'].data[3];
  assert.strictEqual(row[0], 'Spotify');
  assert.strictEqual(row[3], +today.slice(8)); // 沒說幾號就用今天
  assert.match(env.replies[0], /✅ 已新增固定支出/);
  const schema = env.geminiRequests[0].body.generationConfig.responseSchema;
  assert.ok(schema.properties.intent.enum.includes('recurring'));
});

test('分析：比較期間', () => {
  const env = createEnv([], RULES);
  const p = (a, b) => env.context.previousPeriod_(a, b).join('~');
  assert.strictEqual(p('2026-10-01', '2026-10-09'), '2026-09-01~2026-09-09');
  assert.strictEqual(p('2026-10-01', '2026-10-31'), '2026-09-01~2026-09-30');
  assert.strictEqual(p('2026-03-01', '2026-03-31'), '2026-02-01~2026-02-28');
  assert.strictEqual(p('2026-01-01', '2026-01-09'), '2025-12-01~2025-12-09');
  assert.strictEqual(p('2026-10-05', '2026-10-09'), '2026-09-30~2026-10-04');
});

test('分析：總支出、平均、最大支出、和上月同期比較、固定支出分開', () => {
  const env = createEnv([], RULES);
  const add = (date, category, item, amount, recorder, source) =>
    env.context.appendEntries_([{ date, category, item, amount, note: '' }],
      { recorder, userId: 'U', messageId: date + item, source: source || '文字' });
  add('2026-09-03', '餐飲', '午餐', 3000, '爸爸');
  add('2026-09-05', '交通', '加油', 2100, '爸爸');
  add('2026-09-20', '餐飲', '聚餐', 9999, '爸爸'); // 不在上月同期
  add('2026-10-02', '餐飲', '午餐', 3000, '爸爸');
  add('2026-10-03', '交通', '中油', 1800, '爸爸');
  add('2026-10-04', '餐飲', '晚餐', 2200, '媽媽');
  add('2026-10-06', '日用品', '全聯', 1450, '媽媽');
  add('2026-10-05', '其他', '房租', 15000, '🔁 固定支出', '固定支出');

  const text = env.context.formatAnalysis_(env.context.analyze_('2026-10-01', '2026-10-09', '2026-10-09'));
  const expect = [
    '📊 消費分析（10/1～10/9）',
    '💵 總支出 $23,450（日常 $8,450＋固定支出 $15,000）',
    '📅 日常花費平均每天 $939（9 天）',
    '📈 日常花費比上月同期（9/1～9/9）多 $3,350（+66%）',
    '🔥 最大支出：餐飲 $5,200（62%）',
    '💡 餐飲占了日常花費的6成左右，是主要的花費來源。',
    '📌 餐飲比上月同期多 $2,200，是增加最多的項目',
    '・餐飲 $5,200（62%） ▲ +$2,200',
    '・交通 $1,800（21%） ▼ -$300',
    '・日用品 $1,450（17%） ▲ +$1,450',
    '・房租 $15,000',
    '・爸爸 $4,800',
    '・媽媽 $3,650'
  ];
  expect.forEach((line) => assert.ok(text.includes(line), '缺少：' + line + '\n' + text));
  assert.ok(text.indexOf('・爸爸') < text.indexOf('・媽媽'));

  // 上一期沒有資料：不算百分比
  const none = env.context.formatAnalysis_(env.context.analyze_('2026-09-03', '2026-09-05', '2026-10-09'));
  assert.match(none, /前一段時間（8\/31～9\/2）沒有日常花費紀錄，還無法比較/);
  assert.ok(!/增加最多/.test(none));
  assert.ok(!/[▲▼]/.test(none));

  const empty = env.context.formatAnalysis_(env.context.analyze_('2026-08-01', '2026-08-31', '2026-10-09'));
  assert.match(empty, /這段期間沒有紀錄/);
});

test('分析：規則模式的指令與 AI 意圖', () => {
  const env = createEnv([], RULES);
  const a = (t) => {
    const r = env.context.parseWithRules_(t, '2026-10-09');
    return r.intent === 'analysis' ? r.query.start_date + '~' + r.query.end_date : r.intent;
  };
  assert.strictEqual(a('分析'), '2026-10-01~2026-10-09');
  assert.strictEqual(a('本月分析'), '2026-10-01~2026-10-09');
  assert.strictEqual(a('上月分析'), '2026-09-01~2026-09-30');
  assert.strictEqual(a('幫我分析這個月的消費'), '2026-10-01~2026-10-09');
  assert.strictEqual(a('本週分析'), '2026-10-05~2026-10-09');
  assert.strictEqual(a('全聯分析'), 'other');

  env.post({ type: 'text', id: '1', text: '午餐 120' });
  env.post({ type: 'text', id: '2', text: '分析' });
  assert.match(env.replies[1], /^📊 消費分析/);
  assert.match(env.replies[1], /💵 總支出 \$120/);

  const ai = createEnv([{
    intent: 'analysis', entries: [], query: { start_date: '2026-09-01', end_date: '2026-09-30', category: '全部', keyword: '', detail: false }
  }], { GEMINI_API_KEY: 'gm' });
  ai.post({ type: 'text', id: '1', text: '上個月花得比較多嗎' });
  assert.match(ai.replies[0], /📊 消費分析（9\/1～9\/30）/);
  assert.ok(ai.geminiRequests[0].body.generationConfig.responseSchema.properties.intent.enum.includes('analysis'));
});

test('Gemini 一直忙線（503）：文字改用規則辨識，照片和聊天回覆友善訊息', () => {
  const env = createEnv([], { GEMINI_API_KEY: 'gm' });
  const busy = () => env.geminiStatuses.push(503, 503, 503);

  busy();
  env.post({ type: 'text', id: '1', text: '午餐 120' });
  assert.strictEqual(env.geminiRequests.length, 3); // 第一次 + 重試兩次
  assert.strictEqual(env.rows[1][3], '午餐');
  assert.match(env.replies[0], /已記帳/);

  busy();
  env.post({ type: 'text', id: '2', text: '你好' }, { type: 'user', userId: 'Udad' });
  assert.match(env.replies[1], /AI 暫時忙線/);
  assert.match(env.replies[1], /品項 金額/);

  busy();
  env.post({ type: 'text', id: '3', text: '晚上吃什麼' }); // 群組閒聊仍然不回
  assert.strictEqual(env.replies.length, 2);

  busy();
  env.post({ type: 'image', id: '4' });
  assert.match(env.replies[2], /Gemini 暫時忙線，照片暫時無法辨識/);
  assert.ok(!/"error"/.test(env.replies[2])); // 不再把原始 JSON 丟給家人

  env.geminiStatuses.push(400);
  env.post({ type: 'text', id: '5', text: '中油 1200' }); // 金鑰錯誤等其他錯誤也會退回規則辨識
  assert.strictEqual(env.rows[2][3], '中油');
});

test('總計行（金額／合計）不重複記帳，並核對明細加總', () => {
  const env = createEnv([], RULES);
  env.post({ type: 'text', id: '1', text: '今天午餐660元\n食材鮭魚菲力275\n鯖魚*2片190\n金額1125元' });
  const ledger = env.rows.slice(1).map((r) => r[3] + '|' + r[2] + '|' + r[4]).join(',');
  assert.strictEqual(ledger, '午餐|餐飲|660,食材鮭魚菲力|餐飲|275,鯖魚*2片|餐飲|190');
  assert.match(env.replies[0], /合計 \$1,125（和你寫的總計相符）/);

  env.post({ type: 'text', id: '2', text: '早餐 80、咖啡 60、合計：150' });
  assert.match(env.replies[1], /⚠️ 你寫的總計是 \$150，和明細加總 \$140 不同/);
  assert.strictEqual(env.rows.length - 1, 5);

  // 單獨一行「金額 100」沒有其他明細時，仍當成一筆記帳
  env.post({ type: 'text', id: '3', text: '金額 100' });
  assert.strictEqual(env.rows.length - 1, 6);
});

test('修改：金額、分類、品項、指定品項、多筆時要求指定', () => {
  const env = createEnv([], RULES);
  const col = (i) => env.rows.slice(1).map((r) => r[i]).join(',');
  env.post({ type: 'text', id: '1', text: '午餐 120' });
  env.post({ type: 'text', id: '2', text: '修改 150' });
  assert.strictEqual(col(4), '150');
  assert.match(env.replies[1], /✏️ 已修改\n・原本：.*午餐 \$120\n・改成：.*午餐 \$150/);

  env.post({ type: 'text', id: '3', text: '修改 交通' });
  assert.strictEqual(col(2), '交通');
  env.post({ type: 'text', id: '4', text: '修改 品項 早午餐' });
  assert.strictEqual(col(3), '早午餐');

  env.post({ type: 'text', id: '5', text: '鮭魚 275、鯖魚 190' });
  env.post({ type: 'text', id: '6', text: '修改 200' });
  assert.match(env.replies[5], /最近一次記了 2 筆，請指定/);
  env.post({ type: 'text', id: '7', text: '鯖魚改成200' });
  assert.strictEqual(col(4), '150,275,200');
  env.post({ type: 'text', id: '8', text: '早午餐改成 餐飲' });
  assert.strictEqual(col(2), '餐飲,餐飲,餐飲');

  // 只能改自己的帳
  env.post({ type: 'text', id: '9', text: '修改 鮭魚 999' }, { type: 'group', groupId: 'G1', userId: 'Umom' });
  assert.match(env.replies[8], /找不到你記的「鮭魚」/);
  assert.strictEqual(col(4), '150,275,200');

  // 聊天不會被當成修改
  const before = env.replies.length;
  env.post({ type: 'text', id: '10', text: '計畫改成明天' });
  assert.strictEqual(env.replies.length, before);
  env.post({ type: 'text', id: '11', text: '修改' }, { type: 'user', userId: 'Udad' });
  assert.match(env.replies[env.replies.length - 1], /修改的用法/);
});

test('刪除：指定品項、日期與金額，「取消」只能單獨使用', () => {
  const env = createEnv([], RULES);
  const items = () => env.rows.slice(1).map((r) => r[3]).join(',');
  env.post({ type: 'text', id: '1', text: '昨天 停車 60' });
  env.post({ type: 'text', id: '2', text: '停車 80' });
  env.post({ type: 'text', id: '3', text: '午餐 120' });
  env.post({ type: 'text', id: '4', text: '早餐 70' });

  env.post({ type: 'text', id: '5', text: '刪除 午餐' });
  assert.strictEqual(items(), '停車,停車,早餐');
  env.post({ type: 'text', id: '6', text: '刪除 昨天 停車 60' });
  assert.strictEqual(env.rows.slice(1).map((r) => r[4]).join(','), '80,70');
  env.post({ type: 'text', id: '7', text: '取消聚餐' });
  assert.strictEqual(items(), '停車,早餐');
  env.post({ type: 'text', id: '8', text: '取消' });
  assert.strictEqual(items(), '停車');
  env.post({ type: 'text', id: '9', text: '刪除 晚餐' });
  assert.match(env.replies[env.replies.length - 1], /找不到你記的「晚餐」/);

  env.context.setup();
  env.post({ type: 'text', id: '10', text: '刪除 房租' });
  assert.match(env.replies[env.replies.length - 1], /固定支出 刪除 房租/);
});

test('近期扣款：列出 14 天內要繳的固定支出', () => {
  const env = createEnv([], RULES);
  env.context.setup();
  const D = (s) => new (vm.runInContext('Date', env.context))(s + 'T00:00:00+08:00');
  const sheet = env.sheets['固定支出'];
  sheet.data[1] = ['房租', 16000, '其他', 12, 1, D('2026-10-12'), '是'];
  sheet.data[2] = ['電費', '', '其他', 9, 2, D('2026-10-09'), '是'];
  sheet.data.push(['Netflix', 390, '其他', 30, 1, D('2026-10-30'), '是']);
  const text = env.context.formatDueSoon_('2026-10-09', 14);
  assert.match(text, /10\/9（今天）電費｜金額不固定\n・10\/12（3 天後）房租｜\$16,000/);
  assert.ok(!/Netflix/.test(text));
  assert.match(text, /已知金額合計 \$16,000/);
  ['近期扣款', '最近要繳什麼', '這週要扣款的有哪些', '扣款提醒'].forEach((t) =>
    assert.ok(env.context.matchDueSoonQuestion_(t), t));
  assert.ok(!env.context.matchDueSoonQuestion_('午餐 120'));
});

test('洞察：查詢顯示最高單筆，分析顯示最高單筆與最高單日', () => {
  const env = createEnv([], RULES);
  const add = (date, item, amount) => env.context.appendEntries_([{ date, category: '餐飲', item, amount, note: '' }],
    { recorder: '爸爸', userId: 'U', messageId: date + item, source: '文字' });
  add('2026-10-02', '午餐', 120);
  add('2026-10-02', '晚餐', 300);
  add('2026-10-05', '聚餐', 380);
  const a = env.context.formatAnalysis_(env.context.analyze_('2026-10-01', '2026-10-09', '2026-10-09'));
  assert.match(a, /🏆 最高單筆：聚餐 \$380（10\/5，爸爸）/);
  assert.match(a, /📆 最高單日：10\/2 \$420/);

  const q = env.context.formatSummary_({ start_date: '2026-10-01', end_date: '2026-10-09', category: '全部', keyword: '', detail: false },
    env.context.summarize_('2026-10-01', '2026-10-09', '全部', ''));
  assert.match(q, /🏆 最高單筆：聚餐 \$380（10-05）/);
});

test('AI 修改意圖', () => {
  const env = createEnv([
    { intent: 'record', entries: [{ date: '2026-10-09', category: '交通', item: '停車', amount: 60, note: '' }], query: noQuery },
    { intent: 'modify', entries: [], query: noQuery,
      modify: { action: 'edit', keyword: '停車', date: '', amount: 0, new_amount: 80, new_category: '', new_item: '' } },
    { intent: 'modify', entries: [], query: noQuery,
      modify: { action: 'delete', keyword: '停車', date: '', amount: 0, new_amount: 0, new_category: '', new_item: '' } }
  ], { GEMINI_API_KEY: 'gm' });
  env.post({ type: 'text', id: '1', text: '停車 60' });
  env.post({ type: 'text', id: '2', text: '剛剛的停車費其實是 80' });
  assert.strictEqual(env.rows[1][4], 80);
  env.post({ type: 'text', id: '3', text: '把停車費那筆刪掉' });
  assert.strictEqual(env.rows.length, 1);
  assert.ok(env.geminiRequests[0].body.generationConfig.responseSchema.properties.intent.enum.includes('modify'));
});

test('沒說用途時記成「未說明」並提醒補上；「你是誰」會自我介紹', () => {
  const env = createEnv([], RULES);
  const one = (t) => {
    const r = env.context.parseWithRules_(t, '2026-10-09');
    return r.intent === 'record' ? r.entries.map((e) => [e.date, e.category, e.item, e.amount].join('|')).join(',') : r.intent;
  };
  assert.strictEqual(one('我今天花了120元'), '2026-10-09|其他|未說明|120');
  assert.strictEqual(one('昨天花了 300'), '2026-10-08|其他|未說明|300');
  assert.strictEqual(one('午餐花了120'), '2026-10-09|餐飲|午餐|120');
  assert.strictEqual(one('我午餐 120'), '2026-10-09|餐飲|午餐|120');

  env.post({ type: 'text', id: '1', text: '我今天花了120元' });
  assert.match(env.replies[0], /👉 這筆用在哪裡？傳「修改 品項 午餐」補上/);
  env.post({ type: 'text', id: '2', text: '修改 品項 午餐' });
  assert.strictEqual(env.rows[1][3], '午餐');

  env.post({ type: 'text', id: '3', text: '你是誰？' }, { type: 'user', userId: 'Udad' });
  assert.match(env.replies[2], /我是家庭記帳機器人/);
  env.post({ type: 'text', id: '4', text: '你會什麼' });
  assert.match(env.replies[3], /點下面的按鈕/);
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
