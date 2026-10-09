// 把 gas/*.gs 合併成 dist/家庭記帳.gs，方便在 Apps Script 編輯器只貼一個檔案。
// 執行：node scripts/bundle.js（修改 gas/ 後要重新執行）
'use strict';

const fs = require('fs');
const path = require('path');

const ORDER = ['Config.gs', 'Sheet.gs', 'Line.gs', 'Rules.gs', 'Parser.gs', 'Gemini.gs', 'Claude.gs', 'Code.gs'];
const root = path.join(__dirname, '..');
const OUT = path.join(root, 'dist', '家庭記帳.gs');

function build() {
  const files = fs.readdirSync(path.join(root, 'gas')).filter((f) => f.endsWith('.gs'));
  const missing = files.filter((f) => !ORDER.includes(f));
  if (missing.length) throw new Error('scripts/bundle.js 的 ORDER 缺少：' + missing.join(', '));

  const parts = ORDER.map((f) =>
    '// ===== ' + f + ' =====\n\n' + fs.readFileSync(path.join(root, 'gas', f), 'utf8').trim() + '\n'
  );
  return '// 自動產生，請勿直接修改。原始檔在 gas/，修改後執行 node scripts/bundle.js\n\n' + parts.join('\n');
}

module.exports = { build, OUT };

if (require.main === module) {
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, build());
  console.log('已產生 ' + path.relative(root, OUT));
}
