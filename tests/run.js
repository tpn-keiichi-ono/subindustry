#!/usr/bin/env node
/**
 * テストの実行（npm test）
 *   npm test                 すべて
 *   npm test -- credential   ファイル名に credential を含むものだけ
 * GAS_VERBOSE=1 を付けると、スクリプトの console 出力も表示する。
 */
'use strict';

// Apps Script のプロジェクトと同じタイムゾーンにそろえる（appsscript.json の timeZone）
process.env.TZ = 'Asia/Tokyo';

const fs = require('fs');
const path = require('path');
const harness = require('./lib/harness');

const filter = process.argv[2] || '';
const files = fs.readdirSync(__dirname)
  .filter(name => name.endsWith('.test.js') && name.includes(filter))
  .sort();

files.forEach(name => {
  harness.setCurrentFile(name);
  require(path.join(__dirname, name));
});

(async () => {
  let failed = 0;
  let lastFile = '';
  for (const t of harness.tests) {
    if (t.file !== lastFile) {
      console.log('\n' + t.file);
      lastFile = t.file;
    }
    try {
      await t.fn();
      console.log('  ✓ ' + t.name);
    } catch (error) {
      failed++;
      console.log('  ✗ ' + t.name);
      console.log(String(error && error.stack || error).split('\n').map(l => '      ' + l).join('\n'));
    }
  }
  const total = harness.tests.length;
  console.log('\n' + (failed ? '✗ ' + failed + ' / ' + total + ' 件が失敗しました。' : '✓ ' + total + ' 件すべて成功しました。'));
  if (!total) console.log('（対象のテストがありません）');
  process.exit(failed || !total ? 1 : 0);
})();
