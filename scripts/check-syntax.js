/** src の .gs と、.html 内の <script> の構文を確かめる（実行はしない）：node scripts/check-syntax.js */
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..', 'src');
let failed = 0;
const names = new Map();
fs.readdirSync(dir).forEach(file => {
  const full = path.join(dir, file);
  const base = path.parse(file).name;
  if (names.has(base)) { console.error('NG 同じ名前のファイル（拡張子違い）: ' + names.get(base) + ' / ' + file); failed++; }
  names.set(base, file);
  try {
    const text = fs.readFileSync(full, 'utf8');
    if (file.endsWith('.gs')) new Function(text);
    if (file.endsWith('.html')) {
      const scripts = [...text.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
      scripts.forEach(code => new Function(code));
    }
    console.log('ok ' + file);
  } catch (error) {
    failed++;
    console.error('NG ' + file + ': ' + error.message);
  }
});
const onOpenCount = fs.readdirSync(dir).filter(f => f.endsWith('.gs'))
  .map(f => (fs.readFileSync(path.join(dir, f), 'utf8').match(/^function onOpen\(/gm) || []).length)
  .reduce((a, b) => a + b, 0);
if (onOpenCount > 1) { console.error('NG onOpen が ' + onOpenCount + ' 個あります（1つにする）'); failed++; }
process.exit(failed ? 1 : 0);
