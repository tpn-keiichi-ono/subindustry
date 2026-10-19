#!/usr/bin/env node
/**
 * Apps Script のファイル（リポジトリ直下の .gs・.html）のチェック（npm run check）
 * - .gs・.html がフォルダの中に無いか（同期でファイル名にフォルダ名が付き、HTML が見つからなくなる）
 * - .gs と .html の <script> の構文
 * - .gs と .html の同じ名前（Apps Script では拡張子が違っても同名不可）
 * - トップレベルの関数・定数の重複（全ファイルが1つのグローバルスコープに読み込まれるため）
 * - onOpen が1つだけか
 * - HTML ファイル名（createTemplateFromFile など）が実在するか
 * - メニュー・google.script.run・トリガーから呼ぶ関数が実在し、末尾 _ でないか
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const errors = [];
const fail = message => errors.push(message);

const files = fs.readdirSync(ROOT).filter(name => !name.startsWith('.') &&
  fs.statSync(path.join(ROOT, name)).isFile());
const gsFiles = files.filter(name => name.endsWith('.gs'));
const htmlFiles = files.filter(name => name.endsWith('.html'));
const read = name => fs.readFileSync(path.join(ROOT, name), 'utf8');

if (!gsFiles.length) fail('リポジトリ直下に .gs ファイルがありません。');
if (!files.includes('appsscript.json')) fail('リポジトリ直下に appsscript.json がありません。');
if (!fs.existsSync(path.join(ROOT, '.claspignore'))) {
  fail('.claspignore がありません（clasp で tests/・scripts/ の .js まで Apps Script に送ってしまいます）。');
}

/* ---------- フォルダの中の .gs・.html ---------- */

// 同期するとファイル名が「フォルダ名/ファイル名」になり、createTemplateFromFile('名前') で見つからなくなる
(function findNested(dir) {
  fs.readdirSync(dir, {withFileTypes: true}).forEach(entry => {
    if (entry.name.startsWith('.') || entry.name === 'node_modules') return;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      findNested(full);
    } else if (dir !== ROOT && /\.(gs|html)$/.test(entry.name)) {
      fail(path.relative(ROOT, full) + ': Apps Script のファイルはリポジトリ直下に置いてください' +
        '（フォルダに入れると、同期したときにファイル名にフォルダ名が付きます）。');
    }
  });
})(ROOT);

/* ---------- 構文 ---------- */

function checkSyntax(code, filename) {
  try {
    new vm.Script(code, {filename});
    return true;
  } catch (error) {
    fail(filename + ': 構文エラー: ' + error.message);
    return false;
  }
}

const gsCode = {};
gsFiles.forEach(name => {
  gsCode[name] = read(name);
  checkSyntax(gsCode[name], name);
});

/** インラインの <script> を取り出す。テンプレートの <?= ?> は値に置き換える。 */
function inlineScripts(html) {
  const out = [];
  const pattern = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = pattern.exec(html))) {
    if (/\bsrc\s*=/.test(m[1])) continue;
    out.push(m[2].replace(/<\?(!?=)?[\s\S]*?\?>/g, (_, out) => (out ? 'null' : '')));
  }
  return out;
}

const htmlScripts = {};
htmlFiles.forEach(name => {
  htmlScripts[name] = inlineScripts(read(name));
  htmlScripts[name].forEach((code, i) => checkSyntax(code, name + ' の <script> ' + (i + 1)));
});

/* ---------- ファイル名の重複 ---------- */

const byBase = new Map();
files.filter(name => /\.(gs|html)$/.test(name)).forEach(name => {
  const base = name.replace(/\.(gs|html)$/, '').toLowerCase();
  if (!byBase.has(base)) byBase.set(base, []);
  byBase.get(base).push(name);
});
byBase.forEach(names => {
  if (names.length > 1) fail('同じ名前のファイルがあります（拡張子・大文字小文字の違いも不可）: ' + names.join(', '));
});

/* ---------- トップレベルの名前の重複 ---------- */

const declared = new Map();   // 名前 → [ファイル名]
const functions = new Set();
gsFiles.forEach(name => {
  gsCode[name].split(/\r?\n/).forEach(line => {
    const m = line.match(/^(?:async\s+)?(function\*?|const|let|var|class)\s+([A-Za-z_$][\w$]*)/);
    if (!m) return;
    if (!declared.has(m[2])) declared.set(m[2], []);
    declared.get(m[2]).push(name);
    if (m[1].startsWith('function')) functions.add(m[2]);
  });
});
declared.forEach((where, name) => {
  if (where.length > 1) fail('「' + name + '」が複数の場所で宣言されています: ' + where.join(', '));
});

const onOpenCount = (declared.get('onOpen') || []).length;
if (onOpenCount !== 1) fail('onOpen はプロジェクト全体で1つだけにしてください（見つかった数: ' + onOpenCount + '）。');

/* ---------- HTML ファイル名 ---------- */

const htmlNames = new Set(htmlFiles.map(name => name.replace(/\.html$/, '')));
gsFiles.forEach(name => {
  const code = gsCode[name];
  const refs = [];
  const call = /create(?:TemplateFromFile|HtmlOutputFromFile)\(\s*'([^']+)'/g;
  const option = /\b\w*[tT]emplate\s*:\s*'([^']+)'/g;
  let m;
  while ((m = call.exec(code))) refs.push(m[1]);
  while ((m = option.exec(code))) refs.push(m[1]);
  refs.forEach(ref => {
    if (!htmlNames.has(ref)) fail(name + ': HTML ファイル「' + ref + '.html」がありません。');
  });
});

/* ---------- 外から呼ばれる関数 ---------- */

function checkCallable(fn, where) {
  if (!functions.has(fn)) fail(where + ': 関数「' + fn + '」が見つかりません。');
  else if (fn.endsWith('_')) fail(where + ': 「' + fn + '」は末尾 _ の内部用関数のため、ここからは呼べません。');
}

// メニュー項目・トリガー
gsFiles.forEach(name => {
  const code = gsCode[name];
  const patterns = [
    /\.addItem\(\s*'[^']*'\s*,\s*'([^']+)'\s*\)/g,
    /newTrigger\(\s*'([^']+)'\s*\)/g,
    /\b\w*[hH]andler\s*:\s*'([^']+)'/g
  ];
  patterns.forEach(pattern => {
    let m;
    while ((m = pattern.exec(code))) checkCallable(m[1], name);
  });
});

/** google.script.run の呼び出し先（with〜 を読み飛ばした最初のメソッド名）を返す。 */
function runTargets(code) {
  const targets = [];
  const marker = 'google.script.run';
  let at = code.indexOf(marker);
  while (at >= 0) {
    let i = at + marker.length;
    for (;;) {
      while (/\s/.test(code[i] || '')) i++;
      if (code[i] !== '.') break;
      i++;
      while (/\s/.test(code[i] || '')) i++;
      const m = /^[A-Za-z_$][\w$]*/.exec(code.slice(i));
      if (!m) break;
      i += m[0].length;
      if (!/^with(SuccessHandler|FailureHandler|UserObject)$/.test(m[0])) {
        targets.push(m[0]);
        break;
      }
      i = skipParens(code, i);
    }
    at = code.indexOf(marker, i);
  }
  return targets;
}

/** i の位置の ( から、対応する ) の次まで進める（文字列・テンプレート文字列は読み飛ばす）。 */
function skipParens(code, i) {
  while (/\s/.test(code[i] || '')) i++;
  if (code[i] !== '(') return i;
  let depth = 0;
  for (; i < code.length; i++) {
    const c = code[i];
    if (c === '"' || c === "'" || c === '`') {
      for (i++; i < code.length && code[i] !== c; i++) if (code[i] === '\\') i++;
    } else if (c === '(') depth++;
    else if (c === ')' && --depth === 0) return i + 1;
  }
  return i;
}

htmlFiles.forEach(name => {
  htmlScripts[name].forEach(code => {
    runTargets(code).forEach(fn => checkCallable(fn, name + ' の google.script.run'));
  });
});

/* ---------- 結果 ---------- */

if (errors.length) {
  console.error('✗ チェックで ' + errors.length + ' 件の問題が見つかりました。');
  errors.forEach(message => console.error('  - ' + message));
  process.exit(1);
}
console.log('✓ Apps Script のファイルのチェックに問題はありません（.gs ' + gsFiles.length + ' / .html ' + htmlFiles.length + ' ファイル）。');
