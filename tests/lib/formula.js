/**
 * テスト用の小さな数式の計算（gas-mock から使う）
 *
 * サービス管理の候補のシート（__REQUEST_LISTS）が使う関数だけを、Google スプレッドシートと同じ考え方で計算する。
 * - 関数：IF・IFERROR・TRANSPOSE・UNIQUE・FILTER・SUMPRODUCT・COUNTA
 * - 演算子：= <> < > <= >= + - * / & と括弧。範囲（'シート'!$A$2:$A のような最終行までの範囲も）・文字列・数値
 * - 比べ方：文字列は大文字・小文字を区別しない。空のセルは "" とも 0 とも等しい。文字列と数値は等しくない
 * - 配列どうしの演算は要素ごと。エラーは {error: '#N/A'} のような値で伝わり、IFERROR で拾える
 * 知らない関数は #NAME? にする（本物と同じく、数式の書き間違いに気づけるように）。
 */
'use strict';

const isError = v => v && typeof v === 'object' && !Array.isArray(v) && 'error' in v;
const err = code => ({error: code});
const isArray = Array.isArray;

function colToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/* ---------------- 字句 ---------------- */

function tokenize(text) {
  const tokens = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let s = '';
      i++;
      for (;;) {
        if (i >= text.length) throw new Error('閉じていない文字列: ' + text);
        if (text[i] === '"') {
          if (text[i + 1] === '"') { s += '"'; i += 2; continue; }
          i++;
          break;
        }
        s += text[i++];
      }
      tokens.push({t: 'str', v: s});
      continue;
    }
    if (ch === "'") {
      let s = '';
      i++;
      for (;;) {
        if (i >= text.length) throw new Error('閉じていないシート名: ' + text);
        if (text[i] === "'") {
          if (text[i + 1] === "'") { s += "'"; i += 2; continue; }
          i++;
          break;
        }
        s += text[i++];
      }
      if (text[i] !== '!') throw new Error('シート名のあとに ! が必要です: ' + text);
      i++;
      tokens.push({t: 'sheet', v: s});
      continue;
    }
    const two = text.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { tokens.push({t: 'op', v: two}); i += 2; continue; }
    if ('=<>+-*/&(),:'.indexOf(ch) >= 0) { tokens.push({t: ch === '(' || ch === ')' || ch === ',' || ch === ':' ? ch : 'op', v: ch}); i++; continue; }
    const num = /^\d+(\.\d+)?/.exec(text.slice(i));
    if (num && !/^\d+[A-Za-z$]/.test(text.slice(i))) { tokens.push({t: 'num', v: Number(num[0])}); i += num[0].length; continue; }
    const word = /^[$A-Za-z_][$A-Za-z0-9_.]*/.exec(text.slice(i));
    if (word) {
      const rest = text.slice(i + word[0].length);
      if (rest[0] === '!') { tokens.push({t: 'sheet', v: word[0]}); i += word[0].length + 1; continue; }
      tokens.push({t: 'word', v: word[0]});
      i += word[0].length;
      continue;
    }
    throw new Error('読めない文字 "' + ch + '": ' + text);
  }
  return tokens;
}

/* ---------------- 構文 ---------------- */

function parse(text) {
  const tokens = tokenize(text.replace(/^=/, ''));
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  const expect = t => {
    const tok = next();
    if (!tok || tok.t !== t) throw new Error(t + ' が必要です: ' + text);
    return tok;
  };

  const binary = (ops, sub) => () => {
    let left = sub();
    while (peek() && peek().t === 'op' && ops.indexOf(peek().v) >= 0) {
      const op = next().v;
      left = {k: 'bin', op, left, right: sub()};
    }
    return left;
  };
  const primary = () => {
    const tok = next();
    if (!tok) throw new Error('式が途中で終わっています: ' + text);
    if (tok.t === 'num') return {k: 'lit', v: tok.v};
    if (tok.t === 'str') return {k: 'lit', v: tok.v};
    if (tok.t === '(') { const e = comparison(); expect(')'); return e; }
    if (tok.t === 'op' && tok.v === '-') return {k: 'neg', e: primary()};
    let sheet = null;
    let word = tok;
    if (tok.t === 'sheet') { sheet = tok.v; word = expect('word'); }
    if (word.t !== 'word') throw new Error('読めない式: ' + text);
    if (!sheet && peek() && peek().t === '(') {
      next();
      const args = [];
      if (peek() && peek().t !== ')') {
        args.push(comparison());
        while (peek() && peek().t === ',') { next(); args.push(comparison()); }
      }
      expect(')');
      return {k: 'call', name: word.v.toUpperCase(), args};
    }
    if (/^(TRUE|FALSE)$/i.test(word.v) && !sheet) return {k: 'lit', v: /^TRUE$/i.test(word.v)};
    let end = null;
    if (peek() && peek().t === ':') { next(); end = expect('word').v; }
    return {k: 'ref', sheet, start: word.v, end};
  };
  const multiplicative = binary(['*', '/'], primary);
  const additive = binary(['+', '-'], multiplicative);
  const concat = binary(['&'], additive);
  const comparison = binary(['=', '<>', '<', '>', '<=', '>='], concat);

  const tree = comparison();
  if (pos !== tokens.length) throw new Error('余分な字句があります: ' + text);
  return tree;
}

/* ---------------- 計算 ---------------- */

const blank = v => v === '' || v == null;

function compare(a, b, op) {
  if (isError(a)) return a;
  if (isError(b)) return b;
  let x = a;
  let y = b;
  if (blank(x) && typeof y === 'number') x = 0;
  if (blank(y) && typeof x === 'number') y = 0;
  if (blank(x) && typeof y === 'boolean') x = false;
  if (blank(y) && typeof x === 'boolean') y = false;
  if (blank(x)) x = '';
  if (blank(y)) y = '';
  if (typeof x !== typeof y) return op === '<>';
  if (typeof x === 'string') { x = x.toLowerCase(); y = y.toLowerCase(); }
  switch (op) {
    case '=': return x === y;
    case '<>': return x !== y;
    case '<': return x < y;
    case '>': return x > y;
    case '<=': return x <= y;
    default: return x >= y;
  }
}

function toNumber(v) {
  if (isError(v)) return v;
  if (blank(v)) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'number') return v;
  return err('#VALUE!');
}

function arith(a, b, op) {
  const x = toNumber(a);
  const y = toNumber(b);
  if (isError(x)) return x;
  if (isError(y)) return y;
  if (op === '+') return x + y;
  if (op === '-') return x - y;
  if (op === '*') return x * y;
  return y === 0 ? err('#DIV/0!') : x / y;
}

/** 配列どうし（または配列と1つの値）を要素ごとに計算する。 */
function elementwise(a, b, fn) {
  if (!isArray(a) && !isArray(b)) return fn(a, b);
  const rows = Math.max(isArray(a) ? a.length : 1, isArray(b) ? b.length : 1);
  const cols = Math.max(isArray(a) ? a[0].length : 1, isArray(b) ? b[0].length : 1);
  const at = (m, r, c) => (isArray(m) ? (m[m.length === 1 ? 0 : r] || [])[m[0].length === 1 ? 0 : c] : m);
  const out = [];
  for (let r = 0; r < rows; r++) {
    const line = [];
    for (let c = 0; c < cols; c++) line.push(fn(at(a, r, c), at(b, r, c)));
    out.push(line);
  }
  return out;
}

const truthy = v => (typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : false);
const flat = v => (isArray(v) ? [].concat(...v) : [v]);
const first = v => (isArray(v) ? v[0][0] : v);

/**
 * sheet（gas-mock の Sheet）の row 行 col 列の数式を計算する。結果は1つの値か2次元の配列。
 * cellValue(sheet, row, col) で、参照したセルの値（数式なら計算した値）を返す。
 */
function evaluate(formula, ctx) {
  const tree = parse(formula);

  const resolve = node => {
    const sheet = node.sheet ? ctx.sheetByName(node.sheet) : ctx.sheet;
    if (!sheet) return err('#REF!');
    const cell = text => {
      const m = /^\$?([A-Z]+)\$?(\d*)$/i.exec(text);
      if (!m) throw new Error('読めない参照: ' + text);
      return {col: colToNum(m[1].toUpperCase()), row: m[2] ? Number(m[2]) : null};
    };
    const a = cell(node.start);
    const b = node.end ? cell(node.end) : a;
    const r1 = a.row || 1;
    const r2 = b.row || sheet.getMaxRows();
    if (!node.end) return ctx.cellValue(sheet, r1, a.col);
    const out = [];
    for (let r = r1; r <= r2; r++) {
      const line = [];
      for (let c = a.col; c <= b.col; c++) line.push(ctx.cellValue(sheet, r, c));
      out.push(line);
    }
    return out;
  };

  const calls = {
    IF: args => {
      const c = first(run(args[0]));
      if (isError(c)) return c;
      return truthy(c) ? run(args[1]) : (args.length > 2 ? run(args[2]) : false);
    },
    IFERROR: args => {
      const v = run(args[0]);
      return isError(v) || (isArray(v) && v.length === 1 && v[0].length === 1 && isError(v[0][0])) ? run(args[1]) : v;
    },
    TRANSPOSE: args => {
      const v = run(args[0]);
      if (isError(v)) return v;
      if (!isArray(v)) return v;
      return v[0].map((_, c) => v.map(line => line[c]));
    },
    UNIQUE: args => {
      const v = run(args[0]);
      if (isError(v) || !isArray(v)) return v;
      const seen = new Set();
      return v.filter(line => {
        const key = JSON.stringify(line);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    },
    FILTER: args => {
      const range = run(args[0]);
      if (isError(range)) return range;
      const rows = isArray(range) ? range : [[range]];
      const conds = args.slice(1).map(run);
      for (const c of conds) if (isError(c)) return c;
      if (conds.some(c => !isArray(c) || c.length !== rows.length)) return err('#VALUE!');
      const out = rows.filter((_, r) => conds.every(c => truthy(c[r][0])));
      return out.length ? out : err('#N/A');
    },
    SUMPRODUCT: args => {
      const arrays = args.map(run);
      for (const a of arrays) if (isError(a)) return a;
      const values = arrays.map(flat);
      let sum = 0;
      values[0].forEach((_, i) => {
        let p = 1;
        values.forEach(v => { const n = toNumber(v[i]); p *= isError(n) ? 0 : n; });
        sum += p;
      });
      return sum;
    },
    COUNTA: args => args.map(run).reduce((n, v) => n + flat(v).filter(x => !blank(x)).length, 0)
  };

  function run(node) {
    switch (node.k) {
      case 'lit': return node.v;
      case 'ref': return resolve(node);
      case 'neg': return elementwise(run(node.e), 0, (x) => arith(0, x, '-'));
      case 'call': return calls[node.name] ? calls[node.name](node.args) : err('#NAME?');
      default: {
        const left = run(node.left);
        const right = run(node.right);
        if (node.op === '&') return elementwise(left, right, (x, y) => (isError(x) ? x : isError(y) ? y : String(blank(x) ? '' : x) + String(blank(y) ? '' : y)));
        if ('+-*/'.indexOf(node.op) >= 0) return elementwise(left, right, (x, y) => arith(x, y, node.op));
        return elementwise(left, right, (x, y) => compare(x, y, node.op));
      }
    }
  }

  return run(tree);
}

module.exports = {evaluate, parse, isError};
