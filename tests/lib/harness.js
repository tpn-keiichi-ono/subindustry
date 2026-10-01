/**
 * 小さなテストの仕組み（依存パッケージなし）
 * 各テストファイルで test('名前', () => { ... }) と書く。実行は tests/run.js。
 */
'use strict';

const tests = [];
let currentFile = '';

function test(name, fn) {
  tests.push({file: currentFile, name, fn});
}

function setCurrentFile(file) {
  currentFile = file;
}

module.exports = {test, tests, setCurrentFile, assert: require('assert')};
