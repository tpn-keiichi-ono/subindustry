/** すべてのテストを順に実行する：node tests/run.js */
const fs = require('fs');
const path = require('path');
let failed = 0;
fs.readdirSync(__dirname).filter(f => f.endsWith('.test.js')).sort().forEach(file => {
  try {
    require(path.join(__dirname, file));
  } catch (error) {
    failed++;
    console.error('FAIL ' + file + '\n' + (error.stack || error));
  }
});
if (failed) { console.error(failed + ' test file(s) failed'); process.exit(1); }
console.log('all tests passed');
