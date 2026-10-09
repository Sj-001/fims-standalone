#!/usr/bin/env node
// Runs every *.test.js in tests/cases/ and exits non-zero if any assertion failed. Each test file
// extracts real functions straight out of client/src/App.jsx or server/lib/sheets.js (see tests/lib/
// extract.js) and runs assertions against them on require — no separate test framework/build step
// needed, just `node tests/run-all.js` (or `npm test`).
const fs = require('fs');
const path = require('path');

const casesDir = path.join(__dirname, 'cases');
const files = fs.readdirSync(casesDir).filter(f => f.endsWith('.test.js')).sort();

let totalPass = 0;
let totalFail = 0;
const failedFiles = [];

for (const file of files) {
  const full = path.join(casesDir, file);
  process.stdout.write(`\n=== ${file} ===\n`);
  try {
    const result = require(full);
    if (result && typeof result.pass === 'number') {
      totalPass += result.pass;
      totalFail += result.fail;
      if (result.fail > 0) failedFiles.push(file);
    } else {
      throw new Error('test file did not export { pass, fail } — did it forget `module.exports = t.summary();`?');
    }
  } catch (e) {
    console.error(`FILE ERRORED: ${file}\n  ${e.stack || e}`);
    totalFail += 1;
    failedFiles.push(file);
  }
}

console.log(`\n${'='.repeat(60)}`);
console.log(`TOTAL: ${totalPass} passed, ${totalFail} failed, across ${files.length} file(s)`);
if (failedFiles.length) {
  console.log(`Failed: ${failedFiles.join(', ')}`);
  process.exit(1);
}
console.log('All tests passed.');
process.exit(0);
