// Shared assertion recorder for every test case file in tests/cases/. Each test file creates one
// recorder, runs its checks against it, and exports the final { pass, fail } summary — run-all.js reads
// that to decide the overall exit code. Deliberately tiny and dependency-free (no test framework
// installed in this repo) so `node tests/run-all.js` is the entire setup needed to run this suite.
function makeRecorder(label) {
  let pass = 0;
  let fail = 0;
  function assertEqual(actual, expected, msg) {
    const a = JSON.stringify(actual);
    const e = JSON.stringify(expected);
    if (a === e) { pass++; return; }
    fail++;
    console.error(`  FAIL [${label}]: ${msg}\n    got:  ${a}\n    want: ${e}`);
  }
  function assertNotEqual(actual, expected, msg) {
    const a = JSON.stringify(actual);
    const e = JSON.stringify(expected);
    if (a !== e) { pass++; return; }
    fail++;
    console.error(`  FAIL [${label}]: ${msg}\n    both: ${a}`);
  }
  function assertTrue(cond, msg) { assertEqual(!!cond, true, msg); }
  function summary() {
    console.log(`  ${pass} passed, ${fail} failed`);
    return { pass, fail };
  }
  return { assertEqual, assertNotEqual, assertTrue, summary };
}
module.exports = { makeRecorder };
