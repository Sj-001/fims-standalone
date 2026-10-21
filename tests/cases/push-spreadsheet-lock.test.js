// THE LIVE BUG, confirmed directly against a real customer's Sheet (DIAMOND STOCK), 2026-10-21:
// pushCustomerSheet (and appendDaburSpecRows) each read the spreadsheet's CURRENT rows, compute an
// insert plan against that snapshot, then write -- and nothing serialized two pushes to the SAME
// spreadsheet against each other. Two overlapping push requests (confirmed as the real trigger here:
// two browser tabs/sessions pushing for the same customer close together) each read the same
// pre-write snapshot and each plan their own inserts against it -- so by the time the second one
// writes, its plan is already stale relative to what the first one just wrote. Real, confirmed
// consequence: a genuine duplicate production row for one item, and a column-shifted row (the
// production number landing in the Dispatch column instead) for a different item in the very same
// push, in the real DIAMOND Google Sheet -- found via the Sheet's own edit history (two writes from
// the service account one second apart, 12:52:12 PM and 12:52:13 PM, both against the pre-either-write
// snapshot).
//
// withSpreadsheetLock closes this: every write that does this same read-then-write dance runs through
// it, keyed by spreadsheetId, so a second write for the same spreadsheet only starts its OWN fresh read
// once the first one's write has actually finished -- never against a snapshot the first write is
// about to invalidate. Different spreadsheetIds must never block each other (an unrelated customer's
// push should never wait on this one), and one write throwing must never wedge the lock for every
// write after it for that same spreadsheet.
const path = require('path');
const { extractBlock } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

const lockSource = extractBlock(
  SHEETS_JS,
  'const pushLocksBySpreadsheetId = new Map();',
  (l) => l.trim() === '}',
);
const withSpreadsheetLock = new Function(`${lockSource}\nreturn withSpreadsheetLock;`)();

const t = makeRecorder('push-spreadsheet-lock');

function deferred() {
  let resolve;
  const promise = new Promise((res) => { resolve = res; });
  return { promise, resolve };
}

async function flushMicrotasks(times = 5) {
  for (let i = 0; i < times; i++) await Promise.resolve();
}

// CommonJS has no top-level await -- the whole suite runs inside one async IIFE, and run-all.js
// already knows how to await a test file's exported Promise (see its own comment on this).
module.exports = (async () => {

/* ===== THE LIVE CASE: two overlapping pushes for the SAME spreadsheetId -- the second must not start
   its own work until the first one's has fully finished. ===== */
await (async () => {
  const order = [];
  const first = deferred();
  const p1 = withSpreadsheetLock('sheet-A', async () => {
    order.push('first-start');
    await first.promise;
    order.push('first-end');
    return 'first-result';
  });
  // Give the first call's synchronous "first-start" push a chance to run before queuing the second.
  await flushMicrotasks();
  const p2 = withSpreadsheetLock('sheet-A', async () => {
    order.push('second-start');
    return 'second-result';
  });
  // The second call must NOT have started yet -- it's queued behind the still-pending first call.
  await flushMicrotasks();
  t.assertEqual(order.join(','), 'first-start', 'LIVE CASE: the second push for the same spreadsheet has not started while the first is still in flight');
  first.resolve();
  const [r1, r2] = await Promise.all([p1, p2]);
  t.assertEqual(order.join(','), 'first-start,first-end,second-start', 'LIVE CASE: the second push only starts once the first has fully finished');
  t.assertEqual(r1, 'first-result', 'the first call\'s own result is returned to its own caller');
  t.assertEqual(r2, 'second-result', 'the second call\'s own result is returned to its own caller, not the first call\'s');
})();

/* ===== SAFETY: a different spreadsheetId must never be blocked by an unrelated one's in-flight push.
   ===== */
await (async () => {
  const order = [];
  const first = deferred();
  const p1 = withSpreadsheetLock('sheet-B', async () => {
    order.push('B-start');
    await first.promise;
    order.push('B-end');
  });
  await flushMicrotasks();
  const p2 = withSpreadsheetLock('sheet-C', async () => {
    order.push('C-start');
    order.push('C-end');
  });
  await p2;
  t.assertEqual(order.join(','), 'B-start,C-start,C-end', 'SAFETY: an unrelated spreadsheetId runs immediately, not queued behind a different spreadsheet\'s in-flight push');
  first.resolve();
  await p1;
})();

/* ===== SAFETY: one push throwing must not permanently wedge the lock -- the next push for the same
   spreadsheet must still run (and run only after the failed one actually finished). ===== */
await (async () => {
  const order = [];
  let threw = false;
  try {
    await withSpreadsheetLock('sheet-D', async () => {
      order.push('fails');
      throw new Error('simulated write failure');
    });
  } catch (e) {
    threw = true;
  }
  t.assertTrue(threw, 'the failing call\'s own caller still sees the real error');
  const result = await withSpreadsheetLock('sheet-D', async () => {
    order.push('succeeds');
    return 'ok';
  });
  t.assertEqual(order.join(','), 'fails,succeeds', 'SAFETY: a later push for the same spreadsheet still runs after an earlier one failed, not wedged forever');
  t.assertEqual(result, 'ok', 'the later, successful call\'s own result still comes through');
})();

return t.summary();
})();
