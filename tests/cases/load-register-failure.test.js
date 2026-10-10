// A real, severe data-loss incident: loadRegister used to swallow any read failure (network blip,
// cold-start, Sheets API hiccup) and silently return [] -- indistinguishable from "this register is
// genuinely empty." On 2026-10-09, one bad read of fims_customer_dispatch on page load seeded React
// state with [] instead of its real multi-week history; the next confirm/push then persisted that
// near-empty state, overwriting the real Sheet's full history with a single new row. Fixed by retrying
// internally, then THROWING instead of returning [] if every attempt still fails -- the caller (the
// main load effect) must never treat "couldn't read" the same as "read it, it's empty."
//
// Exports a Promise (not a plain {pass, fail}) since this genuinely needs real retry-backoff timing to
// exercise the actual code path -- run-all.js awaits a thenable export.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadFn() {
  const sleepBlock = extractBlock(APP_JSX, 'const sleep = (ms, signal) => new Promise((resolve, reject) => {', (l) => l.trim() === '});');
  const loadRegisterBlock = extractBlock(
    APP_JSX, 'async function loadRegister(key, attempt = 0) {',
    '// Deliberately does NOT swallow its own errors',
    { endExclusive: true },
  );
  // Deliberately NOT injected as a function param -- loadRegister must resolve `window` from the
  // ambient global, same as it does in the real browser, so each test's `global.window = {...}`
  // (set before calling loadFn()) is what it actually sees.
  return buildFromSource([sleepBlock, loadRegisterBlock], 'loadRegister');
}

const t = makeRecorder('load-register-failure');

module.exports = (async () => {
  /* ===== 1. Normal success: returns the real data, no retry needed ===== */
  {
    global.window = { storage: { get: async () => ({ value: JSON.stringify([{ id: '1', date: '1.1.26' }]) }) } };
    const loadRegister = loadFn();
    const rows = await loadRegister('fims_customer_dispatch');
    t.assertEqual(rows, [{ id: '1', date: '1.1.26' }], 'normal successful load returns the real rows');
  }

  /* ===== 2. Genuinely empty register (successful fetch, zero rows) still returns [] cleanly -- this
     is a real, valid state (a brand new register) and must not be treated as a failure. ===== */
  {
    global.window = { storage: { get: async () => null } };
    const loadRegister = loadFn();
    const rows = await loadRegister('fims_customer_dispatch');
    t.assertEqual(rows, [], 'a genuinely empty register still loads as [] without throwing');
  }

  /* ===== 3. Transient failure that clears on retry: still returns the real data, doesn't give up early ===== */
  {
    let calls = 0;
    global.window = {
      storage: {
        get: async () => {
          calls++;
          if (calls < 2) throw new Error('simulated transient network blip');
          return { value: JSON.stringify([{ id: '1' }, { id: '2' }]) };
        },
      },
    };
    const loadRegister = loadFn();
    const rows = await loadRegister('fims_customer_dispatch');
    t.assertEqual(rows, [{ id: '1' }, { id: '2' }], 'a transient failure that clears on retry still returns the real data');
    t.assertTrue(calls >= 2, 'retried at least once before succeeding');
  }

  /* ===== 4. THE LIVE INCIDENT: every attempt fails -- must throw, never silently return [] ===== */
  {
    global.window = { storage: { get: async () => { throw new Error('simulated persistent failure'); } } };
    const loadRegister = loadFn();
    let threw = false;
    try {
      await loadRegister('fims_customer_dispatch');
    } catch (e) {
      threw = true;
    }
    t.assertTrue(threw, 'LIVE INCIDENT: a persistent read failure throws instead of silently returning []');
  }

  return t.summary();
})();
