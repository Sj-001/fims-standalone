// Real, directly-confirmed live incident (2026-11-21): a bulk push spanning several customer Sheets
// fired a burst of reads (readPushState per customer, the client's own preview/verify calls, the
// post-push mirror refresh) comfortably enough in one tight window to trip Google's per-user read quota
// — on a push covering nothing unusual in size (a few dozen rows across 4 customers). The actual Google
// error: "Quota exceeded for quota metric 'Read requests' and limit 'Read requests per minute per
// user'...". withGoogleApiRetry/wrapWithRetry fix this by retrying a momentarily-rate-limited call with
// backoff, transparently, so the person pushing never has to think about how many customers or rows one
// action touches.
//
// wrapWithRetry is NOT a plain Object.keys walk over the real googleapis client — confirmed directly,
// against the real client, that its methods (get, batchUpdate, values.batchGet, ...) are NOT
// own-enumerable properties (Object.keys(sheets.spreadsheets) returns only
// `['context', 'developerMetadata', 'sheets', 'values']` — every actual method lives on the prototype).
// An eager Object.keys copy silently produces an object with none of the real methods on it, breaking
// every Sheets API call outright the moment it ships — confirmed as a real bug caught before deploy,
// not a hypothetical. The Proxy-based version here is what actually reaches every real method.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

function loadFns() {
  // endExclusive + the next real declaration as each end anchor — not `(l) => l.trim() === '}'` — since
  // both functions have an inner block (an `if`, a Proxy `get` trap) that closes with its own `}` first;
  // a bare trimmed-'}' anchor would match that one and truncate the real function partway through.
  const retrySrc = extractBlock(SHEETS_JS, 'async function withGoogleApiRetry(fn, { retries = 5, baseDelayMs = 1000 } = {}) {', 'function wrapWithRetry(obj) {', { endExclusive: true });
  const wrapSrc = extractBlock(SHEETS_JS, 'function wrapWithRetry(obj) {', 'function getSheetsClient() {', { endExclusive: true });
  const ns = {};
  // Real timers would make this test slow (exponential backoff) and flaky (real jitter/randomness) --
  // setTimeout is swapped for an instant-resolving stub, isolating the RETRY DECISION logic (does it
  // retry this error or not, how many times) from how long it waits between attempts, which is not what
  // this test is checking.
  const body = `${retrySrc}\n\n${wrapSrc}\nreturn { withGoogleApiRetry, wrapWithRetry };`;
  const fn = new Function('setTimeout', body);
  return fn((cb) => cb());
}

const t = makeRecorder('google-api-retry');
const { withGoogleApiRetry, wrapWithRetry } = loadFns();

module.exports = (async () => {

/* ===== THE EXACT LIVE ERROR SHAPE: a 429 with reason 'rateLimitExceeded' (what a real googleapis
   client error for this looks like) is retried, and eventually succeeds once the underlying call stops
   failing -- exactly the real quota-exceeded-then-recovers case. ===== */
{
  let calls = 0;
  const flaky = async () => {
    calls++;
    if (calls < 3) { const e = new Error('Quota exceeded for quota metric...'); e.code = 429; e.errors = [{ reason: 'rateLimitExceeded' }]; throw e; }
    return 'ok';
  };
  const result = await withGoogleApiRetry(flaky);
  t.assertEqual(result, 'ok', 'LIVE CASE: eventually succeeds once the underlying call stops failing');
  t.assertEqual(calls, 3, 'LIVE CASE: retried exactly as many times as needed, no more');
}

/* ===== SAFETY: a permission error (403) is never retried -- retrying can't fix "not shared with the
   service account", so this must fail immediately, not after several pointless attempts. ===== */
{
  let calls = 0;
  const alwaysForbidden = async () => { calls++; const e = new Error('permission denied'); e.code = 403; throw e; };
  let threw = null;
  try { await withGoogleApiRetry(alwaysForbidden); } catch (e) { threw = e; }
  t.assertTrue(!!threw, 'SAFETY: a 403 still throws');
  t.assertEqual(calls, 1, 'SAFETY: a 403 is never retried -- exactly one attempt');
}

/* ===== SAFETY: a genuinely sustained rate limit (never recovers) still gives up eventually, surfacing
   the real error, rather than retrying forever. ===== */
{
  let calls = 0;
  const neverRecovers = async () => { calls++; const e = new Error('still exceeded'); e.code = 429; throw e; };
  let threw = null;
  try { await withGoogleApiRetry(neverRecovers, { retries: 3, baseDelayMs: 1 }); } catch (e) { threw = e; }
  t.assertTrue(!!threw, 'SAFETY: gives up and throws once retries are exhausted');
  t.assertEqual(calls, 4, 'SAFETY: exactly retries+1 attempts (the original try plus 3 retries), never more');
}

/* ===== SAFETY: a transient 503 is retried too, same as a 429 -- not every retryable failure is a
   quota error specifically. ===== */
{
  let calls = 0;
  const flaky503 = async () => { calls++; if (calls < 2) { const e = new Error('backend error'); e.code = 503; throw e; } return 'ok'; };
  const result = await withGoogleApiRetry(flaky503);
  t.assertEqual(result, 'ok', 'SAFETY: a transient 503 is retried and recovers');
}

/* ===== THE REAL CLIENT SHAPE: wrapWithRetry must reach methods that are NOT own-enumerable properties
   (confirmed directly against the real googleapis client -- see this file's own header comment). A
   plain object standing in for that exact shape (methods on the prototype, not as own keys). ===== */
{
  class FakeSpreadsheets {
    constructor() { this.callLog = []; }
    get(args) { this.callLog.push(['get', args]); return Promise.resolve('got-it'); }
  }
  const fakeClient = new FakeSpreadsheets();
  // Confirms the test's own fake actually reproduces the real shape before trusting what it proves.
  t.assertTrue(!Object.prototype.hasOwnProperty.call(fakeClient, 'get'), 'REAL SHAPE: get is on the prototype, not an own key, same as the real googleapis client');
  const wrapped = wrapWithRetry(fakeClient);
  const result = await wrapped.get({ spreadsheetId: 'abc' });
  t.assertEqual(result, 'got-it', 'REAL SHAPE: a Proxy reaches a prototype method a plain Object.keys walk would miss entirely');
}

/* ===== THE REAL CLIENT SHAPE, NESTED: a method under a nested namespace (spreadsheets.values.batchGet)
   is reached too, and still gets the retry behavior. ===== */
{
  class FakeValues {
    constructor() { this.calls = 0; }
    batchGet() { this.calls++; if (this.calls < 2) { const e = new Error('rate limited'); e.code = 429; throw e; } return Promise.resolve('batch-ok'); }
  }
  class FakeSpreadsheets2 { constructor() { this.values = new FakeValues(); } }
  const wrapped = wrapWithRetry(new FakeSpreadsheets2());
  const result = await wrapped.values.batchGet();
  t.assertEqual(result, 'batch-ok', 'NESTED: a nested namespace method is both reachable and retried on failure');
}

return t.summary();
})();
