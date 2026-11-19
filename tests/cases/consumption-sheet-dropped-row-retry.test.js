// Real, directly-confirmed failure mode (2026-11-19): the exact same already-upright consumption-sheet
// page, re-sent completely unchanged, came back with all 11 real rows on one call and silently dropped
// 5 of them (an interior run, SL.NO 2/3/5/7/11) on another — genuine run-to-run non-determinism in the
// vision model, not anything about the image. The existing interior-gap flag (see
// consumption-sheet-dropped-row-flag.test.js) already detects this shape of failure reliably; what it
// couldn't do on its own is fix it. maybeRetryForDroppedRows (the code under test here) is what turns
// that detection into a real second chance: one fresh extraction call on the exact same image, which a
// direct side-by-side proved can genuinely self-heal the page. It fires ONLY for that specific flag
// reason (DROPPED_ROW_FLAG_MARKER) — a scribble or a leftover/weight mismatch reflects something
// actually ambiguous on the page itself, which asking the model again can't resolve, so those are left
// alone for a person to check instead of being bounced through a pointless extra API call.
//
// Real, directly-reported follow-up (2026-11-20): the SAME non-determinism that lets a retry fix a
// page just as easily lets a retry make it WORSE — a second attempt can drop even more real strips
// than the first. Blindly swapping in whatever the retry returned was itself a live row-dropping bug,
// not a hypothetical. The retry is now only trusted when it extracted at least as many rows as the
// original; otherwise the original (and its gap flag) is kept, exactly as if no retry had run.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff) {
  const blocks = [
    extractLine(APP_JSX, 'const DROPPED_ROW_FLAG_MARKER = '),
    extractBlock(APP_JSX, 'const maybeRetryForDroppedRows = async (', (l) => l.trim() === '};'),
  ];
  return buildFromSource(blocks, 'maybeRetryForDroppedRows', ['extractWithRateLimitBackoff'], [extractWithRateLimitBackoff]);
}

const t = makeRecorder('consumption-sheet-dropped-row-retry');

const GAP_REASON = 'SL.NO 2, 3 seem to be missing — this page has real data both before and after, which usually means a row was skipped during extraction. Check the original sheet and add it by hand if it has real data.';
const OTHER_REASON = 'Leftover (700) is not less than the reel weight (689) — likely landed on the wrong row. Check the original sheet.';

// shapeByRaw: a Map from a specific raw object (by reference) to the rows shape() should return for
// it — lets a test say exactly what the ORIGINAL raw shapes to vs. what the RETRIED raw shapes to,
// since the real code now calls shape() on each separately to compare row counts.
function makeMock(shapeByRaw, retriedResult) {
  const retryCalls = [];
  const extractWithRateLimitBackoff = async (prompt, base64, signal, onWaiting, onTick) => {
    retryCalls.push({ prompt, base64, signal, onWaiting, onTick });
    return retriedResult;
  };
  const shapeCalls = [];
  const shape = (raw) => {
    shapeCalls.push(raw);
    if (!shapeByRaw.has(raw)) throw new Error('shape() called with an unexpected raw object');
    return shapeByRaw.get(raw);
  };
  return { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls };
}

module.exports = (async () => {
  /* ===== 1. Wrong document type: never retried, even with a gap-flagged row — checked by pageConfig.key
     alone, shape() isn't even consulted. ===== */
  {
    const result = { raw: { items: [] }, truncated: false };
    const { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls } = makeMock(
      new Map([[result.raw, [{ flagged: true, flagReason: GAP_REASON }]]]), null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'production_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'WRONG TYPE: result passed through unchanged');
    t.assertEqual(shapeCalls.length, 0, 'WRONG TYPE: shape() never even called');
    t.assertEqual(retryCalls.length, 0, 'WRONG TYPE: never retries');
  }

  /* ===== 2. consumption_sheet, truncated: never retried regardless of flags — truncation already has
     its own, separate handling (see callClaudeExtract). ===== */
  {
    const result = { raw: { items: [] }, truncated: true };
    const { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls } = makeMock(
      new Map([[result.raw, [{ flagged: true, flagReason: GAP_REASON }]]]), null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'TRUNCATED: result passed through unchanged');
    t.assertEqual(shapeCalls.length, 0, 'TRUNCATED: shape() never even called');
    t.assertEqual(retryCalls.length, 0, 'TRUNCATED: never retries');
  }

  /* ===== 3. consumption_sheet, not truncated, no row flagged at all: passed through unchanged. ===== */
  {
    const result = { raw: { items: [] }, truncated: false };
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock(
      new Map([[result.raw, [{ flagged: false, flagReason: '' }, { flagged: false, flagReason: '' }]]]), null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'NO FLAGS: result passed through unchanged');
    t.assertEqual(retryCalls.length, 0, 'NO FLAGS: never retries');
  }

  /* ===== 4. consumption_sheet, not truncated, a row flagged but for a DIFFERENT reason (not the gap
     marker): never retried — that kind of flag needs a person, not another API call. ===== */
  {
    const result = { raw: { items: [] }, truncated: false };
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock(
      new Map([[result.raw, [{ flagged: true, flagReason: OTHER_REASON }]]]), null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'OTHER FLAG: result passed through unchanged');
    t.assertEqual(retryCalls.length, 0, 'OTHER FLAG: never retries for a non-gap flag reason');
  }

  /* ===== 5. Gap flag, retry comes back with MORE rows than the original (the real self-healing case):
     retries exactly once, with the SAME base64/prompt, and the retried raw/truncated wins. ===== */
  {
    const result = { raw: { items: [{ sl_no: '1' }, { sl_no: '4' }] }, truncated: false };
    const retried = { raw: { orientation: 'upright', items: [{ sl_no: '1' }, { sl_no: '2' }, { sl_no: '3' }] }, truncated: false, extra: 'should be dropped' };
    const originalShaped = [
      { flagged: false, flagReason: '' },
      { flagged: true, flagReason: OTHER_REASON },
      { flagged: true, flagReason: GAP_REASON },
    ];
    const retriedShaped = [{ flagged: false, flagReason: '' }, { flagged: false, flagReason: '' }, { flagged: false, flagReason: '' }, { flagged: false, flagReason: '' }];
    const { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls } = makeMock(
      new Map([[result.raw, originalShaped], [retried.raw, retriedShaped]]), retried);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(shapeCalls, [result.raw, retried.raw], 'MORE ROWS: shape() checked both the original AND the retried raw');
    t.assertEqual(retryCalls.length, 1, 'MORE ROWS: retries exactly once');
    t.assertEqual(retryCalls[0].base64, 'B64', 'MORE ROWS: retries with the SAME (already-corrected) image');
    t.assertEqual(retryCalls[0].prompt, 'PROMPT', 'MORE ROWS: retries with the exact same prompt');
    t.assertEqual(out, { raw: retried.raw, truncated: retried.truncated }, 'MORE ROWS: the retried (better) raw/truncated wins');
  }

  /* ===== 6. Gap flag, retry comes back with the SAME row count as the original (flag resolved, no new
     rows needed): the retry still wins — equal counts count as an improvement when the gap is gone. ===== */
  {
    const result = { raw: { items: [{ sl_no: '1' }] }, truncated: false };
    const retried = { raw: { items: [{ sl_no: '1' }] }, truncated: false };
    const originalShaped = [{ flagged: true, flagReason: GAP_REASON }, { flagged: false, flagReason: '' }];
    const retriedShaped = [{ flagged: false, flagReason: '' }, { flagged: false, flagReason: '' }];
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock(
      new Map([[result.raw, originalShaped], [retried.raw, retriedShaped]]), retried);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(retryCalls.length, 1, 'SAME COUNT: still retries once');
    t.assertEqual(out, { raw: retried.raw, truncated: retried.truncated }, 'SAME COUNT: retry still wins on an equal row count');
  }

  /* ===== 7. REAL BUG BEING FIXED: gap flag, but the retry comes back THINNER than the original (the
     second attempt dropped even more real rows) — the original is kept instead, flag and all, rather
     than silently swapping in a worse result. ===== */
  {
    const result = { raw: { items: [{ sl_no: '1' }, { sl_no: '2' }, { sl_no: '3' }] }, truncated: false };
    const retried = { raw: { items: [{ sl_no: '1' }] }, truncated: false };
    const originalShaped = [{ flagged: false }, { flagged: false }, { flagged: true, flagReason: GAP_REASON }];
    const retriedShaped = [{ flagged: false }];
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock(
      new Map([[result.raw, originalShaped], [retried.raw, retriedShaped]]), retried);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(retryCalls.length, 1, 'THINNER RETRY: still only ever retries once, win or lose');
    t.assertEqual(out, result, 'THINNER RETRY: the ORIGINAL (with its gap flag) is kept, not the worse retry');
  }

  /* ===== 8. Gap flag, but the retry ITSELF comes back truncated — never trusted, original kept, same
     reasoning as a thinner retry (a cut-off response is never known to be an improvement). ===== */
  {
    const result = { raw: { items: [{ sl_no: '1' }, { sl_no: '2' }] }, truncated: false };
    const retried = { raw: { items: [{ sl_no: '1' }, { sl_no: '2' }, { sl_no: '3' }] }, truncated: true };
    const originalShaped = [{ flagged: false }, { flagged: true, flagReason: GAP_REASON }];
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock(
      new Map([[result.raw, originalShaped]]), retried);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(retryCalls.length, 1, 'TRUNCATED RETRY: still retries once');
    t.assertEqual(out, result, 'TRUNCATED RETRY: original kept — a cut-off retry is never trusted, even with more raw items');
  }

  return t.summary();
})();
