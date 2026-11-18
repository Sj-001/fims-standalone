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

function makeMock(shapedRows, retriedResult) {
  const retryCalls = [];
  const extractWithRateLimitBackoff = async (prompt, base64, signal, onWaiting, onTick) => {
    retryCalls.push({ prompt, base64, signal, onWaiting, onTick });
    return retriedResult;
  };
  const shapeCalls = [];
  const shape = (raw) => { shapeCalls.push(raw); return shapedRows; };
  return { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls };
}

module.exports = (async () => {
  /* ===== 1. Wrong document type: never retried, even with a gap-flagged row — checked by pageConfig.key
     alone, shape() isn't even consulted. ===== */
  {
    const { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls } = makeMock([{ flagged: true, flagReason: GAP_REASON }], null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'production_sheet', shape };
    const result = { raw: { items: [] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'WRONG TYPE: result passed through unchanged');
    t.assertEqual(shapeCalls.length, 0, 'WRONG TYPE: shape() never even called');
    t.assertEqual(retryCalls.length, 0, 'WRONG TYPE: never retries');
  }

  /* ===== 2. consumption_sheet, truncated: never retried regardless of flags — truncation already has
     its own, separate handling (see callClaudeExtract). ===== */
  {
    const { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls } = makeMock([{ flagged: true, flagReason: GAP_REASON }], null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const result = { raw: { items: [] }, truncated: true };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'TRUNCATED: result passed through unchanged');
    t.assertEqual(shapeCalls.length, 0, 'TRUNCATED: shape() never even called');
    t.assertEqual(retryCalls.length, 0, 'TRUNCATED: never retries');
  }

  /* ===== 3. consumption_sheet, not truncated, no row flagged at all: passed through unchanged. ===== */
  {
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock([{ flagged: false, flagReason: '' }, { flagged: false, flagReason: '' }], null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const result = { raw: { items: [] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'NO FLAGS: result passed through unchanged');
    t.assertEqual(retryCalls.length, 0, 'NO FLAGS: never retries');
  }

  /* ===== 4. consumption_sheet, not truncated, a row flagged but for a DIFFERENT reason (not the gap
     marker): never retried — that kind of flag needs a person, not another API call. ===== */
  {
    const { extractWithRateLimitBackoff, shape, retryCalls } = makeMock([{ flagged: true, flagReason: OTHER_REASON }], null);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const result = { raw: { items: [] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, result, 'OTHER FLAG: result passed through unchanged');
    t.assertEqual(retryCalls.length, 0, 'OTHER FLAG: never retries for a non-gap flag reason');
  }

  /* ===== 5. consumption_sheet, not truncated, at least one row flagged with the gap marker (mixed in
     with an unflagged row and an other-reason-flagged row): retries exactly once, with the SAME base64
     and prompt, and returns the retried raw/truncated. ===== */
  {
    const retried = { raw: { orientation: 'upright', items: [{ sl_no: '1' }] }, truncated: false, extra: 'should be dropped' };
    const shapedRows = [
      { flagged: false, flagReason: '' },
      { flagged: true, flagReason: OTHER_REASON },
      { flagged: true, flagReason: GAP_REASON },
    ];
    const { extractWithRateLimitBackoff, shape, retryCalls, shapeCalls } = makeMock(shapedRows, retried);
    const fn = loadMaybeRetryForDroppedRows(extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet', shape };
    const result = { raw: { items: [{ sl_no: '1' }, { sl_no: '4' }] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(shapeCalls, [result.raw], 'GAP FLAG: shape() was called on the ORIGINAL raw to check for the marker');
    t.assertEqual(retryCalls.length, 1, 'GAP FLAG: retries exactly once');
    t.assertEqual(retryCalls[0].base64, 'B64', 'GAP FLAG: retries with the SAME (already-corrected) image');
    t.assertEqual(retryCalls[0].prompt, 'PROMPT', 'GAP FLAG: retries with the exact same prompt');
    t.assertEqual(out, { raw: retried.raw, truncated: retried.truncated }, 'GAP FLAG: returns only the retried raw/truncated, nothing extra');
  }

  return t.summary();
})();
