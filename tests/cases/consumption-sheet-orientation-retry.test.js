// Real, directly-confirmed root cause (2026-11-18): rendering real uploaded consumption-sheet PDFs
// showed every page scanned with one phone app ("Adobe Scan") comes in physically rotated 90° from
// upright, while pages scanned a different way (older files, same register) are already upright.
// Reading a sideways table is far more error-prone for the vision model than reading an upright one —
// this is the actual, demonstrated mechanism behind rows silently dropping/misaligning and decimals
// going missing, not just a plausible theory. consumption_sheet's own systemPrompt now asks the model
// to report "orientation" BEFORE reading any row data, and to return an empty "items" array whenever
// it isn't "upright" rather than attempt a sideways read. maybeCorrectOrientationAndRetry (the code
// under test here) is what turns that report into an actual fix: it physically rotates the SAME image
// right-side up and re-sends it for a real extraction pass, exactly once. Every other document type's
// raw response never has an "orientation" field, so this is a no-op for all of them.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff) {
  const blocks = [
    extractLine(APP_JSX, 'const ORIENTATION_TO_ROTATE_DEG = {'),
    extractLine(APP_JSX, 'const MAX_ORIENTATION_ATTEMPTS = 3;'),
    extractBlock(APP_JSX, 'const maybeCorrectOrientationAndRetry = async (', (l) => l.trim() === '};'),
  ];
  return buildFromSource(blocks, 'maybeCorrectOrientationAndRetry', ['rotateBase64Image', 'extractWithRateLimitBackoff'], [rotateBase64Image, extractWithRateLimitBackoff]);
}

const t = makeRecorder('consumption-sheet-orientation-retry');

function makeMocks(retriedResult) {
  const rotateCalls = [];
  const retryCalls = [];
  const rotateBase64Image = async (base64, deg) => { rotateCalls.push({ base64, deg }); return `ROTATED(${base64},${deg})`; };
  const extractWithRateLimitBackoff = async (prompt, base64, signal, onWaiting, onTick) => {
    retryCalls.push({ prompt, base64, signal, onWaiting, onTick });
    return retriedResult;
  };
  return { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls };
}

module.exports = (async () => {
  /* ===== 1. Wrong document type: even with a non-upright orientation on the raw payload, a page
     that isn't consumption_sheet is returned completely untouched — no rotate, no retry call. ===== */
  {
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls } = makeMocks(null);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const pageConfig = { key: 'production_sheet' };
    const result = { raw: { orientation: 'rotate_cw', items: [] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, { ...result, base64: 'B64' }, 'WRONG TYPE: result passed through unchanged');
    t.assertEqual(rotateCalls.length, 0, 'WRONG TYPE: never rotates');
    t.assertEqual(retryCalls.length, 0, 'WRONG TYPE: never retries');
  }

  /* ===== 2. consumption_sheet, orientation "upright": passed through unchanged, no rotate/retry —
     this is the common case and must cost nothing extra. ===== */
  {
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls } = makeMocks(null);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet' };
    const result = { raw: { orientation: 'upright', items: [{ sl_no: '1' }] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, { ...result, base64: 'B64' }, 'UPRIGHT: result passed through unchanged');
    t.assertEqual(rotateCalls.length, 0, 'UPRIGHT: never rotates');
    t.assertEqual(retryCalls.length, 0, 'UPRIGHT: never retries');
  }

  /* ===== 3. consumption_sheet, no orientation field at all (defensive: an older/unexpected raw shape)
     -- never crashes, treated like upright. ===== */
  {
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls } = makeMocks(null);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet' };
    const result = { raw: { items: [{ sl_no: '1' }] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, { ...result, base64: 'B64' }, 'NO FIELD: result passed through unchanged, no crash');
    t.assertEqual(rotateCalls.length, 0, 'NO FIELD: never rotates');
  }

  /* ===== 4. consumption_sheet, "rotate_cw": rotates +90, re-extracts with the SAME prompt/signal on
     the ROTATED image, and returns the retried raw/truncated plus the new base64. ===== */
  {
    const retried = { raw: { orientation: 'upright', items: [{ sl_no: '1' }, { sl_no: '2' }] }, truncated: false };
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls } = makeMocks(retried);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const pageConfig = { key: 'consumption_sheet' };
    const result = { raw: { orientation: 'rotate_cw', items: [] }, truncated: false };
    const out = await fn(pageConfig, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(rotateCalls, [{ base64: 'B64', deg: 90 }], 'ROTATE_CW: rotates the original image by +90');
    t.assertEqual(retryCalls.length, 1, 'ROTATE_CW: retries exactly once');
    t.assertEqual(retryCalls[0].base64, 'ROTATED(B64,90)', 'ROTATE_CW: retries with the ROTATED image, not the original');
    t.assertEqual(retryCalls[0].prompt, 'PROMPT', 'ROTATE_CW: retries with the exact same prompt');
    t.assertEqual(out, { raw: retried.raw, truncated: retried.truncated, base64: 'ROTATED(B64,90)' }, 'ROTATE_CW: returns the retried raw/truncated plus the rotated base64');
  }

  /* ===== 5. "rotate_ccw": rotates -90. ===== */
  {
    const retried = { raw: { orientation: 'upright', items: [] }, truncated: false };
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls } = makeMocks(retried);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const out = await fn({ key: 'consumption_sheet' }, { raw: { orientation: 'rotate_ccw' }, truncated: false }, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(rotateCalls, [{ base64: 'B64', deg: -90 }], 'ROTATE_CCW: rotates the original image by -90');
    t.assertEqual(out.base64, 'ROTATED(B64,-90)', 'ROTATE_CCW: returns the rotated base64');
  }

  /* ===== 6. "upside_down": rotates 180. ===== */
  {
    const retried = { raw: { orientation: 'upright', items: [] }, truncated: false };
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls } = makeMocks(retried);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const out = await fn({ key: 'consumption_sheet' }, { raw: { orientation: 'upside_down' }, truncated: false }, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(rotateCalls, [{ base64: 'B64', deg: 180 }], 'UPSIDE_DOWN: rotates the original image by 180');
    t.assertEqual(out.base64, 'ROTATED(B64,180)', 'UPSIDE_DOWN: returns the rotated base64');
  }

  /* ===== 7. Unrecognized orientation string (model hallucinated something else): never guessed at,
     passed through unchanged rather than rotated blindly. ===== */
  {
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls } = makeMocks(null);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const result = { raw: { orientation: 'sideways-ish', items: [{ sl_no: '1' }] }, truncated: false };
    const out = await fn({ key: 'consumption_sheet' }, result, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(out, { ...result, base64: 'B64' }, 'UNRECOGNIZED: passed through unchanged, not guessed at');
    t.assertEqual(rotateCalls.length, 0, 'UNRECOGNIZED: never rotates');
    t.assertEqual(retryCalls.length, 0, 'UNRECOGNIZED: never retries');
  }

  /* ===== 8. The retried call keeps reporting non-upright, call after call (e.g. a genuinely
     unreadable/skewed page) — bounded at MAX_ORIENTATION_ATTEMPTS (3) corrections so this can't loop
     forever, and whatever the LAST attempt returned is used as final. ===== */
  {
    const neverUpright = { raw: { orientation: 'rotate_cw', items: [] }, truncated: false };
    const { rotateBase64Image, extractWithRateLimitBackoff, rotateCalls, retryCalls } = makeMocks(neverUpright);
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const out = await fn({ key: 'consumption_sheet' }, { raw: { orientation: 'rotate_ccw' }, truncated: false }, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(rotateCalls.length, 3, 'CAPPED: stops after MAX_ORIENTATION_ATTEMPTS (3) corrections, not an infinite loop');
    t.assertEqual(retryCalls.length, 3, 'CAPPED: re-extracts 3 times, no more');
    t.assertEqual(out.raw, neverUpright.raw, 'CAPPED: the last (still-imperfect) attempt is used as final, not discarded');
  }

  /* ===== 9. REAL CASE CAPTURED LIVE (2026-11-20, raw API traffic from an actual upload): the model
     reported "rotate_cw" on the original, then — on that corrected copy — reported "upside_down"
     instead of "upright" (a wrong first guess, or a page whose true skew wasn't a clean 90°). The OLD
     single-retry code would have stopped right there and handed back that still-sideways, zero-row
     result as final. The loop must instead rotate AGAIN off the new report and keep going until it
     actually reaches upright. ===== */
  {
    const results = [
      { raw: { orientation: 'upside_down', items: [] }, truncated: false },   // after 1st correction: still wrong
      { raw: { orientation: 'upright', items: [{ sl_no: '1' }] }, truncated: false }, // after 2nd correction: fixed
    ];
    let callIdx = 0;
    const rotateCalls = [];
    const retryCalls = [];
    const rotateBase64Image = async (base64, deg) => { rotateCalls.push({ base64, deg }); return `ROTATED(${base64},${deg})`; };
    const extractWithRateLimitBackoff = async (prompt, base64) => { retryCalls.push({ base64 }); return results[callIdx++]; };
    const fn = loadMaybeCorrectOrientationAndRetry(rotateBase64Image, extractWithRateLimitBackoff);
    const out = await fn({ key: 'consumption_sheet' }, { raw: { orientation: 'rotate_cw', items: [] }, truncated: false }, 'B64', 'PROMPT', 'SIG');
    t.assertEqual(rotateCalls, [{ base64: 'B64', deg: 90 }, { base64: 'ROTATED(B64,90)', deg: 180 }], 'CONVERGES: 2nd rotation is applied to the already-rotated image, by the newly reported delta — not re-rotating the original from scratch');
    t.assertEqual(retryCalls.length, 2, 'CONVERGES: took 2 corrections to reach upright, both actually used');
    t.assertEqual(out.raw.items, [{ sl_no: '1' }], 'CONVERGES: the final (now-upright, real) extraction is returned, not the still-sideways first retry');
    t.assertEqual(out.base64, 'ROTATED(ROTATED(B64,90),180)', 'CONVERGES: returns the final, fully-corrected image');
  }

  return t.summary();
})();
