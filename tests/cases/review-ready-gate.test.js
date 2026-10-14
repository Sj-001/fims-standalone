// Architectural fix, 2026-10-14: the live screenshot evidence showed two symptoms of the SAME root
// cause -- (1) a real production push creating a duplicate block ("Vishesh 100 ML") right next to the
// correct existing one ("100 ML") that the app itself had matched, and (2) a test-environment Pending
// Review showing EVERY row as needing a manual Tab/Block pick, including ones with a perfectly good,
// already-used catalog mapping. Root cause: reviewByCustomer[customer] starts as `undefined` (never
// fetched) and getRealBlocksForTab returns [] identically for "never fetched" and "genuinely empty,
// confirmed real" -- so a fresh page load, or a push that lands before the debounced background
// refreshReview effect completes, can't tell "still checking" apart from "verified, no real block
// exists". isReviewReadyForCustomer makes that distinction explicit so the UI can show a neutral
// "(checking…)" state and pushPendingRows can withhold the row entirely, instead of either screen
// treating "haven't looked yet" as "confirmed nothing's there".
const path = require('path');
const { extractBlock, extractLine } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadIsReviewReadyForCustomer() {
  const block = extractBlock(
    APP_JSX,
    'const isReviewReadyForCustomer = (customer) => {',
    (l) => l.trim() === '};',
  );
  const body = `${block}\nreturn isReviewReadyForCustomer;`;
  const fn = new Function('getCustomerSheetId', 'reviewByCustomer', body);
  return fn;
}

const t = makeRecorder('review-ready-gate');
const buildIsReviewReadyForCustomer = loadIsReviewReadyForCustomer();

function isReady(customer, { sheetId = 'sheet-123', reviewByCustomer = {} } = {}) {
  const getCustomerSheetId = (c) => (c === customer ? sheetId : '');
  const fn = buildIsReviewReadyForCustomer(getCustomerSheetId, reviewByCustomer);
  return fn(customer);
}

/* ===== 1. No customer at all -- never ready (mirrors every other "Unassigned" guard in this codebase). */
{
  t.assertEqual(isReady(''), false, 'EMPTY: no customer is never ready');
}

/* ===== 2. A customer with no linked Sheet at all has nothing to check against -- always ready, matching
   how pushPendingRows already commits straight away for a customer with no getCustomerSheetId. ===== */
{
  t.assertEqual(isReady('DIAMOND', { sheetId: '' }), true, 'NO SHEET: always ready, nothing to verify against');
}

/* ===== 3. THE LIVE BUG: reviewByCustomer never fetched yet for this customer -- NOT ready, must show
   "(checking…)", not a confirmed blank/needs-picker state. ===== */
{
  t.assertEqual(isReady('DIAMOND', { reviewByCustomer: {} }), false, 'NEVER FETCHED: not ready -- this is the exact gap that let a correct match get treated as unmatched');
}

/* ===== 4. A fetch is actively in flight (refreshReview set loading:true) -- still not ready. ===== */
{
  t.assertEqual(isReady('DIAMOND', { reviewByCustomer: { DIAMOND: { loading: true } } }), false, 'IN FLIGHT: not ready while a refresh is actively running');
}

/* ===== 5. Fetch completed (whatever the result) -- ready, the real verified state can now be trusted. */
{
  t.assertEqual(isReady('DIAMOND', { reviewByCustomer: { DIAMOND: { loading: false, existingBlockTitles: ['100 ML'] } } }), true, 'LOADED: ready once the real fetch has actually completed');
}
{
  t.assertEqual(isReady('DIAMOND', { reviewByCustomer: { DIAMOND: { loading: false, existingBlockTitles: [] } } }), true, 'LOADED, GENUINELY EMPTY: ready -- an empty result IS the confirmed real answer once loaded');
}

/* ===== 6. Multiple customers tracked independently -- one customer's loaded state never leaks into
   another's "never fetched" state. ===== */
{
  const reviewByCustomer = { DIAMOND: { loading: false, existingBlockTitles: ['100 ML'] } };
  t.assertEqual(isReady('DIAMOND', { reviewByCustomer }), true, 'ISOLATION: DIAMOND is ready (loaded)');
  t.assertEqual(isReady('ANMOL', { reviewByCustomer }), false, 'ISOLATION: ANMOL is still not ready (never fetched), unaffected by DIAMOND\'s state');
}

module.exports = t.summary();
