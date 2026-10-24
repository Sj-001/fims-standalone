// THE LIVE BUG, confirmed directly against the real test environment, 2026-10-24: pushCustomerSheetNow's
// catch block used to treat ANY network-level failure as "nothing happened" -- never calling onConfirmed,
// leaving the row stuck in Pending Review forever. But Google Sheets writes take a few seconds, long
// enough for the connection to drop (or a cold-started free-tier Render dyno to time out) AFTER the
// server already finished the write. Confirmed directly: "M.D PLAIN" 22.10.26, 538 pieces, genuinely
// present in the real VISHESH STOCK Sheet, permanently stuck pending in the app with no way back.
//
// The fix: on a network-level catch, re-read the real Sheet fresh and check whether the SPECIFIC rows
// this push was trying to send are actually there now, before accepting the failure as final. Two parts
// tested here: computeVerifiedIndexesFromMirror (pure matching logic) and the full catch-block recovery
// flow in pushCustomerSheetNow (mocked fetch/verify).
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

const t = makeRecorder('push-network-error-recovery');

/* ========================================================================
   PART 1: computeVerifiedIndexesFromMirror (pure, no network involved)
   ======================================================================== */
function loadComputeVerifiedIndexes() {
  const block = extractBlock(
    APP_JSX,
    'const computeVerifiedIndexesFromMirror = (wireItemGroups, mirrorRows) => {',
    (l) => l.trim() === '};',
  );
  const normalizeForCatalogMatch = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  return buildFromSource(block, 'computeVerifiedIndexesFromMirror', ['normalizeForCatalogMatch'], [normalizeForCatalogMatch]);
}

// A single-variant, single-row tab claiming only production (dispatch null) -- the exact live shape.
function productionOnlyGroup(tabName, block, date, production) {
  return { tabName, variants: [{ title: block, rows: [[date, 0, production, null, production]] }] };
}

{
  const computeVerifiedIndexesFromMirror = loadComputeVerifiedIndexes();

  // ===== 1. THE LIVE CASE: a production-only row exactly matches the fresh mirror -- verified. =====
  {
    const wireItemGroups = [productionOnlyGroup('M.D', 'M.D PLAIN', '22.10.26', 538)];
    const mirrorRows = [{ sheetTab: 'M.D', block: 'M.D PLAIN', date: '22.10.26', production: 538, dispatch: 0 }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [0], 'LIVE CASE: exact tab+block+date+production match verifies the tab');
  }

  // ===== 2. SAFETY: mirror has the same tab+block+date but a DIFFERENT production value -- not verified. =====
  {
    const wireItemGroups = [productionOnlyGroup('M.D', 'M.D PLAIN', '22.10.26', 538)];
    const mirrorRows = [{ sheetTab: 'M.D', block: 'M.D PLAIN', date: '22.10.26', production: 999, dispatch: 0 }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [], 'SAFETY: a mismatched production value never counts as verified, even for the same tab/block/date');
  }

  // ===== 3. SAFETY: no matching tab/block/date in the mirror at all -- not verified (the row genuinely
  //    was never written, the real "the network truly failed" case). =====
  {
    const wireItemGroups = [productionOnlyGroup('M.D', 'M.D PLAIN', '22.10.26', 538)];
    const mirrorRows = [{ sheetTab: 'M.D', block: 'M.D PLAIN', date: '25.9.26', production: 770, dispatch: 0 }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [], 'SAFETY: a genuinely different date for the same block is never treated as verifying this row');
  }

  // ===== 4. SAFETY: empty mirror (e.g. the verify re-check itself returned nothing) -- never verifies
  //    anything, which is indistinguishable from this fix not existing at all. =====
  {
    const wireItemGroups = [productionOnlyGroup('M.D', 'M.D PLAIN', '22.10.26', 538)];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, []);
    t.assertEqual(result, [], 'SAFETY: an empty mirror read verifies nothing');
  }

  // ===== 5. Multi-row tab: ALL rows in a tab must match for the tab to verify -- one row missing from
  //    the mirror fails the WHOLE tab, same all-or-nothing granularity the real server response uses
  //    per tab (one verdict per tab, never per row). =====
  {
    const wireItemGroups = [{
      tabName: 'M.D',
      variants: [{ title: 'M.D PLAIN', rows: [
        ['22.10.26', 0, 538, null, 538],
        ['23.10.26', 538, 100, null, 638],
      ] }],
    }];
    const mirrorRows = [
      { sheetTab: 'M.D', block: 'M.D PLAIN', date: '22.10.26', production: 538, dispatch: 0 },
      // 23.10.26 row never landed in the mirror -- genuinely still outstanding
    ];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [], 'SAFETY: one unverified row in a tab means the whole tab stays unverified');
  }

  // ===== 6. Multiple independent tabs: each verified separately -- one tab's failure never affects
  //    another tab's verdict. =====
  {
    const wireItemGroups = [
      productionOnlyGroup('M.D', 'M.D PLAIN', '22.10.26', 538),
      productionOnlyGroup('JEERA', 'Jeera Dhamal 32g x144', '23.10.26', 5900),
    ];
    const mirrorRows = [
      { sheetTab: 'M.D', block: 'M.D PLAIN', date: '22.10.26', production: 538, dispatch: 0 },
      // JEERA tab's row never landed
    ];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [0], 'Tab 0 (M.D) verifies independently of tab 1 (JEERA) staying unverified');
  }

  // ===== 7. A dispatch-only row (production null) verifies against the mirror's dispatch field,
  //    completely ignoring whatever the mirror's production value happens to be. =====
  {
    const wireItemGroups = [{ tabName: 'DIAMOND', variants: [{ title: 'Butter Plus', rows: [['10.10.26', 100, null, 40, 60]] }] }];
    const mirrorRows = [{ sheetTab: 'DIAMOND', block: 'Butter Plus', date: '10.10.26', production: 9999, dispatch: 40 }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [0], 'A dispatch-only claim verifies on dispatch alone, production is irrelevant to it');
  }

  // ===== 8. A dual-column row (both production AND dispatch claimed) must match BOTH on the SAME
  //    mirror row -- a mirror row matching only one of the two fields is not good enough. =====
  {
    const wireItemGroups = [{ tabName: 'DIAMOND', variants: [{ title: 'Butter Plus', rows: [['10.10.26', 100, 50, 40, 110]] }] }];
    // Mirror has a row with the right production but wrong dispatch -- must NOT verify.
    const mirrorRowsWrongDispatch = [{ sheetTab: 'DIAMOND', block: 'Butter Plus', date: '10.10.26', production: 50, dispatch: 999 }];
    t.assertEqual(computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRowsWrongDispatch), [], 'SAFETY: right production, wrong dispatch on the same row -- not verified');
    // Mirror has a row matching both -- verifies.
    const mirrorRowsBothMatch = [{ sheetTab: 'DIAMOND', block: 'Butter Plus', date: '10.10.26', production: 50, dispatch: 40 }];
    t.assertEqual(computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRowsBothMatch), [0], 'Both production and dispatch matching the same mirror row verifies');
  }

  // ===== 9. A row claiming NEITHER production nor dispatch (a pure opening/closing scaffolding entry)
  //    has nothing to verify and is trivially treated as fine -- it never blocks a tab from verifying. =====
  {
    const wireItemGroups = [{ tabName: 'M.D', variants: [{ title: 'M.D PLAIN', rows: [['1.1.26', 0, null, null, 0]] }] }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, []);
    t.assertEqual(result, [0], 'A row claiming nothing in either column is trivially verified, even against an empty mirror');
  }

  // ===== 10. Block/tab names matching only after normalization (case/punctuation) still verify -- same
  //     normalizeForCatalogMatch tolerance already used for every other tab/block comparison in this app. =====
  {
    const wireItemGroups = [productionOnlyGroup('m.d', 'M.D. PLAIN', '22.10.26', 538)];
    const mirrorRows = [{ sheetTab: 'M.D', block: 'M D Plain', date: '22.10.26', production: 538, dispatch: 0 }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [0], 'Tab/block names differing only in case and punctuation still match');
  }

  // ===== 11. blockTitleOverride, when present, is what gets matched against -- not the variant's own
  //     title -- mirroring how the real push payload resolves the actual block name sent to the server. =====
  {
    const wireItemGroups = [{
      tabName: 'M.D',
      variants: [{ title: 'M.D PLAIN (raw description)', blockTitleOverride: 'M.D PLAIN', rows: [['22.10.26', 0, 538, null, 538]] }],
    }];
    const mirrorRows = [{ sheetTab: 'M.D', block: 'M.D PLAIN', date: '22.10.26', production: 538, dispatch: 0 }];
    const result = computeVerifiedIndexesFromMirror(wireItemGroups, mirrorRows);
    t.assertEqual(result, [0], 'blockTitleOverride is used for matching, not the raw variant title');
  }
}

/* ========================================================================
   PART 2: full pushCustomerSheetNow catch-block recovery flow
   ======================================================================== */
function loadPushFn({ verifyRowsActuallyWrittenMock, fetchMock }) {
  const BLOCK = extractBlock(
    APP_JSX,
    'const pushCustomerSheetNow = async (customer, productionRows, dispatchRows, routingOverrides = {}, onConfirmed) => {',
    '  // The only way to push has always been confirming a PENDING row',
    { endExclusive: true },
  );
  const getCustomerSheetId = () => 'sheet-id-123';
  const pushInFlightRef = { current: new Set() };
  const pushQueuedRef = { current: new Map() };
  const pushedIdsRef = { current: new Set() };
  const normalizeForCatalogMatch = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const verifyAgainstFreshBlocks = async (customer, itemGroups) => ({ safeItemGroups: itemGroups, newlyUnmatched: [] });
  const pushStatuses = [];
  const setPushStatus = (fn) => pushStatuses.push(fn({}));
  const productionState = { current: [] };
  const setProduction = (updater) => { productionState.current = typeof updater === 'function' ? updater(productionState.current) : updater; };
  const setCustomerDispatch = () => {};
  const persist = () => {};
  const confirmedProductionRows = [];
  const confirmedDispatchRows = [];
  const mirrorCalls = [];
  const replaceCustomerMirrorRows = (customer, rows) => mirrorCalls.push({ customer, rows });
  const updateKnownCountsFromMirrorRows = () => {};
  const refreshReview = () => {};
  const patchCustomerSheetEntry = () => {};
  const setReviewEdits = () => {};
  const windowMock = { confirm: () => true, dispatchEvent: () => {} };
  // One tab per row, named by the row's own id -- same convention as push-queued-retry-dedup.test.js.
  const getEditedPayloadFromRows = (customer, productionRows) => ({
    itemGroups: productionRows.map(r => ({
      tabName: r.id,
      variants: [{ title: r.id, rows: [[r.date || '1.1.26', 0, r.pieces || 1, null, r.pieces || 1]], rowSourceIds: [{ productionIds: [r.id], dispatchIds: [] }] }],
    })),
    unmatched: [],
  });

  const fn = buildFromSource(
    BLOCK, 'pushCustomerSheetNow',
    [
      'getCustomerSheetId', 'pushInFlightRef', 'pushQueuedRef', 'pushedIdsRef', 'getEditedPayloadFromRows',
      'normalizeForCatalogMatch', 'verifyAgainstFreshBlocks', 'verifyRowsActuallyWritten', 'setPushStatus', 'setProduction',
      'setCustomerDispatch', 'persist', 'confirmedProductionRows', 'confirmedDispatchRows',
      'replaceCustomerMirrorRows', 'updateKnownCountsFromMirrorRows', 'refreshReview',
      'patchCustomerSheetEntry', 'setReviewEdits', 'window', 'fetch',
    ],
    [
      getCustomerSheetId, pushInFlightRef, pushQueuedRef, pushedIdsRef, getEditedPayloadFromRows,
      normalizeForCatalogMatch, verifyAgainstFreshBlocks, verifyRowsActuallyWrittenMock, setPushStatus, setProduction,
      setCustomerDispatch, persist, confirmedProductionRows, confirmedDispatchRows,
      replaceCustomerMirrorRows, updateKnownCountsFromMirrorRows, refreshReview,
      patchCustomerSheetEntry, setReviewEdits, windowMock, fetchMock,
    ],
  );
  return { fn, pushStatuses, productionState, mirrorCalls };
}

const networkErrorFetch = async () => { throw new Error('Network error — could not reach the server.'); };

module.exports = (async () => {

// ===== 12. THE LIVE CASE, fixed: fetch fails, but the re-check finds the row genuinely already
//    written -- onConfirmed fires (so the row leaves Pending Review), pushedToSheet gets set, and the
//    status is reported as done, not stuck in error. =====
{
  const onConfirmedCalls = [];
  const verifyRowsActuallyWrittenMock = async (customer, wireItemGroups) => ({
    verifiedIndexes: wireItemGroups.map((g, i) => i), // everything verified
    mirrorRows: [{ sheetTab: 'r1', block: 'r1', date: '22.10.26', production: 538, dispatch: 0 }],
  });
  const { fn, pushStatuses, productionState } = loadPushFn({ verifyRowsActuallyWrittenMock, fetchMock: networkErrorFetch });
  await fn('VISHESH STOCK', [{ id: 'r1', date: '22.10.26', pieces: 538, description: 'M.D PLAIN' }], [], {}, (list) => onConfirmedCalls.push(list));

  t.assertEqual(onConfirmedCalls.length, 1, 'LIVE CASE FIXED: onConfirmed still fires despite the network error, once the re-check confirms the write');
  t.assertEqual(onConfirmedCalls[0], [], 'LIVE CASE FIXED: nothing withheld -- the row genuinely was written');
  const last = pushStatuses[pushStatuses.length - 1]['VISHESH STOCK'];
  t.assertEqual(last.state, 'done', 'LIVE CASE FIXED: status reports done, not stuck in error, once verified');
}

// ===== 13. SAFETY: fetch fails AND the re-check genuinely finds nothing (the real "actually failed"
//    case) -- falls through to EXACTLY the original behavior: onConfirmed never called, status is
//    error. This is the regression guard -- a true failure must still behave exactly as before. =====
{
  const onConfirmedCalls = [];
  const verifyRowsActuallyWrittenMock = async () => ({ verifiedIndexes: [], mirrorRows: null });
  const { fn, pushStatuses } = loadPushFn({ verifyRowsActuallyWrittenMock, fetchMock: networkErrorFetch });
  await fn('VISHESH STOCK', [{ id: 'r1', date: '22.10.26', pieces: 538, description: 'M.D PLAIN' }], [], {}, (list) => onConfirmedCalls.push(list));

  t.assertEqual(onConfirmedCalls.length, 0, 'SAFETY: onConfirmed is never called when the re-check also finds nothing -- a genuine failure behaves exactly as before this fix');
  const last = pushStatuses[pushStatuses.length - 1]['VISHESH STOCK'];
  t.assertEqual(last.state, 'error', 'SAFETY: status stays error for a genuine, unrecoverable failure');
}

// ===== 14. SAFETY: the re-check call itself throws (also offline, server still unreachable) -- must
//    never let that exception escape; falls through to the original error behavior, same as #13. =====
{
  const onConfirmedCalls = [];
  const verifyRowsActuallyWrittenMock = async () => { throw new Error('also unreachable'); };
  let threw = false;
  const { fn, pushStatuses } = loadPushFn({ verifyRowsActuallyWrittenMock, fetchMock: networkErrorFetch });
  try {
    await fn('VISHESH STOCK', [{ id: 'r1', date: '22.10.26', pieces: 538, description: 'M.D PLAIN' }], [], {}, (list) => onConfirmedCalls.push(list));
  } catch (e) { threw = true; }

  t.assertTrue(!threw, 'SAFETY: a throwing re-check must never escape pushCustomerSheetNow itself');
  t.assertEqual(onConfirmedCalls.length, 0, 'SAFETY: nothing confirmed when the re-check itself also fails outright');
  const last = pushStatuses[pushStatuses.length - 1]['VISHESH STOCK'];
  t.assertEqual(last.state, 'error', 'SAFETY: still reported as error, exactly as if this fix did not exist');
}

// ===== 15. PARTIAL RECOVERY: two rows pushed together, the re-check verifies only ONE of them --
//    the verified one is confirmed and marked pushed, the unverified one stays pending, and both facts
//    are visible in the status message. Nothing here ever upgrades an unverified row to confirmed just
//    because ITS SIBLING in the same push happened to verify. =====
{
  const onConfirmedCalls = [];
  const verifyRowsActuallyWrittenMock = async (customer, wireItemGroups) => ({
    verifiedIndexes: [0], // only the FIRST tab (r1) verifies; r2 does not
    mirrorRows: [{ sheetTab: 'r1', block: 'r1', date: '22.10.26', production: 538, dispatch: 0 }],
  });
  const { fn, pushStatuses, productionState } = loadPushFn({ verifyRowsActuallyWrittenMock, fetchMock: networkErrorFetch });
  await fn(
    'VISHESH STOCK',
    [
      { id: 'r1', date: '22.10.26', pieces: 538, description: 'M.D PLAIN' },
      { id: 'r2', date: '23.10.26', pieces: 100, description: 'Kaju Bake' },
    ],
    [], {},
    (list) => onConfirmedCalls.push(list),
  );

  t.assertEqual(onConfirmedCalls.length, 1, 'PARTIAL: onConfirmed still fires exactly once for the whole push');
  t.assertEqual(onConfirmedCalls[0], ['r2'], 'PARTIAL: only the unverified row (r2, by its tab title) is withheld -- r1 is not in this list');
  const last = pushStatuses[pushStatuses.length - 1]['VISHESH STOCK'];
  t.assertEqual(last.state, 'error', 'PARTIAL: overall status is still error while something remains genuinely unconfirmed');
  t.assertTrue(/r2/.test(last.message) && /1 item/.test(last.message), 'PARTIAL: the message names the one verified item and the one still-pending item');
}

// ===== 16. SAFETY: a push with no Google Sheet linked for the customer never even reaches the fetch,
//    so the recovery path is irrelevant there -- confirms this fix didn't disturb that earlier,
//    unrelated early-return. =====
{
  const onConfirmedCalls = [];
  const verifyRowsActuallyWrittenMock = async () => { throw new Error('should never be called'); };
  const BLOCK = extractBlock(
    APP_JSX,
    'const pushCustomerSheetNow = async (customer, productionRows, dispatchRows, routingOverrides = {}, onConfirmed) => {',
    '  // The only way to push has always been confirming a PENDING row',
    { endExclusive: true },
  );
  const fn = buildFromSource(
    BLOCK, 'pushCustomerSheetNow',
    [
      'getCustomerSheetId', 'pushInFlightRef', 'pushQueuedRef', 'pushedIdsRef', 'getEditedPayloadFromRows',
      'normalizeForCatalogMatch', 'verifyAgainstFreshBlocks', 'verifyRowsActuallyWritten', 'setPushStatus', 'setProduction',
      'setCustomerDispatch', 'persist', 'confirmedProductionRows', 'confirmedDispatchRows',
      'replaceCustomerMirrorRows', 'updateKnownCountsFromMirrorRows', 'refreshReview',
      'patchCustomerSheetEntry', 'setReviewEdits', 'window', 'fetch',
    ],
    [
      () => '', { current: new Set() }, { current: new Map() }, { current: new Set() }, () => ({ itemGroups: [], unmatched: [] }),
      (s) => s, async (c, g) => ({ safeItemGroups: g, newlyUnmatched: [] }), verifyRowsActuallyWrittenMock, () => {}, () => {},
      () => {}, () => {}, [], [],
      () => {}, () => {}, () => {},
      () => {}, () => {}, { confirm: () => true }, networkErrorFetch,
    ],
  );
  await fn('NO SHEET CUSTOMER', [{ id: 'r1' }], [], {}, (list) => onConfirmedCalls.push(list));
  t.assertEqual(onConfirmedCalls.length, 0, 'SAFETY: no Sheet ID means an early return before any of this fix\'s code runs at all');
}

return t.summary();
})();
