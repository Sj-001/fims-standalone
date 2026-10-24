// A row queued because a push for the same customer was already in flight gets resent automatically
// once that in-flight push finishes (pushQueuedRef/pushInFlightRef's own retry mechanism, by design --
// see pushCustomerSheetNow's own comment). The gap: the retry never checked whether its own queued rows
// had, in the meantime, already been covered by the push it waited behind -- e.g. a bulk push covering
// rows [A, B, C] is still in flight (rows aren't marked pushed until the server responds, so they're
// still clickable), and row B also gets queued via a second push attempt while the bulk push is still
// running. Once the bulk push succeeds, the retry fires with its queued rows exactly as captured --
// including B, which the bulk push it waited behind already wrote. Real, confirmed consequence: a
// second write for the same row in the real Sheet. Fixed: pushedIdsRef tracks ids the moment the server
// actually confirms them (synchronously, via a ref -- not React state, which wouldn't have propagated
// into this same call's own closure yet), and the retry filters its queued rows against it before
// resending, dropping only what's now stale -- never sweeping in anything extra (that's the exact
// opposite bug pushCustomerSheetNow's own top comment already guards against).
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

const BLOCK = extractBlock(
  APP_JSX,
  'const pushCustomerSheetNow = async (customer, productionRows, dispatchRows, routingOverrides = {}, onConfirmed) => {',
  '  // The only way to push has always been confirming a PENDING row',
  { endExclusive: true },
);

const t = makeRecorder('push-queued-retry-dedup');

function deferred() {
  let resolve;
  const promise = new Promise((res) => { resolve = res; });
  return { promise, resolve };
}

function makeEnv() {
  const fetchCalls = []; // { body, resolve } -- each push's own fetch, resolved under test control
  const pushInFlightRef = { current: new Set() };
  const pushQueuedRef = { current: new Map() };
  const pushedIdsRef = { current: new Set() };
  const getCustomerSheetId = () => 'sheet-id-123';
  const normalizeForCatalogMatch = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const verifyAgainstFreshBlocks = async (customer, itemGroups) => ({ safeItemGroups: itemGroups, newlyUnmatched: [] });
  // None of this file's scenarios exercise the catch block (every fetchMock resolves, never rejects),
  // so this is never actually called -- present only so buildFromSource has a value for the name.
  const verifyRowsActuallyWritten = async () => ({ verifiedIndexes: [], mirrorRows: null });
  const setPushStatus = () => {};
  const setProduction = () => {};
  const setCustomerDispatch = () => {};
  const persist = () => {};
  const confirmedProductionRows = [];
  const confirmedDispatchRows = [];
  const replaceCustomerMirrorRows = () => {};
  const updateKnownCountsFromMirrorRows = () => {};
  const refreshReview = () => {};
  const patchCustomerSheetEntry = () => {};
  const setReviewEdits = () => {};
  const windowMock = { confirm: () => true, dispatchEvent: () => {} };
  // Each row's own id becomes its tab name, so getEditedPayloadFromRows (mocked below) can build a
  // distinct, individually-trackable wireItemGroups entry per row, same as the real app sending one
  // tab per real block.
  const getEditedPayloadFromRows = (customer, productionRows, dispatchRows) => ({
    itemGroups: productionRows.map(r => ({
      tabName: r.id,
      variants: [{ title: r.id, header: [], rows: [[null, 0, 1, null, 1]], rowSourceIds: [{ productionIds: [r.id], dispatchIds: [] }] }],
    })),
    unmatched: [],
  });
  const fetchMock = async (url, opts) => {
    const d = deferred();
    fetchCalls.push({ body: JSON.parse(opts.body), resolve: d.resolve });
    return d.promise;
  };

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
      normalizeForCatalogMatch, verifyAgainstFreshBlocks, verifyRowsActuallyWritten, setPushStatus, setProduction,
      setCustomerDispatch, persist, confirmedProductionRows, confirmedDispatchRows,
      replaceCustomerMirrorRows, updateKnownCountsFromMirrorRows, refreshReview,
      patchCustomerSheetEntry, setReviewEdits, windowMock, fetchMock,
    ],
  );
  return { fn, fetchCalls };
}

async function flushMicrotasks(times = 6) {
  for (let i = 0; i < times; i++) await Promise.resolve();
}

function okResponseFor(wireItemGroups) {
  return { status: 200, ok: true, json: async () => ({ ok: true, results: wireItemGroups.map(g => ({ tab: g.tabName, ok: true, newRows: 1 })), mirrorRows: [] }) };
}

// CommonJS has no top-level await -- the whole suite runs inside one async IIFE, and run-all.js already
// knows how to await a test file's exported Promise.
module.exports = (async () => {

/* ===== THE LIVE CASE: row B is part of an in-flight bulk push (A, B, C); a second push for just [B]
   fires while that's still running, so it queues; once the bulk push succeeds, B must NOT be resent --
   the bulk push already covered it. ===== */
await (async () => {
  const { fn, fetchCalls } = makeEnv();
  const rowA = { id: 'A', description: 'A' };
  const rowB = { id: 'B', description: 'B' };
  const rowC = { id: 'C', description: 'C' };

  const bulkOnConfirmedCalls = [];
  const p1 = fn('DIAMOND', [rowA, rowB, rowC], [], {}, (list) => bulkOnConfirmedCalls.push(list));
  await flushMicrotasks();
  t.assertEqual(fetchCalls.length, 1, 'LIVE CASE: the bulk push (A, B, C) fired its own fetch');

  // While that's still in flight (no response yet), a second push for just row B is attempted -- must
  // queue, not fire its own fetch, since DIAMOND is already in flight.
  const soloOnConfirmedCalls = [];
  const p2 = fn('DIAMOND', [rowB], [], {}, (list) => soloOnConfirmedCalls.push(list));
  await flushMicrotasks();
  t.assertEqual(fetchCalls.length, 1, 'LIVE CASE: the second push for row B queued instead of firing its own overlapping fetch');

  // The bulk push now succeeds -- A, B, and C are all written for real.
  fetchCalls[0].resolve(okResponseFor([{ tabName: 'A' }, { tabName: 'B' }, { tabName: 'C' }]));
  await p1;
  await flushMicrotasks();

  t.assertEqual(fetchCalls.length, 1, 'LIVE CASE: once the bulk push succeeds, row B\'s queued retry does NOT fire a second fetch -- it was already covered');
  t.assertEqual(soloOnConfirmedCalls.length, 1, 'LIVE CASE: row B\'s own onConfirmed still fires (so its local UI state settles), even with nothing left to send');
  t.assertEqual(soloOnConfirmedCalls[0], [], 'LIVE CASE: row B is reported as cleanly confirmed, not withheld -- it genuinely did get written, just by the push it waited behind');
  await p2;
})();

/* ===== SAFETY: if the queued row was NOT part of what the in-flight push covered, the retry must still
   send it for real -- this fix only drops rows that are now genuinely stale, never anything else. =====
   */
await (async () => {
  const { fn, fetchCalls } = makeEnv();
  const rowA = { id: 'A', description: 'A' };
  const rowD = { id: 'D', description: 'D' }; // a genuinely different row, never part of the first push

  const p1 = fn('DIAMOND', [rowA], [], {}, () => {});
  await flushMicrotasks();

  const soloOnConfirmedCalls = [];
  const p2 = fn('DIAMOND', [rowD], [], {}, (list) => soloOnConfirmedCalls.push(list));
  await flushMicrotasks();
  t.assertEqual(fetchCalls.length, 1, 'SAFETY: row D queued behind the in-flight push for A, same as before');

  fetchCalls[0].resolve(okResponseFor([{ tabName: 'A' }]));
  await p1;
  await flushMicrotasks();

  t.assertEqual(fetchCalls.length, 2, 'SAFETY: row D\'s retry DOES fire its own real fetch -- it was never covered by the push it waited behind');
  t.assertEqual(fetchCalls[1].body.itemGroups.map(g => g.tabName), ['D'], 'SAFETY: the retry\'s payload is exactly row D, nothing swept in, nothing dropped that shouldn\'t be');

  fetchCalls[1].resolve(okResponseFor([{ tabName: 'D' }]));
  await p2;
  await flushMicrotasks();
  t.assertEqual(soloOnConfirmedCalls[0], [], 'SAFETY: row D is confirmed normally once its own real push succeeds');
})();

return t.summary();
})();
