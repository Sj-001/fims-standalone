// Real gap, confirmed directly against the live isolated test environment, 2026-10-14: pushCustomerSheetNow
// used to call onConfirmed(unmatched) immediately after the confirm() dialog was accepted -- based ONLY on
// this function's own client-side pre-check (verifyAgainstFreshBlocks), BEFORE the server had even been
// asked. The exact live case: the client's pre-check said "Vishesh 100 ML" matched a real block fine, but
// the actual payload sent to the server (built from the same routingOverrides, via a since-fixed stale-
// draft bug -- see push-unmatched-not-committed.test.js) didn't match what the server's own, independent
// authorizeTabPlansForPush check required. The server correctly refused the whole tab -- but the row had
// ALREADY been marked confirmed a moment earlier (onConfirmed already fired) and silently vanished from
// Pending Review with no real Sheet write behind it, and no easy way back into view.
//
// Fixed by moving onConfirmed to fire only AFTER the real server response is known: on full success, with
// exactly the client pre-check's own unmatched list (nothing more to add, every tab it sent succeeded); on
// a server refusal, merged with every row belonging to a tab the server itself rejected (matched back to
// titles by POSITION -- same rename-proof convention idsToMarkByIndex already uses); on a network error,
// not at all, since nothing is known to have succeeded.
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

const t = makeRecorder('push-confirm-after-server-response');

// itemGroups shaped exactly like the real VISHESH case: one tab, one variant, titled with the row's own
// raw description -- matches what buildCustomerSheetPayloadFromRows actually produces.
const VISHESH_ITEM_GROUPS = [{
  tabName: 'VP PAPER CUP 1',
  variants: [{ title: 'Vishesh 100 ML', header: [], rows: [[null, 0, 300, null, 300]], rowSourceIds: [{ productionIds: ['r1'], dispatchIds: [] }] }],
}];

async function run({ fetchResponse, fetchThrows, preCheckUnmatched = [] }) {
  const onConfirmedCalls = [];
  const onConfirmed = (list) => onConfirmedCalls.push(list);
  const pushStatuses = [];

  const getCustomerSheetId = () => 'sheet-id-123';
  const pushInFlightRef = { current: new Set() };
  const pushQueuedRef = { current: new Map() };
  const pushedIdsRef = { current: new Set() };
  const getEditedPayloadFromRows = () => ({ itemGroups: VISHESH_ITEM_GROUPS, unmatched: preCheckUnmatched });
  const normalizeForCatalogMatch = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const verifyAgainstFreshBlocks = async (customer, itemGroups) => ({ safeItemGroups: itemGroups, newlyUnmatched: [] });
  const setPushStatus = (fn) => pushStatuses.push(fn({}));
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
  const fetchMock = async () => {
    if (fetchThrows) throw new Error('Network error — could not reach the server.');
    return { status: fetchResponse.status || 200, ok: fetchResponse.ok !== false, json: async () => fetchResponse.body };
  };

  const fn = buildFromSource(
    BLOCK, 'pushCustomerSheetNow',
    [
      'getCustomerSheetId', 'pushInFlightRef', 'pushQueuedRef', 'pushedIdsRef', 'getEditedPayloadFromRows',
      'normalizeForCatalogMatch', 'verifyAgainstFreshBlocks', 'setPushStatus', 'setProduction',
      'setCustomerDispatch', 'persist', 'confirmedProductionRows', 'confirmedDispatchRows',
      'replaceCustomerMirrorRows', 'updateKnownCountsFromMirrorRows', 'refreshReview',
      'patchCustomerSheetEntry', 'setReviewEdits', 'window', 'fetch',
    ],
    [
      getCustomerSheetId, pushInFlightRef, pushQueuedRef, pushedIdsRef, getEditedPayloadFromRows,
      normalizeForCatalogMatch, verifyAgainstFreshBlocks, setPushStatus, setProduction,
      setCustomerDispatch, persist, confirmedProductionRows, confirmedDispatchRows,
      replaceCustomerMirrorRows, updateKnownCountsFromMirrorRows, refreshReview,
      patchCustomerSheetEntry, setReviewEdits, windowMock, fetchMock,
    ],
  );

  await fn('VISHESH STOCK', [{ id: 'r1', description: 'Vishesh 100 ML' }], [], {}, onConfirmed);
  return { onConfirmedCalls };
}

module.exports = (async () => {
  // ===== THE LIVE CASE: the server refuses the tab -- onConfirmed must still be called (so the row
  // never gets silently stuck confirmed-but-unpushed), but WITH the row's own description included,
  // so pushPendingRows' commitAll correctly withholds it instead of confirming it anyway. =====
  {
    const { onConfirmedCalls } = await run({
      fetchResponse: { ok: false, status: 502, body: { ok: false, results: [{ tab: 'VP PAPER CUP 1', ok: false, error: 'Refused to create 1 new block without explicit confirmation: Vishesh 100 ML.' }] } },
    });
    t.assertEqual(onConfirmedCalls.length, 1, 'LIVE CASE: onConfirmed is still called exactly once (never left hanging)');
    t.assertEqual(onConfirmedCalls[0], ['Vishesh 100 ML'], 'LIVE CASE: the server-refused row is included, so it stays pending instead of being silently confirmed');
  }

  // ===== NORMAL SUCCESS: every tab the server was asked about came back ok -- onConfirmed fires with
  // exactly the client's own pre-check unmatched list (empty here), confirming the row normally. =====
  {
    const { onConfirmedCalls } = await run({
      fetchResponse: { ok: true, status: 200, body: { ok: true, results: [{ tab: 'VP PAPER CUP 1', ok: true, newRows: 1 }], mirrorRows: [] } },
    });
    t.assertEqual(onConfirmedCalls.length, 1, 'SUCCESS: onConfirmed is called once the server has actually agreed');
    t.assertEqual(onConfirmedCalls[0], [], 'SUCCESS: nothing withheld -- the row the server confirmed gets confirmed');
  }

  // ===== PRE-CHECK UNMATCHED STILL HONORED: a row the client's OWN pre-check already knew wasn't real
  // must still show up in the final onConfirmed list even on a clean server response for everything
  // else -- the merge must never accidentally drop something verifyAgainstFreshBlocks already caught. =====
  {
    const { onConfirmedCalls } = await run({
      preCheckUnmatched: ['Some Other Pre-Check-Failed Item'],
      fetchResponse: { ok: true, status: 200, body: { ok: true, results: [{ tab: 'VP PAPER CUP 1', ok: true, newRows: 1 }], mirrorRows: [] } },
    });
    t.assertEqual(onConfirmedCalls[0], ['Some Other Pre-Check-Failed Item'], 'PRE-CHECK: an item the client pre-check already flagged stays withheld even though the server push itself succeeded');
  }

  // ===== NETWORK ERROR: the fetch itself throws -- nothing is known to have succeeded, so onConfirmed
  // must never be called at all (pushPendingRows' commitAll then never runs -- nothing gets confirmed). =====
  {
    const { onConfirmedCalls } = await run({ fetchThrows: true });
    t.assertEqual(onConfirmedCalls.length, 0, 'NETWORK ERROR: onConfirmed is never called -- nothing is confirmed when the push outcome is unknown');
  }

  return t.summary();
})();
