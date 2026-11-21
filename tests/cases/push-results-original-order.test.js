// Real, directly-confirmed live incident (2026-11-21): a single DIAMOND push sent two tabs together —
// T 50 (blocked, an unauthorized-new-block refusal) and CREAM (authorized, genuinely written). The old
// code built pushCustomerSheet's response as `[...blockedResults, ...written-tab results]` — every
// blocked tab FIRST, regardless of which position it was actually sent in, then every authorized tab
// after, in whatever order it finished writing. The client matches a response entry back to its own rows
// purely by POSITION (see App.jsx's own comment near idsToMarkByIndex: "results[] is built in the exact
// same order tabPlans... at every step of the server pipeline"), so when T 50 was sent FIRST and CREAM
// SECOND, the response still came back [T50-blocked, CREAM-ok] in THAT order — correct, by luck. But when
// CREAM was sent first and T 50 second (itemGroups is built from a plain object's Object.entries, so
// order depends on insertion order, not a fixed rule), the response still came back
// [T50-blocked, CREAM-ok] — now in the WRONG order relative to what was actually sent. The client read
// position 0 (meant for CREAM) as T 50's refusal — leaving CREAM's rows stuck "pending" forever despite
// already being genuinely written to the Sheet — and position 1 (meant for T 50) as CREAM's success —
// silently marking a REFUSED row as confirmed, even though it was never actually written at all (a real,
// silent data-loss risk, not just a cosmetic mislabel). buildOrderedResults fixes this by reassembling the
// response strictly in allTabPlans' own original order, no matter how many tabs were blocked vs. written,
// or in what order the write loop happened to finish them.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

const t = makeRecorder('push-results-original-order');

function build(allTabPlans, blockedResults) {
  const block = extractBlock(SHEETS_JS, 'const buildOrderedResults = (moreResults) => {', (l) => l.trim() === '};');
  return buildFromSource(block, 'buildOrderedResults', ['allTabPlans', 'blockedResults'], [allTabPlans, blockedResults]);
}

/* ===== THE EXACT LIVE CASE: T 50 sent FIRST, CREAM sent SECOND, but T 50 is the one that's blocked —
   must come back in [T50, CREAM] order, i.e. unchanged, since that already matches original order.
   (Confirms the simple case the old code happened to get right isn't broken by the fix.) ===== */
{
  const allTabPlans = [{ tabName: 'T 50' }, { tabName: 'CREAM' }];
  const blockedResults = [{ tab: 'T 50', ok: false, error: 'refused' }];
  const buildOrderedResults = build(allTabPlans, blockedResults);
  const out = buildOrderedResults([{ tab: 'CREAM', ok: true, newRows: 2 }]);
  t.assertEqual(out.map(r => r.tab), ['T 50', 'CREAM'], 'SIMPLE CASE: blocked-first order preserved when that is also the original send order');
}

/* ===== THE EXACT LIVE CASE, REVERSED: CREAM sent FIRST, T 50 sent SECOND — the actual incident. Must
   come back [CREAM, T50], matching the ORIGINAL send order, not [T50, CREAM]. ===== */
{
  const allTabPlans = [{ tabName: 'CREAM' }, { tabName: 'T 50' }];
  const blockedResults = [{ tab: 'T 50', ok: false, error: 'refused' }];
  const buildOrderedResults = build(allTabPlans, blockedResults);
  const out = buildOrderedResults([{ tab: 'CREAM', ok: true, newRows: 2 }]);
  t.assertEqual(out.map(r => r.tab), ['CREAM', 'T 50'], 'LIVE CASE: result order matches the ORIGINAL send order, not "blocked tabs first"');
  t.assertEqual(out[0].ok, true, 'LIVE CASE: position 0 (CREAM, as sent) correctly reports ok — the client marks this row pushed, correctly');
  t.assertEqual(out[1].ok, false, 'LIVE CASE: position 1 (T 50, as sent) correctly reports refused — the client leaves this row pending, correctly, instead of silently confirming a never-written row');
}

/* ===== Three tabs, blocked tab in the MIDDLE of the original order — must land back in the middle, not
   get shoved to the front. ===== */
{
  const allTabPlans = [{ tabName: 'DOODH' }, { tabName: 'T 50' }, { tabName: 'BUTTER' }];
  const blockedResults = [{ tab: 'T 50', ok: false, error: 'refused' }];
  const buildOrderedResults = build(allTabPlans, blockedResults);
  const out = buildOrderedResults([{ tab: 'DOODH', ok: true, newRows: 1 }, { tab: 'BUTTER', ok: true, newRows: 3 }]);
  t.assertEqual(out.map(r => r.tab), ['DOODH', 'T 50', 'BUTTER'], 'MIDDLE: a blocked tab sandwiched between two written tabs stays in its real original position');
}

/* ===== All tabs blocked, none written — degrades to exactly blockedResults, in original order. ===== */
{
  const allTabPlans = [{ tabName: 'A' }, { tabName: 'B' }];
  const blockedResults = [{ tab: 'A', ok: false, error: 'x' }, { tab: 'B', ok: false, error: 'y' }];
  const buildOrderedResults = build(allTabPlans, blockedResults);
  const out = buildOrderedResults([]);
  t.assertEqual(out.map(r => r.tab), ['A', 'B'], 'ALL BLOCKED: order preserved with nothing written at all');
}

/* ===== All tabs written, none blocked — degrades to exactly the write results, in original order
   (the common case, must stay exactly as it always worked). ===== */
{
  const allTabPlans = [{ tabName: 'A' }, { tabName: 'B' }];
  const blockedResults = [];
  const buildOrderedResults = build(allTabPlans, blockedResults);
  const out = buildOrderedResults([{ tab: 'A', ok: true, newRows: 1 }, { tab: 'B', ok: true, newRows: 1 }]);
  t.assertEqual(out.map(r => r.tab), ['A', 'B'], 'ALL WRITTEN: the common, every-tab-ok case is unaffected');
}

/* ===== A tab in allTabPlans with NO matching result at all (defensive — should never happen in
   practice, every tab is either blocked or written) is simply omitted, never crashes, never inserts a
   hole/undefined into the array the client will index into. ===== */
{
  const allTabPlans = [{ tabName: 'A' }, { tabName: 'GHOST' }, { tabName: 'B' }];
  const blockedResults = [];
  const buildOrderedResults = build(allTabPlans, blockedResults);
  const out = buildOrderedResults([{ tab: 'A', ok: true }, { tab: 'B', ok: true }]);
  t.assertEqual(out.map(r => r.tab), ['A', 'B'], 'DEFENSIVE: a tab with no result at all is skipped, not left as a hole');
}

module.exports = t.summary();
