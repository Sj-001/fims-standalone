// Real user complaint, 2026-10-12: "why does it show matched correctly, when its not matched while
// pushing... why do I have to first push thinking its all matched correctly and then you tell me its not
// matched". Root cause: pushCustomerSheetNow's own fresh pre-write check (verifyAgainstFreshBlocks) could
// already tell a row's catalog-resolved block wasn't real, and correctly excluded it from what actually
// got WRITTEN to the Sheet (`unmatched`) -- but pushPendingRows' `onConfirmed`/commitAll callback still
// confirmed EVERY row regardless, including ones whose item ended up in `unmatched`. So the row vanished
// from Pending Review (marked stockConfirmed) even though it was never actually sent, with the only trace
// being a push-status message the user had to separately notice and act on later via "Push outstanding".
// That's a silent bait-and-switch: the user accepts "Send N items", and one of those N quietly didn't
// happen and got hidden from the queue they could still fix it in.
//
// Fixed two ways, both tested here:
//   1. pushPendingRows: commits are now tracked WITH their row's description, and commitAll takes the
//      `unmatched` list pushCustomerSheetNow passes back, skipping confirmStockRow for exactly those rows
//      -- they stay pending, stay visible, still offer the Tab/Block picker (needsTabBlock reads
//      isCatalogEntryBlockReal off the now-corrected reviewByCustomer cache -- see test 3 below).
//   2. pushCustomerSheetNow: the confirm() dialog itself now says which items won't be sent BEFORE the
//      user accepts, instead of only reporting it after the fact.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');
const normalizeForCatalogMatch = (s) => String(s || '')
  .toLowerCase()
  .replace(/pkt\.?/g, '')
  .replace(/\s*[x×*]\s*\d+\s*$/i, '')
  .replace(/(\d)\s*gm\b/g, '$1g')
  .replace(/\bcont\.?\b/g, 'container')
  .replace(/[^a-z0-9]/g, '');

const t = makeRecorder('push-unmatched-not-committed');

/* ===== 1 & 2. pushPendingRows: commitAll only confirms rows NOT in the unmatched list ===== */
{
  const block = extractBlock(
    APP_JSX,
    'const pushPendingRows = (registerKey, rows) => {',
    'const updatePendingCustomer = (registerKey) => (id, value) => {',
    { endExclusive: true },
  );

  function run({ rows, unmatchedOnPush = [], notReadyCustomers = [], pendingTabBlockForms = {}, catalogEntryByDescription = {}, blockIsRealByDescription = {} }) {
    const confirmed = [];
    const pushCalls = [];
    const confirmStockRow = (registerKey, row, customer) => confirmed.push(row.id);
    const isKnownCustomerGuess = () => true;
    const matchCustomer = (r) => r.confirmedCustomer || 'DIAMOND';
    const getCustomerSheetId = () => 'sheet-id-123';
    // Defaults every customer to "ready" (review data already loaded) -- the pre-architectural-fix
    // behavior every existing test here exercises. notReadyCustomers lets a test opt a specific
    // customer INTO the "still checking" state to exercise the new push-time skip-guard.
    const isReviewReadyForCustomer = (customer) => !notReadyCustomers.includes(customer);
    // Defaults to "no catalog entry" for every row, which combined with an empty pendingTabBlockForms
    // (the default) makes every existing test below a no-op for the stale-draft gate -- only tests that
    // explicitly configure these exercise it.
    const getCatalogEntryForItem = (customer, description) => catalogEntryByDescription[description] || null;
    const isCatalogEntryBlockReal = (customer, catalogEntry, description) => !!blockIsRealByDescription[description];
    // Stands in for the real async pushCustomerSheetNow: synchronously invokes onConfirmed with
    // whatever this test configured as "what the fresh check would have rejected" -- exactly mirroring
    // what the real function does after its own confirm()/verifyAgainstFreshBlocks resolve.
    const pushCustomerSheetNow = (customer, productionRows, dispatchRows, routingOverrides, onConfirmed) => {
      pushCalls.push({ customer, productionRows, routingOverrides });
      if (onConfirmed) onConfirmed(unmatchedOnPush);
    };
    const fn = buildFromSource(
      block, 'pushPendingRows',
      ['isKnownCustomerGuess', 'matchCustomer', 'pendingTabBlockForms', 'normalizeForCatalogMatch', 'confirmStockRow', 'getCustomerSheetId', 'pushCustomerSheetNow', 'isReviewReadyForCustomer', 'getCatalogEntryForItem', 'isCatalogEntryBlockReal', 'window'],
      [isKnownCustomerGuess, matchCustomer, pendingTabBlockForms, normalizeForCatalogMatch, confirmStockRow, getCustomerSheetId, pushCustomerSheetNow, isReviewReadyForCustomer, getCatalogEntryForItem, isCatalogEntryBlockReal, { alert: () => {} }],
    );
    fn('production', rows);
    return { confirmed, pushCalls };
  }

  // THE LIVE CASE: three rows pushed together, one of them resolves to an item the fresh check rejects.
  {
    const rows = [
      { id: 'r1', description: 'TOOTI FRUTI 28G * 120 PKT', confirmedCustomer: 'DIAMOND' },
      { id: 'r2', description: 'Butter Bake 130g x30 Packet', confirmedCustomer: 'DIAMOND' },
      { id: 'r3', description: 'Jeera Dhamal 32g x144 Packet', confirmedCustomer: 'DIAMOND' },
    ];
    const { confirmed } = run({ rows, unmatchedOnPush: ['TOOTI FRUTI 28G * 120 PKT'] });
    t.assertEqual(confirmed.sort(), ['r2', 'r3'], 'LIVE CASE: the unmatched row is never confirmed, the genuinely-sent ones are');
  }

  // Nothing unmatched: every row gets confirmed, same as before this fix.
  {
    const rows = [
      { id: 'r1', description: 'Butter Bake 130g x30 Packet', confirmedCustomer: 'DIAMOND' },
      { id: 'r2', description: 'Jeera Dhamal 32g x144 Packet', confirmedCustomer: 'DIAMOND' },
    ];
    const { confirmed } = run({ rows, unmatchedOnPush: [] });
    t.assertEqual(confirmed.sort(), ['r1', 'r2'], 'NORMAL: nothing unmatched means everything gets confirmed as usual');
  }

  // Punctuation-only difference between the row's own description and the unmatched entry must still
  // match (unmatched descriptions come from itemGroups' aggregated description, not always byte-identical
  // to the row's raw text) -- uses the same normalizeForCatalogMatch both sides use everywhere else.
  {
    const rows = [{ id: 'r1', description: 'T 50 32G*120PKT', confirmedCustomer: 'DIAMOND' }];
    const { confirmed } = run({ rows, unmatchedOnPush: ['T 50 32G * 120 PKT'] });
    t.assertEqual(confirmed, [], 'PUNCTUATION: a cosmetically-different but equivalent description still withholds the commit');
  }

  // ARCHITECTURAL FIX, 2026-10-14: a customer whose real-Sheet review data hasn't loaded yet (still
  // "(checking…)" in the UI) must never be pushed at all -- not even sent to pushCustomerSheetNow --
  // since treating a not-yet-verified row as either matched or unmatched is exactly the race that let a
  // correctly-matched real block ("100 ML") get a duplicate created next to it in production.
  {
    const rows = [
      { id: 'r1', description: 'Butter Bake 130g x30 Packet', confirmedCustomer: 'DIAMOND' },
      { id: 'r2', description: 'Jeera Dhamal 32g x144 Packet', confirmedCustomer: 'ANMOL' },
    ];
    const { confirmed, pushCalls } = run({ rows, notReadyCustomers: ['DIAMOND'] });
    t.assertEqual(confirmed, ['r2'], 'NOT READY: the not-yet-checked customer\'s row is withheld, the ready one still pushes');
    t.assertEqual(pushCalls.length, 1, 'NOT READY: pushCustomerSheetNow is never called for the not-yet-checked customer');
    t.assertEqual(pushCalls[0].customer, 'ANMOL', 'NOT READY: only the ready customer reaches pushCustomerSheetNow');
  }
  {
    const rows = [{ id: 'r1', description: 'Butter Bake 130g x30 Packet', confirmedCustomer: 'DIAMOND' }];
    const { confirmed, pushCalls } = run({ rows, notReadyCustomers: ['DIAMOND'] });
    t.assertEqual(confirmed, [], 'NOT READY, ALL: nothing is confirmed when every row\'s customer is still checking');
    t.assertEqual(pushCalls.length, 0, 'NOT READY, ALL: pushCustomerSheetNow is never even called');
  }

  // THE LIVE BUG, confirmed directly against the real test environment, 2026-10-14: "Vishesh 100 ML"
  // showed correctly matched in Pending Review (real catalog block "100 ML", genuinely existing in the
  // real Sheet) and STILL got refused by the server for trying to create a brand-new block literally
  // titled "Vishesh 100 ML" -- the row's own raw description. Root cause: a pendingTabBlockForms draft
  // left over from BEFORE the row became correctly catalog-matched (Sheet Tab picked, Block left blank
  // -- which the picker itself labels "blank = new block", correct for a genuinely new item at the
  // time) was still being read and trusted unconditionally, with no check for whether the row still
  // actually needed it.
  {
    const rows = [{ id: 'r1', description: 'Vishesh 100 ML', confirmedCustomer: 'VISHESH STOCK' }];
    const { pushCalls } = run({
      rows,
      pendingTabBlockForms: { r1: { sheetGroup: 'VP PAPER CUP 1', block: '' } },
      catalogEntryByDescription: { 'Vishesh 100 ML': { item: 'Vishesh 100 ML', sheetGroup: 'VP PAPER CUP 1', block: '100 ML' } },
      blockIsRealByDescription: { 'Vishesh 100 ML': true },
    });
    t.assertEqual(pushCalls[0].routingOverrides, {}, 'LIVE CASE: a stale draft is ignored once the row is genuinely catalog-matched -- no bogus self-titled block override reaches the server');
  }
  // The exact same draft, but the row GENUINELY still needs it (no catalog match, or the catalog's
  // block genuinely isn't real) -- must still work exactly as before, picking up the person's explicit
  // choice. The stale-draft fix must never break a draft that's still actually needed.
  {
    const rows = [{ id: 'r1', description: 'Some New Item', confirmedCustomer: 'VISHESH STOCK' }];
    const { pushCalls } = run({
      rows,
      pendingTabBlockForms: { r1: { sheetGroup: 'SOME NEW TAB', block: 'Some New Block' } },
      catalogEntryByDescription: {},
      blockIsRealByDescription: {},
    });
    t.assertEqual(pushCalls[0].routingOverrides, { somenewitem: { sheetGroup: 'SOME NEW TAB', block: 'Some New Block' } }, 'STILL NEEDED: a draft for a genuinely unmatched item is still honored exactly as before');
  }
}

/* ===== 3. pushCustomerSheetNow: confirm() dialog states unmatched items up front, onConfirmed gets the
   unmatched list (not called bare) ===== */
{
  const dialogLine = extractLine(APP_JSX, 'const confirmMessage = unmatched.length');
  const block = extractBlock(APP_JSX, (l) => l === dialogLine, "if (!window.confirm(confirmMessage)) return;");
  const onConfirmedLine = extractLine(APP_JSX, 'if (onConfirmed) onConfirmed(unmatched);');

  function run({ totalRowsToSend, customer, unmatched, confirmReturns = true }) {
    let shownMessage = null;
    let onConfirmedCalled = false;
    let onConfirmedArg;
    const windowMock = { confirm: (msg) => { shownMessage = msg; return confirmReturns; } };
    const onConfirmed = (arg) => { onConfirmedCalled = true; onConfirmedArg = arg; };
    const body = `${block}\n${onConfirmedLine}`;
    const fn = new Function('totalRowsToSend', 'customer', 'unmatched', 'window', 'onConfirmed', body);
    fn(totalRowsToSend, customer, unmatched, windowMock, onConfirmed);
    return { shownMessage, onConfirmedCalled, onConfirmedArg };
  }

  {
    const { shownMessage, onConfirmedCalled, onConfirmedArg } = run({
      totalRowsToSend: 5, customer: 'DIAMOND', unmatched: ['TOOTI FRUTI 28G * 120 PKT'],
    });
    t.assertTrue(shownMessage.includes('will NOT be sent'), 'DIALOG: warns about the unmatched item before the user accepts');
    t.assertTrue(shownMessage.includes('TOOTI FRUTI 28G * 120 PKT'), 'DIALOG: names the specific unmatched item');
    t.assertTrue(onConfirmedCalled, 'onConfirmed still runs once the dialog is accepted');
    t.assertEqual(onConfirmedArg, ['TOOTI FRUTI 28G * 120 PKT'], 'onConfirmed receives the unmatched list, not called bare');
  }
  {
    const { shownMessage, onConfirmedArg } = run({ totalRowsToSend: 3, customer: 'DIAMOND', unmatched: [] });
    t.assertTrue(!shownMessage.includes('will NOT be sent'), 'NORMAL: no unmatched warning when everything matches');
    t.assertEqual(onConfirmedArg, [], 'NORMAL: onConfirmed still receives the (empty) unmatched list');
  }
  {
    const { onConfirmedCalled } = run({ totalRowsToSend: 3, customer: 'DIAMOND', unmatched: [], confirmReturns: false });
    t.assertTrue(!onConfirmedCalled, 'CANCEL: declining the dialog never calls onConfirmed at all — nothing gets confirmed');
  }
}

module.exports = t.summary();
