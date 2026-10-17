// THE LIVE BUG, confirmed directly against the real test Production Register, 2026-10-17: syncFromSheet
// (pulls rows added directly in a customer's real Sheet back into the local register) used to append
// straight to setProduction/setCustomerDispatch with ZERO dedup check against the register it was
// appending to -- unlike every other path that ever adds a row (addRows, gated by rowsMatchForDedup).
// A real duplicate landed this way: two Production Register rows for 16.10.26, "N 100 Jumbo Container"/
// "N 100 Jumbo Lid" (540 pieces each, still pending confirmation) sat right alongside a second pair,
// "N 100 JUMBO CONTAINER"/"N 100 JUMBO LID" (540 each, pulledFromSheet:true, stockConfirmed:true) --
// same real entries, pulled back in because the known-row-count baseline for that Sheet block had
// fallen one row behind the real Sheet (see syncFromSheet's own comment: it judges "anything new?"
// purely by row count, never content). The baseline drifting is still possible; this fix is that once
// it does, the pulled-in row is no longer blindly appended -- it goes through the exact same
// rowsMatchForDedup gate (DEDUP_FIELDS.production: date/party/description/customerHint/pieces/dispatch)
// every extraction-confirm row already goes through, via addRows, so it's silently recognized as a row
// already in the register -- case difference and all, since rowsMatchForDedup's description comparison
// already normalizes case (normalizeForCatalogMatch).
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadAddRows(registerState, persisted) {
  const dedupFieldsBlock = extractBlock(APP_JSX, 'const DEDUP_FIELDS = {', (l) => l.trim() === '};');
  const applyAbbreviationsBlock = extractBlock(APP_JSX, 'const applyAbbreviations = (s) => {', (l) => l.trim() === '};');
  const normalizeForCatalogMatchBlock = extractBlock(
    APP_JSX,
    'const normalizeForCatalogMatch = (s) => applyAbbreviations(s || \'\')',
    ".replace(/[^a-z0-9]/g, '');"
  );
  const blankWildcardFieldsBlock = extractLine(APP_JSX, 'const BLANK_WILDCARD_DEDUP_FIELDS = {');
  const dedupFieldValueBlock = extractLine(APP_JSX, 'const dedupFieldValue = (row, f) =>');
  const rowsMatchForDedupBlock = extractBlock(APP_JSX, 'const rowsMatchForDedup = (registerKey, a, b) => {', (l) => l.trim() === '};');
  const dedupRegistersBlock = extractLine(APP_JSX, 'const DEDUP_REGISTERS = new Set(');
  const addRowsBlock = extractBlock(APP_JSX, 'const addRows = (registerKey, rawRows) => {', (l) => l.trim() === '};');

  const blocks = [
    dedupFieldsBlock, applyAbbreviationsBlock, normalizeForCatalogMatchBlock,
    blankWildcardFieldsBlock, dedupFieldValueBlock, rowsMatchForDedupBlock,
    dedupRegistersBlock, addRowsBlock,
  ];

  const registerSetters = {};
  Object.keys(registerState).forEach(key => {
    registerSetters[key] = (fn) => { registerState[key] = fn(registerState[key]); };
  });
  const persist = (key, rows) => { persisted[key] = rows; };

  return buildFromSource(blocks, 'addRows', ['abbreviations', 'registerState', 'registerSetters', 'persist'], [[], registerState, registerSetters, persist]);
}

const t = makeRecorder('sync-from-sheet-dedup');

/* ===== THE LIVE CASE: a syncFromSheet-shaped pull (uppercase, pulledFromSheet) of a row already
   sitting in the register (Title Case, still pending) must be recognized as the same real entry and
   skipped, not appended as a second row. ===== */
{
  const registerState = {
    production: [{
      id: '1792255997465-p05o088', date: '16.10.26', party: '', description: 'N 100 Jumbo Container',
      customerHint: '', pieces: 540, dispatch: 0, stockConfirmed: false, confirmedCustomer: '', pushedToSheet: '',
    }],
  };
  const persisted = {};
  const addRows = loadAddRows(registerState, persisted);
  const pulledFromSheet = [{
    id: 'new-pulled-1', date: '16.10.26', party: '', description: 'N 100 JUMBO CONTAINER', customerHint: '',
    pieces: 540, dispatch: 0, confirmedCustomer: 'BINDAL STOCK 1.08.26', stockConfirmed: true, pushedToSheet: true, pulledFromSheet: true,
  }];
  const skipped = addRows('production', pulledFromSheet);
  t.assertEqual(skipped, 1, 'LIVE CASE: the sync-pulled uppercase row is recognized as a duplicate of the already-pending Title Case row');
  t.assertEqual(registerState.production.length, 1, 'LIVE CASE: the register still has exactly the one original row -- no second copy was appended');
}

/* ===== A genuinely new sync-pulled row (nothing like it already in the register) still gets added --
   this fix must not make syncFromSheet stop working for real new rows. ===== */
{
  const registerState = { production: [] };
  const persisted = {};
  const addRows = loadAddRows(registerState, persisted);
  const pulledFromSheet = [{
    id: 'new-pulled-2', date: '16.10.26', party: '', description: 'N 100 JUMBO LID', customerHint: '',
    pieces: 540, dispatch: 0, confirmedCustomer: 'BINDAL STOCK 1.08.26', stockConfirmed: true, pushedToSheet: true, pulledFromSheet: true,
  }];
  const skipped = addRows('production', pulledFromSheet);
  t.assertEqual(skipped, 0, 'a genuinely new sync-pulled row is not skipped');
  t.assertEqual(registerState.production.length, 1, 'the genuinely new row is actually added to the register');
  t.assertEqual(registerState.production[0].pulledFromSheet, true, 'the added row still carries pulledFromSheet:true, unchanged by addRows');
}

/* ===== SAFETY: a sync-pulled row for a genuinely different real entry (different pieces) on the same
   date/description must still be added, not wrongly merged into an unrelated existing row. ===== */
{
  const registerState = {
    production: [{
      id: 'old1', date: '16.10.26', party: '', description: 'N 100 Jumbo Lid',
      customerHint: '', pieces: 600, dispatch: 0, stockConfirmed: false, confirmedCustomer: '', pushedToSheet: '',
    }],
  };
  const persisted = {};
  const addRows = loadAddRows(registerState, persisted);
  const pulledFromSheet = [{
    id: 'new-pulled-3', date: '16.10.26', party: '', description: 'N 100 JUMBO LID', customerHint: '',
    pieces: 540, dispatch: 0, confirmedCustomer: 'BINDAL STOCK 1.08.26', stockConfirmed: true, pushedToSheet: true, pulledFromSheet: true,
  }];
  const skipped = addRows('production', pulledFromSheet);
  t.assertEqual(skipped, 0, 'SAFETY: a different real quantity on the same date/description is not silently merged away');
  t.assertEqual(registerState.production.length, 2, 'SAFETY: both the existing 600-piece row and the new 540-piece row are kept');
}

module.exports = t.summary();
