// Real gap, confirmed directly against the live isolated test environment, 2026-10-14: after the
// "still checking" architectural fix shipped, a Pending Review row for "IT 500 Container" (BINDAL STOCK
// 1.08.26) with a perfectly good Product Catalog mapping (sheetGroup "IT 500") STILL showed the picker
// forever, never resolving to matched. Root cause, found by calling the real preview endpoint directly:
// refreshReview's own probe set only ever asked the Sheet about a tab name that was either already
// confirmed/pushed before for this customer, or that a person had manually started typing into a
// picker draft -- a catalog-known-but-never-before-pushed item's own suggested tab was never probed at
// all, so getRealBlocksForTab always got [] for it and isCatalogEntryBlockReal always said "not real",
// forever, independent of loading state. Confirmed live: asking the real BINDAL Sheet for tab "IT 500"
// directly (bypassing the app) returned existingBlockTitles including "IT 500 CONTAINER" -- the data was
// always there, it just never got asked for. This test locks in that every pending row's catalog-
// suggested sheetGroup is included in the probe set, regardless of whether a manual draft exists yet.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

const BLOCK = extractBlock(
  APP_JSX,
  'const pendingRowsForCustomer = [...pendingProductionRows, ...pendingDispatchRows].filter(row => {',
  'const newProbeNames = probeTabNames;',
  { endExclusive: true },
);

const t = makeRecorder('review-probe-catalog-tab');

function run({ customer, rows, catalog = {}, drafts = {}, isKnownCustomerGuess, matchCustomer, getCatalogEntryForItem }) {
  return buildFromSource(
    BLOCK, 'probeTabNames',
    ['customer', 'pendingProductionRows', 'pendingDispatchRows', 'isKnownCustomerGuess', 'matchCustomer', 'unmatchedAssignForms', 'assignForms', 'pendingTabBlockForms', 'getCatalogEntryForItem'],
    [
      customer, rows, [],
      isKnownCustomerGuess || (() => true),
      matchCustomer || ((r) => r.confirmedCustomer || customer),
      {}, {}, drafts,
      getCatalogEntryForItem || ((c, description) => catalog[description] || null),
    ],
  );
}

/* ===== THE LIVE CASE: a pending row with a catalog mapping but no manual draft yet must still probe
   the catalog's own sheetGroup -- this is the exact BINDAL "IT 500 Container" case. ===== */
{
  const rows = [{ id: 'r1', description: 'IT 500 Container', confirmedCustomer: 'BINDAL STOCK 1.08.26' }];
  const catalog = { 'IT 500 Container': { sheetGroup: 'IT 500', block: '' } };
  const probes = run({ customer: 'BINDAL STOCK 1.08.26', rows, catalog });
  t.assertTrue(probes.includes('IT 500'), 'LIVE CASE: the catalog-suggested tab is probed with no manual draft at all');
}

/* ===== A row with no catalog entry contributes nothing new -- never probes an empty/undefined
   sheetGroup, same as every other source in this array (.filter(Boolean)). ===== */
{
  const rows = [{ id: 'r1', description: 'Totally Unknown Item', confirmedCustomer: 'BINDAL STOCK 1.08.26' }];
  const probes = run({ customer: 'BINDAL STOCK 1.08.26', rows, catalog: {} });
  t.assertEqual(probes, [], 'NO CATALOG ENTRY: nothing probed, no blank string sneaks into the list');
}

/* ===== A manual draft and a catalog suggestion for the SAME row both end up probed (deduped) -- the
   new source is additive, it never replaces the existing draft-based probing. ===== */
{
  const rows = [{ id: 'r1', description: 'IT 500 Container', confirmedCustomer: 'BINDAL STOCK 1.08.26' }];
  const catalog = { 'IT 500 Container': { sheetGroup: 'IT 500', block: '' } };
  const drafts = { r1: { sheetGroup: 'IT 500' } };
  const probes = run({ customer: 'BINDAL STOCK 1.08.26', rows, catalog, drafts });
  t.assertEqual(probes, ['IT 500'], 'ADDITIVE: draft and catalog suggestion agree, deduped to one entry');
}

/* ===== A manual draft that DIFFERS from the catalog suggestion (person overriding the suggested tab)
   still gets its own real tab probed too -- the picker's own options must stay live, not just the
   catalog's original guess. ===== */
{
  const rows = [{ id: 'r1', description: 'IT 500 Container', confirmedCustomer: 'BINDAL STOCK 1.08.26' }];
  const catalog = { 'IT 500 Container': { sheetGroup: 'IT 500', block: '' } };
  const drafts = { r1: { sheetGroup: 'SOME OTHER TAB' } };
  const probes = run({ customer: 'BINDAL STOCK 1.08.26', rows, catalog, drafts });
  t.assertTrue(probes.includes('IT 500'), 'OVERRIDE: the catalog tab is still probed');
  t.assertTrue(probes.includes('SOME OTHER TAB'), 'OVERRIDE: the manually-typed tab is still probed too');
}

/* ===== The catalog lookup is keyed by the exact customer the block runs for -- an entry that only
   exists under a DIFFERENT customer in the mock never leaks into this customer's probe set. ===== */
{
  const rows = [{ id: 'r1', description: 'IT 500 Container', confirmedCustomer: 'BINDAL STOCK 1.08.26' }];
  const getCatalogEntryForItem = (c, description) => (c === 'SOMEONE ELSE' ? { sheetGroup: 'IT 500' } : null);
  const probes = run({ customer: 'BINDAL STOCK 1.08.26', rows, getCatalogEntryForItem });
  t.assertEqual(probes, [], 'CUSTOMER-SCOPED: a catalog entry under a different customer never leaks in');
}

module.exports = t.summary();
