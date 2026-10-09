// A catalog entry's resolved block can point at a block that doesn't actually exist in the customer's
// real Sheet tab (typo, stale mapping, a punctuation mismatch like the normalizeTabKey case) -- nothing
// used to check that, so the row sailed through Pending Review with no picker and silently got a
// brand-new duplicate block on push. Two-part fix, both covered here:
//   - isCatalogEntryBlockReal: extends needsTabBlock in Pending Review to force the picker when a
//     catalog entry's block matches nothing real in a tab that already has real blocks.
//   - the push-time hard backstop inside buildCustomerSheetPayloadFromRows: routes to `unmatched`
//     instead of creating a block, unless the item was just explicitly routed via routingOverrides.
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

const t = makeRecorder('catalog-block-ambiguity');

/* ===== 1. isCatalogEntryBlockReal ===== */
{
  const block = extractBlock(
    APP_JSX,
    'const isCatalogEntryBlockReal = (customer, catalogEntry, description) => {',
    (l) => l.trim() === '};',
  );
  function run(customer, catalogEntry, description, realBlocksFor) {
    const getRealBlocksForTab = (c, sg) => realBlocksFor[`${c}||${sg}`] || [];
    const fn = buildFromSource(block, 'isCatalogEntryBlockReal', ['getRealBlocksForTab', 'normalizeForCatalogMatch'], [getRealBlocksForTab, normalizeForCatalogMatch]);
    return fn(customer, catalogEntry, description);
  }

  // *** THE EXACT REPORTED CASE ***: catalog points at "T 50 32G * 120 PKT new" (no parens), real tab
  // only has "T 50 32G * 120 PKT (new)" (with parens) -- normalizeForCatalogMatch strips all
  // punctuation, so these DO match once compared properly.
  t.assertEqual(
    run('DIAMOND', { sheetGroup: 'T 50', block: 'T 50 32G * 120 PKT new' }, 'irrelevant',
      { 'DIAMOND||T 50': ['T 50 32G * 120 PKT', 'T 50 64G * 60 PKT', 'T 50 32G * 120 PKT (new)'] }),
    true, 'EXACT CASE: punctuation-only difference is recognized as real, no picker forced'
  );
  t.assertEqual(
    run('DIAMOND', { sheetGroup: 'T 50', block: 'Totally Different Block' }, 'irrelevant',
      { 'DIAMOND||T 50': ['T 50 32G * 120 PKT', 'T 50 64G * 60 PKT'] }),
    false, 'GENUINE MISMATCH: a block that matches nothing real forces the picker'
  );
  t.assertEqual(run('DIAMOND', null, 'X', {}), true, 'NO ENTRY: returns true, caller handles it separately');
  t.assertEqual(
    run('DIAMOND', { sheetGroup: 'CREAM', block: '' }, 'Cream 30g (new)', { 'DIAMOND||CREAM': ['CREAM 30G (NEW)'] }),
    true, 'BLANK BLOCK: falls back to description, matches the real block via normalization'
  );
  t.assertEqual(
    run('DIAMOND', { sheetGroup: 'NEWTAB', block: 'Whatever New Item' }, 'irrelevant', {}),
    true, 'BRAND NEW TAB: no real blocks yet -> not ambiguous, first-ever block is fine'
  );
}

/* ===== 2. buildCustomerSheetPayloadFromRows push-time safety net ===== */
{
  const startLine = extractLine(APP_JSX, "if (!sheetGroup) { unmatched.push(g.description || '(blank description)'); return; }");
  const block = extractBlock(
    APP_JSX,
    (l) => l === startLine,
    (l) => l.includes('const itemGroups = Object.entries(tabsMap)'),
    { endExclusive: true },
  );

  function run({ customer, groups, sheetGroupByItem, blockByItem, routingOverrides, realBlocksFor }) {
    const getRealBlocksForTab = (c, sg) => realBlocksFor[`${c}||${sg}`] || [];
    const unmatched = [];
    const tabsMap = {};
    const fn = new Function(
      'customer', 'groups', 'sheetGroupByItem', 'blockByItem', 'routingOverrides', 'unmatched', 'tabsMap',
      'normalizeForCatalogMatch', 'getRealBlocksForTab', 'normalizeDateToDots',
      `groups.forEach(g => {
        const key = normalizeForCatalogMatch(g.description);
        const sheetGroup = sheetGroupByItem[key];
        ${block}
      return { unmatched, tabsMap };`
    );
    return fn(customer, groups, sheetGroupByItem, blockByItem, routingOverrides, unmatched, tabsMap, normalizeForCatalogMatch, getRealBlocksForTab, (d) => d);
  }

  {
    const groups = [{ description: 'T 50 32G * 120 PKT new', customer: 'DIAMOND', ledger: [] }];
    const { unmatched, tabsMap } = run({
      customer: 'DIAMOND', groups,
      sheetGroupByItem: { [normalizeForCatalogMatch('T 50 32G * 120 PKT new')]: 'T 50' },
      blockByItem: { [normalizeForCatalogMatch('T 50 32G * 120 PKT new')]: 'T 50 32G * 120 PKT new' },
      routingOverrides: {},
      realBlocksFor: { 'DIAMOND||T 50': ['T 50 32G * 120 PKT', 'T 50 64G * 60 PKT', 'T 50 32G * 120 PKT (new)'] },
    });
    t.assertEqual(unmatched.length, 0, 'PARENS: client-side normalizer already reconciles this, proceeds normally');
    t.assertEqual(Object.keys(tabsMap).length, 1, 'PARENS: pushes correctly, matching the real "(new)" block');
  }
  {
    const groups = [{ description: 'Some Other Totally Unrelated Item', customer: 'DIAMOND', ledger: [] }];
    const key = normalizeForCatalogMatch('Some Other Totally Unrelated Item');
    const { unmatched, tabsMap } = run({
      customer: 'DIAMOND', groups,
      sheetGroupByItem: { [key]: 'T 50' }, blockByItem: { [key]: 'Some Other Totally Unrelated Item' }, routingOverrides: {},
      realBlocksFor: { 'DIAMOND||T 50': ['T 50 32G * 120 PKT', 'T 50 64G * 60 PKT', 'T 50 32G * 120 PKT (new)'] },
    });
    t.assertEqual(Object.keys(tabsMap).length, 0, 'PUSH SAFETY: nothing gets queued to push for an item that matches no real block');
    t.assertEqual(unmatched.length, 1, 'PUSH SAFETY: the item is reported as unmatched instead of silently creating a block');
  }
  {
    const groups = [{ description: 'Genuinely Brand New Item', customer: 'DIAMOND', ledger: [] }];
    const key = normalizeForCatalogMatch('Genuinely Brand New Item');
    const { unmatched, tabsMap } = run({
      customer: 'DIAMOND', groups,
      sheetGroupByItem: { [key]: 'T 50' }, blockByItem: {},
      routingOverrides: { [key]: { sheetGroup: 'T 50', block: 'Genuinely Brand New Item' } },
      realBlocksFor: { 'DIAMOND||T 50': ['T 50 32G * 120 PKT'] },
    });
    t.assertEqual(unmatched.length, 0, 'EXPLICIT ROUTE: a just-picked routingOverride is trusted, not blocked');
    t.assertEqual(Object.keys(tabsMap).length, 1, 'EXPLICIT ROUTE: proceeds to push normally');
  }
  {
    const groups = [{ description: 'T 50 64G * 60 PKT', customer: 'DIAMOND', ledger: [] }];
    const key = normalizeForCatalogMatch('T 50 64G * 60 PKT');
    const { unmatched } = run({
      customer: 'DIAMOND', groups, sheetGroupByItem: { [key]: 'T 50' }, blockByItem: {}, routingOverrides: {},
      realBlocksFor: { 'DIAMOND||T 50': ['T 50 64G * 60 PKT'] },
    });
    t.assertEqual(unmatched.length, 0, 'NORMAL: a correctly-matching item is never blocked');
  }
}

module.exports = t.summary();
