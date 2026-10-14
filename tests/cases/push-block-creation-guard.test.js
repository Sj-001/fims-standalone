// Real gap, confirmed directly against the live isolated test environment, 2026-10-14: the server's own
// pushCustomerSheet had ZERO validation of its own — it created whatever block title a caller sent it,
// no questions asked. Every "never auto-create a block" fix this session (isCatalogEntryBlockReal, the
// push-time sync check, verifyAgainstFreshBlocks) was CLIENT-SIDE ONLY — a real guarantee only as long
// as the client's own logic happened to be bug-free that day, which it repeatedly wasn't (TOOTI FRUTI,
// the staleness race, the "empty tab" exception). Proven live: a hand-crafted POST straight to
// /api/customer-sheets/push with a block title that didn't exist anywhere in the real Sheet created it
// without complaint. authorizeTabPlansForPush closes that gap server-side, so the rule holds no matter
// what reaches this endpoint, not just when the client happens to behave.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

function loadAuthorizeTabPlansForPush() {
  const blocks = [
    extractLine(SHEETS_JS, 'const normalizeCellStr = (v) =>'),
    extractBlock(SHEETS_JS, 'const serialToDateStr = (n) => {', (l) => l.trim() === '};'),
    extractLine(SHEETS_JS, 'const normalizeDateCell = (v) =>'),
    extractBlock(SHEETS_JS, 'function canonicalDateKey(v) {', 'function normalizeDateDisplay(v) {', { endExclusive: true }),
    extractBlock(SHEETS_JS, 'function parseExistingBlocks(grid) {', 'function planChronologicalInserts(existingRowsOrdered, newRows) {', { endExclusive: true }),
    extractBlock(SHEETS_JS, 'const TAB_KEY_ALIASES = [', (l) => l.trim() === '];'),
    extractBlock(SHEETS_JS, 'const normalizeTabKey = (name) => {', (l) => l.trim() === '};'),
    extractBlock(SHEETS_JS, 'function authorizeTabPlansForPush(tabPlans, previousValuesByTab, exemptBlockTitles) {', (l) => l.trim() === '}'),
  ];
  return buildFromSource(blocks, 'authorizeTabPlansForPush');
}

const t = makeRecorder('push-block-creation-guard');
const authorizeTabPlansForPush = loadAuthorizeTabPlansForPush();

// A real block, as it would come back from a batchGet read: a 'Date' header cell with the title one row
// above it. Minimal 2-column shape is enough for parseExistingBlocks to recognize a block.
function gridWithBlocks(titles) {
  const titleRow = [];
  const headerRow = [];
  titles.forEach((title, i) => { titleRow[i * 2] = title; headerRow[i * 2] = 'Date'; headerRow[i * 2 + 1] = 'Production'; });
  return [titleRow, headerRow];
}

/* ===== 1. THE EXACT LIVE CASE: a block title with no matching existing block anywhere in the tab, and
   nothing exempting it, is refused — the whole tab is blocked, nothing gets through. ===== */
{
  const tabPlans = [{ tabName: 'T 50', variants: [{ title: 'ZZZ TEST FAKE BLOCK DELETE ME', header: [], rows: [] }] }];
  const previousValuesByTab = { 'T 50': gridWithBlocks(['T 50 32G * 120 PKT', 'T 50 64G * 60 PKT', 'T 50 32G * 120 PKT (new)']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 0, 'LIVE CASE: a genuinely unmatched block title is never authorized');
  t.assertEqual(blockedResults.length, 1, 'LIVE CASE: the tab is reported as blocked');
  t.assertTrue(!blockedResults[0].ok, 'LIVE CASE: blocked result is not ok');
  t.assertTrue(blockedResults[0].error.includes('ZZZ TEST FAKE BLOCK DELETE ME'), 'LIVE CASE: the error names the specific unauthorized title');
}

/* ===== 2. A variant that matches a real existing block (even with cosmetic punctuation/spacing
   differences, via normalizeTabKey) is authorized normally — the ordinary, correct push path. ===== */
{
  const tabPlans = [{ tabName: 'T 50', variants: [{ title: 'T 50 32G*120PKT (new)', header: [], rows: [] }] }];
  const previousValuesByTab = { 'T 50': gridWithBlocks(['T 50 32G * 120 PKT (new)']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 1, 'NORMAL: a title matching a real existing block (modulo formatting) is authorized');
  t.assertEqual(blockedResults.length, 0, 'NORMAL: nothing is blocked');
}

/* ===== 3. A variant with no matching existing block, but explicitly exempted (the person's own
   deliberate pick) IS authorized -- this is what a real routingOverrides pick looks like server-side. ===== */
{
  const tabPlans = [{ tabName: 'T 50', variants: [{ title: 'Genuinely New Variant', header: [], rows: [] }] }];
  const previousValuesByTab = { 'T 50': gridWithBlocks(['T 50 32G * 120 PKT']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, ['Genuinely New Variant']);
  t.assertEqual(authorized.length, 1, 'EXPLICIT PICK: an exempted title is authorized even with no matching existing block');
  t.assertEqual(blockedResults.length, 0, 'EXPLICIT PICK: nothing is blocked');
}

/* ===== 4. POLICY: a brand-new tab (no entry in previousValuesByTab at all, i.e. it doesn't exist yet)
   still requires every variant to be exempt -- no "nothing to be ambiguous with yet" exception, same
   policy as everywhere else this rule is enforced. ===== */
{
  const tabPlans = [{ tabName: 'BRAND NEW TAB', variants: [{ title: 'First Item Ever', header: [], rows: [] }] }];
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, {}, []);
  t.assertEqual(authorized.length, 0, 'POLICY: a brand-new tab with zero real blocks still requires explicit exemption');
  t.assertEqual(blockedResults.length, 1, 'POLICY: blocked, not silently created');
}
{
  const tabPlans = [{ tabName: 'BRAND NEW TAB', variants: [{ title: 'First Item Ever', header: [], rows: [] }] }];
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, {}, ['First Item Ever']);
  t.assertEqual(authorized.length, 1, 'POLICY: ...but an explicitly exempted item still creates the tab\'s first block fine');
  t.assertEqual(blockedResults.length, 0, 'POLICY: not blocked when deliberately picked');
}

/* ===== 5. FAIL-SAFE: a tab with a MIX of authorized and unauthorized variants is blocked IN FULL --
   nothing in that tab is written, not even the variants that would have been fine on their own. ===== */
{
  const tabPlans = [{
    tabName: 'T 50',
    variants: [
      { title: 'T 50 32G * 120 PKT', header: [], rows: [] }, // matches a real block
      { title: 'Some Unauthorized Title', header: [], rows: [] }, // does not
    ],
  }];
  const previousValuesByTab = { 'T 50': gridWithBlocks(['T 50 32G * 120 PKT']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 0, 'FAIL-SAFE: the whole tab is blocked, including the variant that would have matched fine');
  t.assertEqual(blockedResults.length, 1, 'FAIL-SAFE: reported once, for the whole tab');
}

/* ===== 6. Multiple tabs: an unauthorized variant in ONE tab never blocks a completely separate,
   correctly-authorized tab in the same push. ===== */
{
  const tabPlans = [
    { tabName: 'T 50', variants: [{ title: 'Unauthorized Here', header: [], rows: [] }] },
    { tabName: 'BUTTER', variants: [{ title: 'BUTTER PLUS 28G*140PKT', header: [], rows: [] }] },
  ];
  const previousValuesByTab = {
    'T 50': gridWithBlocks(['T 50 32G * 120 PKT']),
    'BUTTER': gridWithBlocks(['BUTTER PLUS 28G*140PKT']),
  };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 1, 'ISOLATION: the correctly-authorized tab still goes through');
  t.assertEqual(authorized[0].tabName, 'BUTTER', 'ISOLATION: specifically the authorized one');
  t.assertEqual(blockedResults.length, 1, 'ISOLATION: only the unauthorized tab is blocked');
  t.assertEqual(blockedResults[0].tab, 'T 50', 'ISOLATION: specifically the unauthorized one');
}

/* ===== 7. THE LIVE BUG: a variant whose own item title differs from its real block's title
   (blockTitleOverride) must be authorized against the OVERRIDE, not the raw title -- confirmed directly
   against the real VISHESH STOCK test Sheet: item "Vishesh 100 ML" correctly resolves to the real,
   pre-existing block "100 ML", but was refused every time because this function used to key on v.title
   alone, ignoring blockTitleOverride entirely (the one field every other block-matching step in
   server/lib/sheets.js already treats as taking priority). ===== */
{
  const tabPlans = [{
    tabName: 'VP PAPER CUP 1',
    variants: [{ title: 'Vishesh 100 ML', blockTitleOverride: '100 ML', header: [], rows: [] }],
  }];
  const previousValuesByTab = { 'VP PAPER CUP 1': gridWithBlocks(['100 ML', 'DIA DOME']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 1, 'LIVE CASE: authorized against the real block the override points at, not the item\'s own name');
  t.assertEqual(blockedResults.length, 0, 'LIVE CASE: not blocked -- "100 ML" genuinely exists');
}
/* ===== 8. The flip side: a blockTitleOverride that does NOT match any real block must still be
   refused, exactly like a plain title would be -- the override isn't a bypass, just a different key to
   check. ===== */
{
  const tabPlans = [{
    tabName: 'VP PAPER CUP 1',
    variants: [{ title: 'Vishesh 100 ML', blockTitleOverride: 'Totally Made Up Block', header: [], rows: [] }],
  }];
  const previousValuesByTab = { 'VP PAPER CUP 1': gridWithBlocks(['100 ML']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 0, 'FLIP SIDE: an override pointing at nothing real is still refused');
  t.assertEqual(blockedResults.length, 1, 'FLIP SIDE: reported as blocked');
  t.assertTrue(blockedResults[0].error.includes('Totally Made Up Block'), 'FLIP SIDE: the error names the override title, not the item\'s own name');
}
/* ===== 9. A variant with NO blockTitleOverride at all still authorizes against its own title, exactly
   as before -- the override is additive, never required. ===== */
{
  const tabPlans = [{ tabName: 'T 50', variants: [{ title: 'T 50 32G * 120 PKT', header: [], rows: [] }] }];
  const previousValuesByTab = { 'T 50': gridWithBlocks(['T 50 32G * 120 PKT']) };
  const { authorized, blockedResults } = authorizeTabPlansForPush(tabPlans, previousValuesByTab, []);
  t.assertEqual(authorized.length, 1, 'NO OVERRIDE: still authorizes correctly against the plain title');
  t.assertEqual(blockedResults.length, 0, 'NO OVERRIDE: not blocked');
}

module.exports = t.summary();
