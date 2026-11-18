// Real, repeated failure confirmed directly against the live test deploy (2026-11-18): a real
// consumption-report page where SL.NO 1 through 20 ALL genuinely had handwritten data had several
// strips from the MIDDLE of that run silently dropped by the extraction model across four separate
// runs (6 or 7 of the 20 real rows came back each time) -- with nothing in the output visibly wrong.
// Tried strengthening the prompt's own "count every strip" self-check first; that backfired badly
// (made the drop WORSE, down to 4-7 rows even on a page the original prompt had read correctly), so it
// was reverted. This is the actual fix: a deterministic, code-level gap check in consumption_sheet's
// shape() that can't change what the model extracts, only flags the row right after an interior SL.NO
// gap (one with real rows on both sides of it) for a human to check against the original sheet --
// since on this form a genuinely unused part of the grid is always a TRAILING run, never sandwiched.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

// A plain `(l) => l.trim() === '}'` end anchor stops at the FIRST brace at any indentation, not the
// function's real closing one -- normalizeDateToDots has several nested `if (m) { ... }` blocks of its
// own. This tracks brace depth across lines instead (every `{{...}}`-style regex/template interpolation
// in this file's extracted functions is self-balanced, so a naive per-character count is reliable here).
function endAtDepthZero(startDepth) {
  let depth = startDepth;
  return (l) => {
    for (const ch of l) { if (ch === '{') depth++; else if (ch === '}') depth--; }
    return depth === 0;
  };
}

function loadConsumptionShape() {
  const blocks = [
    extractBlock(APP_JSX, 'const num = (v) => {', (l) => l.trim() === '};'),
    extractLine(APP_JSX, 'const MONTH_ABBR_TO_NUM = {'),
    extractBlock(APP_JSX, 'function normalizeDateToDots(raw) {', endAtDepthZero(1)),
    extractLine(APP_JSX, 'const REAL_DATE_RE ='),
    extractBlock(APP_JSX, 'function fillDittoDates(rows) {', (l) => l.trim() === '}'),
    extractBlock(APP_JSX, 'function normalizeNumericStr(raw) {', (l) => l.trim() === '}'),
    extractBlock(APP_JSX, "key: 'consumption_sheet',", "key: 'production_sheet',", { endExclusive: true }),
  ];
  // The consumption_sheet block starts mid-object (`key: '...',`) and ends with its own closing `},`
  // followed by the next entry's bare opening `{` (endExclusive drops only the matched `key:` line, not
  // the brace right before it) -- same pattern tests/cases/dabur-spec-continuation-page.test.js uses.
  const entryBlockRaw = blocks[blocks.length - 1];
  const entryBody = `{\n${entryBlockRaw.replace(/\{\s*$/, '').trimEnd().replace(/,$/, '')}`;
  const rest = blocks.slice(0, -1);
  return buildFromSource(
    [...rest, `const entry = ${entryBody};\nreturn entry.shape;`],
    'entry',
    ['genId'],
    [() => 'test-id'],
  );
}

const t = makeRecorder('consumption-sheet-dropped-row-flag');
const shape = loadConsumptionShape();

function item(sl_no, extra) { return { sl_no, date: '1/10/2026', shade: 'NS', gsm: '140', size: '59.5', weight_consumed: 900, ...extra }; }

/* ===== 1. THE LIVE CASE: SL.NO 1-20 all extracted correctly (no gap) -- nothing extra flagged. ===== */
{
  const items = Array.from({ length: 20 }, (_, i) => item(String(i + 1)));
  const rows = shape({ date: '1/10/2026', items });
  t.assertEqual(rows.filter(r => r.flagged).length, 0, 'FULL SET: no gap, nothing flagged for a missing row');
}

/* ===== 2. THE EXACT LIVE FAILURE: SL.NO 1,2,8,9,14,15,20 only -- real data on both sides of several
   gaps (3-7, 10-13, 16-19) -- every row right after an interior gap gets flagged. ===== */
{
  const items = ['1', '2', '8', '9', '14', '15', '20'].map(n => item(n));
  const rows = shape({ date: '1/10/2026', items });
  const flaggedSlNos = rows.filter(r => r.flagged).map(r => r.sl_no);
  t.assertEqual(flaggedSlNos.sort(), ['14', '20', '8'], 'LIVE FAILURE: the row right after each interior gap is flagged (2->8, 9->14, 15->20)');
  t.assertTrue(rows.find(r => r.sl_no === '8').flagReason.includes('3') && rows.find(r => r.sl_no === '8').flagReason.includes('7'), 'LIVE FAILURE: the flag names the actual missing range (3 through 7)');
}

/* ===== 3. SAFETY: a genuinely partial day (SL.NO 1-10 real, 11-20 truly blank and never extracted) is
   a TRAILING gap, not an interior one -- nothing flagged, this is the normal, expected pattern. ===== */
{
  const items = Array.from({ length: 10 }, (_, i) => item(String(i + 1)));
  const rows = shape({ date: '1/10/2026', items });
  t.assertEqual(rows.filter(r => r.flagged).length, 0, 'SAFETY: a trailing blank tail (11-20 never extracted) is never flagged as a gap');
}

/* ===== 4. SAFETY: the "0 printed in place of 4" quirk -- SL.NO 1,2,3,0,5 is a complete, contiguous run
   on this form's own numbering, not a gap at "4". ===== */
{
  const items = ['1', '2', '3', '0', '5'].map(n => item(n));
  const rows = shape({ date: '1/10/2026', items });
  t.assertEqual(rows.filter(r => r.flagged).length, 0, 'SAFETY: "0" filling the printed "4" slot is not mistaken for a gap');
}

/* ===== 5. SAFETY: a single missing row (SL.NO 5 skipped between 1-4 and 6-10) is still caught, message
   names just the one missing number, singular wording. ===== */
{
  const items = ['1', '2', '3', '0', '6', '7'].map(n => item(n));
  const rows = shape({ date: '1/10/2026', items });
  const flagged = rows.find(r => r.flagged);
  t.assertTrue(!!flagged, 'SAFETY: a single-row interior gap is still caught');
  t.assertEqual(flagged.sl_no, '6', 'SAFETY: flagged on the row right after the single gap');
  t.assertTrue(flagged.flagReason.includes('seems to be missing'), 'SAFETY: singular wording for exactly one missing SL.NO');
}

/* ===== 6. SAFETY: this new check never overrides an existing flag (e.g. a leftover>=weight flag) --
   additive only, never clobbers a more specific reason already set for that row. ===== */
{
  const items = ['1', '2', '8'].map(n => item(n));
  items[2].leftover_weight = 950; // >= weight_consumed (900) -> already flagged by the existing check
  const rows = shape({ date: '1/10/2026', items });
  const row8 = rows.find(r => r.sl_no === '8');
  t.assertTrue(row8.flagged, 'SAFETY: still flagged (by the pre-existing leftover check)');
  t.assertTrue(row8.flagReason.includes('reel weight'), 'SAFETY: the pre-existing, more specific flag reason is kept, not overwritten by the gap message');
}

/* ===== 7. SAFETY: a row with an illegible/unparseable SL.NO is skipped for gap-detection purposes,
   never crashes, and never itself gets a spurious gap flag. ===== */
{
  const items = [item('1'), item(''), item('3')];
  const rows = shape({ date: '1/10/2026', items });
  t.assertEqual(rows.length, 3, 'SAFETY: an illegible SL.NO row is still kept, not dropped');
  t.assertTrue(!rows[1].flagged, 'SAFETY: the illegible-SL.NO row itself is not flagged by the gap check');
}

module.exports = t.summary();
