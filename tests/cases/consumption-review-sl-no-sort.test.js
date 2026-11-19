// Real, directly-reported problem (2026-11-20): the consumption review table (and the confirmed
// Consumption Register itself) sorts by SL.NO so cross-checking a row against the physical sheet is
// predictable even though the model's own row order isn't stable run to run (confirmed directly — the
// same page came back in reverse SL.NO order on one extraction). But the OLD sort compared the literal
// digit, so a row printed "0" (this form's own stand-in for "4", confirmed on every sample seen) sorted
// to the very FRONT of the table instead of into its real 4th position — exactly backwards from the
// physical page, which is the opposite of what the sort exists to do. EditableTable's new
// slNoZeroMeansFour prop (true only for the consumption register, see its two call sites in
// client/src/App.jsx) fixes this without changing sort behavior for any other register's sl_no column,
// if one is ever added.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadSlNoCompare(slNoZeroMeansFour) {
  const sortValueBlock = extractBlock(APP_JSX, 'const slNoSortValue = (slNo) => {', (l) => l.trim() === '};');
  const compareBlock = extractBlock(APP_JSX, 'const slNoCompare = (a, b) => {', (l) => l.trim() === '};');
  return buildFromSource([sortValueBlock, compareBlock], 'slNoCompare', ['slNoZeroMeansFour'], [slNoZeroMeansFour]);
}

const t = makeRecorder('consumption-review-sl-no-sort');

function row(sl_no) { return { sl_no }; }

/* ===== 1. slNoZeroMeansFour: false (every register except consumption, including consumption's own
   table before this fix) — "0" sorts as literal 0, to the front, same as always. ===== */
{
  const cmp = loadSlNoCompare(false);
  const rows = ['1', '2', '3', '0', '5', '6'].map(row);
  rows.sort(cmp);
  t.assertEqual(rows.map(r => r.sl_no), ['0', '1', '2', '3', '5', '6'], 'OFF: literal "0" sorts first, unchanged behavior for non-consumption tables');
}

/* ===== 2. slNoZeroMeansFour: true — "0" sorts as if it were "4", landing in its real physical
   position, matching the exact live case (SL.NO 1,2,3,0,5,6,...,20 printed on the real sheet). ===== */
{
  const cmp = loadSlNoCompare(true);
  const rows = ['1', '2', '3', '0', '5', '6'].map(row);
  rows.sort(cmp);
  t.assertEqual(rows.map(r => r.sl_no), ['1', '2', '3', '0', '5', '6'], 'ON: "0" sorts into its real 4th-position slot, matching the physical page top-to-bottom');
}

/* ===== 3. slNoZeroMeansFour: true, full 20-row page, fed in REVERSED (worst-case) extraction order —
   confirmed directly as a real response shape (one run came back 20,19,...,1) — still sorts correctly
   into physical reading order. ===== */
{
  const cmp = loadSlNoCompare(true);
  const rows = ['20', '19', '18', '17', '16', '15', '14', '13', '12', '11', '10', '9', '8', '7', '6', '5', '0', '3', '2', '1'].map(row);
  rows.sort(cmp);
  t.assertEqual(
    rows.map(r => r.sl_no),
    ['1', '2', '3', '0', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20'],
    'ON, REVERSED INPUT: still lands in correct physical top-to-bottom order regardless of model output order'
  );
}

/* ===== 4. slNoZeroMeansFour: true, a genuinely illegible/non-numeric SL.NO — never crashes, falls
   back to the string compare exactly as before (an edge case this fix must not break). ===== */
{
  const cmp = loadSlNoCompare(true);
  const rows = ['1', '', '3'].map(row);
  let threw = false;
  try { rows.sort(cmp); } catch (e) { threw = true; }
  t.assertTrue(!threw, 'ON, ILLEGIBLE: sorting a blank/non-numeric SL.NO never throws');
}

module.exports = t.summary();
