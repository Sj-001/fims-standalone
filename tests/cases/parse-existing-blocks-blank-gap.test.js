// Real corruption, confirmed directly against a real customer sheet, 2026-11-03: parseExistingBlocks
// used to stop scanning a block's existing rows at the FIRST blank date cell, on the assumption that a
// block's blanks are always trailing padding (GENERATED_PAD_ROWS writes blank-dated filler specifically
// so this scan stops right after the real data). That assumption broke when a row in the MIDDLE of a
// block's real data went blank -- a person clearing a row's cells by hand instead of deleting the row
// structurally empties its content without shifting anything up. The scan hit that one blank row,
// stopped immediately, and reported the block as having ZERO existing rows even though 20 real rows sat
// right below it -- so the next push inserted a new row at the very top, shoving every real row down by
// one (reproduced live: existing=[] logged for a block that visibly had 20 rows in the real sheet).
// Tolerating a run of blank rows before giving up fixes this without ever mistaking genuine trailing
// padding (hundreds of rows long) for more real data.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

function loadParseExistingBlocks() {
  const blocks = [
    extractLine(SHEETS_JS, 'const normalizeCellStr = (v) =>'),
    extractBlock(SHEETS_JS, 'const serialToDateStr = (n) => {', (l) => l.trim() === '};'),
    extractLine(SHEETS_JS, 'const normalizeDateCell = (v) =>'),
    extractBlock(SHEETS_JS, 'function canonicalDateKey(v) {', 'function normalizeDateDisplay(v) {', { endExclusive: true }),
    extractBlock(SHEETS_JS, 'function parseExistingBlocks(grid) {', 'function planChronologicalInserts(existingRowsOrdered, newRows) {', { endExclusive: true }),
  ];
  return buildFromSource(blocks, 'parseExistingBlocks');
}

const t = makeRecorder('parse-existing-blocks-blank-gap');
const parseExistingBlocks = loadParseExistingBlocks();

// Minimal real-shaped grid: title row, header row, then whatever data rows the test supplies.
function gridFor(title, dataRows) {
  const grid = [];
  grid[0] = [title];
  grid[1] = ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'];
  dataRows.forEach((row, i) => { grid[2 + i] = row; });
  return grid;
}

/* ===== THE LIVE CASE: row 3 (the block's first row) is blank, rows 4-23 hold 20 real dates below it
   -- exactly the real incident. Must find all 20 rows, not report the block as empty. ===== */
{
  const realDates = ['23.9.26', '25.9.26', '3.10.26', '5.10.26', '5.10.26', '6.10.26', '10.10.26', '12.10.26', '12.10.26', '13.10.26', '13.10.26', '14.10.26', '14.10.26', '17.10.26', '19.10.26', '21.10.26', '22.10.26', '23.10.26', '27.10.26', '31.10.26'];
  const dataRows = [['', 0, 0, 0, 0], ...realDates.map(d => [d, 0, 100, 0, 100])];
  const { blocks } = parseExistingBlocks(gridFor('butter bake 130gm x 30', dataRows));
  t.assertEqual(blocks.length, 1, 'LIVE CASE: the block itself is still found');
  t.assertEqual(blocks[0].existingRowsOrdered.length, 20, 'LIVE CASE: all 20 real rows below the single blank are found, not reported as zero');
  t.assertEqual(blocks[0].existingRowsOrdered[0].dateKey, '2026-09-23', 'LIVE CASE: the first REAL row found is 23.9.26, not the blank row above it');
  t.assertEqual(blocks[0].nextRowIdx, 23, 'LIVE CASE: nextRowIdx lands one past the last real row (row index 22, 0-indexed), not right after the blank');
}

/* ===== SAFETY: genuine trailing padding (GENERATED_PAD_ROWS-style, hundreds of blank-dated rows right
   after the real data) must still correctly end the block there -- this fix must never start treating
   padding as "a gap with more data to find." ===== */
{
  const realDates = ['1.10.26', '5.10.26', '6.10.26'];
  const padding = Array.from({ length: 50 }, () => ['', 0, 0, 0]);
  const dataRows = [...realDates.map(d => [d, 0, 100, 0, 100]), ...padding];
  const { blocks } = parseExistingBlocks(gridFor('some item', dataRows));
  t.assertEqual(blocks[0].existingRowsOrdered.length, 3, 'SAFETY: only the 3 real rows are found -- 50 rows of trailing padding never get swept in');
  t.assertEqual(blocks[0].nextRowIdx, 5, 'SAFETY: nextRowIdx lands right after the real data (row index 4, 0-indexed), not somewhere inside the padding');
}

/* ===== SAFETY: a 4-row gap in the middle is tolerated (below the 5-row threshold) -- real data found
   after it is still picked up, and the gap rows themselves are not reported as real rows. ===== */
{
  const dataRows = [
    ['1.10.26', 0, 100, 0, 100],
    ['', 0, 0, 0], ['', 0, 0, 0], ['', 0, 0, 0], ['', 0, 0, 0], // 4 blank rows -- under the threshold
    ['10.10.26', 0, 200, 0, 200],
  ];
  const { blocks } = parseExistingBlocks(gridFor('some item', dataRows));
  t.assertEqual(blocks[0].existingRowsOrdered.length, 2, 'SAFETY: a 4-row gap is tolerated -- both real rows, before and after it, are found');
  t.assertEqual(blocks[0].existingRowsOrdered.map(r => r.dateKey), ['2026-10-01', '2026-10-10'], 'SAFETY: the two real rows are exactly the ones on either side of the gap');
  t.assertEqual(blocks[0].nextRowIdx, 8, 'SAFETY: nextRowIdx lands after the real row past the gap, not inside the gap');
}

/* ===== SAFETY: a 5-row gap hits the threshold and ends the block there -- this is the actual boundary
   the user asked for ("scan for 5 rows to consider a block empty"), confirmed exactly at the edge. ===== */
{
  const dataRows = [
    ['1.10.26', 0, 100, 0, 100],
    ['', 0, 0, 0], ['', 0, 0, 0], ['', 0, 0, 0], ['', 0, 0, 0], ['', 0, 0, 0], // 5 blank rows -- at the threshold
    ['10.10.26', 0, 200, 0, 200],
  ];
  const { blocks } = parseExistingBlocks(gridFor('some item', dataRows));
  t.assertEqual(blocks[0].existingRowsOrdered.length, 1, 'SAFETY: a 5-row gap hits the threshold -- the block is considered to end there');
  t.assertEqual(blocks[0].existingRowsOrdered[0].dateKey, '2026-10-01', 'SAFETY: only the row before the 5-row gap is found');
  t.assertEqual(blocks[0].nextRowIdx, 3, 'SAFETY: nextRowIdx lands right after the single real row, before the gap -- the row past the gap is never seen');
}

/* ===== SAFETY: a genuinely empty block (blank from row 1) still reports zero rows -- this fix changes
   WHEN scanning gives up, not whether a truly empty block is still recognized as empty. ===== */
{
  const dataRows = Array.from({ length: 10 }, () => ['', 0, 0, 0]);
  const { blocks } = parseExistingBlocks(gridFor('brand new item', dataRows));
  t.assertEqual(blocks[0].existingRowsOrdered.length, 0, 'SAFETY: a block with no real data anywhere is still correctly reported as empty');
  t.assertEqual(blocks[0].nextRowIdx, 2, 'SAFETY: nextRowIdx for a genuinely empty block stays right after the header');
}

module.exports = t.summary();
