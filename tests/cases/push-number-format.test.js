// Real bug, confirmed directly against a real customer sheet, 2026-10-31: a brand-new block's Opening
// cell showed "30-Dec" and its Closing cell showed "3-Mar" instead of the plain numbers they actually
// held (0 and 4080) -- Google Sheets stores a date as a day-count from Dec 30 1899, so those displayed
// strings were exactly that correct number rendered through whatever numberFormat happened to already
// be sitting in that spot on the sheet. This app had never once set numberFormat anywhere before
// writing, so a cell's display depended entirely on the sheet's own leftover formatting history --
// confirmed in the sheet itself, not guessed. buildNumberFormatRequestsForPatches resets every written
// cell to a plain NUMBER format, same repeatCell technique already used for the highlight requests, so
// display no longer depends on what a given spot in a given sheet happened to be formatted as before.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

function loadBuildNumberFormatRequestsForPatches() {
  const block = extractBlock(
    SHEETS_JS,
    'function buildNumberFormatRequestsForPatches(sheetId, patches) {',
    (l) => l.trim() === '}',
  );
  return buildFromSource(block, 'buildNumberFormatRequestsForPatches');
}

const t = makeRecorder('push-number-format');
const buildNumberFormatRequestsForPatches = loadBuildNumberFormatRequestsForPatches();

/* ===== THE LIVE CASE: a brand-new block's first data row -- Date/Opening/Production/Dispatch/Closing,
   exactly the shape computeMergePatches writes for a genuinely new block -- gets a numberFormat reset
   covering the row's full width, anchored at the patch's own position. ===== */
{
  const patches = [
    { startRow0: 2, startCol0: 20, values: [['\'29.10.26', 0, 4080, 0, '=V3+W3-X3']], isNewRow: true },
  ];
  const requests = buildNumberFormatRequestsForPatches(123, patches);
  t.assertEqual(requests.length, 1, 'LIVE CASE: one repeatCell request per patch');
  const r = requests[0].repeatCell;
  t.assertEqual(r.range.sheetId, 123, 'LIVE CASE: targets the real tab\'s sheetId');
  t.assertEqual(r.range.startRowIndex, 2, 'LIVE CASE: row matches the patch exactly');
  t.assertEqual(r.range.endRowIndex, 3, 'LIVE CASE: one row tall, matching a single data row');
  t.assertEqual(r.range.startColumnIndex, 20, 'LIVE CASE: column matches the patch exactly (Date column included)');
  t.assertEqual(r.range.endColumnIndex, 25, 'LIVE CASE: spans the full 5-column Date/Opening/Production/Dispatch/Closing width');
  t.assertEqual(r.cell.userEnteredFormat.numberFormat.type, 'NUMBER', 'LIVE CASE: resets to plain NUMBER, not DATE');
  t.assertEqual(r.fields, 'userEnteredFormat.numberFormat', 'LIVE CASE: touches only numberFormat -- never background color, borders, or anything else already on the cell');
}

/* ===== SAFETY: the Date column is included in the reset range, but that's harmless -- forceTextValue's
   leading apostrophe makes the Date cell's value explicitly TEXT, and Sheets never reinterprets an
   explicit-TEXT cell's display based on numberFormat, so resetting it alongside the numeric columns in
   one request can never make a real date display wrong. ===== */
{
  const patches = [{ startRow0: 5, startCol0: 0, values: [['\'1.10.26', '=D4', 100, 0, '=B5+C5-D5']], isNewRow: true }];
  const requests = buildNumberFormatRequestsForPatches(1, patches);
  t.assertEqual(requests[0].repeatCell.range.startColumnIndex, 0, 'SAFETY: the reset range starts at the Date column, not skipped');
}

/* ===== MULTIPLE PATCHES: a push touching several rows (new block data rows, plus the 1-cell
   Opening/Closing formula-chain repairs on a pre-existing row) gets one request per patch, each scoped
   to exactly that patch's own range -- never a blanket whole-tab reset that could touch cells this push
   didn't write. ===== */
{
  const patches = [
    { startRow0: 2, startCol0: 20, values: [['\'29.10.26', 0, 4080, 0, '=V3+W3-X3']], isNewRow: true },
    { startRow0: 2, startCol0: 1, values: [['=E1']] }, // Opening formula-chain repair on an existing row
    { startRow0: 2, startCol0: 4, values: [['=B3+C3-D3']] }, // Closing formula-chain repair, same row
  ];
  const requests = buildNumberFormatRequestsForPatches(9, patches);
  t.assertEqual(requests.length, 3, 'MULTIPLE PATCHES: one request per patch, not merged or deduped');
  t.assertEqual(requests[1].repeatCell.range.startColumnIndex, 1, 'MULTIPLE PATCHES: the Opening repair request is scoped to its own single column');
  t.assertEqual(requests[1].repeatCell.range.endColumnIndex, 2, 'MULTIPLE PATCHES: the Opening repair request is exactly 1 column wide');
}

/* ===== SAFETY: no sheetId (tab didn't exist yet when metadata was fetched, or some other edge case
   buildHighlightRequestsForPatches already guards against identically) returns no requests rather than
   crashing or sending a request Sheets would reject. ===== */
{
  t.assertEqual(buildNumberFormatRequestsForPatches(undefined, [{ startRow0: 0, startCol0: 0, values: [[1]] }]).length, 0, 'SAFETY: undefined sheetId -> no requests');
  t.assertEqual(buildNumberFormatRequestsForPatches(null, [{ startRow0: 0, startCol0: 0, values: [[1]] }]).length, 0, 'SAFETY: null sheetId -> no requests');
  t.assertEqual(buildNumberFormatRequestsForPatches(5, []).length, 0, 'SAFETY: no patches -> no requests');
  t.assertEqual(buildNumberFormatRequestsForPatches(5, undefined).length, 0, 'SAFETY: undefined patches -> no requests, not a crash');
}

module.exports = t.summary();
