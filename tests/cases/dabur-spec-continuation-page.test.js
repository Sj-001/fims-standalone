// Real gap, confirmed directly against real 2-page Dabur spec sheets the user uploaded ("CFB Bucket
// Pack (1Lx4) RGP'26", "CFB Real Standup Pouch 3L RGP'26"): every page of an uploaded PDF is extracted
// completely independently (see runExtraction/extractQueue in App.jsx) -- there was never any merging
// across pages anywhere in the pipeline. For these two documents, the partition spec lives ONLY on
// page 2, which has no item header printed on it at all (just a few more rows continuing page 1's
// table). Page 1 correctly produced a row with everything EXCEPT partition data; page 2, having no
// item_code to key a row off, produced nothing at all -- its partition data was silently discarded,
// never reaching the row it belonged to.
//
// isContinuationPage/mergeContinuationFields (on the dabur_pm_spec_cutting DOCUMENT_TYPES entry) fix
// this: a page with no item_code but SOME partition/plate data is recognized as a continuation instead
// of "nothing here", and its fields are folded into the preceding page's row from the same upload
// (tryMergeContinuationPage, tested against the real field shapes seen on these exact documents).
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

const numLine = extractBlock(APP_JSX, 'const num = (v) => {', (l) => l.trim() === '};');
const entryBlockRaw = extractBlock(
  APP_JSX,
  "key: 'dabur_pm_spec_cutting',",
  "key: 'dabur_po',",
  { endExclusive: true },
);
// entryBlockRaw starts mid-object (at the `key:` line, missing the entry's own opening `{`) and ends
// with the entry's own closing `},` followed by the NEXT entry's bare opening `{` (endExclusive only
// drops the matched "key: 'dabur_po'," line itself, not the brace right before it) -- strip that
// trailing `{`, then strip the array's own trailing comma, before wrapping it back into a real object
// literal so `const entry = { ... };` is valid.
const entryBody = `{\n${entryBlockRaw.replace(/\{\s*$/, '').trimEnd().replace(/,$/, '')}`;

const daburSpecEntry = buildFromSource(
  [numLine, `const entry = ${entryBody};\nreturn entry;`],
  'entry',
  ['genId'],
  [() => 'test-id'],
);

const t = makeRecorder('dabur-spec-continuation-page');

/* ===== isContinuationPage ===== */
{
  // THE LIVE CASE: "CFB Bucket Pack" page 2 -- no item_code, real partition data present.
  const page2Raw = {
    item_code: '', item_name: '',
    partition_tier_count: 1, partition_longer_length_mm: 200, partition_longer_height_mm: 257,
    partition_longer_qty_total: 2, partition_shorter_width_mm: 0, partition_shorter_height_mm: 0,
    partition_shorter_qty_total: 0, paper_comb_partitions: 'Paper specs same as outer box',
    no_of_plate: 0,
  };
  t.assertTrue(daburSpecEntry.isContinuationPage(page2Raw), 'LIVE CASE: a page with no item_code but real partition data is recognized as a continuation');
}
{
  // A genuinely blank/irrelevant page (nothing on it at all) must NOT be treated as a continuation --
  // still correctly falls through to "nothing here, skip".
  const blankRaw = { item_code: '', item_name: '', partition_tier_count: 0, no_of_plate: 0 };
  t.assertTrue(!daburSpecEntry.isContinuationPage(blankRaw), 'BLANK PAGE: no item_code AND no partition/plate data -- not a continuation');
}
{
  // A normal, complete primary page (has its own item_code) is never treated as a continuation, even
  // if it also happens to have partition data on the same page.
  const normalRaw = { item_code: '1223310843', item_name: 'CFB Bucket Pack', partition_tier_count: 1, no_of_plate: 0 };
  t.assertTrue(!daburSpecEntry.isContinuationPage(normalRaw), 'NORMAL PAGE: has its own item_code -- never a continuation regardless of its own content');
}
{
  // A continuation page whose only addition is Central Plate data (no partition) must still be
  // recognized -- the check must not require partition data specifically.
  const plateOnlyRaw = { item_code: '', item_name: '', partition_tier_count: 0, no_of_plate: 1 };
  t.assertTrue(daburSpecEntry.isContinuationPage(plateOnlyRaw), 'PLATE-ONLY CONTINUATION: no_of_plate alone is enough to recognize a continuation');
}

/* ===== mergeContinuationFields: THE LIVE CASE, exact field shapes from "CFB Bucket Pack (1Lx4) RGP'26" ===== */
{
  // Page 1's raw payload, as the model would have produced it -- real box dims, paper combination,
  // genuinely NO partition/plate data (none of it is printed on page 1 of this document).
  const page1Raw = {
    item_code: '1223310843', item_name: "CFB Bucket Pack (1Lx4) RGP'26",
    no_of_ply: 5, box_length_min: 615, box_length_max: 621, box_width_min: 202, box_width_max: 208,
    box_height_min: 262, box_height_max: 268, paper_comb: '150gsm SK x 5 Ply, SK – Semi Kraft',
    partition_tier_count: 0, partition_longer_length_mm: 0, partition_longer_height_mm: 0,
    partition_longer_qty_total: 0, partition_shorter_width_mm: 0, partition_shorter_height_mm: 0,
    partition_shorter_qty_total: 0, paper_comb_partitions: '',
    no_of_plate: 0, paper_comb_plate: '', plate_length_mm: 0, plate_width_mm: 0,
  };
  // Page 2's raw payload -- "Number of Sets of Partition: No. of partitions - 2 Vertical Plates,
  // Vertical Plate Size - 200mm (W) x 257mm (H)" -- real text from the real document.
  const page2Raw = {
    item_code: '', item_name: '',
    partition_tier_count: 1, partition_longer_length_mm: 200, partition_longer_height_mm: 257,
    partition_longer_qty_total: 2, partition_shorter_width_mm: 0, partition_shorter_height_mm: 0,
    partition_shorter_qty_total: 0, paper_comb_partitions: 'Paper specs same as outer box',
    no_of_plate: 0, paper_comb_plate: '', plate_length_mm: 0, plate_width_mm: 0,
  };
  const merged = daburSpecEntry.mergeContinuationFields(page1Raw, page2Raw);
  t.assertEqual(merged.item_code, '1223310843', 'LIVE CASE: the merged row keeps page 1\'s own item_code');
  t.assertEqual(merged.box_length_min, 615, 'LIVE CASE: the merged row keeps page 1\'s own box dimensions, untouched');
  t.assertEqual(merged.partition_tier_count, 1, 'LIVE CASE: the merged row picks up page 2\'s partition tier count');
  t.assertEqual(merged.partition_longer_length_mm, 200, 'LIVE CASE: the merged row picks up page 2\'s partition dimensions');
  t.assertEqual(merged.partition_longer_height_mm, 257, 'LIVE CASE: the merged row picks up page 2\'s partition dimensions');

  // And the fully-shaped row (what actually lands in the register) must show real, non-zero partition
  // data -- not silently blank, which is exactly the live bug this confirms is fixed.
  const shapedRows = daburSpecEntry.shape(merged);
  t.assertEqual(shapedRows.length, 1, 'LIVE CASE: the merged raw still shapes into exactly one row');
  t.assertEqual(shapedRows[0].no_of_partitions, 1, 'LIVE CASE: the final row\'s partition count is no longer zero');
  t.assertTrue(shapedRows[0].partition_longer_piece !== '', 'LIVE CASE: the final row\'s partition piece size is no longer blank');
}

/* ===== mergeContinuationFields: a plate-only continuation must never wipe out partition data a
   DIFFERENT continuation page already merged in (field-group isolation). ===== */
{
  const alreadyMergedRaw = {
    item_code: 'X', partition_tier_count: 1, partition_longer_length_mm: 200, partition_longer_height_mm: 257,
    partition_longer_qty_total: 2, partition_shorter_width_mm: 0, partition_shorter_height_mm: 0,
    partition_shorter_qty_total: 0, paper_comb_partitions: 'Paper specs same as outer box',
    no_of_plate: 0, paper_comb_plate: '', plate_length_mm: 0, plate_width_mm: 0,
  };
  const plateOnlyPageRaw = {
    item_code: '', partition_tier_count: 0, no_of_plate: 1, paper_comb_plate: '3-ply, 120/120/120',
    plate_length_mm: 300, plate_width_mm: 225,
  };
  const merged = daburSpecEntry.mergeContinuationFields(alreadyMergedRaw, plateOnlyPageRaw);
  t.assertEqual(merged.partition_tier_count, 1, 'ISOLATION: a plate-only continuation never wipes out partition data already merged in');
  t.assertEqual(merged.no_of_plate, 1, 'ISOLATION: the plate-only continuation\'s own data is still picked up');
  t.assertEqual(merged.plate_length_mm, 300, 'ISOLATION: plate dimensions are picked up correctly');
}

module.exports = t.summary();
