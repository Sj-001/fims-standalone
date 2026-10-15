// Two real additions to the Dabur PM Spec (cutting sheet) extraction, requested directly: Compression
// Strength and Bursting Strength weren't captured at all (Compression Strength was explicitly on the
// prompt's own ignore list; Bursting Strength was simply never mentioned), and sheet_size_inch was
// displayed as "length x width" when the factory's own convention reads it "width x length". Both
// confirmed against the shape() function extracted straight out of the real dabur_pm_spec_cutting
// DOCUMENT_TYPES entry, the same extraction harness dabur-spec-continuation-page.test.js uses.
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
const entryBody = `{\n${entryBlockRaw.replace(/\{\s*$/, '').trimEnd().replace(/,$/, '')}`;

const daburSpecEntry = buildFromSource(
  [numLine, `const entry = ${entryBody};\nreturn entry;`],
  'entry',
  ['genId'],
  [() => 'test-id'],
);

const t = makeRecorder('dabur-spec-strength-and-sheet-size');

// Real box dimensions from "CFB Bucket Pack (1Lx4) RGP'26" (Length 615-621mm avg 618, Width 202-208mm
// avg 205, Height 262-268mm avg 265) -- a full, otherwise-ordinary raw payload used as the base for
// every case below so each test only has to vary what it's actually checking.
const baseRaw = () => ({
  item_code: '1223310843', item_name: "CFB Bucket Pack (1Lx4) RGP'26",
  no_of_ply: 5, box_length_min: 615, box_length_max: 621, box_width_min: 202, box_width_max: 208,
  box_height_min: 262, box_height_max: 268, paper_comb: '150gsm SK x 5 Ply, SK – Semi Kraft',
  partition_tier_count: 0, no_of_plate: 0,
});

/* ===== Compression Strength / Bursting Strength pass straight through to the final row ===== */
{
  const raw = { ...baseRaw(), compression_strength: 'NLT 300 kgf (with partition)', bursting_strength: 'NLT 11.5 Kg / cm²' };
  const rows = daburSpecEntry.shape(raw);
  t.assertEqual(rows.length, 1, 'produces exactly one row');
  t.assertEqual(rows[0].compression_strength, 'NLT 300 kgf (with partition)', 'compression_strength carries through exactly as given');
  t.assertEqual(rows[0].bursting_strength, 'NLT 11.5 Kg / cm²', 'bursting_strength carries through exactly as given');
}
// The split-column real case (CFB Real Standup Pouch 3L): SPECIFICATION cell alone just says "Not Less
// Than", with the actual number only in a separate TARGET column -- the prompt's own field notes ask
// the model to combine these into one string itself; shape() just needs to pass whatever comes back
// through unchanged, without silently dropping it for being an unusual phrasing.
{
  const raw = { ...baseRaw(), compression_strength: 'Not Less Than 350 KG', bursting_strength: '' };
  const rows = daburSpecEntry.shape(raw);
  t.assertEqual(rows[0].compression_strength, 'Not Less Than 350 KG', 'the combined (SPECIFICATION + TARGET column) form passes through unchanged');
  t.assertEqual(rows[0].bursting_strength, '', 'a genuinely blank bursting_strength stays blank, not undefined or null');
}
{
  // Missing from raw entirely (e.g. an older cached extraction, or a page where the row just isn't
  // present) must never throw -- defaults to an empty string, same as every other optional text field
  // in this shape() function.
  const raw = baseRaw();
  const rows = daburSpecEntry.shape(raw);
  t.assertEqual(rows[0].compression_strength, '', 'missing compression_strength defaults to empty string, never throws');
  t.assertEqual(rows[0].bursting_strength, '', 'missing bursting_strength defaults to empty string, never throws');
}

/* ===== sheet_size_inch: width x length, not length x width ===== */
{
  // Real "CFB Bucket Pack" box: avg L=618mm/24.33in, W=205mm/8.07in, H=265mm/10.43in.
  // Sheet_len = (L_in + W_in) x 2 + 1.5 = (24.33+8.07)x2+1.5 = 66.3 (length)
  // Sheet_wid = (W_in + H_in) x floor(46 / (W_in+H_in)) = (8.07+10.43) x floor(46/18.5) = 18.5 x 2 = 37 (width)
  const raw = baseRaw();
  const rows = daburSpecEntry.shape(raw);
  const [widthPart, lengthPart] = rows[0].sheet_size_inch.split(' x ').map(Number);
  t.assertTrue(widthPart < lengthPart, `LIVE CASE: sheet_size_inch "${rows[0].sheet_size_inch}" reads width x length (the smaller "Sheet_wid" value first, the larger "Sheet_len" value second) -- matches the real computed values (~37 x ~66.3)`);
  t.assertTrue(Math.abs(widthPart - 37) < 1, 'the width-first number is the real computed Sheet_wid (~37), not Sheet_len');
  t.assertTrue(Math.abs(lengthPart - 66.3) < 1, 'the length-second number is the real computed Sheet_len (~66.3), not Sheet_wid');
}

module.exports = t.summary();
