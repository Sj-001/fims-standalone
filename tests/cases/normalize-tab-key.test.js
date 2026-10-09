// normalizeTabKey (server/lib/sheets.js) matches a push's item title against existing Sheet block
// titles. Root cause of a real duplicate block ("T 50 32G * 120 PKT (new)" vs "...new" with no
// parens): it never stripped punctuation, only whitespace, so a parenthesis difference alone made two
// identical blocks look like two different ones and the second got silently recreated. Fixed by adding
// a final strip-all-non-alphanumeric step. This file guards that fix plus every alias/pattern this
// function is already relied on for, so a future change to this function can't silently re-break one
// while fixing another.
const path = require('path');
const { extractBlock, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const SHEETS_JS = path.join(__dirname, '..', '..', 'server', 'lib', 'sheets.js');

function loadNormalizeTabKey() {
  const block = extractBlock(
    SHEETS_JS,
    'const TAB_KEY_ALIASES = [',
    (l) => l.trim() === '};',
  );
  return buildFromSource(block, 'normalizeTabKey');
}

const t = makeRecorder('normalize-tab-key');
const normalizeTabKey = loadNormalizeTabKey();

// *** THE EXACT REPORTED CASE ***
t.assertEqual(normalizeTabKey('T 50 32G * 120 PKT (new)'), normalizeTabKey('T 50 32G * 120 PKT new'), 'PARENS: "(new)" and "new" now match');

// The ORIGINAL (parens-less, no "new" at all) block must stay DISTINCT from the "(new)" variant.
t.assertNotEqual(normalizeTabKey('T 50 32G * 120 PKT'), normalizeTabKey('T 50 32G * 120 PKT (new)'), 'PARENS: the original (no "new") stays distinct from the "(new)" variant');
t.assertNotEqual(normalizeTabKey('T 50 32G * 120 PKT'), normalizeTabKey('T 50 32G * 120 PKT new'), 'PARENS: the original also stays distinct from the parens-less "new" variant');

t.assertEqual(normalizeTabKey('CREAM 30G*144PKT (new)'), normalizeTabKey('CREAM 30G*144PKT new'), 'PARENS: CREAM (new) variant also matches');

// --- Regression: every previously-confirmed alias/pattern still works exactly as before ---
t.assertEqual(normalizeTabKey('Jeera Dhamal 70g'), normalizeTabKey('jeera 70g'), 'REGR: dhamal stripped, still matches');
t.assertEqual(normalizeTabKey('N200 Jumbo Container'), normalizeTabKey('N200 J CONTAINER'), 'REGR: J -> Jumbo alias still works');
t.assertEqual(normalizeTabKey('E 900 Container'), normalizeTabKey('E900 CONT.'), 'REGR: CONT. -> Container alias still works');
t.assertEqual(normalizeTabKey('Cream Burst 30g x140 pkt'), normalizeTabKey('CREAM 30G*140PKT'), 'REGR: burst-stripping + pack-count stripping still works together');
t.assertEqual(normalizeTabKey('T-GEL CONT'), normalizeTabKey('T GEL CONTAINER'), 'REGR: hyphen-to-space + CONT alias still works');
t.assertEqual(normalizeTabKey('Blacko 32g x120 new'), normalizeTabKey('Blacko 32g x120 new'), 'REGR: identical strings still match (sanity)');
t.assertNotEqual(normalizeTabKey('Blacko 32g x120 new'), normalizeTabKey('Blacko 32g'), 'REGR: "x120 new" variant still stays distinct from the base "Blacko 32g" item');
t.assertEqual(normalizeTabKey('65GM x 70 pkt'), normalizeTabKey('65g'), 'REGR: gm->g + trailing pack-count-with-unit stripping still works');

module.exports = t.summary();
