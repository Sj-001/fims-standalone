// Covers "upload the same page twice, new rows appended below old ones": old rows already in a
// register must never silently reappear as duplicates, and never slip through completely unflagged
// either. Two layers, both tested here: dedupKeyForRow (exact match -> silently dropped, see addRows/
// dropAlreadyConfirmedDuplicates) and looseDateDescKey (date+description match with numbers ignored ->
// flagged for a person to check, see flagLikelyReReadDuplicates).
//
// Written after a real regression: a physical row already confirmed in the Production Register
// ("Butter Plus 28gm x 140 Pkt") came back unflagged as a brand-new row on re-extraction, because that
// day's catalog-match produced its canonical spelling instead ("BUTTER PLUS 28G*140PKT") — cosmetically
// different enough that the OLD looseDateDescKey (plain lowercase, no punctuation/unit normalization)
// never recognized the two as possibly the same entry, so the row never even reached the "flag it"
// branch. Confirmed live: 2026-10-09, against the real Production Register (date 8.10.26 upload,
// customer "anmol stock").
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

function loadDedupFunctions(registerState, abbreviations = []) {
  const dedupFieldsBlock = extractBlock(APP_JSX, 'const DEDUP_FIELDS = {', (l) => l.trim() === '};');
  const rowDedupKeyBlock = extractBlock(APP_JSX, 'function rowDedupKey(fields, row) {', (l) => l.trim() === '}');
  const applyAbbreviationsBlock = extractBlock(APP_JSX, 'const applyAbbreviations = (s) => {', (l) => l.trim() === '};');
  const normalizeForCatalogMatchBlock = extractBlock(
    APP_JSX,
    'const normalizeForCatalogMatch = (s) => applyAbbreviations(s || \'\')',
    ".replace(/[^a-z0-9]/g, '');"
  );
  const dedupKeyForRowBlock = extractBlock(APP_JSX, 'const dedupKeyForRow = (registerKey, row) => {', (l) => l.trim() === '};');
  const blankWildcardFieldsBlock = extractLine(APP_JSX, 'const BLANK_WILDCARD_DEDUP_FIELDS = {');
  const dedupFieldValueBlock = extractLine(APP_JSX, 'const dedupFieldValue = (row, f) =>');
  const rowsMatchForDedupBlock = extractBlock(APP_JSX, 'const rowsMatchForDedup = (registerKey, a, b) => {', (l) => l.trim() === '};');
  const disagreeingDedupFieldsBlock = extractBlock(APP_JSX, 'const disagreeingDedupFields = (registerKey, a, b) => {', (l) => l.trim() === '};');
  const dedupFieldLabelsBlock = extractBlock(APP_JSX, 'const DEDUP_FIELD_LABELS = {', (l) => l.trim() === '};');
  const looseDateDescKeyBlock = extractLine(APP_JSX, 'const looseDateDescKey = (row) =>');
  const structuredLooseFieldsBlock = extractBlock(APP_JSX, 'const STRUCTURED_LOOSE_KEY_FIELDS = {', (l) => l.trim() === '};');
  const structuredLooseKeyBlock = extractBlock(
    APP_JSX,
    'const structuredLooseKey = (registerKey, row) =>',
    (l) => l.trim().endsWith(".join('||');"),
  );
  const descriptionNoiseWordsBlock = extractLine(APP_JSX, 'const DESCRIPTION_NOISE_WORDS = new Set(');
  const significantWordsBlock = extractBlock(
    APP_JSX, 'const significantWords = (desc) => new Set(', (l) => l.trim() === ');',
  );
  const numericFingerprintFieldsBlock = extractLine(APP_JSX, 'const NUMERIC_FINGERPRINT_FIELDS = {');
  const findNumericFingerprintMatchBlock = extractBlock(
    APP_JSX, 'const findNumericFingerprintMatch = (registerKey, row, existing) => {', (l) => l.trim() === '};',
  );
  const flagLikelyReReadDuplicatesBlock = extractBlock(
    APP_JSX,
    'const flagLikelyReReadDuplicates = (registerKey, rows, extraExisting = []) => {',
    'const dropAlreadyConfirmedDuplicates = (registerKey, rows, extraExisting = []) => {',
    { endExclusive: true },
  );
  const dropAlreadyConfirmedDuplicatesBlock = extractBlock(
    APP_JSX,
    'const dropAlreadyConfirmedDuplicates = (registerKey, rows, extraExisting = []) => {',
    (l) => l.trim() === '};'
  );

  const blocks = [
    dedupFieldsBlock, rowDedupKeyBlock, applyAbbreviationsBlock, normalizeForCatalogMatchBlock,
    dedupKeyForRowBlock, blankWildcardFieldsBlock, dedupFieldValueBlock, rowsMatchForDedupBlock,
    disagreeingDedupFieldsBlock, dedupFieldLabelsBlock,
    looseDateDescKeyBlock, structuredLooseFieldsBlock, structuredLooseKeyBlock,
    descriptionNoiseWordsBlock, significantWordsBlock, numericFingerprintFieldsBlock, findNumericFingerprintMatchBlock,
  ];

  const dedupKeyForRow = buildFromSource(blocks, 'dedupKeyForRow', ['abbreviations'], [abbreviations]);
  const rowsMatchForDedup = buildFromSource(blocks, 'rowsMatchForDedup', ['abbreviations'], [abbreviations]);
  const looseDateDescKey = buildFromSource(blocks, 'looseDateDescKey', ['abbreviations'], [abbreviations]);
  const flagLikelyReReadDuplicates = buildFromSource(
    [...blocks, flagLikelyReReadDuplicatesBlock],
    'flagLikelyReReadDuplicates',
    ['abbreviations', 'registerState'],
    [abbreviations, registerState],
  );
  const dropAlreadyConfirmedDuplicates = buildFromSource(
    [...blocks, dropAlreadyConfirmedDuplicatesBlock],
    'dropAlreadyConfirmedDuplicates',
    ['abbreviations', 'registerState'],
    [abbreviations, registerState],
  );
  return { dedupKeyForRow, rowsMatchForDedup, looseDateDescKey, flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates };
}

const t = makeRecorder('dedup-reextraction');

/* ===== 1. The exact live regression: cosmetic re-spelling of an already-confirmed row must be flagged, not silent ===== */
{
  const confirmed = [{
    id: 'old1', date: '7.10.26', party: '', description: 'Butter Plus 28gm x 140 Pkt',
    customerHint: '', pieces: 5000, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ production: confirmed });
  // Same physical row, re-extracted with the catalog's own canonical spelling instead -- same date,
  // same real item, genuinely different-looking description text, and (for this check) a DIFFERENT
  // piece count so it's not an exact match either -- this is exactly what should get flagged.
  const reExtracted = [{
    id: 'new1', date: '7.10.26', description: 'BUTTER PLUS 28G*140PKT', pieces: 5500, dispatch: 0,
  }];
  const result = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(result[0].flagged, 'LIVE BUG: cosmetically re-spelled re-extraction of a confirmed row gets flagged for review');
}

/* ===== 2. Identical re-extraction (true re-upload, no drift at all) is silently dropped, not flagged ===== */
{
  const confirmed = [{
    id: 'old1', date: '7.10.26', party: '', description: 'BUTTER PLUS 28G*140PKT',
    customerHint: '', pieces: 5000, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { dropAlreadyConfirmedDuplicates, flagLikelyReReadDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '7.10.26', description: 'BUTTER PLUS 28G*140PKT', pieces: 5000, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'exact re-upload is not flagged (nothing disagrees)');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'exact re-upload is silently dropped from the pre-confirm preview');
}

/* ===== 3. A genuinely different item/date must never be flagged just because it's unrelated ===== */
{
  const confirmed = [{
    id: 'old1', date: '7.10.26', party: '', description: 'BUTTER PLUS 28G*140PKT',
    customerHint: '', pieces: 5000, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ production: confirmed });
  const newItem = [{ id: 'new1', date: '8.10.26', description: 'Coconut Dream 31gm x 120', pieces: 2000, dispatch: 0 }];
  const result = flagLikelyReReadDuplicates('production', newItem);
  t.assertTrue(!result[0].flagged, 'a genuinely new item on a genuinely new date is never flagged');
}

/* ===== 4. A real misread-number case (same day, same catalog spelling, numbers disagree) still flags, as before ===== */
{
  const confirmed = [{
    id: 'old1', date: '3.10.26', party: '', description: 'Jeera Dhamal 32g x144 Packet',
    customerHint: '', pieces: 7560, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ production: confirmed });
  const misread = [{ id: 'new1', date: '3.10.26', description: 'Jeera Dhamal 32g x144 Packet', pieces: 7569, dispatch: 0 }];
  const result = flagLikelyReReadDuplicates('production', misread);
  t.assertTrue(result[0].flagged, 'REGRESSION: identical description, disagreeing numbers, still flags as before');
}

/* ===== 5. Exact DEDUP_FIELDS match (addRows' own final backstop) stays byte-exact, not loosened ===== */
{
  const { dedupKeyForRow } = loadDedupFunctions({});
  const a = { date: '7.10.26', party: '', description: 'Butter Plus 28gm x 140 Pkt', customerHint: '', pieces: 5000, dispatch: 0 };
  const b = { date: '7.10.26', party: '', description: 'BUTTER PLUS 28G*140PKT', customerHint: '', pieces: 5000, dispatch: 0 };
  t.assertNotEqual(
    dedupKeyForRow('production', a), dedupKeyForRow('production', b),
    'exact dedup key (the silent-drop path) is NOT loosened by this fix -- still requires byte-exact description'
  );
}

/* ===== 6. Raw Material In / Consumption structured loose-match is untouched by this change ===== */
{
  const confirmedReel = [{
    id: 'r1', date: '1.10.26', mill: 'Ashoka', reel_no: '12', size: '34', unit: 'Inch', gsm: '180', bf: '18',
    shade: 'NS', weight_kg: 450, consumed: '',
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ rawMaterialIn: confirmedReel });
  const misreadReelNo = [{ id: 'r2', date: '1.10.26', mill: 'Ashoka', reel_no: '19', size: '34', unit: 'Inch', gsm: '180', bf: '18', shade: 'NS', weight_kg: 450 }];
  const result = flagLikelyReReadDuplicates('rawMaterialIn', misreadReelNo);
  t.assertTrue(result[0].flagged, 'REGRESSION: structured-register loose match (reel_no disagreement) still flags as before');
}

/* ===== 7. party/customerHint are no longer part of a Production row's identity at all, per explicit
   instruction: a blank party (confirmed) vs ANY re-extraction reading of the same otherwise-identical
   row -- stray text, a real name, anything -- must NOT be flagged, and must be silently dropped from
   the preview. Originally added for a narrower "blank vs populated" tolerance (Production Register,
   "Butter Bake 130g x30 Packet" 3.10.26/4000 pieces, re-upload read a bracketed note as "Vijyant" that
   the original confirmed row never had); now covers any party reading at all, not just a blank one,
   since party was dropped from DEDUP_FIELDS.production entirely. ===== */
{
  const confirmed = [{
    id: 'old1', date: '3.10.26', party: '', description: 'Butter Bake 130g x30 Packet',
    customerHint: '', pieces: 4000, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates, dedupKeyForRow } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '3.10.26', party: 'Vijyant', description: 'Butter Bake 130g x30 Packet', customerHint: '', pieces: 4000, dispatch: 0 }];
  t.assertEqual(
    dedupKeyForRow('production', confirmed[0]), dedupKeyForRow('production', reExtracted[0]),
    'sanity: the byte-exact key now agrees despite party blank vs "Vijyant" -- party is no longer part of the key at all'
  );
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'blank-vs-populated party on an otherwise-identical row is NOT flagged');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'silently dropped from the preview, same as a byte-exact re-upload');
}

/* ===== 8. THE LIVE FIX, per explicit instruction, 2026-10-22: two DIFFERENT non-blank party values on
   an otherwise-identical row must NOT flag either -- this inverts the old "SAFETY" test that used to
   require exactly the opposite. A page-header party name can be misread once and read correctly every
   time after (confirmed live: "N 100 Jumbo Container"/"N 100 Jumbo Lid" 16.10.26, misread as party
   "रिटलाई" once, correctly "सिलाई" every time since), which meant two readings of the SAME physical
   row would disagree on party forever, with no way for that disagreement to ever resolve itself. party
   was dropped from DEDUP_FIELDS.production entirely rather than given another tolerance rule, since
   it was never what made a Production row the row it is -- date, item, and quantity are. A genuinely
   different second order sharing the same date/item/quantity is a real, if rare, edge case this no
   longer catches -- an explicit, accepted trade-off, not an oversight. ===== */
{
  const confirmed = [{
    id: 'old1', date: '3.10.26', party: 'Vijyant', description: 'Butter Bake 130g x30 Packet',
    customerHint: '', pieces: 4000, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  const differentParty = [{ id: 'new1', date: '3.10.26', party: 'Suresh', description: 'Butter Bake 130g x30 Packet', customerHint: '', pieces: 4000, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', differentParty);
  t.assertTrue(!flagged[0].flagged, 'LIVE FIX: two different non-blank party values no longer flag -- party is not part of a Production row\'s identity');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'LIVE FIX: silently dropped, same as any other party reading of the same date/item/quantity');
}

/* ===== 9b. THE REAL FAILURE MODE: a confirmed row loaded back from the register/Sheet has its numeric
   fields as STRINGS ("1490"), while a freshly-extracted row has them as real JS numbers (1490) -- test
   7 above used a number literal on both sides and could never have caught this. Confirmed live:
   2026-10-10, Production Register, "IT 500 Container" 8.10.26/1490 pieces, blank party vs "सिलाई" --
   this exact row stayed flagged live even after the party-tolerance fix shipped, because `1490 ===
   "1490"` is false under strict equality, so the pieces field alone was reported as a disagreement no
   matter what the party tolerance did. ===== */
{
  const confirmed = [{
    id: 'old1', date: '8.10.26', party: '', description: 'IT 500 Container',
    customerHint: '', pieces: '1490', dispatch: '0', stockConfirmed: true, confirmedCustomer: 'BINDAL STOCK 1.08.26',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '8.10.26', party: 'सिलाई', description: 'IT 500 Container', customerHint: '', pieces: 1490, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'LIVE BUG: string "1490" (register) vs number 1490 (fresh extraction) is recognized as the same value, not flagged');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'LIVE BUG: silently dropped from the preview once the type mismatch no longer masks the match');
}

/* ===== 9c. THE DESCRIPTION-NORMALIZATION ASYMMETRY: looseDateDescKey (decides "is this even a
   candidate match") already used the aggressive normalizeForCatalogMatch, but rowsMatchForDedup's own
   description comparison (decides "silently drop" vs "flag") only used mild applyAbbreviations --
   meaning a description differing ONLY in punctuation/spacing/pack-count notation between the confirmed
   row and a fresh re-extraction would be found as a candidate, then STILL get flagged on that formatting
   difference alone, every single re-upload, forever. Using the same normalizer in both places closes
   this. A genuine word-level difference ("(new)" present or absent) must still survive and still
   disagree -- tested alongside so this fix can't accidentally start merging real variants. ===== */
{
  const confirmed = [{
    id: 'old1', date: '7.10.26', party: '', description: 'T 50 32G*120PKT',
    customerHint: '', pieces: 3200, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'DIAMOND',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '7.10.26', description: 'T 50 32G * 120 PKT', pieces: 3200, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'ASYMMETRY FIX: punctuation/spacing-only description difference is not flagged');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'ASYMMETRY FIX: silently dropped as the same real row');
}
{
  // POLICY CHANGE, confirmed live 2026-10-13: this used to be the expected, deliberate behavior --
  // "(new)" present/absent was treated as a genuine word-level variant, never matched. Turned out wrong
  // for the actual live case: "T 50 32G * 120 PKT new" (confirmed) vs "T 50 32G * 120 PKT" (re-read) was
  // the SAME physical row -- the "new" was an illegible handwritten colour note, not a real distinguishing
  // "new" designation, and OCR simply didn't catch it on the second read. Same date + same exact piece
  // count is already a strong real-world fingerprint; once that agrees, a trailing-qualifier-only
  // difference in the description (one reading has it, the other doesn't) is now tolerated and silently
  // dropped, not flagged and not treated as a plain new item. See rowsMatchForDedup's prefix tolerance.
  const confirmed = [{
    id: 'old1', date: '7.10.26', party: '', description: 'T 50 32G*120PKT new',
    customerHint: '', pieces: 3200, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'DIAMOND',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '7.10.26', description: 'T 50 32G * 120 PKT', pieces: 3200, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'LIVE CASE (T50): trailing-qualifier-only difference, same date/pieces, is not flagged');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'LIVE CASE (T50): silently dropped as the same real row');
}
/* ===== 9d. THE SECOND LIVE INSTANCE of the same bug, different item: "Blacko 32g x120 (New)" vs
   "Blacko 32g x120" -- confirms the fix generalizes, not a one-off special case for "T 50". ===== */
{
  const confirmed = [{
    id: 'old1', date: '5.9.26', party: '', description: 'Blacko 32g x120 (New)',
    customerHint: '', pieces: 4240, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '5.9.26', description: 'Blacko 32g x120', pieces: 4240, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'LIVE CASE (Blacko): trailing-qualifier-only difference, same date/pieces, is not flagged');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'LIVE CASE (Blacko): silently dropped as the same real row');
}
/* ===== 9e. SAFETY: the prefix tolerance only kicks in once date+pieces already agree -- a genuinely
   different variant on a genuinely different day (the real BLACKO 32g x120-OLD vs (New) case, which
   really are two distinct catalog items) must never be silently merged just because one description is a
   prefix of the other. ===== */
{
  const confirmed = [{
    id: 'old1', date: '22.9.26', party: '', description: 'Blacko 32g x120-OLD',
    customerHint: '', pieces: 6440, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'DIAMOND',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  // Different day, different quantity -- a genuinely separate batch, happens to share the prefix-able base name.
  const newBatch = [{ id: 'new1', date: '5.9.26', description: 'Blacko 32g x120 (New)', pieces: 4240, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', newBatch);
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 1, 'SAFETY: a genuinely different day+quantity variant is never silently merged just by a description prefix relationship');
}
/* ===== 9f. THE SECOND REAL ROOT CAUSE, confirmed live 2026-10-13: a description abbreviated
   DIFFERENTLY between two reads (not just a trailing qualifier) -- "CREAM 60G*70PKT" (itself a shortened
   read, already confirmed) vs "Cream Burst 60g x70 Packet" (today's correct, fuller read). Same date,
   same exact piece count, but "cream60g" and "creamburst60gpacket" share no prefix relationship at all,
   so Fix B's tolerance can't catch this one -- it needs the separate word-overlap check
   (findNumericFingerprintMatch), and deliberately only ever FLAGS this, never silently drops it, since a
   bare word-overlap is a much weaker signal than an exact prefix match. ===== */
{
  const confirmed = [{
    id: 'old1', date: '22.9.26', party: '', description: 'CREAM 60G*70PKT',
    customerHint: '', pieces: 2944, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'DIAMOND',
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '22.9.26', description: 'Cream Burst 60g x70 Packet', pieces: 2944, dispatch: 0 }];
  const result = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(result[0].flagged, 'LIVE CASE (CREAM): same date/quantity, differently-abbreviated description, is flagged for review');
  t.assertTrue(/read differently/.test(result[0].flagReason), 'LIVE CASE (CREAM): flag reason explains it is a possible same-entry/different-reading case');
}
/* ===== 9g. SAFETY: the word-overlap flag requires BOTH date and exact quantity to already agree -- two
   unrelated items that happen to share one common word (e.g. both have "Butter" in the name) on
   different days/quantities must never be flagged just because of that shared word. ===== */
{
  const confirmed = [{
    id: 'old1', date: '2.9.26', party: '', description: 'Butter Bake 130g x30 Packet',
    customerHint: '', pieces: 7200, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'anmol stock 01.08.26',
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ production: confirmed });
  const unrelated = [{ id: 'new1', date: '2.9.26', description: 'Butter Plus 28g x140 Packet', pieces: 1940, dispatch: 0 }];
  const result = flagLikelyReReadDuplicates('production', unrelated);
  t.assertTrue(!result[0].flagged, 'SAFETY: a genuinely different item sharing one word is not flagged -- quantity disagrees, so the word-overlap check never even applies');
}

/* ===== 9. SAFETY: the blank-wildcard tolerance is scoped to production only, not customerDispatch --
   a Buyer name off a printed invoice is real financial data, not a sometimes-missed bracketed note, so
   a blank-vs-populated party there must still flag like every other field. ===== */
{
  const confirmed = [{
    id: 'old1', date: '9.10.26', invoice_no: 'INV-1', party: '', buyer_order_no: '',
    description: 'IT 500 Jumbo Container', quantity: 500, rate: 10, amount: 5000,
  }];
  const { flagLikelyReReadDuplicates } = loadDedupFunctions({ customerDispatch: confirmed });
  const reExtracted = [{ id: 'new1', date: '9.10.26', invoice_no: 'INV-1', party: 'Some Buyer Ltd', buyer_order_no: '', description: 'IT 500 Jumbo Container', quantity: 500, rate: 10, amount: 5000 }];
  const result = flagLikelyReReadDuplicates('customerDispatch', reExtracted);
  t.assertTrue(result[0].flagged, 'SAFETY: customerDispatch party is NOT given the blank-wildcard tolerance -- still flags');
}

/* ===== 10. THE LIVE BUG, confirmed directly against the real production register (main), 2026-10-21,
   and ITS OWN FIX, per explicit instruction, 2026-10-22: a page-header party name misread once
   ("रिटलाई") and correctly the next time ("सिलाई") -- same date, same item, same pieces -- used to get
   flagged every single re-extraction forever, because that disagreement could never self-resolve (the
   row genuinely IS the same physical row; only the party reading drifts). The fix removed party from
   DEDUP_FIELDS.production entirely: date, item, and pieces are what make a Production row the row it
   is, not a page-header party name. So this exact live case must now be silently matched and dropped,
   not flagged at all. ===== */
{
  const confirmed = [{
    id: 'old1', date: '16.10.26', party: 'रिटलाई', description: 'N 100 Jumbo Container',
    customerHint: '', pieces: 540, dispatch: 0, stockConfirmed: true, confirmedCustomer: 'BINDAL STOCK 1.08.26',
  }];
  const { flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates } = loadDedupFunctions({ production: confirmed });
  const reExtracted = [{ id: 'new1', date: '16.10.26', party: 'सिलाई', description: 'N 100 Jumbo Container', pieces: 540, dispatch: 0 }];
  const flagged = flagLikelyReReadDuplicates('production', reExtracted);
  t.assertTrue(!flagged[0].flagged, 'FIXED: the रिटलाई/सिलाई party disagreement no longer flags -- party is not part of a Production row\'s identity anymore');
  const dropped = dropAlreadyConfirmedDuplicates('production', flagged);
  t.assertEqual(dropped.length, 0, 'FIXED: silently dropped from the preview, never surfaced for a person to re-check, since date/item/pieces all agree');
}

module.exports = t.summary();
