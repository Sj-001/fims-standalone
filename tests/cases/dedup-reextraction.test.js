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
  const looseDateDescKeyBlock = extractLine(APP_JSX, 'const looseDateDescKey = (row) =>');
  const structuredLooseFieldsBlock = extractBlock(APP_JSX, 'const STRUCTURED_LOOSE_KEY_FIELDS = {', (l) => l.trim() === '};');
  const structuredLooseKeyBlock = extractBlock(
    APP_JSX,
    'const structuredLooseKey = (registerKey, row) =>',
    (l) => l.trim().endsWith(".join('||');"),
  );
  const flagLikelyReReadDuplicatesBlock = extractBlock(
    APP_JSX,
    'const flagLikelyReReadDuplicates = (registerKey, rows, extraExisting = []) => {',
    (l) => l.trim() === '};'
  );
  const dropAlreadyConfirmedDuplicatesBlock = extractBlock(
    APP_JSX,
    'const dropAlreadyConfirmedDuplicates = (registerKey, rows, extraExisting = []) => {',
    (l) => l.trim() === '};'
  );

  const blocks = [
    dedupFieldsBlock, rowDedupKeyBlock, applyAbbreviationsBlock, normalizeForCatalogMatchBlock,
    dedupKeyForRowBlock, looseDateDescKeyBlock, structuredLooseFieldsBlock, structuredLooseKeyBlock,
  ];

  const dedupKeyForRow = buildFromSource(blocks, 'dedupKeyForRow', ['abbreviations'], [abbreviations]);
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
  return { dedupKeyForRow, looseDateDescKey, flagLikelyReReadDuplicates, dropAlreadyConfirmedDuplicates };
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

module.exports = t.summary();
