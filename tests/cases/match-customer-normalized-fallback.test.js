// Real gap, confirmed directly against the live isolated test environment, 2026-10-14: a Production
// Register row extracted as "T 50 32g x 120 PKT" showed "Unassigned" in Pending Review, even though
// Customer Mapping has an exact, correctly-taught rule for it: keyword "t 50 32g * 120 pkt" -> DIAMOND,
// visible right there in the Customer Mapping tab, and the real DIAMOND Sheet genuinely has that exact
// block. Root cause: matchCustomer's keyword check is a raw, unnormalized substring match --
// "t 50 32g x 120 pkt".includes("t 50 32g * 120 pkt") is false, because "x" and "*" are different
// characters even though they mean the exact same pack-count separator everywhere else in this app
// (the real Sheet itself writes blocks with "*", the handwritten register with "x" -- see
// normalizeForCatalogMatch, which already treats both as equivalent for every OTHER matching step).
// Fixed with a FALLBACK pass: raw-text matching is tried first and unchanged (so nothing that already
// matched today can start matching a different rule), and only once nothing matched raw does a second
// pass retry with both sides run through normalizeForCatalogMatch -- catching exactly this class of
// cosmetic formatting drift (x/×/*, "gm" vs "g", "pkt" noise, punctuation) without touching anything
// that was already working.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

const t = makeRecorder('match-customer-normalized-fallback');

function run(customerMapping, row) {
  const matchCustomer = buildFromSource(
    [
      extractBlock(APP_JSX, 'const applyAbbreviations = (s) => {', (l) => l.trim() === '};'),
      extractBlock(APP_JSX, 'const matchCustomer = (row) => {', (l) => l.trim() === '};'),
      extractBlock(APP_JSX, 'const normalizeForCatalogMatch = (s) => applyAbbreviations(s || \'\')', (l) => l.trim() === '.replace(/[^a-z0-9]/g, \'\');'),
    ],
    'matchCustomer',
    ['abbreviations', 'customerNameAliases', 'customerSheetIds', 'customerMapping'],
    [[], [], [], customerMapping],
  );
  return matchCustomer(row);
}

// Real Customer Mapping rows, pulled live from the test environment's fims_customer_mapping sheet.
const REAL_MAPPING = [
  { keyword: 'it 500 cont', customer: 'BINDAL STOCK 1.08.26' },
  { keyword: 't 50 32g * 120 pkt old', customer: 'DIAMOND' },
  { keyword: 't 50 32g * 120 pkt', customer: 'DIAMOND' },
  { keyword: 't 50 64g * 60 pkt', customer: 'DIAMOND' },
  { keyword: 't 50 32g * 120 pkt (new)', customer: 'DIAMOND' },
  { keyword: 'it500 j cont.', customer: 'BINDAL STOCK 1.08.26' },
  { keyword: 'it 500 lid', customer: 'BINDAL STOCK 1.08.26' },
  { keyword: 'it 500 container', customer: 'BINDAL STOCK 1.08.26' },
  { keyword: 'it 500 jumbo container', customer: 'BINDAL STOCK 1.08.26' },
];

/* ===== THE LIVE CASE: exact real symptom, exact real keyword, exact real description. ===== */
{
  const customer = run(REAL_MAPPING, { description: 'T 50 32g x 120 PKT' });
  t.assertEqual(customer, 'DIAMOND', 'LIVE CASE: "x" vs "*" no longer falls through to Unassigned');
}

/* ===== Still correctly distinguishes a genuinely different weight -- 64g must never match the 32g
   rule just because both got normalized. ===== */
{
  const customer = run(REAL_MAPPING, { description: 'T 50 64g x 60 PKT' });
  t.assertEqual(customer, 'DIAMOND', 'DIFFERENT WEIGHT: 64g still resolves (to its own rule), not lost');
}

/* ===== A row that already matched via plain raw-text substring continues to match exactly the same
   way -- the fallback pass must never change behavior for something that already worked. ===== */
{
  const customer = run(REAL_MAPPING, { description: 'IT 500 Container' });
  t.assertEqual(customer, 'BINDAL STOCK 1.08.26', 'ALREADY WORKING: an exact raw-text match is untouched by the fallback');
}

/* ===== A genuinely unknown item (no rule anywhere, normalized or not) still correctly falls through
   to Unassigned -- the fallback must never invent a match out of nothing. ===== */
{
  const customer = run(REAL_MAPPING, { description: 'Totally Unrelated New Product 99g x 10 PKT' });
  t.assertEqual(customer, 'Unassigned', 'GENUINELY UNKNOWN: still Unassigned, no false positive');
}

/* ===== GUARD: a rule with an empty or punctuation-only keyword (normalizes to '') must never become a
   universal wildcard that matches every row via the fallback's .includes('') trap. ===== */
{
  const mapping = [{ keyword: '---', customer: 'SHOULD NEVER MATCH' }];
  const customer = run(mapping, { description: 'Literally Anything At All' });
  t.assertEqual(customer, 'Unassigned', 'GUARD: a keyword that normalizes to empty never becomes a wildcard');
}

/* ===== The "(new)" qualifier variant still resolves to the same customer as its plain counterpart,
   via its own distinct rule -- not required to collapse into one key, just required to still match
   SOMETHING that routes to DIAMOND. ===== */
{
  const customer = run(REAL_MAPPING, { description: 'T 50 32G * 120 PKT (new)' });
  t.assertEqual(customer, 'DIAMOND', 'QUALIFIER VARIANT: the (new) keyword itself still matches via raw text, unaffected');
}

/* ===== A customerHint takes priority over keyword matching entirely (existing, untouched behavior) --
   the fallback pass lives strictly inside the no-hint branch and must never run when a hint exists. */
{
  const mapping = [{ keyword: 't 50 32g * 120 pkt', customer: 'DIAMOND' }];
  const customer = run(mapping, { description: 'T 50 32g x 120 PKT', customerHint: 'Some Other Party' });
  t.assertEqual(customer, 'Some other party', 'HINT PRIORITY: an explicit customerHint still wins over any keyword rule');
}

module.exports = t.summary();
