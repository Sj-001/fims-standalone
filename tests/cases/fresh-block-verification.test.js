// A real duplicate block was created ("TOOTI FRUTI 28G * 120 PKT", brand new, right alongside the real
// "...PKT (new)" block the Known Product Catalog's own stored block title should have matched) even
// after two earlier rounds of fixes to this exact class of bug (normalizeTabKey punctuation stripping;
// isCatalogEntryBlockReal + the first push-time safety net). Root cause, confirmed live: the synchronous
// safety check inside buildCustomerSheetPayloadFromRows compares against getRealBlocksForTab, which only
// knows about whatever a SEPARATE, debounced background effect (refreshReview) last happened to fetch
// into reviewByCustomer -- nothing guarantees that has fetched anything for a given tab by the time a
// push fires. When it hasn't, getRealBlocksForTab returns [], identical to "this tab genuinely has no
// blocks yet", and the push goes through unchecked. Same class of bug as the loadRegister data-loss
// incident: treating "haven't read this yet" as "confirmed empty."
//
// Fixed with verifyAgainstFreshBlocks: a second, synchronously-fresh real-vs-resolved-block check, run
// with an awaited fetch immediately before the actual push write, using data that cannot be stale
// because it was just fetched for this exact push. This file tests that function directly, mocking
// `fetch` and `normalizeForCatalogMatch`.
const path = require('path');
const { extractBlock, extractLine, buildFromSource } = require('../lib/extract');
const { makeRecorder } = require('../lib/assert');

const APP_JSX = path.join(__dirname, '..', '..', 'client', 'src', 'App.jsx');

const normalizeForCatalogMatch = (s) => String(s || '')
  .toLowerCase()
  .replace(/pkt\.?/g, '')
  .replace(/\s*[x×*]\s*\d+\s*$/i, '')
  .replace(/(\d)\s*gm\b/g, '$1g')
  .replace(/\bcont\.?\b/g, 'container')
  .replace(/[^a-z0-9]/g, '');

function loadFn(sheetIdsByCustomer, fetchImpl) {
  const getCustomerSheetIdLine = extractLine(APP_JSX, 'const getCustomerSheetId = (customer) =>');
  // endExclusive + the next real declaration as the end anchor — not `(l) => l.trim() === '};'` — since
  // this function now has its own inner helper (failClosed) that closes with its own `};` first; a bare
  // trimmed-'};'  anchor would match that one and truncate the real function partway through.
  const verifyBlock = extractBlock(
    APP_JSX,
    'const verifyAgainstFreshBlocks = async (customer, itemGroups, exemptBlockTitles) => {',
    'const pushInFlightRef = useRef(new Set());',
    { endExclusive: true },
  );
  const customerSheetIds = Object.entries(sheetIdsByCustomer).map(([customer, sheetId]) => ({ customer, sheetId }));
  return buildFromSource(
    [getCustomerSheetIdLine, verifyBlock],
    'verifyAgainstFreshBlocks',
    ['customerSheetIds', 'normalizeForCatalogMatch', 'fetch'],
    [customerSheetIds, normalizeForCatalogMatch, fetchImpl],
  );
}

const t = makeRecorder('fresh-block-verification');

module.exports = (async () => {
  /* ===== 1. THE EXACT LIVE CASE: TOOTI FRUTI resolves to a block that doesn't exist, caught by a fresh
     check even though the synchronous pass (not under test here) would have missed it. ===== */
  {
    const fetchImpl = async () => ({
      ok: true,
      json: async () => ({
        ok: true,
        tabs: [{ existingBlockTitles: ['TOOTI FRUTI 28G * 120 PKT (new)'] }],
      }),
    });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'TOOTI FRUTI', variants: [{ title: 'TOOTI FRUTI 28G * 120 PKT' }] }];
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, new Set());
    t.assertEqual(safeItemGroups.length, 0, 'LIVE CASE: the mismatched block is pulled out of the push entirely');
    t.assertEqual(newlyUnmatched, ['TOOTI FRUTI 28G * 120 PKT'], 'LIVE CASE: reported as unmatched instead of silently creating a block');
  }

  /* ===== 2. POLICY CHANGE: a tab with zero real blocks yet is NO LONGER automatically fine -- per
     explicit instruction, a new block is only ever created when a person deliberately assigns it, full
     stop, no "nothing to be ambiguous with yet" exception (that exact exception is what let the TOOTI
     FRUTI case in test 1 through the FIRST version of this fix). ===== */
  {
    const fetchImpl = async () => ({ ok: true, json: async () => ({ ok: true, tabs: [{ existingBlockTitles: [] }] }) });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'NEW TAB', variants: [{ title: 'Genuinely New Item' }] }];
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, new Set());
    t.assertEqual(safeItemGroups.length, 0, 'POLICY CHANGE: an empty tab no longer auto-proceeds, pulled out of the push');
    t.assertEqual(newlyUnmatched, ['Genuinely New Item'], 'POLICY CHANGE: reported as unmatched, waiting for a deliberate assignment');
  }

  /* ===== 2b. ...but an item explicitly routed in this same action (routingOverrides / exemptBlockTitles)
     still proceeds even into a genuinely empty tab -- that IS the deliberate assignment the policy asks
     for, not an automatic fallback. ===== */
  {
    const fetchImpl = async () => ({ ok: true, json: async () => ({ ok: true, tabs: [{ existingBlockTitles: [] }] }) });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'NEW TAB', variants: [{ title: 'Genuinely New Item', blockTitleOverride: 'Genuinely New Item' }] }];
    const exempt = new Set([normalizeForCatalogMatch('Genuinely New Item')]);
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, exempt);
    t.assertEqual(safeItemGroups[0].variants.length, 1, 'an explicit pick still creates the first block for a genuinely empty tab');
    t.assertEqual(newlyUnmatched.length, 0, 'not reported as unmatched -- this was a deliberate assignment');
  }

  /* ===== 3. A correctly-matching block proceeds normally ===== */
  {
    const fetchImpl = async () => ({
      ok: true,
      json: async () => ({ ok: true, tabs: [{ existingBlockTitles: ['T 50 32G * 120 PKT'] }] }),
    });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'T 50', variants: [{ title: 'T 50 32G * 120 PKT' }] }];
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, new Set());
    t.assertEqual(safeItemGroups[0].variants.length, 1, 'a correctly-matching block is never pulled out');
    t.assertEqual(newlyUnmatched.length, 0, 'not reported as unmatched');
  }

  /* ===== 4. SAFETY: a variant explicitly routed in this same action (routingOverrides) is exempt, even
     if the fresh check would otherwise flag it -- a deliberate just-made choice is never second-guessed. ===== */
  {
    const fetchImpl = async () => ({
      ok: true,
      json: async () => ({ ok: true, tabs: [{ existingBlockTitles: ['Something Else Entirely'] }] }),
    });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'T 50', variants: [{ title: 'Genuinely New Variant', blockTitleOverride: 'Genuinely New Variant' }] }];
    const exempt = new Set([normalizeForCatalogMatch('Genuinely New Variant')]);
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, exempt);
    t.assertEqual(safeItemGroups[0].variants.length, 1, 'SAFETY: an explicit routingOverride pick is trusted, not blocked');
    t.assertEqual(newlyUnmatched.length, 0, 'SAFETY: not reported as unmatched');
  }

  /* ===== 5. Multiple variants in the same tab: only the genuinely mismatched one is pulled out ===== */
  {
    const fetchImpl = async () => ({
      ok: true,
      json: async () => ({ ok: true, tabs: [{ existingBlockTitles: ['Butter Bake 130g x30 Packet'] }] }),
    });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{
      tabName: 'BUTTER',
      variants: [
        { title: 'Butter Bake 130g x30 Packet' },
        { title: 'Butter Bake 999g x30 Packet (fake)' },
      ],
    }];
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, new Set());
    t.assertEqual(safeItemGroups[0].variants.length, 1, 'the matching variant is kept');
    t.assertEqual(safeItemGroups[0].variants[0].title, 'Butter Bake 130g x30 Packet', 'the correct one is kept');
    t.assertEqual(newlyUnmatched, ['Butter Bake 999g x30 Packet (fake)'], 'only the mismatched one is reported unmatched');
  }

  /* ===== 6. FAILS CLOSED, not open: a real, live incident (2026-11-21) -- a Google Sheets API quota
     error struck this exact fetch mid-push. The OLD behavior here was "fall back to treating the input
     as safe (the synchronous check's own prior decision stands)" -- which is exactly how a block the
     synchronous pass had wrongly let through (its own cache stale or wrong) sailed past the ONE check
     specifically meant to catch that, and got created for real. A thrown exception now means every
     non-exempt variant is treated as UNVERIFIED and pulled out, not "probably fine" -- it stays pending
     for the person to retry once the real Sheet can actually be checked again. ===== */
  {
    const fetchImpl = async () => { throw new Error('simulated network failure'); };
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'T 50', variants: [{ title: 'T 50 32G * 120 PKT' }] }];
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, new Set());
    t.assertEqual(safeItemGroups.length, 0, 'FAILS CLOSED: a verification failure pulls the unverified variant out of the push entirely');
    t.assertEqual(newlyUnmatched, ['T 50 32G * 120 PKT'], 'FAILS CLOSED: reported as unmatched (unverified), not silently trusted');
  }

  /* ===== 6b. FAILS CLOSED, same non-ok-response case (a 502 from a quota-exceeded error, exactly what
     the real incident's HTTP response looked like) -- not just a thrown exception. ===== */
  {
    const fetchImpl = async () => ({ ok: false, status: 502, json: async () => ({ error: 'Quota exceeded for quota metric...' }) });
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'VAAYU', variants: [{ title: 'VAAYU 250 ML' }] }];
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, new Set());
    t.assertEqual(safeItemGroups.length, 0, 'FAILS CLOSED (502): a non-ok response pulls the unverified variant out, same as a thrown exception');
    t.assertEqual(newlyUnmatched, ['VAAYU 250 ML'], 'FAILS CLOSED (502): reported as unmatched, exactly the real incident');
  }

  /* ===== 6c. ...but a variant explicitly picked THIS session (exemptBlockTitles) still goes through even
     when the fresh check itself fails -- that exemption never depended on this fetch succeeding, since a
     person just made that exact choice moments ago, independent of whether the Sheet could be re-read. ===== */
  {
    const fetchImpl = async () => { throw new Error('simulated network failure'); };
    const verify = loadFn({ DIAMOND: 'sheet-id-123' }, fetchImpl);
    const itemGroups = [{ tabName: 'T 50', variants: [{ title: 'Freshly Picked Item', blockTitleOverride: 'Freshly Picked Item' }] }];
    const exempt = new Set([normalizeForCatalogMatch('Freshly Picked Item')]);
    const { safeItemGroups, newlyUnmatched } = await verify('DIAMOND', itemGroups, exempt);
    t.assertEqual(safeItemGroups[0].variants.length, 1, 'FAILS CLOSED, EXEMPT: a this-session deliberate pick still proceeds even when the fetch itself fails');
    t.assertEqual(newlyUnmatched.length, 0, 'FAILS CLOSED, EXEMPT: not reported as unmatched');
  }

  /* ===== 7. No Sheet ID for this customer: nothing to verify against, falls back cleanly ===== */
  {
    const verify = loadFn({}, async () => { throw new Error('should not be called'); });
    const itemGroups = [{ tabName: 'T 50', variants: [{ title: 'T 50 32G * 120 PKT' }] }];
    const { safeItemGroups } = await verify('NO SHEET CUSTOMER', itemGroups, new Set());
    t.assertEqual(safeItemGroups, itemGroups, 'no sheet id means nothing to verify -- falls back cleanly, no fetch attempted');
  }

  return t.summary();
})();
