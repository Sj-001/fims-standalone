// Generic "array of flat objects" <-> Google Sheets tab bridge. Each register (production,
// customerDispatch, etc.) becomes one tab, with a header row built from the union of keys across
// all rows and one spreadsheet row per array item — so opening the Sheet directly shows real,
// readable columns (Date, Description, Pieces, ...), not opaque JSON blobs. This is generic on
// purpose: it doesn't hardcode each register's column list, so adding a new field in the frontend
// doesn't require a matching backend change.
//
// Auth: uses a Google service account (see README) via a JWT client — no OAuth login flow, no
// browser consent screen, safe for a fully unattended backend.
//
// Write quota: Google's default is 60 write requests/minute PER USER, per project (300/minute per
// project overall) — https://developers.google.com/workspace/sheets/api/limits. Every function here
// is written to spend exactly ONE write call per save, no matter how large the data: writeTab and
// writeBlocksTab each do at most one values.update call (see the "clear via padding" note below).
// The other half of the fix — not calling these functions once per row edit in the first place — is
// on the frontend (debounced saves).

const { google } = require('googleapis');

let _sheetsClient = null;

function getCredentials() {
  // Accept either the raw JSON (GOOGLE_SERVICE_ACCOUNT_JSON) or a base64-encoded version
  // (GOOGLE_SERVICE_ACCOUNT_JSON_BASE64) — base64 is offered because some hosting UIs mangle
  // multi-line/quote-heavy env var values, and base64 sidesteps that entirely.
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  const b64 = process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64;
  if (b64) {
    try { return JSON.parse(Buffer.from(b64, 'base64').toString('utf8')); }
    catch (e) { throw new Error(`GOOGLE_SERVICE_ACCOUNT_JSON_BASE64 is not valid base64-encoded JSON: ${e.message}`); }
  }
  if (raw) {
    try { return JSON.parse(raw); }
    catch (e) { throw new Error(`GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON: ${e.message}`); }
  }
  throw new Error('Set either GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_JSON_BASE64.');
}

function getSheetsClient() {
  if (_sheetsClient) return _sheetsClient;
  const credentials = getCredentials();
  const auth = new google.auth.JWT({
    email: credentials.client_email,
    key: credentials.private_key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  _sheetsClient = google.sheets({ version: 'v4', auth });
  return _sheetsClient;
}

function getSpreadsheetId() {
  const id = process.env.GOOGLE_SHEET_ID;
  if (!id) throw new Error('GOOGLE_SHEET_ID is not set.');
  return id;
}

// The service account's own email — this is what every customer's Google Sheet needs to be shared
// with (as at least Viewer, Editor if it'll also be pushed to) before this app can read or write it.
// Surfaced via /api/service-account-email so the frontend can show it directly in the Customer Sheets
// tab, instead of making someone go dig it out of a Render env var or a downloaded JSON key file.
function getServiceAccountEmail() {
  try { return getCredentials().client_email || ''; } catch (e) { return ''; }
}

// Turns a raw Google API error into something a non-engineer can actually act on. By far the most
// common failure here — for BOTH import and push — is pasting a Sheet ID that simply hasn't been
// shared with the service account yet. Google's own error for that case is just "The caller does not
// have permission" with a 403, which gives no hint of what to actually do about it. This detects that
// specific case (and the "wrong ID" 404 case) and rewrites the message into the exact fix, including
// the real service account email pulled from this app's own credentials — not a placeholder.
function friendlyGoogleError(e) {
  const raw = (e && e.message) || 'unknown error';
  const code = e && (e.code || (e.response && e.response.status));
  const isPermission = code === 403 || /does not have permission|permission_denied/i.test(raw);
  if (isPermission) {
    const email = getServiceAccountEmail();
    return `That Google Sheet hasn't been shared with this app yet. Open the Sheet → Share → add ${email || "this app's service account (its email is in your GOOGLE_SERVICE_ACCOUNT_JSON's \"client_email\" field)"} as an Editor → try again.`;
  }
  const isNotFound = code === 404 || /requested entity was not found/i.test(raw);
  if (isNotFound) {
    return `No Google Sheet found with that ID. Double-check you copied the full ID from the Sheet's URL — the part between "/d/" and "/edit".`;
  }
  return raw;
}

// Sheet tab names have the same restrictions as Excel (no \/?*[] , 31 char max) — reuse the same
// sanitizing approach the artifact version used for .xlsx export, so register keys map predictably.
function sanitizeTabName(name) {
  let base = String(name || '').replace(/[\\/*?:[\]]/g, '').trim();
  if (!base) base = 'Sheet';
  return base.slice(0, 31);
}

// One read call that returns each tab's title AND its current allocated grid size
// (rowCount/columnCount). The grid size is what makes it possible to "clear" a tab's stale content
// in the SAME call as writing new data (see writeValuesClearingStale below), instead of a separate
// values.clear call — cutting every save from 2-3 write requests down to 1.
async function getSheetMeta(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))',
  });
  const byTitle = {};
  (meta.data.sheets || []).forEach(s => {
    byTitle[s.properties.title] = {
      sheetId: s.properties.sheetId,
      rowCount: (s.properties.gridProperties && s.properties.gridProperties.rowCount) || 1000,
      columnCount: (s.properties.gridProperties && s.properties.gridProperties.columnCount) || 26,
    };
  });
  return byTitle;
}

// Creates the tab if it doesn't exist yet. Returns its current {rowCount, columnCount} either way
// (a freshly created tab gets Google Sheets' default new-sheet grid, 1000x26, which this also
// returns so the caller doesn't need a second read to find out).
async function ensureTabExists(sheets, spreadsheetId, tabName) {
  const sheetMeta = await getSheetMeta(sheets, spreadsheetId);
  if (sheetMeta[tabName]) return sheetMeta[tabName];
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] },
  });
  return { rowCount: 1000, columnCount: 26 };
}

// Writes `values` (a 2D array) starting at A1, padding it with blank rows/columns up to the tab's
// current grid size first — so any leftover cells from a previous, longer save get overwritten with
// blanks in this SAME request, instead of needing a separate values.clear call beforehand. This is
// the one-write-per-save trick referenced at the top of this file.
async function writeValuesClearingStale(sheets, spreadsheetId, tabName, values, gridSize) {
  const colCount = Math.max(gridSize.columnCount, ...values.map(r => r.length), 1);
  const rowCount = Math.max(gridSize.rowCount, values.length);
  const padded = new Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    const src = values[i] || [];
    const row = new Array(colCount).fill('');
    for (let j = 0; j < src.length; j++) row[j] = src[j] === undefined || src[j] === null ? '' : src[j];
    padded[i] = row;
  }
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${tabName}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: padded },
  });
}

// Reads a tab and returns an array of plain objects, one per data row, keyed by the header row.
// Returns [] if the tab doesn't exist yet (nothing saved there yet) or is empty — this is a normal,
// expected state for a brand-new register, not an error.
async function readTab(tabKey) {
  const sheets = getSheetsClient();
  const spreadsheetId = getSpreadsheetId();
  const tabName = sanitizeTabName(tabKey);
  let values;
  try {
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range: tabName });
    values = resp.data.values || [];
  } catch (e) {
    // "Unable to parse range" (400) means the tab doesn't exist yet — treat as empty, not fatal.
    if (e.code === 400 || (e.errors && e.errors[0] && /Unable to parse range/i.test(e.errors[0].message || ''))) {
      return [];
    }
    throw e;
  }
  if (!values.length) return [];
  const [header, ...rows] = values;
  return rows.map(row => {
    const obj = {};
    header.forEach((key, i) => { obj[key] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

// Overwrites a tab's entire contents with `rows` — matches the artifact's original storage
// semantics exactly (window.storage.set always wrote the full current array, never an incremental
// diff), so the frontend's save logic needed no restructuring, just a different destination.
// Costs exactly ONE write call (plus one more, only the first time a given tab is created).
async function writeTab(tabKey, rows) {
  const sheets = getSheetsClient();
  const spreadsheetId = getSpreadsheetId();
  const tabName = sanitizeTabName(tabKey);
  const gridSize = await ensureTabExists(sheets, spreadsheetId, tabName);
  if (!Array.isArray(rows) || !rows.length) {
    await writeValuesClearingStale(sheets, spreadsheetId, tabName, [], gridSize);
    return;
  }
  // Header = union of keys across all rows, 'id' pinned first if present, rest in first-seen order.
  const keySet = new Set();
  rows.forEach(r => Object.keys(r || {}).forEach(k => keySet.add(k)));
  const keys = Array.from(keySet);
  const header = keys.includes('id') ? ['id', ...keys.filter(k => k !== 'id')] : keys;
  const values = [header, ...rows.map(r => header.map(k => (r[k] === undefined || r[k] === null) ? '' : String(r[k])))];
  await writeValuesClearingStale(sheets, spreadsheetId, tabName, values, gridSize);
}

// Writes a tab as a stack of titled tables — one block per item, each block being a title row, a
// blank row, a header row, then that item's data rows — matching the real customer stock files
// (BINDAL STOCK.xlsx, DIAMOND.xlsx, anmol stock dec 22.xlsx) this app was built to replace, and the
// same layout already used for the in-app "Download .xlsx" export. Used for the per-customer stock
// tabs (Bindal, Diamond, Anmol, ...), which are fully computed from the Production Register and
// Customer Dispatch Bills — this is a write-only mirror for convenience, never read back into the
// app's own state. Costs exactly ONE write call, same as writeTab.
async function writeBlocksTab(tabKey, blocks) {
  const sheets = getSheetsClient();
  const spreadsheetId = getSpreadsheetId();
  const tabName = sanitizeTabName(tabKey);
  const gridSize = await ensureTabExists(sheets, spreadsheetId, tabName);
  const values = [];
  (blocks || []).forEach((b, i) => {
    if (i > 0) values.push([]); // blank separator row between items
    values.push([b.title || '']);
    values.push([]);
    values.push(b.header || []);
    (b.rows || []).forEach(r => values.push(r));
  });
  await writeValuesClearingStale(sheets, spreadsheetId, tabName, values, gridSize);
}

// --- Customer Sheets: push computed data to a customer's OWN external spreadsheet (not this app's
// main GOOGLE_SHEET_ID) — reproduces the real BINDAL STOCK.xlsx / DIAMOND.xlsx / anmol stock dec 22
// files' actual layout, verified by opening those files directly: one tab per base item, with every
// variant of that item as its own small table sitting side by side (title row, header row, dated
// rows), separated by one blank spacer column — not stacked vertically like writeBlocksTab above.
// Same service-account auth works for any spreadsheet ID, as long as that spreadsheet has been
// shared with the service account's email (same one-time step as the main sheet).

// Builds one tab's full 2D grid from a list of variant blocks placed side by side.
// variants: [{ title, header: [...], rows: [[...], ...] }]
function buildSideBySideGrid(variants) {
  const blocks = (Array.isArray(variants) ? variants : []).map(v => ({
    title: v.title || '',
    header: (v.header && v.header.length) ? v.header : ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'],
    rows: Array.isArray(v.rows) ? v.rows : [],
  }));
  if (!blocks.length) return [[]];
  const widths = blocks.map(b => Math.max(b.header.length, ...b.rows.map(r => (r || []).length), 1));
  const totalWidth = widths.reduce((s, w) => s + w, 0) + (blocks.length - 1); // +1 blank spacer col between blocks
  const maxDataRows = Math.max(0, ...blocks.map(b => b.rows.length));
  const totalHeight = 2 + maxDataRows; // title row + header row + data rows
  const grid = Array.from({ length: totalHeight }, () => new Array(totalWidth).fill(''));
  let colOffset = 0;
  blocks.forEach((b, i) => {
    grid[0][colOffset] = b.title;
    b.header.forEach((h, ci) => { grid[1][colOffset + ci] = h; });
    b.rows.forEach((r, ri) => {
      (r || []).forEach((val, ci) => { grid[2 + ri][colOffset + ci] = (val === undefined || val === null) ? '' : val; });
    });
    colOffset += widths[i] + 1;
  });
  return grid;
}

const normalizeCellStr = (v) => (v === undefined || v === null) ? '' : String(v).trim();

// Google Sheets serial date (epoch 1899-12-30) -> "DD/MM/YYYY". Most real date cells in the customer
// sheets are typed as literal text ("20-Dec", "6.1.23", etc.), but verified directly against the real
// Anmol sheet: some blocks (e.g. Butter Bake 130g) store the date column as an actual Sheets date
// value instead. Read with UNFORMATTED_VALUE that comes back as a bare number like 44920 — without
// converting it, that number leaks straight into the review screen instead of a date.
const serialToDateStr = (n) => {
  const ms = Math.round((n - 25569) * 86400 * 1000); // 25569 = days between 1899-12-30 and 1970-01-01
  const d = new Date(ms);
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${d.getUTCFullYear()}`;
};
// Same as normalizeCellStr, but for a date-column cell specifically: converts a numeric date serial
// to a readable date string first, instead of stringifying the raw number.
const normalizeDateCell = (v) => (typeof v === 'number' && isFinite(v)) ? serialToDateStr(v) : normalizeCellStr(v);

// Parses an EXISTING tab's grid into its variant blocks, using the same "find the row containing a
// 'date' cell" heuristic used for import — so a merge-push can tell, per variant, exactly which dates
// already have a row (whether the app put it there on a previous push, or a person typed it in by
// hand) and exactly where that block's data currently ends, without touching any of it.
function parseExistingBlocks(grid) {
  const rows = Array.isArray(grid) ? grid : [];
  let headerRowIdx = -1;
  for (let r = 0; r < rows.length; r++) {
    if ((rows[r] || []).some(c => normalizeCellStr(c).toLowerCase() === 'date')) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) return { headerRowIdx: -1, blocks: [] };
  const headerRow = rows[headerRowIdx] || [];
  const dateColIdxs = [];
  headerRow.forEach((cell, c) => { if (normalizeCellStr(cell).toLowerCase() === 'date') dateColIdxs.push(c); });
  const blocks = dateColIdxs.map((startCol, i) => {
    const nextStart = dateColIdxs[i + 1] !== undefined ? dateColIdxs[i + 1] : Math.max(headerRow.length, startCol + 1);
    let width = 1;
    for (let c = startCol + 1; c < nextStart; c++) {
      if (normalizeCellStr(headerRow[c])) width = c - startCol + 1; else break;
    }
    let title = '';
    for (let back = 1; back <= 3 && headerRowIdx - back >= 0; back++) {
      const v = (rows[headerRowIdx - back] || [])[startCol];
      if (normalizeCellStr(v)) { title = normalizeCellStr(v); break; }
    }
    const existingDates = new Set();
    let nextRowIdx = headerRowIdx + 1;
    let lastRowValues = null;
    for (let r = headerRowIdx + 1; r < rows.length; r++) {
      const dateStr = normalizeDateCell((rows[r] || [])[startCol]);
      if (!dateStr) break; // first blank date cell ends this block's existing data
      existingDates.add(dateStr);
      nextRowIdx = r + 1;
      lastRowValues = [];
      for (let c = startCol; c < startCol + width; c++) {
        const raw = (rows[r] || [])[c];
        lastRowValues.push(c === startCol ? dateStr : raw);
      }
    }
    return { title, startCol, width, nextRowIdx, existingDates, lastRowValues };
  });
  return { headerRowIdx, blocks };
}

// Grows an EXISTING tab's grid instead of replacing it — this is the "merge, never touch existing
// rows, no duplicates" push mode. For each variant being pushed:
//   - if a block with that exact title already exists: every row already there (whether the app put
//     it there before, or someone typed it in directly) is copied through completely unchanged; only
//     rows for dates NOT already present get appended directly below the block's last existing row.
//   - if no block with that title exists yet: a brand new block is appended to the right of every
//     other block in the tab, aligned to the tab's own title/header row convention — identical to a
//     first-ever push for that variant.
// If the tab has no parseable blocks at all (first-ever push to a blank tab), every variant falls
// into the "no existing block" case, which reproduces buildSideBySideGrid's layout exactly.
//
// Balance carry-forward: a variant's incoming rows carry the app's own from-scratch opening/closing
// balance (computed purely from what's in this app's Production Register + Dispatch Bills). That's
// only correct if the app's register covers a customer's ENTIRE history — for a customer sheet with
// years of prior manual entries the app never saw, trusting that absolute number would silently
// disagree with the real balance already sitting in the sheet. So for existing blocks, only the
// PRODUCTION and DISPATCH quantities from the incoming rows are trusted as-is (they come straight
// from confirmed register entries); the opening/closing balance is recomputed to continue seamlessly
// from whatever this block's actual last existing closing balance is, read straight from the sheet.
// Converts a 0-indexed column number into its A1 letter(s) — 0->A, 25->Z, 26->AA, etc.
function colLetter(colIdx0) {
  let n = colIdx0 + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Wrapping a date string in a leading apostrophe is the standard Sheets convention for "store this
// exactly as typed, never auto-detect it as a number or date" — needed because the write below uses
// valueInputOption USER_ENTERED (required so the Opening/Closing formula strings actually become live
// formulas instead of literal text), and USER_ENTERED otherwise applies the same auto-parsing Sheets
// would to anything a person types. Every real customer file's date column is already plain text
// (verified: reading it back via UNFORMATTED_VALUE returns strings like "5.1.24", not date serial
// numbers, which is what a real auto-converted date cell would return) — this keeps new rows
// byte-for-byte consistent with that existing convention.
const forceTextValue = (v) => `'${v === undefined || v === null ? '' : v}`;

// Computes the MINIMAL set of targeted range writes needed to append new rows to each variant's
// block — this never reads-then-rewrites a single pre-existing cell, which matters for two reasons a
// full-tab overwrite would silently break: (1) a pre-existing row's Opening/Closing formulas would
// get flattened into their computed static value the instant this app pushed to that tab even once,
// permanently breaking the "edit an old row, everything below it recalculates" behavior the real
// files depend on; (2) any fixed cell reference elsewhere in the sheet (most importantly the summary
// tab's own formulas, e.g. ='N200'!E109) would keep working, since the cell it points at is never
// touched, only ever grown into.
// New rows link into the chain purely by REFERENCE — Opening = "=<previous row's Closing cell>",
// Closing = "=Opening+Production-Dispatch" — matching the exact convention already used throughout
// every real customer file, so the app's own rows are indistinguishable from ones a person typed by
// hand. A brand-new block (no existing block with this title) gets the same convention from its very
// first row, seeded with a literal 0 opening balance since there's no prior row to reference yet.
// Returns an array of { startRow0, startCol0, values } — 0-indexed, one entry per contiguous range
// that needs writing (title/header/data are separate ranges for a new block; just data for an
// existing one). The caller turns each into an A1 range and, separately, a highlight request — same
// structured data drives both, so there's no risk of the two ever disagreeing about what's "new".
function computeMergePatches(existingGrid, variants) {
  const { headerRowIdx, blocks } = parseExistingBlocks(existingGrid);
  const usedHeaderRowIdx = headerRowIdx === -1 ? 1 : headerRowIdx; // title row 0, header row 1 by default
  let rightmostCol = blocks.reduce((max, b) => Math.max(max, b.startCol + b.width), -1);
  const patches = [];

  (Array.isArray(variants) ? variants : []).forEach(v => {
    const key = normalizeTabKey(v.title);
    const match = blocks.find(b => normalizeTabKey(b.title) === key);
    const header = (v.header && v.header.length) ? v.header : ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'];
    const incomingRows = Array.isArray(v.rows) ? v.rows : [];

    if (match) {
      const newRows = incomingRows.filter(r => !match.existingDates.has(normalizeCellStr((r || [])[0])));
      if (!newRows.length) return;
      const openingCol = colLetter(match.startCol + 1);
      const prodCol = colLetter(match.startCol + 2);
      const dispCol = colLetter(match.startCol + 3);
      const closingCol = colLetter(match.startCol + match.width - 1);
      let prevRow1 = match.nextRowIdx; // 0-indexed row (match.nextRowIdx - 1) -> 1-indexed = match.nextRowIdx
      const values = newRows.map((r, i) => {
        const row = r || [];
        const thisRow1 = match.nextRowIdx + i + 1;
        const out = [
          forceTextValue(row[0] || ''),
          `=${closingCol}${prevRow1}`,
          Number(row[2]) || 0,
          Number(row[3]) || 0,
          `=${openingCol}${thisRow1}+${prodCol}${thisRow1}-${dispCol}${thisRow1}`,
        ];
        for (let ci = 5; ci < row.length; ci++) out.push(row[ci]);
        prevRow1 = thisRow1;
        return out;
      });
      patches.push({ startRow0: match.nextRowIdx, startCol0: match.startCol, values });
      // If a SECOND variant in this same push also resolves to this block (two catalog entries
      // mapped to the same real block, or two differently-worded register lines for the same item),
      // it must continue appending after THESE rows, not start over at the same spot and overwrite
      // them. Advancing the block's own bookkeeping here — same object `blocks.find` will return next
      // time — is what makes that safe.
      match.nextRowIdx += newRows.length;
      newRows.forEach(r => match.existingDates.add(normalizeCellStr((r || [])[0])));
    } else {
      const startCol = rightmostCol === -1 ? 0 : rightmostCol + 1;
      const width = Math.max(header.length, ...incomingRows.map(r => (r || []).length), 1);
      const openingCol = colLetter(startCol + 1);
      const prodCol = colLetter(startCol + 2);
      const dispCol = colLetter(startCol + 3);
      const closingCol = colLetter(startCol + width - 1);
      const dataStartRow0 = usedHeaderRowIdx + 1;
      patches.push({ startRow0: usedHeaderRowIdx - 1, startCol0: startCol, values: [[v.title || '']] });
      patches.push({ startRow0: usedHeaderRowIdx, startCol0: startCol, values: [header] });
      let prevRow1 = null;
      const values = incomingRows.map((r, i) => {
        const row = r || [];
        const thisRow1 = dataStartRow0 + i + 1;
        const out = [
          forceTextValue(row[0] || ''),
          i === 0 ? 0 : `=${closingCol}${prevRow1}`,
          Number(row[2]) || 0,
          Number(row[3]) || 0,
          `=${openingCol}${thisRow1}+${prodCol}${thisRow1}-${dispCol}${thisRow1}`,
        ];
        for (let ci = 5; ci < row.length; ci++) out.push(row[ci]);
        prevRow1 = thisRow1;
        return out;
      });
      if (values.length) patches.push({ startRow0: dataStartRow0, startCol0: startCol, values });
      rightmostCol = startCol + width;
      blocks.push({ title: v.title, startCol, width, nextRowIdx: dataStartRow0 + values.length, existingDates: new Set(), lastRowValues: null });
    }
  });
  return patches;
}

// Highlight color for the "what's new since last push" formatting below — a warm tint matching this
// app's own accent color, applied ONLY to the exact cells a patch just wrote. There's no more "reset
// everything to white first" step: since a push never touches a pre-existing cell anymore, there's
// nothing to fade — a cell keeps its highlight from whichever push actually last wrote it, and a
// person can always clear formatting by hand for a cell they've since confirmed/reviewed.
const HIGHLIGHT_COLOR = { red: 0.941, green: 0.886, blue: 0.769 };

// Turns the same patches that were just written into the matching repeatCell highlight requests —
// built from the identical structured data (startRow0/startCol0/values), so the highlighted range can
// never disagree with what was actually written.
function buildHighlightRequestsForPatches(sheetId, patches) {
  if (sheetId === undefined || sheetId === null) return [];
  return (patches || []).map(p => {
    const height = p.values.length;
    const width = Math.max(0, ...p.values.map(r => r.length));
    return {
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: p.startRow0, endRowIndex: p.startRow0 + height,
          startColumnIndex: p.startCol0, endColumnIndex: p.startCol0 + width,
        },
        cell: { userEnteredFormat: { backgroundColor: HIGHLIGHT_COLOR } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    };
  });
}

// Real customer spreadsheets routinely have tab names with trailing spaces or inconsistent casing
// ("hit & run ", "summary ", "SUMMARY", "T GEL ") — verified directly against BINDAL STOCK.xlsx,
// DIAMOND.xlsx, and anmol stock dec 22.xlsx, where MOST category tabs have a trailing space. Matching
// tab names with plain exact-string equality against a trimmed/sanitized proposed name misses those
// real tabs entirely, so the app used to think they didn't exist and created a second, near-identical
// tab (e.g. "hit & run" alongside the real "hit & run ") instead of writing into the one already
// there. This normalizes for comparison ONLY — the real tab's exact original name/spacing is always
// what gets written to and returned, never a normalized or re-cased version of it.
// This is used for BOTH levels of matching: the top-level category tab name, AND (inside
// parseExistingBlocks/computeMergePatches) a variant's title against an existing side-by-side block's
// title within that tab. The second use is why this also strips a trailing pack/carton-count
// annotation — a real block is titled just "Kaju Bake 65g", but the incoming variant's title is
// whatever the Production Register wrote that day, e.g. "Kaju Bake 65g x60". Without stripping it
// here too, the block match would fail even though the block genuinely exists, wrongly treating it as
// a brand-new item and creating a duplicate block right next to the real one on push.
//
// Verified directly against the real Anmol sheet (2026-08-11): the same pack-count annotation shows
// up in a growing list of different handwritten conventions across real block titles — "x60", "×60",
// "* 60pkt", "60 PKT", "60pkt", "x 60 PPKT" (note the double P — a typo, not a different unit) — and
// real block titles also vary in whitespace ("64 g" vs "64g") and unit spelling ("65GM" vs "65g").
// None of these differences represent a different physical item, so all of them have to be normalized
// away or matching silently fails and a real existing block gets reported as "new".
// Known word-level naming differences between what the Production Register calls an item and what
// the real Sheet calls it — confirmed directly by the customer, not guessed. Unlike the pack-count/
// spacing cleanup below, these are genuine different words (not formatting variants), so they can't be
// inferred automatically; each entry here was explicitly confirmed. E.g. the customer confirmed
// "Jeera Dhamal 70g"/"35g" in the register is the same item as the "jeera 70g"/"35g" block in the real
// sheet — the Sheet never uses the word "Dhamal" at all.
const TAB_KEY_ALIASES = [
  [/\bdhamal\b/gi, ''],
  // Bindal's real Sheet abbreviates "Jumbo" as a standalone "J" (e.g. "N200 J CONTAINER" vs. the
  // Production Register's "N200 Jumbo Container"). Confirmed directly against the real Sheet. Note:
  // Bindal also has "N 100 J XL Container" in the catalog, where that "J" means something else (it's
  // a distinct item from "N 100 Jumbo Container") — this alias will turn that spaced catalog entry
  // into "N 100 Jumbo XL Container" too, which won't match its real block ("N100J XL CONTAINER", no
  // space, so this alias's \b never fires on it). The catalog has a second, unspaced entry for the
  // same item ("N100J XL CONTAINER") that still matches correctly, so nothing actually breaks — but
  // flagging it here in case a Production Register entry ever logs it with the space.
  [/\bj\b/gi, 'jumbo'],
  // Bindal's E900 tab abbreviates "Container" as "CONT." (e.g. "E900 CONT." vs. the Production
  // Register's "E 900 Container"). Confirmed directly against the real Sheet.
  [/\bcont\.?(?=\s|$)/gi, 'container'],
  // Diamond's real Sheet only says "Burst" on one of the four Cream blocks ("CREAM BURST
  // 30G*140PKT") — the 60G*70PKT and 60G*60PKT blocks are just "CREAM ...", no "Burst" — while the
  // Production Register calls all Cream items "Cream Burst ...". Stripping the word (rather than
  // requiring it) lets both spellings converge on the same key. Confirmed directly against the real
  // Sheet, not guessed.
  [/\bburst\b/gi, ''],
];
const normalizeTabKey = (name) => {
  let s = String(name || '').trim().toLowerCase();
  TAB_KEY_ALIASES.forEach(([pattern, replacement]) => { s = s.replace(pattern, replacement); });
  // Trailing "x60ppkt" / "×60 pkt" / "*60" style pack-count marker: an explicit multiplier symbol
  // (x/×/*) is an unambiguous signal that whatever follows is a pack/carton count, never part of a
  // real product code — so strip the whole tail regardless of exactly how the unit word after the
  // digits is spelled (pkt/ppkt/pcs/nos/ctn/etc. all show up in the real sheets, with typos).
  s = s.replace(/[x×*]\s*\d+\s*[a-z]{0,8}\.?\s*$/i, '').trim();
  // Trailing "60 PKT" / "60pkt" style pack-count with NO multiplier symbol — ambiguous in general (a
  // bare number+word could be a real product code, e.g. "N200 Jumbo"), so only strip a known list of
  // packaging-unit words here rather than any word.
  s = s.replace(/\d+\s*(pkt|pkts|ppkt|ppkts|pcs|nos|ctn|ctns|box|boxes|bag|bags|unit|units)\.?\s*$/i, '').trim();
  // Unit spelling: "65GM" and "65g" mean the same thing.
  s = s.replace(/(\d)\s*gm\b/gi, '$1g');
  // Collapse all remaining whitespace so "64 g" and "64g" compare equal.
  s = s.replace(/\s+/g, '');
  return s;
};

// Resolves each tabPlan's proposed name against tabs that already exist in the target spreadsheet:
// if a tab already exists whose name matches case/whitespace-insensitively, the plan is rewritten to
// point at that tab's REAL name (so nothing new gets created and nothing pre-existing is duplicated).
// Only a genuinely new category (no match at all — a brand-new item, or a truly blank spreadsheet)
// gets a freshly created tab, using the sanitized proposed name.
function resolveTabPlansAgainstExisting(tabPlans, existingMeta) {
  const existingByKey = {};
  Object.keys(existingMeta).forEach(realName => { existingByKey[normalizeTabKey(realName)] = realName; });
  const resolved = tabPlans.map(p => {
    const key = normalizeTabKey(p.tabName);
    const realExistingName = existingByKey[key];
    return { ...p, tabName: realExistingName || sanitizeTabName(p.tabName) };
  });
  const missing = Array.from(new Set(resolved.filter(p => !existingMeta[p.tabName]).map(p => p.tabName)));
  return { resolved, missing };
}

// Pushes every item-group tab to an external spreadsheet in one go, appending only new rows and
// never touching a single pre-existing cell (see computeMergePatches above) — then highlights exactly
// those new cells. The "summary" tab is deliberately never part of this at all: every real customer
// file's summary tab is a handful of live formulas (one per item, pointing at that item's block's
// last row), which keep calculating correctly on their own for as long as this app only ever GROWS a
// block downward and never touches what's already there. Writing anything to that tab, even a
// same-looking refresh, would replace those formulas with dead static numbers.
// Sequence: read tab metadata once, create any missing tabs in a single batchUpdate (capturing their
// new numeric sheetId from the response), read every EXISTING tab's current values in one batchGet
// (this is what computeMergePatches needs to know which dates already have a row and where each
// block's data currently ends), write every tab's new cells in ONE values.batchUpdate covering every
// tab's patches at once, then highlight all the new cells in one final batchUpdate. That's 4 API
// calls total for the whole push, regardless of how many tabs or rows — comfortably inside Google's
// 60 writes/minute/user quota for an infrequent, manually-triggered action.
// itemGroups: [{ tabName, variants: [{ title, header, rows }] }]
async function pushCustomerSheet(spreadsheetId, itemGroups) {
  const { sheets, existingMeta, tabPlans, missing, previousValuesByTab } = await readPushState(spreadsheetId, itemGroups);
  if (missing.length) {
    const createResp = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: missing.map(title => ({ addSheet: { properties: { title } } })) },
    });
    (createResp.data.replies || []).forEach(reply => {
      if (reply.addSheet) {
        const p = reply.addSheet.properties;
        existingMeta[p.title] = { sheetId: p.sheetId, rowCount: 1000, columnCount: 26 };
      }
    });
  }
  const results = [];
  const formatRequests = [];
  const dataUpdates = [];
  const patchesByTab = {};
  for (const plan of tabPlans) {
    const previousGrid = previousValuesByTab[plan.tabName] || [];
    const patches = computeMergePatches(previousGrid, plan.variants);
    patchesByTab[plan.tabName] = patches;
    patches.forEach(p => {
      const startColL = colLetter(p.startCol0);
      const endColL = colLetter(p.startCol0 + Math.max(0, ...p.values.map(r => r.length)) - 1);
      const range = `${plan.tabName}!${startColL}${p.startRow0 + 1}:${endColL}${p.startRow0 + p.values.length}`;
      dataUpdates.push({ range, values: p.values });
    });
    results.push({ tab: plan.tabName, ok: true, newRows: patches.reduce((s, p) => s + p.values.length, 0) });
  }
  if (dataUpdates.length) {
    try {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: { valueInputOption: 'USER_ENTERED', data: dataUpdates },
      });
    } catch (e) {
      const msg = friendlyGoogleError(e);
      return tabPlans.map(plan => ({ tab: plan.tabName, ok: false, error: msg }));
    }
  }
  tabPlans.forEach(plan => {
    const sheetId = existingMeta[plan.tabName] && existingMeta[plan.tabName].sheetId;
    formatRequests.push(...buildHighlightRequestsForPatches(sheetId, patchesByTab[plan.tabName]));
  });
  if (formatRequests.length) {
    try {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: formatRequests } });
    } catch (e) {
      // Highlighting is cosmetic on top of a push that already succeeded — surface it as a per-tab
      // warning rather than failing results that already reported ok:true above.
      console.error('Customer sheet highlight formatting error:', e);
      results.push({ tab: '(highlighting)', ok: false, error: `Data pushed fine, but highlighting changed cells failed: ${friendlyGoogleError(e)}` });
    }
  }
  return results;
}

// Dry-run counterpart to pushCustomerSheet — reads the SAME "before" state and identifies the SAME
// new rows (reusing parseExistingBlocks so there's no risk of the preview ever disagreeing with what
// an actual push would do), but never writes anything and returns plain numbers instead of formula
// strings, since a human reviewing a diff wants to see "what will this balance become," not raw
// Sheets formula syntax. Used to populate the review-before-push screen: shows, per item, the last
// row already in the sheet (for continuity/context) next to the new row(s) about to be appended.
function previewCustomerSheet(previousValuesByTab, tabPlans, missing) {
  return tabPlans.map(plan => {
    const previousGrid = previousValuesByTab[plan.tabName] || [];
    const { blocks } = parseExistingBlocks(previousGrid);
    const variants = (plan.variants || []).map(v => {
      const key = normalizeTabKey(v.title);
      const match = blocks.find(b => normalizeTabKey(b.title) === key);
      const incomingRows = Array.isArray(v.rows) ? v.rows : [];
      const newRows = match ? incomingRows.filter(r => !match.existingDates.has(normalizeCellStr((r || [])[0]))) : incomingRows;
      const lastExisting = (match && match.lastRowValues)
        ? { date: match.lastRowValues[0], closing: Number(match.lastRowValues[match.width - 1]) || 0 }
        : null;
      let running = lastExisting ? lastExisting.closing : 0;
      const rows = newRows.map(r => {
        const row = r || [];
        const opening = running;
        const production = Number(row[2]) || 0;
        const dispatch = Number(row[3]) || 0;
        running = opening + production - dispatch;
        return { date: row[0] || '', opening, production, dispatch, closing: running };
      });
      // Same reasoning as computeMergePatches: if another variant in this same customer's payload
      // also resolves to this block, it needs to see these rows as already staged — both so its own
      // running balance continues from here instead of the real sheet's last row, and so it doesn't
      // re-offer the same dates as "new".
      if (match && rows.length) {
        newRows.forEach(r => match.existingDates.add(normalizeCellStr((r || [])[0])));
        const synthesized = new Array(match.width).fill(null);
        synthesized[0] = rows[rows.length - 1].date;
        synthesized[match.width - 1] = rows[rows.length - 1].closing;
        match.lastRowValues = synthesized;
      }
      return { title: v.title, isNewBlock: !match, lastExisting, rows };
    });
    return { tabName: plan.tabName, isNewTab: missing.includes(plan.tabName), variants };
  });
}

// Shared read phase for both the preview and the real push — resolves tab names against what
// already exists and reads every pre-existing tab's current values once. Kept as one function so the
// two code paths are guaranteed to be looking at identical "before" state.
async function readPushState(spreadsheetId, itemGroups) {
  const sheets = getSheetsClient();
  let tabPlans = (Array.isArray(itemGroups) ? itemGroups : []).map(g => ({
    tabName: g.tabName || 'Sheet',
    variants: g.variants || [],
  }));
  const existingMeta = await getSheetMeta(sheets, spreadsheetId);
  const { resolved, missing } = resolveTabPlansAgainstExisting(tabPlans, existingMeta);
  tabPlans = resolved;
  const preExistingTabNames = tabPlans.map(p => p.tabName).filter(t => existingMeta[t] && !missing.includes(t));
  const previousValuesByTab = {};
  if (preExistingTabNames.length) {
    const readResp = await sheets.spreadsheets.values.batchGet({
      spreadsheetId, ranges: preExistingTabNames, valueRenderOption: 'UNFORMATTED_VALUE',
    });
    (readResp.data.valueRanges || []).forEach((vr, i) => { previousValuesByTab[preExistingTabNames[i]] = vr.values || []; });
  }
  return { sheets, existingMeta, tabPlans, missing, previousValuesByTab, existingTabNames: Object.keys(existingMeta) };
}

// --- Customer Sheets: IMPORT direction — read an existing customer's own Google Sheet (any
// spreadsheet ID the user pastes in, not just ones this app created) and derive its item list, with
// which tab each item belongs to, so it can be added to the Product Catalog and Customer Mapping.
// This is the read-side counterpart to pushCustomerSheet above; it's only ever used to populate the
// catalog, never to overwrite anything in this app's own data automatically.
//
// The extraction heuristic was verified by hand against the three real customer files this app was
// built to replace (BINDAL STOCK.xlsx, DIAMOND.xlsx, anmol stock dec 22.xlsx — every one of their
// ~35 item tabs, side by side variant blocks included) before being written as code: within a tab,
// find the row containing a cell that reads exactly "date" (case-insensitive) — every real file uses
// that as the ledger header, however differently the rest of the tab is labeled ("balance" vs.
// "CLOSING BAL.", different starting rows/columns, etc.). Each column in that row that says "date"
// marks the start of one variant's block; that variant's title sits in the row(s) directly above
// (searched upward a few rows to tolerate an occasional blank spacer row).
function extractItemsFromTabGrid(tabName, grid) {
  let headerRowIdx = -1;
  for (let r = 0; r < grid.length; r++) {
    if ((grid[r] || []).some(c => String(c || '').trim().toLowerCase() === 'date')) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    // No recognizable ledger header in this tab — fall back to treating the whole tab as one item,
    // named after the tab itself, rather than silently skipping it.
    return [{ item: tabName, sheetGroup: tabName }];
  }
  const headerRow = grid[headerRowIdx] || [];
  const items = [];
  headerRow.forEach((cell, c) => {
    if (String(cell || '').trim().toLowerCase() !== 'date') return;
    let title = '';
    for (let back = 1; back <= 3 && headerRowIdx - back >= 0; back++) {
      const v = (grid[headerRowIdx - back] || [])[c];
      if (v !== undefined && v !== null && String(v).trim()) { title = String(v).trim(); break; }
    }
    if (title) items.push({ item: title, sheetGroup: tabName });
  });
  if (!items.length) return [{ item: tabName, sheetGroup: tabName }];
  return items;
}

// Reads every non-"summary" tab of the given spreadsheet and returns the full item list, plus the
// spreadsheet's own title (used as a suggested customer name — editable by the person importing it,
// same as the existing file-upload import flow already lets them edit the suggested name).
async function importCustomerSheet(spreadsheetId) {
  const sheets = getSheetsClient();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'properties.title,sheets.properties.title',
  });
  const spreadsheetTitle = (meta.data.properties && meta.data.properties.title) || '';
  const allTabTitles = (meta.data.sheets || []).map(s => s.properties.title);
  const dataTabTitles = allTabTitles.filter(t => !/summary/i.test(t));
  if (!dataTabTitles.length) {
    return { spreadsheetTitle, tabCount: allTabTitles.length, items: [] };
  }
  const resp = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges: dataTabTitles });
  const items = [];
  (resp.data.valueRanges || []).forEach((vr, idx) => {
    const tabName = dataTabTitles[idx];
    const grid = vr.values || [];
    extractItemsFromTabGrid(tabName, grid).forEach(it => items.push(it));
  });
  return { spreadsheetTitle, tabCount: allTabTitles.length, items };
}

// --- HTTP handlers ---
async function getServiceAccountEmailHandler(req, res) {
  res.json({ email: getServiceAccountEmail() });
}

async function importCustomerSheetHandler(req, res) {
  try {
    const { spreadsheetId } = req.body || {};
    if (!spreadsheetId || !String(spreadsheetId).trim()) {
      return res.status(400).json({ error: 'spreadsheetId is required.' });
    }
    const result = await importCustomerSheet(String(spreadsheetId).trim());
    res.json({ ok: true, ...result });
  } catch (e) {
    console.error('Customer sheet import error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

async function pushCustomerSheetHandler(req, res) {
  try {
    const { spreadsheetId, itemGroups } = req.body || {};
    if (!spreadsheetId || !String(spreadsheetId).trim()) {
      return res.status(400).json({ error: 'spreadsheetId is required.' });
    }
    if (!Array.isArray(itemGroups)) {
      return res.status(400).json({ error: 'itemGroups must be an array.' });
    }
    const results = await pushCustomerSheet(String(spreadsheetId).trim(), itemGroups);
    const failed = results.filter(r => !r.ok);
    res.status(failed.length ? 502 : 200).json({ ok: failed.length === 0, results });
  } catch (e) {
    console.error('Customer sheet push error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

// Dry-run version of the handler above — same request shape, but never writes anything. Backs the
// review-before-push screen: called every time the computed itemGroups for a customer change, so the
// screen always shows an accurate, up-to-date diff against what's REALLY in that customer's Sheet
// right now (not a locally-guessed one), without ever risking a write as a side effect of just
// looking at the page.
async function previewCustomerSheetHandler(req, res) {
  try {
    const { spreadsheetId, itemGroups } = req.body || {};
    if (!spreadsheetId || !String(spreadsheetId).trim()) {
      return res.status(400).json({ error: 'spreadsheetId is required.' });
    }
    if (!Array.isArray(itemGroups)) {
      return res.status(400).json({ error: 'itemGroups must be an array.' });
    }
    const { tabPlans, missing, previousValuesByTab, existingTabNames } = await readPushState(String(spreadsheetId).trim(), itemGroups);
    const tabs = previewCustomerSheet(previousValuesByTab, tabPlans, missing);
    res.json({ ok: true, tabs, existingTabNames });
  } catch (e) {
    console.error('Customer sheet preview error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

async function getTab(req, res) {
  try {
    const rows = await readTab(req.params.tab);
    res.json({ rows });
  } catch (e) {
    console.error(`Sheets read error [${req.params.tab}]:`, e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

async function putTab(req, res) {
  try {
    const rows = (req.body && Array.isArray(req.body.rows)) ? req.body.rows : [];
    await writeTab(req.params.tab, rows);
    res.json({ ok: true, count: rows.length });
  } catch (e) {
    console.error(`Sheets write error [${req.params.tab}]:`, e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

async function putBlocksTab(req, res) {
  try {
    const blocks = (req.body && Array.isArray(req.body.blocks)) ? req.body.blocks : [];
    await writeBlocksTab(req.params.tab, blocks);
    res.json({ ok: true, count: blocks.length });
  } catch (e) {
    console.error(`Sheets blocks-write error [${req.params.tab}]:`, e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

// TEMP DEBUG — added to track down why parseExistingBlocks/extractItemsFromTabGrid report "no date
// header found" on N200/N150/E900/T GEL/TIJORI HANDLE while structurally-identical tabs (E130/N100/
// IT500) parse fine. Returns exactly what sheets.spreadsheets.values.get hands back for one tab, raw,
// so the real API response can be inspected instead of guessed at. REMOVE after root cause is found.
async function debugRawTabHandler(req, res) {
  try {
    const { spreadsheetId, tabName } = req.body || {};
    if (!spreadsheetId || !tabName) return res.status(400).json({ error: 'spreadsheetId and tabName are required.' });
    const sheets = getSheetsClient();
    const meta = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties(sheetId,title,index,gridProperties,hidden)',
    });
    const sheetMetaEntry = (meta.data.sheets || []).find(s => s.properties.title === tabName);
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: tabName,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const values = resp.data.values || [];
    res.json({
      requestedRange: resp.data.range,
      sheetProperties: sheetMetaEntry ? sheetMetaEntry.properties : null,
      rowCount: values.length,
      rowLengths: values.map(r => r.length),
      first6Rows: values.slice(0, 6).map(r => r.map(c => ({ value: c, type: typeof c, jsonEscaped: JSON.stringify(c) }))),
    });
  } catch (e) {
    console.error('debugRawTabHandler error:', e);
    res.status(500).json({ error: String(e && e.message || e), stack: e && e.stack });
  }
}

module.exports = { readTab, writeTab, writeBlocksTab, getTab, putTab, putBlocksTab, pushCustomerSheetHandler, previewCustomerSheetHandler, importCustomerSheetHandler, getServiceAccountEmailHandler, debugRawTabHandler };