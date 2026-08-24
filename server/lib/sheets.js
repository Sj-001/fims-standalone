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

// Converts a 1-based column number to its A1 letter(s) (1 -> A, 26 -> Z, 27 -> AA, ...).
function columnNumberToLetters(n) {
  let s = '';
  let num = n;
  while (num > 0) {
    const rem = (num - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s || 'A';
}

// Builds an UNAMBIGUOUS whole-sheet A1 range for a tab, e.g. "'N200'!A1:L1000". Passing just a bare
// tab name as a values.get/batchGet range (what this file used to do everywhere) is dangerous: if the
// name happens to look like a cell reference — letters immediately followed by digits, no space, e.g.
// "N200", "E900", "N150" — the Sheets API silently reads it as THAT CELL on the spreadsheet's default
// sheet instead of "the whole tab with this name," and returns 0 rows with no error at all. Confirmed
// directly against the live API (a bare "N200" request came back tagged "'summary '!N200" — i.e. cell
// N200 on the first/default tab). Quoting the sheet name AND appending an explicit A1 range removes
// the ambiguity entirely, verified against the same tabs that were previously misread. Every place
// that requests a tab by name for a batchGet/get call MUST go through this, not a bare title string.
function wholeSheetRangeA1(title, meta) {
  const rowCount = (meta && meta.rowCount) || 1000;
  const columnCount = (meta && meta.columnCount) || 26;
  const lastCol = columnNumberToLetters(columnCount);
  const safeTitle = String(title).replace(/'/g, "''");
  return `'${safeTitle}'!A1:${lastCol}${rowCount}`;
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

const MONTH_ABBR_TO_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
// Produces a canonical "YYYY-MM-DD" key for DUPLICATE-ROW DETECTION ONLY — never for what actually
// gets written or displayed (that stays exactly as normalizeDateCell/normalizeCellStr already produce
// it). The same real date shows up written in several different formats depending on where a row came
// from — dot "18.7.26" (this app's own pushes, and most of what's already hand-typed in these Sheets),
// slash "18/07/2026" (Production Register OCR), dash+month "18-Jul-26" (Customer Dispatch Bill
// extraction), or a numeric Sheets date serial. Comparing those as raw trimmed strings (what this used
// to do) means they never match each other, so the SAME real transaction gets pushed a second time as
// a "new" row instead of being recognized as already in the sheet — confirmed directly: Bindal's T GEL
// tab ended up with both "18.07.26" and "18-Jul-26" as separate rows for what was really one dispatch.
// Every date comparison used to decide "is this row already here" MUST go through this, not through
// normalizeDateCell/normalizeCellStr directly.
function canonicalDateKey(v) {
  if (typeof v === 'number' && isFinite(v)) {
    const ms = Math.round((v - 25569) * 86400 * 1000); // 25569 = days between 1899-12-30 and 1970-01-01
    const d = new Date(ms);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  const s = normalizeCellStr(v);
  if (!s) return '';
  const toKey = (d, mo, y) => {
    const yyyy = String(y).length <= 2 ? (Number(y) + 2000) : Number(y);
    return `${yyyy}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  };
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (m) return toKey(m[1], m[2], m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) return toKey(m[1], m[2], m[3]);
  m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (m) {
    const mo = MONTH_ABBR_TO_NUM[m[2].toLowerCase()];
    if (mo) return toKey(m[1], mo, m[3]);
  }
  // Unrecognized format — fall back to the literal trimmed string, same behavior as before this fix.
  return s.toLowerCase();
}

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
    // Per-date values actually sitting in the real sheet — lets a caller tell "this date already has a
    // row AND it's the same numbers" apart from "this date already has a row, but it disagrees with
    // what we're about to send." A date-only match used to be treated as good enough, which is exactly
    // how two real dispatch bills that happened to share a date with an already-present row got
    // silently dropped forever, never pushed and never flagged. Opening/closing are captured for every
    // row (not just the last one) so a brand-new row that needs to land BETWEEN two existing dates can
    // still chain its Opening off the real row directly before it, and so a dispatch-only entry whose
    // date already exists can show the sheet's own real opening/production instead of a guessed one.
    // Left unset for a date when the block is too narrow (width < 4) to even have a dispatch column —
    // that's "can't verify," not "matches," and is handled as such by classifyIncomingRows.
    const existingValuesByDate = new Map();
    // Every real row's {rowIdx, dateKey}, in physical top-to-bottom order — the reference list a new
    // row's date gets checked against to find exactly where it belongs chronologically, instead of
    // always landing after whatever the last physical row happens to be.
    const existingRowsOrdered = [];
    let nextRowIdx = headerRowIdx + 1;
    let lastRowValues = null;
    for (let r = headerRowIdx + 1; r < rows.length; r++) {
      const dateStr = normalizeDateCell((rows[r] || [])[startCol]);
      if (!dateStr) break; // first blank date cell ends this block's existing data
      const dateKey = canonicalDateKey((rows[r] || [])[startCol]);
      existingDates.add(dateKey);
      existingRowsOrdered.push({ rowIdx: r, dateKey });
      if (width >= 4) {
        existingValuesByDate.set(dateKey, {
          production: Number((rows[r] || [])[startCol + 2]) || 0,
          dispatch: Number((rows[r] || [])[startCol + 3]) || 0,
          opening: Number((rows[r] || [])[startCol + 1]) || 0,
          closing: Number((rows[r] || [])[startCol + width - 1]) || 0,
          rowIdx: r,
        });
      }
      nextRowIdx = r + 1;
      lastRowValues = [];
      for (let c = startCol; c < startCol + width; c++) {
        const raw = (rows[r] || [])[c];
        lastRowValues.push(c === startCol ? dateStr : raw);
      }
    }
    return { title, startCol, width, nextRowIdx, existingDates, existingValuesByDate, existingRowsOrdered, lastRowValues };
  });
  return { headerRowIdx, blocks };
}

// Shared by both the dry-run preview and the real push, so the two can never disagree about which
// incoming rows are genuinely new, which are correctly already-there, which have a real gap worth
// filling, and which are a MISMATCH — a date that already has a row in the real sheet, but with
// different numbers than what we're about to send. A mismatch is never pushed (an existing NON-BLANK
// cell is never touched) and never silently treated as "already handled" either. Every single incoming
// row gets a status back — nothing is ever dropped from the result — so the UI can show the real entry
// for every row, always, and just flag the ones that need a second look instead of hiding them behind
// a summary.
//
// Production and dispatch are judged by the SAME rule now, symmetrically — both are filled in at
// different times in real life (production the same day, dispatch whenever the truck actually leaves,
// sometimes days later), so a column that's currently 0 or blank in the real sheet is never treated as
// a disagreement, for either column: it just means "not entered yet," not a conflict. Only a column
// that ALREADY holds a real, non-zero value that disagrees with what we have is a genuine mismatch —
// flagged, never auto-touched. A column sitting at 0/blank with a real incoming value is "fillable":
// safe to write into later, without ever overwriting anything that was actually there.
//
// A row's own production/dispatch cell can be `null` — not 0 — meaning this specific upload has NO
// opinion on that column at all (a dispatch bill has no production figure on it, full stop; a
// production-only entry has no dispatch figure). That's different from a genuine, confirmed 0, and
// must never be compared against the real sheet as if it were one — a pure dispatch upload disagreeing
// with the sheet's real production number would otherwise get flagged as a "mismatch" for a column it
// never actually claimed anything about.
function classifyIncomingRows(match, incomingRows) {
  return (Array.isArray(incomingRows) ? incomingRows : []).map(r => {
    const row = r || [];
    const dateKey = canonicalDateKey(row[0]);
    if (!match || !match.existingDates.has(dateKey)) return { row, status: 'new' };
    const hasProduction = row[2] !== null && row[2] !== undefined && row[2] !== '';
    const hasDispatch = row[3] !== null && row[3] !== undefined && row[3] !== '';
    const expected = { production: Number(row[2]) || 0, dispatch: Number(row[3]) || 0 };
    const existingVals = match.existingValuesByDate.get(dateKey);
    if (!existingVals) return { row, status: 'unverifiable', expected };
    const productionConflict = hasProduction && existingVals.production !== 0 && existingVals.production !== expected.production;
    const dispatchConflict = hasDispatch && existingVals.dispatch !== 0 && existingVals.dispatch !== expected.dispatch;
    if (productionConflict || dispatchConflict) {
      return { row, status: 'mismatch', existing: existingVals, expected };
    }
    const fillProduction = hasProduction && existingVals.production === 0 && expected.production !== 0;
    const fillDispatch = hasDispatch && existingVals.dispatch === 0 && expected.dispatch !== 0;
    if (fillProduction || fillDispatch) {
      return { row, status: 'fillable', existing: existingVals, expected, fillProduction, fillDispatch, rowIdx: existingVals.rowIdx };
    }
    return { row, status: 'duplicate', existing: existingVals };
  });
}

// Works out exactly where every brand-new row needs to physically land so a block stays in true
// chronological order, instead of always tacking new rows onto the end regardless of their date — the
// bug that scrambled real blocks (e.g. IT 500 CONTAINER landing 1,6,10,11 Aug then 3,4,4,12 Aug) once
// dates got confirmed/pushed out of order. `existingRowsOrdered` is the block's real rows in physical
// top-to-bottom order; `newRows` is `[{ row, dateKey }, ...]`. Returns groups in ascending physical-
// target order: `{ beforeRowIdx: <original 0-indexed row to insert directly above> | null (null = goes
// after everything, a plain trailing append, no insert needed), rows: [...] }`. A "new" row's date is
// guaranteed absent from existingDates (that's what makes it 'new'), so there's never a same-date tie
// to resolve here.
function planChronologicalInserts(existingRowsOrdered, newRows) {
  const sorted = [...newRows].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const groups = [];
  sorted.forEach(nr => {
    // First existing row (in physical/chronological order) whose date falls after this new one — the
    // new row belongs directly above it. None found means every existing date is earlier: append.
    const target = existingRowsOrdered.find(er => er.dateKey > nr.dateKey);
    const beforeRowIdx = target ? target.rowIdx : null;
    const last = groups[groups.length - 1];
    if (last && last.beforeRowIdx === beforeRowIdx) last.rows.push(nr);
    else groups.push({ beforeRowIdx, rows: [nr] });
  });
  return groups;
}

// Turns the groups from planChronologicalInserts into two things a caller needs: (1) the FINAL 0-indexed
// row each group's first new row will land on, once every insertion (including groups above it) has
// actually happened — pure arithmetic, independent of what order the real Sheets insertRange calls fire
// in; and (2) a function that maps any OTHER already-existing row's original index (e.g. a 'fillable'
// row that needs its own single-cell update) to its final index after these same insertions shift it
// down. Groups are already in ascending original-target order (append/null last) by construction.
function finalizeInsertPlacement(nextRowIdx, groups) {
  let cumulative = 0;
  groups.forEach(g => {
    const originalTarget = g.beforeRowIdx === null ? nextRowIdx : g.beforeRowIdx;
    g.originalTarget = originalTarget;
    g.finalStartRow = originalTarget + cumulative;
    cumulative += g.rows.length;
  });
  const totalInserted = cumulative;
  // A previously-existing row shifts down by the combined size of every group whose insertion point is
  // at or above it (rows literally pushed down by rows inserted above/at their position); a group whose
  // target is BELOW it (or the trailing append, which is never above anything real) never affects it.
  const adjustExistingRowIdx = (rowIdx) => rowIdx + groups.reduce((s, g) => s + (g.beforeRowIdx !== null && g.beforeRowIdx <= rowIdx ? g.rows.length : 0), 0);
  return { groups, totalInserted, adjustExistingRowIdx };
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
  // One entry per variant: exactly where its block ended up and how far its just-written rows reached
  // (row1 = 1-indexed sheet row number of the LAST row just written, real or phantom-padding). Lets a
  // caller building a summary-tab formula reference (generateCustomerSheetStructure) point at a precise
  // cell without re-deriving this same column/row math itself.
  const placements = [];
  // Collected across every variant — a date that already has a row but disagrees with what we
  // computed. Never pushed (existing rows are never touched), always reported.
  const mismatches = [];
  // Real "insert N blank rows before row X" structural requests — one per contiguous group of new rows
  // that has to land somewhere other than the very end of the block, so the block stays in TRUE
  // chronological order instead of new dates always landing after whatever's already there regardless
  // of date. Turned into actual Sheets API requests by the caller (pushCustomerSheet), which has the
  // tab's real numeric sheetId; this function only knows column/row math.
  const insertRequests = [];

  // Processes every variant that resolves to the SAME block TOGETHER, as one combined batch, rather
  // than one at a time — critical for correctness, not just tidiness: if variant A gets its rows
  // placed first and variant B (sharing the same block) later needs to insert a row chronologically
  // BEFORE one of A's already-placed rows, A's Opening formula must reference B's new row as its real
  // predecessor. Handling them one at a time and patching row numbers afterward would shift WHERE a
  // row lands but leave its formula still pointing at the wrong (pre-insert) predecessor, silently
  // dropping B's row out of the balance chain. Treating every entry destined for this block as one
  // unified list of new rows before computing any formula avoids that entirely.
  function processBlockGroup(match, entries) {
    const perEntry = entries.map(({ v, incomingRows }) => ({ v, classified: classifyIncomingRows(match, incomingRows) }));
    perEntry.forEach(({ v, classified }) => {
      const variantMismatches = classified
        .filter(c => c.status === 'mismatch' || c.status === 'unverifiable')
        .map(c => ({ date: c.row[0], reason: c.status === 'unverifiable' ? 'unverifiable' : 'value_mismatch', existing: c.existing, expected: c.expected }));
      if (variantMismatches.length) mismatches.push({ title: v.title, mismatches: variantMismatches });
    });

    // Every 'new' row across every entry sharing this block, combined into one chronological placement
    // pass — tagged with which entry (vi) it came from, purely so `placements` can still report a
    // separate lastWrittenRow1 per original variant/title afterward.
    const allNew = [];
    perEntry.forEach(({ classified }, vi) => {
      classified.filter(c => c.status === 'new').forEach(c => allNew.push({ row: c.row, dateKey: canonicalDateKey(c.row[0]), vi }));
    });

    const hasRealPriorRow = match.existingDates.size > 0;
    const groups = planChronologicalInserts(match.existingRowsOrdered, allNew);
    const { totalInserted, adjustExistingRowIdx } = finalizeInsertPlacement(match.nextRowIdx, groups);

    // Fillable: the sheet already has a row for this date with a genuinely blank/0 Production and/or
    // Dispatch cell — write ONLY that one cell, exactly like a person filling in a gap by hand. Never
    // touches Opening or the other column on that row, so a live Closing formula recalculates on its
    // own. Its original rowIdx has to be adjusted for any inserts landing above it from this same
    // combined batch.
    perEntry.forEach(({ classified }) => {
      classified.filter(c => c.status === 'fillable').forEach(c => {
        const finalRowIdx = adjustExistingRowIdx(c.rowIdx);
        const k = canonicalDateKey(c.row[0]);
        const prevVals = match.existingValuesByDate.get(k);
        const nextVals = { ...prevVals };
        if (c.fillProduction) {
          patches.push({ startRow0: finalRowIdx, startCol0: match.startCol + 2, values: [[c.expected.production]] });
          nextVals.production = c.expected.production;
        }
        if (c.fillDispatch) {
          patches.push({ startRow0: finalRowIdx, startCol0: match.startCol + 3, values: [[c.expected.dispatch]] });
          nextVals.dispatch = c.expected.dispatch;
        }
        match.existingValuesByDate.set(k, nextVals);
      });
    });

    if (!allNew.length) return;
    const openingCol = colLetter(match.startCol + 1);
    const prodCol = colLetter(match.startCol + 2);
    const dispCol = colLetter(match.startCol + 3);
    const closingCol = colLetter(match.startCol + match.width - 1);

    // Reversed on purpose: groups are in ascending (chronological) order, but each insert physically
    // shifts everything below its own target down, so executing them bottom-to-top (highest original
    // row index first) is what keeps every OTHER group's original target index valid when its own
    // turn comes — the caller issues insertRequests in exactly this array order, in one batchUpdate.
    [...groups].reverse().forEach(g => {
      if (g.beforeRowIdx !== null) {
        insertRequests.push({ startCol0: match.startCol, width: match.width, beforeRowIdx: g.beforeRowIdx, count: g.rows.length });
      }
    });

    const values = [];
    const newExistingRows = []; // {rowIdx, dateKey} for every newly-written row, at its FINAL position
    const lastRow0ByEntry = {};
    groups.forEach((g, gi) => {
      g.rows.forEach((nr, ri) => {
        const row = nr.row || [];
        const row0 = g.finalStartRow + ri;
        const thisRow1 = row0 + 1;
        // Only the very first row ever written into a genuinely empty block has no real row above it
        // to reference — every other row, whether inserted mid-block or appended at the end, always
        // has SOMETHING correct directly above it once the real inserts have physically happened
        // (Sheets auto-adjusts every formula that referenced a shifted cell, so referencing by FINAL
        // row number here is always correct — and since every row from every entry sharing this block
        // was placed in this SAME pass, that's true across entries too, not just within one).
        const isVeryFirstOfBlock = gi === 0 && ri === 0 && !hasRealPriorRow;
        const opening = isVeryFirstOfBlock ? 0 : `=${closingCol}${thisRow1 - 1}`;
        const out = [
          forceTextValue(row[0] || ''),
          opening,
          Number(row[2]) || 0,
          Number(row[3]) || 0,
          `=${openingCol}${thisRow1}+${prodCol}${thisRow1}-${dispCol}${thisRow1}`,
        ];
        for (let ci = 5; ci < row.length; ci++) out.push(row[ci]);
        values.push({ row0, out });
        newExistingRows.push({ rowIdx: row0, dateKey: nr.dateKey });
        lastRow0ByEntry[nr.vi] = row0;
      });
    });
    // Each new row is its own single-row patch (rather than one contiguous block) since chronological
    // insertion can scatter them across several disjoint gaps in the same block, not just one
    // trailing range.
    values.forEach(({ row0, out }) => patches.push({ startRow0: row0, startCol0: match.startCol, values: [out] }));

    // Whenever a group of new rows lands directly ABOVE a pre-existing row (beforeRowIdx !== null),
    // that pre-existing row's own Opening cell still references whatever USED to be its predecessor —
    // Google Sheets only auto-adjusts a formula's cell reference when the cell it points AT physically
    // moves; it has no way to know a newly-inserted row should now become the logical predecessor, since
    // neither the existing row's old predecessor nor its own formula content changed. Left alone, that
    // one stale link corrupts every row's balance below it, permanently, even though the rest of the
    // chain (including the just-written new rows above it) is internally consistent — confirmed against
    // a real customer sheet where exactly this left a pre-existing row's Opening pointing two rows too
    // far back after new rows were correctly inserted above it. Rewritten here as a fresh live formula
    // (never a flattened static number, same self-referencing convention as every row this app writes)
    // chaining from the new group's real last row, so the sheet keeps behaving like a person edited an
    // old row and everything below it recalculated. Closing is rewritten too, in case that existing
    // row's Closing was never a formula to begin with — otherwise fixing Opening alone wouldn't help.
    groups.forEach(g => {
      if (g.beforeRowIdx === null) return;
      const existingRow0 = adjustExistingRowIdx(g.beforeRowIdx);
      const existingRow1 = existingRow0 + 1;
      const lastNewRow1 = g.finalStartRow + g.rows.length;
      patches.push({ startRow0: existingRow0, startCol0: match.startCol + 1, values: [[`=${closingCol}${lastNewRow1}`]] });
      patches.push({ startRow0: existingRow0, startCol0: match.startCol + match.width - 1, values: [[`=${openingCol}${existingRow1}+${prodCol}${existingRow1}-${dispCol}${existingRow1}`]] });
    });

    perEntry.forEach(({ v }, vi) => {
      const lastRow0 = lastRow0ByEntry[vi] !== undefined ? lastRow0ByEntry[vi] : match.nextRowIdx - 1;
      placements.push({ title: v.title, startCol0: match.startCol, width: match.width, lastWrittenRow1: lastRow0 + 1 });
    });

    // If a LATER variant in this same push also resolves to this block via the "brand new title"
    // path below (blocks.find picking up a block created earlier in THIS push), it needs to see this
    // block's TRUE post-insert state — including every existing row's shifted position — not the
    // pre-insert snapshot. Advancing the same `match` object here (the same one `blocks.find` returns
    // next time) is what makes that safe.
    match.nextRowIdx += totalInserted;
    match.existingRowsOrdered = match.existingRowsOrdered
      .map(er => ({ rowIdx: adjustExistingRowIdx(er.rowIdx), dateKey: er.dateKey }))
      .concat(newExistingRows)
      .sort((a, b) => a.rowIdx - b.rowIdx);
    allNew.forEach(nr => {
      match.existingDates.add(nr.dateKey);
      match.existingValuesByDate.set(nr.dateKey, { production: Number((nr.row || [])[2]) || 0, dispatch: Number((nr.row || [])[3]) || 0 });
    });
  }

  // Group every variant that already matches a real block upfront, so two variants sharing one block
  // (two catalog entries mapped to the same real item, or two differently-worded register lines) are
  // always processed together via processBlockGroup — see its comment for why that matters.
  const existingGroups = new Map(); // match -> [{ v, incomingRows }]
  const rest = [];
  (Array.isArray(variants) ? variants : []).forEach(v => {
    // blockTitleOverride: set when the person explicitly picked which real block a not-yet-matched
    // item belongs to (see the block-picker in the review UI) — takes priority over the item's own
    // title for matching purposes ONLY; v.title itself (the item's real name) is left untouched so it
    // keeps working as the stable key the client's review edits are stored under.
    const key = normalizeTabKey(v.blockTitleOverride || v.title);
    const match = blocks.find(b => normalizeTabKey(b.title) === key);
    const incomingRows = Array.isArray(v.rows) ? v.rows : [];
    if (match) {
      if (!existingGroups.has(match)) existingGroups.set(match, []);
      existingGroups.get(match).push({ v, incomingRows });
    } else {
      rest.push({ v, incomingRows });
    }
  });
  existingGroups.forEach((entries, match) => processBlockGroup(match, entries));

  rest.forEach(({ v, incomingRows }) => {
    // Same override-aware key as above — and the actual new block's title text, if one truly gets
    // created below, also has to honor it: if the person explicitly typed a name via "+ New block",
    // that's the name that should land in the Sheet, not silently the item's own title.
    const blockTitle = v.blockTitleOverride || v.title;
    const key = normalizeTabKey(blockTitle);
    const header = (v.header && v.header.length) ? v.header : ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'];
    // A PRECEDING variant with no pre-existing block, but the SAME title, may have already created
    // this exact block earlier in this same loop (matching the original append-only behavior) — route
    // it through the normal combined-block path instead of creating a second, duplicate block.
    const rematch = blocks.find(b => normalizeTabKey(b.title) === key);
    if (rematch) { processBlockGroup(rematch, [{ v, incomingRows }]); return; }
    {
      const startCol = rightmostCol === -1 ? 0 : rightmostCol + 1;
      const width = Math.max(header.length, ...incomingRows.map(r => (r || []).length), 1);
      const openingCol = colLetter(startCol + 1);
      const prodCol = colLetter(startCol + 2);
      const dispCol = colLetter(startCol + 3);
      const closingCol = colLetter(startCol + width - 1);
      const dataStartRow0 = usedHeaderRowIdx + 1;
      patches.push({ startRow0: usedHeaderRowIdx - 1, startCol0: startCol, values: [[blockTitle || '']] });
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
      placements.push({ title: v.title, startCol0: startCol, width, lastWrittenRow1: dataStartRow0 + values.length });
      rightmostCol = startCol + width;
      const newExistingRows = incomingRows.map((r, i) => ({ rowIdx: dataStartRow0 + i, dateKey: canonicalDateKey((r || [])[0]) }));
      const newExistingValues = new Map(incomingRows.map((r, i) => [canonicalDateKey((r || [])[0]), { production: Number((r || [])[2]) || 0, dispatch: Number((r || [])[3]) || 0, rowIdx: dataStartRow0 + i }]));
      blocks.push({
        title: blockTitle, startCol, width, nextRowIdx: dataStartRow0 + values.length,
        existingDates: new Set(newExistingRows.map(er => er.dateKey)), existingValuesByDate: newExistingValues,
        existingRowsOrdered: newExistingRows, lastRowValues: null,
      });
    }
  });
  return { patches, placements, mismatches, insertRequests };
}

// Highlight color for the "what's new since last push" formatting below — a warm tint matching this
// app's own accent color, applied to the exact cells a patch just wrote — and ONLY those, so the
// highlight always means "new since the last push," never "new at some unknown point in the past."
// Every push first clears any leftover highlight from an EARLIER push (see readHighlightedCellsByTab
// and buildUnhighlightRequests below) before applying this push's own, so highlighting never just
// keeps accumulating forever.
const HIGHLIGHT_COLOR = { red: 0.941, green: 0.886, blue: 0.769 };
const NO_HIGHLIGHT_COLOR = { red: 1, green: 1, blue: 1 };
const HIGHLIGHT_COLOR_TOLERANCE = 0.01;

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
const colorsClose = (a, b) => Math.abs((a.red || 0) - b.red) < HIGHLIGHT_COLOR_TOLERANCE
  && Math.abs((a.green || 0) - b.green) < HIGHLIGHT_COLOR_TOLERANCE
  && Math.abs((a.blue || 0) - b.blue) < HIGHLIGHT_COLOR_TOLERANCE;
// Reads which cells in the given (already-existing) tabs currently carry THIS app's highlight color —
// never any other formatting a customer's own sheet template might have — so a fresh push can clear
// exactly last push's highlight and nothing else. One read call for every tab about to be written to.
async function readHighlightedCellsByTab(sheets, spreadsheetId, tabNames, existingMeta) {
  const realTabs = (tabNames || []).filter(t => existingMeta[t]);
  if (!realTabs.length) return {};
  const resp = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: realTabs.map(t => wholeSheetRangeA1(t, existingMeta[t])),
    fields: 'sheets(properties(title),data(startRow,startColumn,rowData(values(userEnteredFormat(backgroundColor)))))',
  });
  const result = {};
  (resp.data.sheets || []).forEach(s => {
    const title = s.properties.title;
    const cells = [];
    (s.data || []).forEach(d => {
      const rowOffset = d.startRow || 0;
      const colOffset = d.startColumn || 0;
      (d.rowData || []).forEach((rd, ri) => {
        (rd.values || []).forEach((cell, ci) => {
          const bg = cell.userEnteredFormat && cell.userEnteredFormat.backgroundColor;
          if (bg && colorsClose(bg, HIGHLIGHT_COLOR)) cells.push({ row: rowOffset + ri, col: colOffset + ci });
        });
      });
    });
    result[title] = cells;
  });
  return result;
}
// Merges a row's highlighted column indices into contiguous [start, end] spans, so an appended
// ledger row (columns highlighted wall-to-wall) costs one repeatCell request, not one per cell.
function runLengthEncodeColumns(cols) {
  const sorted = Array.from(cols).sort((a, b) => a - b);
  const spans = [];
  let start = null, prev = null;
  sorted.forEach(c => {
    if (start === null) { start = c; prev = c; return; }
    if (c === prev + 1) { prev = c; return; }
    spans.push([start, prev]);
    start = c; prev = c;
  });
  if (start !== null) spans.push([start, prev]);
  return spans;
}
// Resets exactly the cells readHighlightedCellsByTab found still carrying an earlier push's highlight
// back to no fill — never a blanket whole-tab reset, so any OTHER formatting the customer's own sheet
// template has (header shading, borders, whatever) is never touched.
function buildUnhighlightRequests(sheetId, highlightedCells) {
  if (sheetId === undefined || sheetId === null || !highlightedCells || !highlightedCells.length) return [];
  const colsByRow = new Map();
  highlightedCells.forEach(({ row, col }) => {
    if (!colsByRow.has(row)) colsByRow.set(row, new Set());
    colsByRow.get(row).add(col);
  });
  const requests = [];
  colsByRow.forEach((cols, row) => {
    runLengthEncodeColumns(cols).forEach(([c0, c1]) => {
      requests.push({
        repeatCell: {
          range: { sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: c0, endColumnIndex: c1 + 1 },
          cell: { userEnteredFormat: { backgroundColor: NO_HIGHLIGHT_COLOR } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
    });
  });
  return requests;
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
  // Fold hyphens/dashes into a plain space FIRST, before any alias or whitespace-collapse logic runs.
  // Confirmed against a real mismatch: the Production Register spells one Bindal item "T-GEL CONT"
  // (hyphen) while the real Sheet's block is titled "T GEL CONTAINER" (space) — since only whitespace
  // got collapsed below, the hyphen survived and the two never converged to the same key. Treating
  // -/–/— exactly like a space (rather than stripping it outright) keeps this symmetric with how
  // spaces are already handled everywhere else in this function.
  s = s.replace(/[-‐-―]/g, ' ');
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
  // Blanket "pkt"/"pkts" strip, wherever it appears — not just trailing-with-a-digit-prefix like the
  // two rules above require. Real block titles have shown up with the word landing mid-string or
  // followed by more text (e.g. a trailing "(BOWL)" after the count), which the anchored `$` rules
  // above never reach. Deliberately NO word-boundary requirement before it, same reasoning as the
  // client's normalizeForCatalogMatch: it's routinely glued straight onto a digit with no space
  // ("70pkt", "40pkt"), and a digit/letter join isn't a regex word boundary anyway. "pkt"/"pkts" (and
  // the "ppkt" typo seen in real sheets) is never part of a real item's identity, so it's always safe
  // to drop outright. Confirmed per explicit customer instruction: "always ignore the pkt/pkts string
  // whenever matching."
  s = s.replace(/p+kts?\.?/gi, '');
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
  const structuralInsertRequests = [];
  for (const plan of tabPlans) {
    const previousGrid = previousValuesByTab[plan.tabName] || [];
    const { patches, placements, mismatches, insertRequests } = computeMergePatches(previousGrid, plan.variants);
    patchesByTab[plan.tabName] = patches;
    // Real "make room" requests — inserting blank rows so a chronologically-earlier new date lands
    // BEFORE whatever's already physically below it, instead of always after. Must happen (and finish)
    // before any values get written, since every row target in `patches` already assumes these inserts
    // have already happened. Scoped to the exact column range of the block they belong to, so a block
    // sitting side-by-side with others in the same tab never shifts rows that aren't its own.
    const sheetId = existingMeta[plan.tabName] && existingMeta[plan.tabName].sheetId;
    (insertRequests || []).forEach(ir => {
      structuralInsertRequests.push({
        insertRange: {
          range: { sheetId, startRowIndex: ir.beforeRowIdx, endRowIndex: ir.beforeRowIdx + ir.count, startColumnIndex: ir.startCol0, endColumnIndex: ir.startCol0 + ir.width },
          shiftDimension: 'ROWS',
        },
      });
    });
    patches.forEach(p => {
      const startColL = colLetter(p.startCol0);
      const endColL = colLetter(p.startCol0 + Math.max(0, ...p.values.map(r => r.length)) - 1);
      const range = `${plan.tabName}!${startColL}${p.startRow0 + 1}:${endColL}${p.startRow0 + p.values.length}`;
      dataUpdates.push({ range, values: p.values });
    });
    // placements is included alongside the usual ok/newRows summary purely so a caller building
    // structure for a brand-new customer (generateCustomerSheetStructure) can construct a precise
    // summary-tab formula reference per item without re-deriving this column/row math itself. Ignored
    // by every other existing caller of pushCustomerSheet.
    // mismatches: dates that already had a row in the sheet whose numbers didn't match what we sent —
    // never written (existing rows are never touched), always reported, so a push can never silently
    // report success while quietly leaving a real discrepancy behind.
    results.push({ tab: plan.tabName, ok: true, newRows: patches.reduce((s, p) => s + p.values.length, 0), placements, mismatches });
  }
  if (structuralInsertRequests.length) {
    try {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: structuralInsertRequests } });
    } catch (e) {
      const msg = friendlyGoogleError(e);
      return tabPlans.map(plan => ({ tab: plan.tabName, ok: false, error: msg }));
    }
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
  // Clear whatever's left highlighted from an EARLIER push before applying this one's — otherwise
  // highlighting only ever accumulates and stops meaning "new since last push." Only tabs that already
  // existed before this push can have a leftover highlight to clear; a tab just created above is empty.
  // Best-effort: a failure here shouldn't block this push's own (successful) highlight from applying.
  let highlightedCellsByTab = {};
  try {
    const preExistingTabNames = tabPlans.map(p => p.tabName).filter(t => !missing.includes(t));
    highlightedCellsByTab = await readHighlightedCellsByTab(sheets, spreadsheetId, preExistingTabNames, existingMeta);
  } catch (e) {
    console.error('Customer sheet pre-push highlight read failed — skipping un-highlight this push:', e);
  }
  tabPlans.forEach(plan => {
    const sheetId = existingMeta[plan.tabName] && existingMeta[plan.tabName].sheetId;
    formatRequests.push(...buildUnhighlightRequests(sheetId, highlightedCellsByTab[plan.tabName]));
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

// --- Customer Sheets: GENERATE a brand-new customer's full tab/item-block/summary structure inside a
// spreadsheet the person already created and shared with the service account themselves. The service
// account can't create+own a new spreadsheet on its own — verified directly: plain service accounts
// have no Drive storage of their own, so `spreadsheets.create` comes back 403 PERMISSION_DENIED — but
// writing into an already-shared blank sheet needs nothing beyond the ordinary Sheets permission this
// app already has for every other customer sheet, so that's the split: the person creates+shares an
// empty Sheet (the same one-time step already required for every customer), this fills it in.
//
// Every item block is pre-seeded with GENERATED_PAD_ROWS "phantom" carry-forward rows (blank date, 0
// production/dispatch, live Opening/Closing formulas chained to the row above) — the exact authoring
// convention already verified directly in the real BINDAL STOCK / DIAMOND / anmol files (rows sitting
// well past the last real transaction that just keep repeating the same balance). This is what lets the
// generated summary tab's formula point at a fixed row far below any real data and still show the
// correct live balance from day one: computeMergePatches stops scanning a block's "real" data at the
// first blank-date cell, so every phantom row is invisible to it — a real push later overwrites them
// from the top down exactly like it would write brand-new rows, and every phantom row still below
// correctly references "the row directly above me," so the chain self-heals with zero extra work. No
// separate re-pointing of the summary formula is ever needed unless GENERATED_PAD_ROWS worth of real
// transactions actually gets used up on one item, same limit every hand-built customer file already has.
const GENERATED_PAD_ROWS = 300;
const SUMMARY_HEADER = ['s. no.', 'item name', 'quantity.'];

function quoteTabNameForFormula(tabName) {
  return `'${String(tabName || '').replace(/'/g, "''")}'`;
}

// Finds an existing tab whose title matches /summary/i (same convention importCustomerSheet already
// uses to exclude it from the item list), or null if this spreadsheet doesn't have one yet.
function findSummaryTabName(existingMeta) {
  return Object.keys(existingMeta).find(t => /summary/i.test(t)) || null;
}

// Writes (creating the tab if needed) one row per generated item into the summary tab, appending after
// whatever's already there. Known limitation: if an existing summary tab ends with a hand-typed "Total"
// row, new rows get appended AFTER it rather than re-inserted above it — acceptable for this feature's
// primary case (a brand-new summary tab, written fresh, with no Total row yet); re-running generate
// against an already-in-use customer sheet may need that Total row moved back down by hand afterward.
async function writeSummaryTab(sheets, spreadsheetId, existingMeta, items) {
  let summaryTabName = findSummaryTabName(existingMeta);
  let startRow0 = 1; // 0-indexed — row right after the header, for a brand-new summary tab
  let nextSerial = 1;
  if (!summaryTabName) {
    summaryTabName = 'summary';
    const createResp = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: summaryTabName, index: 0 } } }] },
    });
    const props = createResp.data.replies[0].addSheet.properties;
    existingMeta[props.title] = { sheetId: props.sheetId, rowCount: 1000, columnCount: 26 };
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${summaryTabName}!A1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [SUMMARY_HEADER] },
    });
  } else {
    const readResp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: wholeSheetRangeA1(summaryTabName, existingMeta[summaryTabName]),
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const rows = readResp.data.values || [];
    let lastRow = 0; // 0-indexed count of rows already used, header included
    for (let r = 1; r < rows.length; r++) {
      if (normalizeCellStr((rows[r] || [])[1])) { lastRow = r; nextSerial++; }
    }
    startRow0 = rows.length ? lastRow + 1 : 1;
  }
  const values = items.map((it, i) => [nextSerial + i, it.item, it.summaryFormula]);
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${summaryTabName}!A${startRow0 + 1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

// categories: [{ name: 'IT 500', items: ['IT 500 LID', 'IT 500 CONTAINER'] }, ...]
async function generateCustomerSheetStructure(spreadsheetId, categories) {
  const id = String(spreadsheetId || '').trim();
  const cats = (Array.isArray(categories) ? categories : [])
    .map(c => ({ name: String((c && c.name) || '').trim(), items: (Array.isArray(c && c.items) ? c.items : []).map(i => String(i || '').trim()).filter(Boolean) }))
    .filter(c => c.name && c.items.length);
  if (!cats.length) throw new Error('At least one category with at least one item is required.');

  const padRow = () => ['', 0, 0, 0];
  const itemGroups = cats.map(c => ({
    tabName: c.name,
    variants: c.items.map(item => ({
      title: item,
      header: ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'],
      rows: Array.from({ length: GENERATED_PAD_ROWS }, padRow),
    })),
  }));

  const pushResults = await pushCustomerSheet(id, itemGroups);
  const failed = pushResults.filter(r => !r.ok);
  if (failed.length) throw new Error(failed.map(f => `${f.tab}: ${f.error}`).join('; '));

  const items = [];
  pushResults.forEach(r => {
    (r.placements || []).forEach(p => {
      const closingCol = colLetter(p.startCol0 + p.width - 1);
      items.push({
        item: p.title,
        sheetGroup: r.tab,
        summaryFormula: `=${quoteTabNameForFormula(r.tab)}!${closingCol}${p.lastWrittenRow1}`,
      });
    });
  });

  const sheets = getSheetsClient();
  const existingMeta = await getSheetMeta(sheets, id);
  await writeSummaryTab(sheets, id, existingMeta, items);

  return { tabCount: cats.length, items: items.map(({ item, sheetGroup }) => ({ item, sheetGroup })) };
}

async function generateCustomerSheetStructureHandler(req, res) {
  try {
    const { spreadsheetId, categories } = req.body || {};
    if (!spreadsheetId || !String(spreadsheetId).trim()) {
      return res.status(400).json({ error: 'spreadsheetId is required.' });
    }
    if (!Array.isArray(categories) || !categories.length) {
      return res.status(400).json({ error: 'categories must be a non-empty array.' });
    }
    const result = await generateCustomerSheetStructure(String(spreadsheetId).trim(), categories);
    res.json({ ok: true, ...result });
  } catch (e) {
    console.error('Customer sheet generate-structure error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

// Dry-run counterpart to pushCustomerSheet — reads the SAME "before" state and classifies the SAME
// rows (reusing parseExistingBlocks so there's no risk of the preview ever disagreeing with what an
// actual push would do), but never writes anything and returns plain numbers instead of formula
// strings, since a human reviewing a diff wants to see "what will this balance become," not raw
// Sheets formula syntax. EVERY row for the item is returned here — never just the new ones — each
// tagged with its status ('new' / 'duplicate' / 'mismatch' / 'unverifiable'), so the review screen can
// show the real entry for every single row and just flag the ones that need a look, instead of ever
// summarizing a row away as plain text.
// Builds the exact display rows for one variant, in TRUE chronological order — merging our own
// classified rows with the sheet's real timeline. This is what makes "fetch the opening/production
// from the real row instead of computing it" and "a brand-new row's opening comes from whatever
// precedes it chronologically" both true on the preview screen, not just on the actual push:
//   - A date we already have a classified row for (any status) shows the SHEET's own opening/closing
//     directly (never this app's from-scratch running total) — a 'fillable' column shows what we're
//     about to write there instead of the sheet's current 0/blank, with Closing recomputed only enough
//     to reflect that one fill, exactly like the real Closing formula will once the write lands.
//   - A brand-new date chains its Opening off whatever the closing balance truly was immediately
//     before it — which might be another new row earlier in this same batch, or a real sheet row this
//     app has never confirmed a matching entry for at all (a hand-typed row with no ledger counterpart
//     still has to be walked past so its closing carries forward correctly, even though it isn't shown
//     as one of "our" rows).
function buildDisplayRows(match, classified) {
  const rows = [];
  const classifiedByDate = new Map(classified.map(c => [canonicalDateKey(c.row[0]), c]));
  const sheetOnlyDates = (match ? match.existingRowsOrdered : [])
    .map(er => er.dateKey)
    .filter(dk => !classifiedByDate.has(dk));
  const timeline = classified.map(c => canonicalDateKey(c.row[0]))
    .concat(sheetOnlyDates)
    .sort((a, b) => a.localeCompare(b));
  let lastClosing = 0;
  timeline.forEach(dateKey => {
    const c = classifiedByDate.get(dateKey);
    if (!c) {
      // A real sheet row for a date we have no ledger entry for at all — nothing of ours to show, but
      // its closing balance still has to carry forward into whatever we chain after it.
      const vals = match.existingValuesByDate.get(dateKey);
      if (vals && vals.closing !== undefined) lastClosing = vals.closing;
      return;
    }
    const row = c.row;
    const production0 = Number(row[2]) || 0;
    const dispatch0 = Number(row[3]) || 0;
    if (c.status === 'new') {
      const opening = lastClosing;
      const closing = opening + production0 - dispatch0;
      rows.push({ date: row[0] || '', opening, production: production0, dispatch: dispatch0, closing, status: c.status, existing: null });
      lastClosing = closing;
    } else {
      const ex = c.existing || {};
      const production = c.fillProduction ? c.expected.production : (ex.production !== undefined ? ex.production : production0);
      const dispatch = c.fillDispatch ? c.expected.dispatch : (ex.dispatch !== undefined ? ex.dispatch : dispatch0);
      const opening = ex.opening !== undefined ? ex.opening : lastClosing;
      const closing = c.status === 'fillable' ? (opening + production - dispatch) : (ex.closing !== undefined ? ex.closing : (opening + production - dispatch));
      rows.push({ date: row[0] || '', opening, production, dispatch, closing, status: c.status, existing: c.existing || null, fillProduction: !!c.fillProduction, fillDispatch: !!c.fillDispatch });
      lastClosing = closing;
    }
  });
  return rows;
}

function previewCustomerSheet(previousValuesByTab, tabPlans, missing) {
  return tabPlans.map(plan => {
    const previousGrid = previousValuesByTab[plan.tabName] || [];
    const { blocks } = parseExistingBlocks(previousGrid);
    // Same reasoning as computeMergePatches: two variants that resolve to the SAME real block (two
    // catalog entries for one item, or two differently-worded register lines) have to be walked as ONE
    // combined chronological timeline, not one at a time — otherwise a variant processed first could
    // show a stale Opening/Closing that never accounts for the other variant's date landing in between,
    // even though preview and an actual push must never disagree about what the real numbers are.
    const resultByVariant = new Map(); // v -> rows[]
    const existingGroups = new Map(); // match -> [v, ...]
    const restVariants = [];
    (plan.variants || []).forEach(v => {
      // Same blockTitleOverride priority as the real push (computeMergePatches) — see its comment.
      const key = normalizeTabKey(v.blockTitleOverride || v.title);
      const match = blocks.find(b => normalizeTabKey(b.title) === key);
      if (match) {
        if (!existingGroups.has(match)) existingGroups.set(match, []);
        existingGroups.get(match).push(v);
      } else {
        restVariants.push(v);
      }
    });
    const processGroup = (match, vs) => {
      const combined = [];
      vs.forEach((v, vi) => {
        const incomingRows = Array.isArray(v.rows) ? v.rows : [];
        classifyIncomingRows(match, incomingRows).forEach(c => combined.push({ ...c, vi }));
      });
      combined.sort((a, b) => canonicalDateKey(a.row[0]).localeCompare(canonicalDateKey(b.row[0])));
      const rows = buildDisplayRows(match, combined);
      vs.forEach((v, vi) => resultByVariant.set(v, rows.filter((r, i) => combined[i].vi === vi)));
      // Preview never issues a real Sheets write, so the exact row index doesn't matter here
      // (rowIdx: -1) — only the dateKey ordering classifyIncomingRows/planChronologicalInserts rely on.
      if (match && rows.length) {
        rows.forEach((r, i) => {
          const k = canonicalDateKey(combined[i].row[0]);
          match.existingDates.add(k);
          match.existingValuesByDate.set(k, { production: r.production, dispatch: r.dispatch, opening: r.opening, closing: r.closing, rowIdx: -1 });
        });
        const newDateKeys = combined.filter(c => c.status === 'new').map(c => canonicalDateKey(c.row[0]));
        match.existingRowsOrdered = match.existingRowsOrdered
          .concat(newDateKeys.map(dateKey => ({ rowIdx: -1, dateKey })))
          .sort((a, b) => a.dateKey.localeCompare(b.dateKey));
      }
    };
    existingGroups.forEach((vs, match) => processGroup(match, vs));
    restVariants.forEach(v => {
      const key = normalizeTabKey(v.blockTitleOverride || v.title);
      const rematch = blocks.find(b => normalizeTabKey(b.title) === key);
      if (rematch) { processGroup(rematch, [v]); return; }
      const incomingRows = Array.isArray(v.rows) ? v.rows : [];
      resultByVariant.set(v, buildDisplayRows(null, classifyIncomingRows(null, incomingRows)));
    });
    const variants = (plan.variants || []).map(v => {
      const key = normalizeTabKey(v.blockTitleOverride || v.title);
      const match = blocks.find(b => normalizeTabKey(b.title) === key);
      return { title: v.title, isNewBlock: !match, rows: resultByVariant.get(v) || [] };
    });
    // Real, pre-existing block titles in this tab ONLY — `blocks` here is never mutated with a
    // fabricated "new block" entry during preview (unlike the real push), so this is exactly the list
    // the client's block-picker dropdown needs to offer for a not-yet-matched item, nothing invented.
    return { tabName: plan.tabName, isNewTab: missing.includes(plan.tabName), existingBlockTitles: blocks.map(b => b.title), variants };
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
      spreadsheetId,
      ranges: preExistingTabNames.map(t => wholeSheetRangeA1(t, existingMeta[t])),
      valueRenderOption: 'UNFORMATTED_VALUE',
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

// Flattens a tab's every real block/date into one array of {sheetTab, block, date, opening,
// production, dispatch, closing} — the shape the client's "Customer Sheets Mirror" register stores
// (one row per real ledger line, across every block in the tab), so the global search box can find
// anything in a customer's real Sheet without a live API call per search. Reuses parseExistingBlocks
// (already correctly finds every block/date in a tab) rather than re-deriving block/row detection —
// same reasoning as computeMergePatches above, kept as one shared implementation.
function extractMirrorRowsFromGrid(tabName, grid) {
  const { blocks } = parseExistingBlocks(grid);
  const rows = [];
  blocks.forEach(b => {
    if (b.width < 4) return; // narrower than Date/Opening/Production/Dispatch/Closing — nothing to mirror
    b.existingRowsOrdered.forEach(er => {
      const vals = b.existingValuesByDate.get(er.dateKey);
      if (!vals) return;
      const rawDate = normalizeDateCell((grid[er.rowIdx] || [])[b.startCol]);
      rows.push({ sheetTab: tabName, block: b.title, date: rawDate, opening: vals.opening, production: vals.production, dispatch: vals.dispatch, closing: vals.closing });
    });
  });
  return rows;
}

// Reads every non-"summary" tab of the given spreadsheet and returns the full item list (for the
// Product Catalog) AND every real ledger row flattened for the Customer Sheets Mirror (for global
// search) — both derived from the SAME single read, so seeding/refreshing the mirror never costs an
// extra API call beyond what importing already does. spreadsheetTitle is used as a suggested customer
// name — editable by the person importing it, same as the existing file-upload import flow already
// lets them edit the suggested name. valueRenderOption UNFORMATTED_VALUE matters here specifically for
// the mirror's numbers (a formatted "1,220" would not parse back to 1220 as a plain Number).
async function importCustomerSheet(spreadsheetId) {
  const sheets = getSheetsClient();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'properties.title,sheets.properties(title,gridProperties(rowCount,columnCount))',
  });
  const spreadsheetTitle = (meta.data.properties && meta.data.properties.title) || '';
  const allTabTitles = (meta.data.sheets || []).map(s => s.properties.title);
  const gridPropsByTitle = {};
  (meta.data.sheets || []).forEach(s => { gridPropsByTitle[s.properties.title] = s.properties.gridProperties || {}; });
  const dataTabTitles = allTabTitles.filter(t => !/summary/i.test(t));
  if (!dataTabTitles.length) {
    return { spreadsheetTitle, tabCount: allTabTitles.length, items: [], mirrorRows: [] };
  }
  // See wholeSheetRangeA1 above: a bare tab name is ambiguous with a cell reference (e.g. "N200",
  // "E900") and the Sheets API will silently read the wrong thing. Always quote + range every tab.
  const resp = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges: dataTabTitles.map(t => wholeSheetRangeA1(t, gridPropsByTitle[t])),
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const items = [];
  const mirrorRows = [];
  (resp.data.valueRanges || []).forEach((vr, idx) => {
    const tabName = dataTabTitles[idx];
    const grid = vr.values || [];
    extractItemsFromTabGrid(tabName, grid).forEach(it => items.push(it));
    extractMirrorRowsFromGrid(tabName, grid).forEach(r => mirrorRows.push(r));
  });
  return { spreadsheetTitle, tabCount: allTabTitles.length, items, mirrorRows };
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
    // Refresh the Customer Sheets Mirror with what's REALLY in the sheet now, post-write — a fresh
    // read (via the same importCustomerSheet used at import time) rather than reconstructing values
    // from the push's own patches, since those carry live formula strings (e.g. "=E5") for
    // Opening/Closing, not the evaluated numbers Google Sheets has now actually computed. Best-effort:
    // a refresh failure here must never mask the push's own (already-succeeded) results.
    let mirrorRows = [];
    try {
      const fresh = await importCustomerSheet(String(spreadsheetId).trim());
      mirrorRows = fresh.mirrorRows;
    } catch (e) {
      console.error('Post-push Customer Sheets Mirror refresh failed (push itself still succeeded):', e);
    }
    res.status(failed.length ? 502 : 200).json({ ok: failed.length === 0, results, mirrorRows });
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
    // Fire-and-forget, AFTER responding: keeps the Raw Material Pivot tab in sync automatically
    // whenever raw material data itself saves, so there's no separate button to remember to click.
    // Runs post-response so a slow or failed pivot rebuild never delays or breaks the actual save the
    // person is waiting on — a failure here is logged, not surfaced as if the save itself failed.
    if (req.params.tab === sanitizeTabName('fims_raw_material_in')) {
      rebuildRawMaterialPivot().catch(e => console.error('Auto pivot rebuild failed:', e));
    }
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

// --- main-sheet maintenance: list/delete tabs, build the Raw Material pivot ---
// Only used to LABEL every tab in the report the app shows before anyone deletes anything — never to
// auto-delete. "known": a register/lookup the app itself reads or writes and something on an app page
// is built from. "internal": bookkeeping the app needs but that never renders as a table anywhere —
// safe to offer as a one-click default. Anything else comes back "unrecognized" so a person can look
// it over themselves rather than this list silently deciding for them. A customer's own generated
// Sheet structure (via generateCustomerSheetStructureHandler) lives in THAT customer's spreadsheet,
// never this one, so it never appears here at all.
const KNOWN_APP_TAB_KEYS = [
  'fims_raw_material_in', 'fims_consumption', 'fims_production', 'fims_customer_dispatch',
  'fims_dabur_specs', 'fims_dabur_po', 'fims_dabur_dispatch',
  'fims_product_catalog', 'fims_customer_mapping', 'fims_customer_sheet_ids', 'fims_customer_name_aliases',
  'fims_abbreviations',
];
const INTERNAL_ONLY_TAB_KEYS = ['fims_training_examples', 'fims_customer_sheets_mirror'];

async function listTabsHandler(req, res) {
  try {
    const sheets = getSheetsClient();
    const spreadsheetId = getSpreadsheetId();
    const meta = await getSheetMeta(sheets, spreadsheetId);
    const known = new Set(KNOWN_APP_TAB_KEYS.map(sanitizeTabName));
    const internalOnly = new Set(INTERNAL_ONLY_TAB_KEYS.map(sanitizeTabName));
    const tabs = Object.keys(meta).map(title => ({
      title,
      kind: internalOnly.has(title) ? 'internal' : known.has(title) ? 'known' : 'unrecognized',
    }));
    res.json({ tabs });
  } catch (e) {
    console.error('List tabs error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

// Deletes exactly the tab names given in the request body — nothing implicit, nothing pattern-
// matched. The app only ever sends this the names a person reviewed and confirmed in the Settings
// screen (defaulting to the two INTERNAL_ONLY_TAB_KEYS above, extendable to any "unrecognized" tab
// they explicitly pick). A name that doesn't currently exist is reported back, not treated as an error
// — deleting something already gone is a no-op, not a failure.
async function deleteTabsHandler(req, res) {
  try {
    const tabNames = (req.body && Array.isArray(req.body.tabNames)) ? req.body.tabNames : [];
    if (!tabNames.length) return res.status(400).json({ error: 'tabNames must be a non-empty array.' });
    const sheets = getSheetsClient();
    const spreadsheetId = getSpreadsheetId();
    const meta = await getSheetMeta(sheets, spreadsheetId);
    const requests = [];
    const notFound = [];
    tabNames.forEach(name => {
      const m = meta[name];
      if (m) requests.push({ deleteSheet: { sheetId: m.sheetId } });
      else notFound.push(name);
    });
    if (requests.length) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
    }
    res.json({ deleted: tabNames.filter(n => !notFound.includes(n)), notFound });
  } catch (e) {
    console.error('Delete tabs error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

// Builds (or fully rebuilds) a native Google Sheets pivot table in its own tab, sourced from the Raw
// Material In tab — Rows grouped Size then GSM (each level gets Sheets' own built-in per-value filter
// dropdown right in the pivot header, which is what actually gives "filter by size and gsm" rather
// than anything this app has to build itself), Values showing total weight and entry count. Reads the
// live header row to find the size/gsm/weight_kg/date column offsets rather than assuming a fixed
// layout, since writeTab's header is a union built from whatever keys rows happen to carry. Always
// deletes and recreates the destination tab first so re-running this (e.g. after the source data grew)
// never stacks a second pivot table or errors on "cell already contains a pivot table" — safe to run
// as many times as it likes, which matters since putTab below calls this automatically on every Raw
// Material In save, not from a button. Returns {skipped: true} rather than throwing when there's
// nothing to pivot yet (an empty/just-cleared register) — a normal, expected state, not a failure.
const RAW_MATERIAL_PIVOT_TAB_NAME = 'Raw Material Pivot';
async function rebuildRawMaterialPivot() {
  const sheets = getSheetsClient();
  const spreadsheetId = getSpreadsheetId();
  const sourceTabName = sanitizeTabName('fims_raw_material_in');
  const destTabName = sanitizeTabName(RAW_MATERIAL_PIVOT_TAB_NAME);
  const meta = await getSheetMeta(sheets, spreadsheetId);
  const sourceMeta = meta[sourceTabName];
  if (!sourceMeta) return { skipped: true, reason: 'no-source-tab' };

  const headerResp = await sheets.spreadsheets.values.get({
    spreadsheetId, range: wholeSheetRangeA1(sourceTabName, { rowCount: 1, columnCount: sourceMeta.columnCount }),
  });
  const header = (headerResp.data.values && headerResp.data.values[0]) || [];
  const sizeIdx = header.indexOf('size');
  const gsmIdx = header.indexOf('gsm');
  const weightIdx = header.indexOf('weight_kg');
  const dateIdx = header.indexOf('date');
  if (sizeIdx === -1 || gsmIdx === -1) return { skipped: true, reason: 'no-data' };

  if (meta[destTabName]) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ deleteSheet: { sheetId: meta[destTabName].sheetId } }] },
    });
  }
  const addResp = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: destTabName } } }] },
  });
  const destSheetId = addResp.data.replies[0].addSheet.properties.sheetId;

  const values = [{
    pivotTable: {
      source: {
        sheetId: sourceMeta.sheetId,
        startRowIndex: 0,
        startColumnIndex: 0,
        endRowIndex: sourceMeta.rowCount,
        endColumnIndex: sourceMeta.columnCount,
      },
      rows: [
        { sourceColumnOffset: sizeIdx, showTotals: true, sortOrder: 'ASCENDING' },
        { sourceColumnOffset: gsmIdx, showTotals: true, sortOrder: 'ASCENDING' },
      ],
      values: [
        { summarizeFunction: 'SUM', sourceColumnOffset: weightIdx, name: 'Total Weight (kg)' },
        { summarizeFunction: 'COUNTA', sourceColumnOffset: dateIdx !== -1 ? dateIdx : sizeIdx, name: 'Entries' },
      ],
      valueLayout: 'HORIZONTAL',
    },
  }];
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        updateCells: {
          rows: [{ values }],
          start: { sheetId: destSheetId, rowIndex: 0, columnIndex: 0 },
          fields: 'pivotTable',
        },
      }],
    },
  });
  return { tabName: destTabName };
}

// Manual-trigger endpoint, kept as a fallback/debug path (e.g. to force a rebuild without a new save)
// even though nothing in the app's UI calls it — the pivot now rebuilds automatically, see putTab.
async function createRawMaterialPivotHandler(req, res) {
  try {
    const result = await rebuildRawMaterialPivot();
    if (result.skipped) {
      return res.status(400).json({ error: 'Raw Material In has no size/gsm data yet — upload at least one mill slip first.' });
    }
    res.json({ ok: true, ...result });
  } catch (e) {
    console.error('Create raw material pivot error:', e);
    res.status(502).json({ error: friendlyGoogleError(e) });
  }
}

module.exports = { readTab, writeTab, writeBlocksTab, getTab, putTab, putBlocksTab, pushCustomerSheetHandler, previewCustomerSheetHandler, importCustomerSheetHandler, getServiceAccountEmailHandler, generateCustomerSheetStructureHandler, listTabsHandler, deleteTabsHandler, createRawMaterialPivotHandler };