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
    fields: 'sheets.properties(title,gridProperties(rowCount,columnCount))',
  });
  const byTitle = {};
  (meta.data.sheets || []).forEach(s => {
    byTitle[s.properties.title] = {
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

// Builds the "summary" tab: S.No / Item Name / Quantity, one row per variant, plus a TOTAL row —
// matching the summary sheet found in every one of the real customer files (verified directly).
function buildSummaryGrid(rows) {
  const grid = [['S. No.', 'Item Name', 'Quantity']];
  let total = 0;
  (Array.isArray(rows) ? rows : []).forEach((r, i) => {
    const qty = Number(r.quantity) || 0;
    grid.push([i + 1, r.item || '', qty]);
    total += qty;
  });
  grid.push(['', 'TOTAL', total]);
  return grid;
}

// Pushes every item-group tab plus a summary tab to an external spreadsheet in one go. Reads tab
// metadata once, creates any missing tabs in a single batchUpdate, then writes each tab — so a
// 15-tab customer costs about 1 read + 1 batchUpdate + 15 writes, comfortably inside Google's
// 60 writes/minute/user quota for what's an infrequent, manually-triggered action.
// itemGroups: [{ tabName, variants: [{ title, header, rows }] }]
// summary: { rows: [{ item, quantity }] } (optional)
async function pushCustomerSheet(spreadsheetId, itemGroups, summary) {
  const sheets = getSheetsClient();
  const tabPlans = (Array.isArray(itemGroups) ? itemGroups : []).map(g => ({
    tabName: sanitizeTabName(g.tabName || 'Sheet'),
    grid: buildSideBySideGrid(g.variants || []),
  }));
  if (summary && Array.isArray(summary.rows)) {
    tabPlans.push({ tabName: 'summary', grid: buildSummaryGrid(summary.rows) });
  }
  const existingMeta = await getSheetMeta(sheets, spreadsheetId);
  const missing = Array.from(new Set(tabPlans.filter(p => !existingMeta[p.tabName]).map(p => p.tabName)));
  if (missing.length) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: missing.map(title => ({ addSheet: { properties: { title } } })) },
    });
  }
  const results = [];
  for (const plan of tabPlans) {
    try {
      const gridSize = existingMeta[plan.tabName] || { rowCount: 1000, columnCount: 26 };
      await writeValuesClearingStale(sheets, spreadsheetId, plan.tabName, plan.grid, gridSize);
      results.push({ tab: plan.tabName, ok: true });
    } catch (e) {
      results.push({ tab: plan.tabName, ok: false, error: e.message || 'unknown error' });
    }
  }
  return results;
}

// --- HTTP handlers ---
async function pushCustomerSheetHandler(req, res) {
  try {
    const { spreadsheetId, itemGroups, summary } = req.body || {};
    if (!spreadsheetId || !String(spreadsheetId).trim()) {
      return res.status(400).json({ error: 'spreadsheetId is required.' });
    }
    if (!Array.isArray(itemGroups)) {
      return res.status(400).json({ error: 'itemGroups must be an array.' });
    }
    const results = await pushCustomerSheet(String(spreadsheetId).trim(), itemGroups, summary);
    const failed = results.filter(r => !r.ok);
    res.status(failed.length ? 502 : 200).json({ ok: failed.length === 0, results });
  } catch (e) {
    console.error('Customer sheet push error:', e);
    res.status(502).json({ error: `Could not push to Google Sheets: ${e.message || 'unknown error'}` });
  }
}

async function getTab(req, res) {
  try {
    const rows = await readTab(req.params.tab);
    res.json({ rows });
  } catch (e) {
    console.error(`Sheets read error [${req.params.tab}]:`, e);
    res.status(502).json({ error: `Could not read from Google Sheets: ${e.message || 'unknown error'}` });
  }
}

async function putTab(req, res) {
  try {
    const rows = (req.body && Array.isArray(req.body.rows)) ? req.body.rows : [];
    await writeTab(req.params.tab, rows);
    res.json({ ok: true, count: rows.length });
  } catch (e) {
    console.error(`Sheets write error [${req.params.tab}]:`, e);
    res.status(502).json({ error: `Could not write to Google Sheets: ${e.message || 'unknown error'}` });
  }
}

async function putBlocksTab(req, res) {
  try {
    const blocks = (req.body && Array.isArray(req.body.blocks)) ? req.body.blocks : [];
    await writeBlocksTab(req.params.tab, blocks);
    res.json({ ok: true, count: blocks.length });
  } catch (e) {
    console.error(`Sheets blocks-write error [${req.params.tab}]:`, e);
    res.status(502).json({ error: `Could not write to Google Sheets: ${e.message || 'unknown error'}` });
  }
}

module.exports = { readTab, writeTab, writeBlocksTab, getTab, putTab, putBlocksTab, pushCustomerSheetHandler };