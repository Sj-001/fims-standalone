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

// Highlight colors for the "what changed since last push" formatting below. HIGHLIGHT is a warm
// tint (matches this app's own accent color, so it doesn't look like a jarring alert) applied to
// any cell whose value is new or different from what was in the sheet before this push; NORMAL is
// plain white, applied to everything else FIRST so last push's highlights don't linger forever.
const HIGHLIGHT_COLOR = { red: 0.941, green: 0.886, blue: 0.769 };
const NORMAL_COLOR = { red: 1, green: 1, blue: 1 };

// Diffs a tab's about-to-be-written grid against what was actually in that tab before this push,
// and returns the batchUpdate sub-requests that (a) reset the tab's whole used range back to plain
// white, then (b) highlight only the cells that are new or changed. Cells that are blank in the new
// grid are never highlighted even if they differ from before (a value being removed isn't "new
// data" — in practice ledgers only ever grow, they don't retroactively blank out old rows). Adjacent
// changed cells in the same row are merged into one range instead of one request per cell, so a
// freshly appended ledger row (the overwhelmingly common case) costs one highlight request, not five.
function buildHighlightRequestsForTab(sheetId, grid, prevGrid) {
  if (sheetId === undefined || sheetId === null) return [];
  const rows = Array.isArray(grid) ? grid : [];
  const prev = Array.isArray(prevGrid) ? prevGrid : [];
  const totalRows = Math.max(rows.length, prev.length, 1);
  const totalCols = Math.max(1, ...rows.map(r => (r || []).length), ...prev.map(r => (r || []).length));
  const requests = [{
    repeatCell: {
      range: { sheetId, startRowIndex: 0, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: totalCols },
      cell: { userEnteredFormat: { backgroundColor: NORMAL_COLOR } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  }];
  const cellStr = (v) => (v === undefined || v === null) ? '' : String(v);
  for (let r = 0; r < totalRows; r++) {
    const oldRow = prev[r] || [];
    const newRow = rows[r] || [];
    let runStart = -1;
    for (let c = 0; c <= totalCols; c++) {
      const changed = c < totalCols && cellStr(newRow[c]) !== '' && cellStr(newRow[c]) !== cellStr(oldRow[c]);
      if (changed && runStart === -1) runStart = c;
      if (!changed && runStart !== -1) {
        requests.push({
          repeatCell: {
            range: { sheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: runStart, endColumnIndex: c },
            cell: { userEnteredFormat: { backgroundColor: HIGHLIGHT_COLOR } },
            fields: 'userEnteredFormat.backgroundColor',
          },
        });
        runStart = -1;
      }
    }
  }
  return requests;
}

// Pushes every item-group tab plus a summary tab to an external spreadsheet in one go, then
// highlights whatever changed. Sequence: read tab metadata once, create any missing tabs in a single
// batchUpdate (capturing their new numeric sheetId from the response), read every EXISTING tab's
// current values in one batchGet (this is the "before" snapshot the highlight diff needs), write
// each tab's new values, then apply ALL tabs' highlight formatting in one final batchUpdate. That's
// 4 API calls total plus one values.update per tab — still just a handful of calls for an infrequent,
// manually-triggered action, comfortably inside Google's 60 writes/minute/user quota.
// itemGroups: [{ tabName, variants: [{ title, header, rows }] }]
// summary: { rows: [{ item, quantity }] } (optional)
// Real customer spreadsheets routinely have tab names with trailing spaces or inconsistent casing
// ("hit & run ", "summary ", "SUMMARY", "T GEL ") — verified directly against BINDAL STOCK.xlsx,
// DIAMOND.xlsx, and anmol stock dec 22.xlsx, where MOST category tabs have a trailing space. Matching
// tab names with plain exact-string equality against a trimmed/sanitized proposed name misses those
// real tabs entirely, so the app used to think they didn't exist and created a second, near-identical
// tab (e.g. "hit & run" alongside the real "hit & run ") instead of writing into the one already
// there. This normalizes for comparison ONLY — the real tab's exact original name/spacing is always
// what gets written to and returned, never a normalized or re-cased version of it.
const normalizeTabKey = (name) => String(name || '').trim().toLowerCase();

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

async function pushCustomerSheet(spreadsheetId, itemGroups, summary) {
  const sheets = getSheetsClient();
  let tabPlans = (Array.isArray(itemGroups) ? itemGroups : []).map(g => ({
    tabName: g.tabName || 'Sheet',
    grid: buildSideBySideGrid(g.variants || []),
  }));
  if (summary && Array.isArray(summary.rows)) {
    tabPlans.push({ tabName: 'summary', grid: buildSummaryGrid(summary.rows) });
  }
  const existingMeta = await getSheetMeta(sheets, spreadsheetId);
  let { resolved, missing } = resolveTabPlansAgainstExisting(tabPlans, existingMeta);
  tabPlans = resolved;
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
  // "Before" snapshot for the highlight diff — only tabs that already had data are worth reading;
  // a tab we just created above is empty, so its previous grid is simply []. UNFORMATTED_VALUE
  // matters here: with the default FORMATTED_VALUE, a cell showing "1,220" wouldn't string-match the
  // raw new value "1220" and would be wrongly flagged as changed every single push.
  const preExistingTabNames = tabPlans.map(p => p.tabName).filter(t => existingMeta[t] && !missing.includes(t));
  const previousValuesByTab = {};
  if (preExistingTabNames.length) {
    try {
      const readResp = await sheets.spreadsheets.values.batchGet({
        spreadsheetId, ranges: preExistingTabNames, valueRenderOption: 'UNFORMATTED_VALUE',
      });
      (readResp.data.valueRanges || []).forEach((vr, i) => { previousValuesByTab[preExistingTabNames[i]] = vr.values || []; });
    } catch (e) {
      console.error('Customer sheet pre-push read (for highlight diff) failed — pushing without highlights:', e);
    }
  }
  const results = [];
  const formatRequests = [];
  for (const plan of tabPlans) {
    try {
      const gridSize = existingMeta[plan.tabName] || { rowCount: 1000, columnCount: 26 };
      await writeValuesClearingStale(sheets, spreadsheetId, plan.tabName, plan.grid, gridSize);
      results.push({ tab: plan.tabName, ok: true });
      const sheetId = existingMeta[plan.tabName] && existingMeta[plan.tabName].sheetId;
      formatRequests.push(...buildHighlightRequestsForTab(sheetId, plan.grid, previousValuesByTab[plan.tabName] || []));
    } catch (e) {
      results.push({ tab: plan.tabName, ok: false, error: friendlyGoogleError(e) });
    }
  }
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

module.exports = { readTab, writeTab, writeBlocksTab, getTab, putTab, putBlocksTab, pushCustomerSheetHandler, importCustomerSheetHandler, getServiceAccountEmailHandler };