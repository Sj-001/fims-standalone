// Generic "array of flat objects" <-> Google Sheets tab bridge. Each register (production,
// customerDispatch, etc.) becomes one tab, with a header row built from the union of keys across
// all rows and one spreadsheet row per array item — so opening the Sheet directly shows real,
// readable columns (Date, Description, Pieces, ...), not opaque JSON blobs. This is generic on
// purpose: it doesn't hardcode each register's column list, so adding a new field in the frontend
// doesn't require a matching backend change.
//
// Auth: uses a Google service account (see README) via a JWT client — no OAuth login flow, no
// browser consent screen, safe for a fully unattended backend.

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

async function listTabTitles(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties.title' });
  return (meta.data.sheets || []).map(s => s.properties.title);
}

async function ensureTabExists(sheets, spreadsheetId, tabName) {
  const titles = await listTabTitles(sheets, spreadsheetId);
  if (titles.includes(tabName)) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] },
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
async function writeTab(tabKey, rows) {
  const sheets = getSheetsClient();
  const spreadsheetId = getSpreadsheetId();
  const tabName = sanitizeTabName(tabKey);
  await ensureTabExists(sheets, spreadsheetId, tabName);
  await sheets.spreadsheets.values.clear({ spreadsheetId, range: tabName });
  if (!Array.isArray(rows) || !rows.length) return;
  // Header = union of keys across all rows, 'id' pinned first if present, rest in first-seen order.
  const keySet = new Set();
  rows.forEach(r => Object.keys(r || {}).forEach(k => keySet.add(k)));
  const keys = Array.from(keySet);
  const header = keys.includes('id') ? ['id', ...keys.filter(k => k !== 'id')] : keys;
  const values = [header, ...rows.map(r => header.map(k => (r[k] === undefined || r[k] === null) ? '' : String(r[k])))];
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${tabName}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values },
  });
}

// --- HTTP handlers ---
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

module.exports = { readTab, writeTab, getTab, putTab };
