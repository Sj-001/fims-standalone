import React, { useState, useEffect, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload, Image as ImageIcon, Package, Boxes, Search, Truck, ClipboardList,
  FileSpreadsheet, Download, CheckCircle2, XCircle, Trash2, Loader2,
  AlertCircle, LayoutDashboard, FileText, Archive, ListChecks, Plus, RefreshCw, Link2, Info, Check,
  ChevronLeft, ChevronRight
} from 'lucide-react';
/* ============================== helpers ============================== */
const num = (v) => {
  const n = parseFloat(String(v ?? '').replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
};
const genId = () => `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
async function resizeImageToBase64(file, maxDim = 1500, quality = 0.85) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) { height = Math.round((height * maxDim) / width); width = maxDim; }
          else { width = Math.round((width * maxDim) / height); height = maxDim; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width; canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);
        resolve({ dataUrl: canvas.toDataURL('image/jpeg', quality), base64: canvas.toDataURL('image/jpeg', quality).split(',')[1] });
      };
      img.onerror = () => reject(new Error('IMAGE_DECODE_FAILED'));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error('FILE_READ_FAILED'));
    reader.readAsDataURL(file);
  });
}
let pdfjsLoadPromise = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (pdfjsLoadPromise) return pdfjsLoadPromise;
  pdfjsLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      try {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        resolve(window.pdfjsLib);
      } catch (e) { reject(e); }
    };
    script.onerror = () => reject(new Error('PDFJS_LOAD_FAILED'));
    document.body.appendChild(script);
  });
  return pdfjsLoadPromise;
}
function canvasToScaled(canvas, maxDim, quality) {
  let { width, height } = canvas;
  let out = canvas;
  if (width > maxDim || height > maxDim) {
    let w = width, h = height;
    if (w > h) { h = Math.round((h * maxDim) / w); w = maxDim; } else { w = Math.round((w * maxDim) / h); h = maxDim; }
    const resized = document.createElement('canvas');
    resized.width = w; resized.height = h;
    resized.getContext('2d').drawImage(canvas, 0, 0, w, h);
    out = resized;
  }
  const dataUrl = out.toDataURL('image/jpeg', quality);
  return { dataUrl, base64: dataUrl.split(',')[1] };
}
// Detects which printed copy a PDF page is (Original/Duplicate/Triplicate/Extra/e-Way Bill) by reading
// the PDF's own text layer — free, instant, no API call. Only works for real digital PDFs (which is
// exactly how dispatch bills come in); photographed JPG/PNG pages have no text layer, so this always
// returns 'unknown' for those and nothing changes for them. Used to skip sending duplicate/e-Way Bill
// pages to the vision API at all for dispatch-bill uploads, since sending all of them was burning
// through the rate limit on a single multi-copy bill for pages we were always going to discard anyway.
function detectCopyLabel(text) {
  const t = (text || '').replace(/\s+/g, ' ');
  if (/e-?way\s*bill\s*details/i.test(t)) return 'ewaybill';
  if (/original\s*for\s*recipient/i.test(t)) return 'original';
  if (/duplicate\s*for\s*transporter/i.test(t)) return 'duplicate';
  if (/triplicate\s*for\s*supplier/i.test(t)) return 'triplicate';
  if (/extra\s*copy/i.test(t)) return 'extra';
  if (/quadruplicate/i.test(t)) return 'extra';
  return 'unknown';
}
async function pdfFileToPages(file, maxDim = 1500, quality = 0.85) {
  const pdfjsLib = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width; canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    const { dataUrl, base64 } = canvasToScaled(canvas, maxDim, quality);
    let copyLabel = 'unknown';
    try {
      const textContent = await page.getTextContent();
      copyLabel = detectCopyLabel(textContent.items.map(it => it.str).join(' '));
    } catch (e) { /* no text layer (scanned PDF) — leave as unknown, nothing gets auto-skipped */ }
    pages.push({ dataUrl, base64, page: i, copyLabel });
  }
  return pages;
}
// Triggers a real file download of a workbook. Only ever called from a direct, explicit click on the
// "Download .xlsx" button inside the export modal (never automatically) — each call is therefore its
// own fresh user gesture, which is what browsers require before they'll reliably allow a download.
// Earlier versions of this app ALSO called window.open() as an automatic "fallback" any time the artifact
// was running inside an iframe — which, inside Claude.ai, is always. That meant every single export
// silently fired two download attempts back-to-back (the <a download> click, then a second one via
// window.open), which is exactly the kind of back-to-back automatic download pattern Chrome and other
// browsers throttle after the first one in a session — the second, third, etc. get silently swallowed
// with no error thrown, which is consistent with "the first export I tried worked, later ones didn't."
// That automatic second attempt has been removed; this function now makes exactly one attempt per click.
function downloadWorkbook(wb, filename) {
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  try {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
}
// Tab-separated text pastes into Excel/Google Sheets as real columns — this is the export path that is
// guaranteed to work no matter what the browser does with actual file downloads, because selecting text
// and pressing Ctrl+C is a native browser action, not a scriptable API call that a sandbox can block.
function toTSV(rows, columns) {
  const esc = (v) => String(v ?? '').replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
  const header = columns.map(c => esc(c.label)).join('\t');
  const body = rows.map(r => columns.map(c => esc(r[c.key])).join('\t')).join('\n');
  return rows.length ? `${header}\n${body}` : header;
}
// For a multi-sheet export (e.g. one item per customer-stock sheet), stacks every sheet's title + TSV
// block into ONE pasteable string, separated by a blank line. This exists so a 15-36-item customer
// stock export doesn't require selecting the dropdown and copy-pasting once per item — one paste gets
// everything, landing as stacked title/header/data blocks in a single Excel sheet (not separate tabs,
// but every row of every item is present in one action). The per-item dropdown view is kept alongside
// this for anyone who specifically wants to paste one item into its own existing tab.
function toMultiBlockTSV(sheets) {
  return sheets.map(s => `${s.name}\n${toTSV(s.rows, s.columns)}`).join('\n\n');
}
// Builds a valid, UNIQUE Excel worksheet name from arbitrary text: strips the characters Excel forbids
// in sheet names (\ / ? * [ ] :), enforces the 31-character limit, falls back to "Sheet" if nothing is
// left, and de-dupes case-insensitively (Excel treats "Kaju Bake" and "kaju bake" as the same name) by
// appending _2, _3, ... Tested against duplicate, empty, overlong, unicode, and symbol-only names.
function sanitizeSheetName(rawName, usedNamesLower) {
  let base = String(rawName || '').replace(/[\\/*?:[\]]/g, '').trim();
  if (!base) base = 'Sheet';
  base = base.slice(0, 31);
  let candidate = base;
  let n = 2;
  while (usedNamesLower.has(candidate.toLowerCase())) {
    const suffix = `_${n}`;
    candidate = base.slice(0, Math.max(1, 31 - suffix.length)) + suffix;
    n++;
  }
  usedNamesLower.add(candidate.toLowerCase());
  return candidate;
}
// A worksheet with a title line above the table (used for per-customer stock ledgers). Built as one
// array-of-arrays so the title never overwrites real header cells — an earlier version wrote the title
// into cell A1 with sheet_add_aoa AFTER json_to_sheet had already put "Date" there, silently clobbering
// the header of every stock export.
function buildTitledSheet(title, rows, columns) {
  const header = columns.map(c => c.label);
  const body = rows.map(r => columns.map(c => r[c.key] ?? ''));
  return XLSX.utils.aoa_to_sheet([[title], [], header, ...body]);
}
function buildTableSheet(rows, columns) {
  const data = rows.length ? rows.map(r => { const o = {}; columns.forEach(c => { o[c.label] = r[c.key]; }); return o; }) : [{}];
  return XLSX.utils.json_to_sheet(data);
}
// Builds the actual .xlsx workbook from the modal's sheet list, on demand, only when the person clicks
// "Download .xlsx" — kept separate from the copy-paste view (which is built instantly and can't fail)
// so a workbook-building problem never blocks the one export path that's guaranteed to work.
function buildWorkbookFromSheets(sheets) {
  const wb = XLSX.utils.book_new();
  const usedNamesLower = new Set();
  sheets.forEach(s => {
    const ws = s.kind === 'titled' ? buildTitledSheet(s.title || s.name, s.rows, s.columns) : buildTableSheet(s.rows, s.columns);
    const safeName = sanitizeSheetName(s.name, usedNamesLower);
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  });
  return wb;
}
function repairTruncatedJson(text) {
  // Best-effort recovery when a response got cut off mid-object: trim back to the
  // last fully-closed brace/bracket, then close whatever's still left open.
  const lastGood = Math.max(text.lastIndexOf('}'), text.lastIndexOf(']'));
  if (lastGood === -1) throw new Error('UNPARSEABLE');
  let s = text.slice(0, lastGood + 1);
  const stack = [];
  for (const ch of s) {
    if (ch === '{' || ch === '[') stack.push(ch);
    else if (ch === '}' && stack[stack.length - 1] === '{') stack.pop();
    else if (ch === ']' && stack[stack.length - 1] === '[') stack.pop();
  }
  while (stack.length) s += (stack.pop() === '{' ? '}' : ']');
  s = s.replace(/,(\s*[}\]])/g, '$1');
  return JSON.parse(s);
}
// Standalone-hosting version: extraction always goes through OUR OWN backend at /api/extract,
// which holds the real Anthropic API key server-side and forwards to Anthropic. The browser never
// sees the key at all — there's no "own key vs free proxy" branch anymore because there's only one
// path now, and it's never rate-limited by a shared quota (the whole point of this migration).
// A 401 here means the login session expired; dispatching 'fims-unauthorized' sends the app back
// to the login screen instead of showing a confusing extraction error.
async function callClaudeExtract(systemPrompt, base64Image, signal) {
  const res = await fetch('/api/extract', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    credentials: 'include',
    signal,
    body: JSON.stringify({ systemPrompt, base64Image }),
  });
  if (res.status === 401) {
    window.dispatchEvent(new Event('fims-unauthorized'));
    throw new Error('EXTRACT_HTTP_401: session expired, please log in again');
  }
  if (!res.ok) {
    const bodyText = await res.text().catch(() => '');
    const err = new Error(`EXTRACT_HTTP_${res.status}: ${bodyText.slice(0, 300)}`);
    err.status = res.status;
    // surface a Retry-After hint if the server sent one — used to size the backoff wait
    const retryAfter = res.headers && res.headers.get ? res.headers.get('retry-after') : null;
    if (retryAfter) err.retryAfterMs = (parseFloat(retryAfter) || 0) * 1000;
    throw err;
  }
  const data = await res.json();
  const block = (data.content || []).find(b => b.type === 'text');
  if (!block) throw new Error('EXTRACT_EMPTY: no text block in response');
  let clean = block.text.trim();
  clean = clean.replace(/^```(json)?/i, '').replace(/```$/,'').trim();
  const firstBrace = Math.min(
    ...['[', '{'].map(ch => { const i = clean.indexOf(ch); return i === -1 ? Infinity : i; })
  );
  if (firstBrace > 0 && firstBrace !== Infinity) clean = clean.slice(firstBrace);
  // Anthropic's own signal for "I was cut off by max_tokens before I finished" — confirmed directly:
  // a dense production-register page came back with only 1 of ~33 real rows, another with 23 of 31,
  // both silently "successful" because the truncated JSON still happened to parse (or repair-parse)
  // cleanly. Trusting stop_reason here, rather than only reacting to a JSON.parse failure, is what
  // catches BOTH failure shapes: a response cut off cleanly between two complete items (parses fine,
  // just missing everything after) and one cut off mid-item (fails to parse, needs repairTruncatedJson).
  const truncated = data.stop_reason === 'max_tokens';
  try {
    return { raw: JSON.parse(clean), truncated };
  } catch (parseErr) {
    // response was cut off badly enough that even a well-formed prefix doesn't exist — needing the
    // salvage repair is itself evidence of truncation, regardless of what stop_reason said.
    return { raw: repairTruncatedJson(clean), truncated: true };
  }
}
/* ============================== storage ============================== */
const STORAGE_KEYS = {
  rawMaterialIn: 'fims_raw_material_in',
  consumption: 'fims_consumption',
  production: 'fims_production',
  customerDispatch: 'fims_customer_dispatch',
  daburSpecs: 'fims_dabur_specs',
  daburPO: 'fims_dabur_po',
  daburDispatch: 'fims_dabur_dispatch',
  // A local mirror of every customer's real Sheet content (every ledger row across every block/tab),
  // refreshed whenever a Sheet ID is imported or pushed to — see confirmSheetImport/pushCustomerSheetNow.
  // Stored as just another flat register (same generic tab-bridge everything else here uses), so global
  // search can look through every customer's real Sheet data locally, without a live API call per search.
  customerSheetsMirror: 'fims_customer_sheets_mirror',
};
const CATALOG_KEY = 'fims_product_catalog';
// Exact item names, one per customer — used to correct handwriting misreads during extraction (e.g.
// "g" vs "9", "&" vs "8"). Starts EMPTY on purpose: every customer's items get added by importing
// their Google Sheet ID (see the Customer Sheets tab) rather than being hardcoded here, so this app
// works for whichever customers you actually have, not just the three it happened to be built
// against. `sheetGroup` is the tab name this item lands under in that customer's own Google Sheet —
// several variants of the same base item (e.g. "IT 500 Lid" and "IT 500 Container") share one tab.
const catalogEntries = (customer, pairs) => pairs.map(([item, sheetGroup]) => ({ id: genId(), customer, item, sheetGroup }));
const DEFAULT_PRODUCT_CATALOG = [];
const CUSTOMER_MAPPING_KEY = 'fims_customer_mapping';
// Which external Google Sheet ID each customer's own stock sheet lives at (see the "Customer
// Sheets" tab) — a separate, customer-owned spreadsheet, never a tab on the main FIMS sheet. Also
// doubles as the "already imported" registry for the Customer Sheets CRUD workflow (tracks when
// each customer's sheet was last imported from / pushed to).
const CUSTOMER_SHEET_IDS_KEY = 'fims_customer_sheet_ids';
// Maps the raw text a dispatch bill's Party field (or a production row's bracketed customer) was
// actually written as, to the one real known customer it means — e.g. "Bindal technopolymer pvt.
// ltd." really is "BINDAL STOCK 1.08.26". Checked before the fuzzy substring-containment guess in
// matchCustomer, so once taught, a legal-name variant resolves straight to the right customer
// instead of spinning up a disconnected duplicate with no Sheet ID and no visible data anywhere.
const CUSTOMER_NAME_ALIASES_KEY = 'fims_customer_name_aliases';
// Lower-priority fallback keywords — only used when a ledger entry doesn't exactly match one of the
// catalog's known item names above (e.g. a pack-size variant that isn't in the catalog yet). Starts
// empty along with the catalog above; importing a customer sheet adds both an exact-name rule per
// item and a broader fallback rule automatically (see confirmSheetImport).
const FALLBACK_CUSTOMER_MAPPING = [];
// Word-level shorthand a person can teach the app themselves ("J" -> "Jumbo", "CONT" -> "Container")
// instead of asking for a code change every time a new abbreviation shows up on a handwritten sheet.
// Applied in two places: (1) fed into the extraction prompt so Claude writes the expanded form
// directly into "description" when it recognizes one of these, and (2) run as a text substitution
// inside normalizeVariant/normalizeForCatalogMatch below, so even an entry that DIDN'T get expanded
// at extraction time (typed by hand, or extracted before this abbreviation was taught) still matches
// the right catalog item / customer-stock variant. Starts empty — nothing is assumed.
const ABBREVIATIONS_KEY = 'fims_abbreviations';
const DEFAULT_ABBREVIATIONS = [];
function buildAbbreviationsText(abbreviations) {
  if (!abbreviations || !abbreviations.length) return '  (none taught yet)';
  return abbreviations.map(a => `  "${a.short}" = "${a.long}"`).join('\n');
}
// The mapping actually used: one EXACT-item-name rule per catalog entry (checked first, most reliable),
// then the broader fallback keywords above for anything that doesn't hit an exact match.
const DEFAULT_CUSTOMER_MAPPING = [
  ...DEFAULT_PRODUCT_CATALOG.map(c => ({ id: genId(), keyword: c.item.toLowerCase(), customer: c.customer })),
  ...FALLBACK_CUSTOMER_MAPPING,
];
function buildCatalogText(catalog) {
  if (!catalog || !catalog.length) return '  (no known items yet)';
  const byCustomer = {};
  catalog.forEach(c => { (byCustomer[c.customer] = byCustomer[c.customer] || []).push(c.item); });
  return Object.entries(byCustomer).map(([customer, items]) => `  ${customer}: ${items.join(', ')}`).join('\n');
}
// A row coming back from Google Sheets is always made of strings — every cell, no matter what was
// written into it. That's fine for most fields (the `num()` helper already tolerates numeric
// strings), but `stockConfirmed` is checked as a real boolean (`!row.stockConfirmed`), and the
// string `"false"` is truthy in JS — so without this fix, every row would look permanently
// "confirmed" after a single round trip through the Sheet. This coerces just that one known
// boolean field back to a real boolean on the way in; everything else passes through untouched.
function coerceStorageValue(row) {
  if (!row || typeof row !== 'object') return row;
  const out = { ...row };
  if ('stockConfirmed' in out) out.stockConfirmed = out.stockConfirmed === true || out.stockConfirmed === 'true';
  return out;
}
// window.storage shim: the original Claude.ai artifact persisted data through a built-in
// `window.storage.get/set/delete(key, ephemeral)` API baked into the artifact sandbox. That API
// doesn't exist outside Claude.ai, so this shim reimplements the same interface on top of our own
// backend's /api/sheets/:tab endpoint (which in turn reads/writes a real Google Sheet tab) — every
// other function in this file that calls window.storage needed zero further changes.
// The backend's sheets endpoint speaks in arrays of flat row objects (so the Sheet stays a normal,
// readable spreadsheet with real columns) — this shim bridges that to the original blob-string
// interface, JSON-stringifying/parsing on the way through, and coercing row types on read.
window.storage = {
  async get(key) {
    const res = await fetch(`/api/sheets/${encodeURIComponent(key)}`, { credentials: 'include' });
    if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return null; }
    if (!res.ok) throw new Error(`Sheets read failed: HTTP ${res.status}`);
    const data = await res.json();
    const rows = Array.isArray(data.rows) ? data.rows.map(coerceStorageValue) : [];
    if (!rows.length) return null;
    return { value: JSON.stringify(rows) };
  },
  async set(key, value) {
    let rows = [];
    try { rows = JSON.parse(value); } catch (e) { rows = []; }
    if (!Array.isArray(rows)) rows = [];
    const res = await fetch(`/api/sheets/${encodeURIComponent(key)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'include',
      body: JSON.stringify({ rows }),
    });
    if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return; }
    if (!res.ok) throw new Error(`Sheets write failed: HTTP ${res.status}`);
  },
  async delete(key) {
    // Deleting = writing an empty array, which clears the tab's contents (the tab itself stays).
    await window.storage.set(key, '[]');
  },
};
async function loadRegister(key) {
  try {
    const r = await window.storage.get(key, false);
    if (r && r.value) return JSON.parse(r.value);
    return [];
  } catch (e) { return []; }
}
async function saveRegister(key, rows) {
  try { await window.storage.set(key, JSON.stringify(rows), false); } catch (e) { /* noop */ }
}
const TRAINING_KEY = 'fims_training_examples';
const MAX_EXAMPLES_STORED = 15;
const MAX_EXAMPLES_IN_PROMPT = 6;
// Rate-limit handling: how long to wait before automatically retrying a page that got rate-limited,
// growing each time in case the window hasn't reset yet. After exhausting these, we stop and let the
// person resume manually via "Retry remaining" — that way we never spin forever silently.
const RATE_LIMIT_BACKOFFS_MS = [20000, 40000, 60000];
// A 502/503/504 here comes from Render's own proxy in front of OUR backend, never from Anthropic —
// it means the backend itself was momentarily unreachable (mid-restart/redeploy, or waking back up
// from an idle sleep on plans that do that), not a problem with the image or the request. That kind
// of gap is usually much shorter than a rate-limit window, hence its own, snappier schedule.
const SERVER_ERROR_BACKOFFS_MS = [5000, 10000, 20000];
// Pacing gap between consecutive requests in a batch. Our backend uses a dedicated, spend-capped
// Anthropic API key (1,000 requests/minute standard limit) instead of Claude.ai's shared free quota,
// so this just needs to avoid firing a literal burst in the same instant — no long throttling needed.
const BATCH_REQUEST_GAP_MS = 400;
const sleep = (ms, signal) => new Promise((resolve, reject) => {
  const timer = setTimeout(resolve, ms);
  if (!signal) return;
  if (signal.aborted) { clearTimeout(timer); reject(new DOMException('Cancelled', 'AbortError')); return; }
  signal.addEventListener('abort', () => { clearTimeout(timer); reject(new DOMException('Cancelled', 'AbortError')); }, { once: true });
});
const isCancelled = (e) => e && e.name === 'AbortError';
// Training examples are stored as one JS object (keyed by document type), not an array of flat rows
// like every other register — so they don't fit the generic array-shaped window.storage shim above.
// Instead this talks to the same /api/sheets/:tab backend endpoint directly, wrapping the whole
// object as a single row with one 'blob' column. It'll show up in the Sheet as one JSON cell rather
// than readable columns, which is fine — these are internal correction examples, not data anyone
// needs to read or edit directly in the spreadsheet.
async function loadTraining() {
  try {
    const res = await fetch(`/api/sheets/${TRAINING_KEY}`, { credentials: 'include' });
    if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return {}; }
    if (!res.ok) return {};
    const data = await res.json();
    const row = Array.isArray(data.rows) && data.rows[0];
    if (row && row.blob) return JSON.parse(row.blob);
    return {};
  } catch (e) { return {}; }
}
async function saveTraining(obj) {
  try {
    await fetch(`/api/sheets/${TRAINING_KEY}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'include',
      body: JSON.stringify({ rows: [{ blob: JSON.stringify(obj) }] }),
    });
  } catch (e) { /* noop */ }
}
// ex.after === null means the row was deleted entirely during review, not edited — a person decided
// it should never have been extracted in the first place (e.g. a blank pre-ruled line with no real
// data, a duplicate, a total row). That's a distinct kind of lesson from a field-level misread, so it
// gets its own sentence rather than being described as a "correction" to a value.
function buildPromptWithTraining(basePrompt, examples) {
  if (!examples || !examples.length) return basePrompt;
  const recent = examples.slice(-MAX_EXAMPLES_IN_PROMPT);
  const examplesText = recent.map((ex, i) => {
    if (ex.after === null) {
      return `Example ${i + 1}:\nExtracted (WRONG — a person deleted this row entirely because it should never have been output at all): ${JSON.stringify(ex.before)}\nDo not produce a row like this — recognize the same underlying situation (e.g. a blank line, a duplicate, a total row) and skip it.`;
    }
    return `Example ${i + 1}:\nExtracted (this had a mistake): ${JSON.stringify(ex.before)}\nHuman-corrected (this is right): ${JSON.stringify(ex.after)}`;
  }).join('\n\n');
  return `${basePrompt}
LEARNED CORRECTIONS — a person has corrected real extraction mistakes on this exact document type before. Each item below shows a row as it was first extracted (with a mistake) and how a human corrected it, OR a row a human deleted entirely because it shouldn't have been extracted at all. Study the pattern behind each one — what kind of value was misread, which field it belongs in, when a "row" isn't really a row — and apply that same fix logic to this new document. Do not copy exact values into unrelated rows; only apply the underlying pattern.
${examplesText}`;
}
// Every date the app touches should end up in the Sheet's own dot-separated D.M.YY convention
// (e.g. "16.7.26") — extraction sources disagree though: Production Register OCR tends to read as
// "02/01/2026" (slash), Customer Dispatch Bill extraction reads as "16-Jul-26" (dash + month
// abbreviation), and manually-typed rows are already dot-formatted. Fixing this only at push time
// let three different formats sit mixed together in a live sheet (confirmed directly in Bindal's
// N200 tab) and let same-day duplicate rows slip past the "already in the sheet" check, since a
// duplicate check that compares raw strings never sees "18.07.26" and "18-Jul-26" as the same date.
// This normalizer now runs immediately when a row is shaped from a fresh extraction — see each
// DOCUMENT_TYPES `shape()` below — so a mismatched format is never even created in the first place.
// It's also still called again right before a push (see buildCustomerSheetPayload) as a safety net
// for anything that entered the ledger some other way (a manually typed/edited row, for instance).
// Anything unrecognized is left exactly as-is rather than guessed at, so a genuinely new format
// shows up untouched (and visibly odd) instead of silently mangled.
const MONTH_ABBR_TO_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function normalizeDateToDots(raw) {
  const s = String(raw || '').trim();
  if (!s) return s;
  // Dot-separated ("5.1.24", "13.08.24", "23.07.2026") — normalize even when it already matches this
  // pattern, since a leading-zero month/day ("08" not "8") or a 4-digit year ("2026" not "26") would
  // otherwise pass through completely untouched and sit inconsistently next to rows that came out the
  // "clean" way — this is exactly what let "23.07.2026" and "12.4.26" end up side by side in the same
  // register. Number(...) strips the leading zero; the year is always truncated to its last 2 digits.
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (m) {
    const yy = m[3].length > 2 ? m[3].slice(-2) : m[3];
    return `${Number(m[1])}.${Number(m[2])}.${yy}`;
  }
  // Slash-separated DD/MM/YYYY or D/M/YY (what the Production Register's OCR tends to produce).
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const yy = m[3].length > 2 ? m[3].slice(-2) : m[3];
    return `${Number(m[1])}.${Number(m[2])}.${yy}`;
  }
  // Dash + month-abbreviation ("16-Jul-26", what Customer Dispatch Bill extraction produces).
  m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (m) {
    const monthNum = MONTH_ABBR_TO_NUM[m[2].toLowerCase()];
    if (monthNum) {
      const yy = m[3].length > 2 ? m[3].slice(-2) : m[3];
      return `${Number(m[1])}.${monthNum}.${yy}`;
    }
  }
  // Unrecognized format — don't guess, leave untouched.
  return s;
}
// Real chronological ordering for D.M.YY / D.M.YYYY dot-formatted dates. A plain string compare
// (what the ledger sort used before) sorts "11.8.26" before "3.8.26" because '1' < '3' character by
// character, even though August 3rd is chronologically first — confirmed live in the Customer Sheets
// review screen (Hit & Run 30g showed 1.08.26, 11.8.26, 3.8.26, 5.8.26 in that literal order). Opening/
// closing balances are a running total down the ledger, so a misordered row corrupts every balance
// after it. Parses into a zero-padded YYYY-MM-DD sort key instead; anything that isn't a clean D.M.YY
// date sorts to the end rather than crashing or silently miscomparing.
function dateSortKey(raw) {
  const s = String(raw || '').trim();
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (!m) return `9999-99-99::${s}`;
  const day = m[1].padStart(2, '0');
  const month = m[2].padStart(2, '0');
  const year = m[3].length === 2 ? `20${m[3]}` : m[3];
  return `${year}-${month}-${day}`;
}
const REAL_DATE_RE = /^\d{1,2}\.\d{1,2}\.\d{2,4}$/;
// Rows genuinely representing the SAME real-world line, extracted twice (the same page uploaded twice,
// or included in two overlapping photos), should never both land in a register — see addRows, which
// uses this to skip a newly-extracted row that already exactly matches one already there. Matched by
// every real data field EXCEPT id/confirmation-state (id, stockConfirmed, confirmedCustomer, flagged,
// flagReason — which legitimately differ for what's otherwise the same entry). Deliberately EXACT, not
// fuzzy: a looser match (say, just date+description) risks silently dropping a genuinely different row
// that happens to share those two fields, which is a far worse failure than occasionally missing a
// duplicate whose OCR reading varied slightly between two extractions of the same page.
const DEDUP_EXCLUDE_FIELDS = new Set(['id', 'stockConfirmed', 'confirmedCustomer', 'flagged', 'flagReason']);
function rowDedupKey(row) {
  return Object.keys(row).filter(k => !DEDUP_EXCLUDE_FIELDS.has(k)).sort()
    .map(k => `${k}=${typeof row[k] === 'number' ? row[k] : String(row[k] ?? '').trim().toLowerCase()}`)
    .join('|');
}
// Safety net for ditto marks (a tick, quote mark, "11", "//", "do", a dash, etc. — handwritten
// shorthand for "same date as the row above"). The extraction prompt already asks Claude to resolve
// these itself, but that's a probabilistic instruction — this is a deterministic backstop that runs
// on every extraction regardless of whether the model actually followed it. Any row whose date, after
// normalizeDateToDots, still isn't a real-looking D.M.YY date (i.e. the ditto symbol came through
// literally, or the field was left blank) gets filled in from the nearest row ABOVE it that does have
// a real date — exactly what a ditto mark means on the page. The very first row can't borrow from
// anything above it, so an unresolved ditto there is left as-is (visibly wrong, not silently guessed).
function fillDittoDates(rows) {
  let lastDate = '';
  rows.forEach(r => {
    if (REAL_DATE_RE.test((r.date || '').trim())) { lastDate = r.date; }
    else if (lastDate) { r.date = lastDate; }
  });
  return rows;
}
/* ============================== document type configs ============================== */
const DOCUMENT_TYPES = [
  {
    key: 'mill_slip',
    label: 'Mill Packing Slip (raw material inward)',
    hint: 'Printed packing slip from a paper mill listing cardboard reels — e.g. Ashoka Pulp & Paper, Hanumant Paper Mills, Sardhana Papers, Shree Sidhbali.',
    register: 'rawMaterialIn',
    systemPrompt: `You read printed packing slips from paper mills that supply cardboard reels to a corrugated box factory. Extract every reel line item on the slip — there may be 2 to 80+ reels, do not stop early, and do not include "Total"/"Grand Total" summary rows. Return ONLY one JSON object in this exact shape, with the mill name and date given ONCE at the top level, not repeated per item:
{"mill":"mill or company name","date":"slip date, DD/MM/YYYY","items":[{"reel_no":"reel number as string","size":"size only, e.g. 55 or 59.5","unit":"the size's unit of measure as printed, e.g. Inch or CM — leave empty string only if truly not shown anywhere on the slip","gsm":"GSM only, e.g. 140","bf":"BF only, e.g. 18","shade":"shade code if present, e.g. NS or GY, else empty string","weight_kg":number}]}
IMPORTANT — column format varies between mills, read carefully:
- Some slips have separate BF / GSM / SIZE / Unit columns already split out — just copy each value into its matching field. If there's a single "Unit" column applying to all rows, use that same value for every item.
- Other slips (e.g. Ashoka Pulp & Paper) print a single combined column, sometimes mislabeled "SIZE * GSM", where each value is actually three numbers joined by "*" in the order Size*GSM*BF — for example "34*180*18" means size=34, gsm=180, bf=18, and "59.5*140*18" means size=59.5, gsm=140, bf=18. Do NOT put the whole compound string into the "size" field — split it into its three separate size/gsm/bf values. Do not assume which number is which; the pattern is always Size, then GSM, then BF, in that left-to-right order.
- SHADE: some slips have their own dedicated "Shade" column (e.g. values like GY, NS) — use that value directly. Other slips have no separate shade column, but the GSM value has a letter code stuck to the end of it, e.g. "120NS" — in that case the number is the GSM (120) and the letters are the shade (NS); split them into "gsm":"120" and "shade":"NS". A third pattern: some slips group reels under bold section headings instead of a shade column — sometimes the heading directly names a shade code as one of its parts (e.g. "Kraft Paper - 25 - GY" means every row under it is shade GY, and the "25" there is the BF, not a separate thing to worry about if BF already has its own column), and sometimes the heading only names a paper type/colour with no explicit code (e.g. "BR Golden Yellow" or "Golden Yellow" means shade GY, "Kraft Paper" alone with no code means shade NS/Natural Shade). Either way, apply the heading's shade to every row listed under it until the next heading appears. Only do this when you are confident in the mapping — if a heading is unfamiliar or ambiguous, leave "shade" as an empty string rather than guessing. If there is genuinely no shade information anywhere for a row (no column, no code, no heading), also leave "shade" as an empty string.`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      // consumed starts empty on purpose — it's a placeholder for a future feature that matches this
      // reel/slip against consumption reports and records the date it actually got used up; nothing
      // populates it yet, so every fresh row just carries the field so it exists to edit/fill by hand.
      return raw.items.map(it => ({ id: genId(), date: normalizeDateToDots(raw.date || ''), mill: raw.mill || '', reel_no: it.reel_no || '', size: it.size || '', unit: it.unit || '', gsm: it.gsm || '', bf: it.bf || '', shade: it.shade || '', weight_kg: num(it.weight_kg), consumed: '' }));
    },
  },
  {
    key: 'consumption_sheet',
    label: 'Daily Consumption Report (handwritten)',
    hint: 'The handwritten daily sheet workers use to record raw material consumed — columns are usually Shade, GSM, Size, Weight, and Balance Left.',
    register: 'consumption',
    systemPrompt: `You read a handwritten daily raw-material consumption report from a corrugated box factory. The sheet has Hindi column headers. Most rows share one date written once at the top. Return ONLY one JSON object:
{"date":"as written at the top, DD/MM/YYYY","items":[{"sl_no":"this row's own SL.NO./serial number exactly as printed in that column, as a string — see rules below","shade":"shade code — see rules below","gsm":"the ग्रा / GMS column value","size":"the साइज़ / Size column value","weight_consumed":"the वजन / Weight column value, as a number","date_override":"only include this if a specific row has a different date than the header, else omit","flag":"only if you couldn't reliably read this row's weight — e.g. a column looks cut off/missing, or there's a crossed-out number and you're unsure which replacement is correct — a short reason why, else omit"}]}
IMPORTANT:
- SKIP TRULY BLANK LINES. This sheet is pre-ruled with a fixed number of numbered SL.NO. slots/columns, and it is completely normal for a real day's report to only fill in the first several and leave the rest of the printed grid blank. A line that has ONLY a pre-printed SL.NO. with no shade, no GSM, no size, and no weight actually written on it is not a row of data — do not output an item for it. Only extract a line that has genuine handwritten content in it.
- SL.NO IS YOUR ROW-ALIGNMENT ANCHOR. This sheet's rows are frequently laid out as narrow vertical strips (SL.NO increasing left-to-right or in whatever direction the page actually runs) with every strip looking nearly identical (same date, same shade, similar GSM) — the kind of layout where it's extremely easy for a value to silently slide into the wrong row. To prevent that: for EACH row, first locate its SL.NO in the image, then read every other field (shade, GSM, size, weight) from that exact same physical strip/line — never from a neighboring one — before moving to the next SL.NO. After you've extracted every row, check your own output: the sl_no values you produced should appear once each, in the same order they run on the page, with no duplicates and no unexplained gaps. If you notice a duplicate or out-of-sequence SL.NO in what you were about to return, that is a sign a value drifted between rows — go back and re-read those specific rows from the image before finalizing, rather than submitting a duplicated or skipped row.
- The column that looks like it's labeled "S/K" is actually the SHADE column, not a party name or code. Decode its handwritten values: "S.K" or "SK" means shade NS (Sada Kraft / natural shade). "G.Y" or "GY" means shade GY (Golden Yellow). If you see a different value you don't recognize, copy it as written rather than forcing it into NS or GY.
- There is no BF field on this document — do not invent one. What might look like a stray extra column is the Size column.
- IGNORE the टुकड़ा / Tukda column entirely — it is not tracked by this app, do not extract it even if it has a value.
- CROSSED-OUT / CORRECTED VALUES: this is a real working register — a number is sometimes struck through with the corrected number written right next to it in that same row. Use only the number that is NOT crossed out; ignore the struck-through one entirely. This stays within one row — never borrow a number from the row above or below because a cell looks messy.
- MISSING/CUT-OFF COLUMNS: if a column genuinely isn't visible for a row (cropped off the scan, torn page, etc.), do not invent a value by reading unrelated nearby text. Leave that field empty/0 and set "flag" to a short reason instead.
- Interpret unclear handwriting as best you can; if a value is genuinely illegible leave it as an empty string rather than guessing wildly. If the SL.NO itself is illegible, still extract the row's other fields as best you can and set "flag" to say the serial number wasn't legible — do not drop the row.`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      const rows = raw.items.map(it => ({
        id: genId(), sl_no: it.sl_no != null ? String(it.sl_no) : '', date: normalizeDateToDots(it.date_override || raw.date || ''), shade: it.shade || '', size: it.size || '', gsm: it.gsm || '',
        weight_consumed: num(it.weight_consumed),
        flagged: !!(it.flag && String(it.flag).trim()), flagReason: (it.flag || '').trim(),
      }));
      return fillDittoDates(rows);
    },
  },
  {
    key: 'production_sheet',
    label: 'Production Register (handwritten)',
    hint: 'The handwritten daily production register — covers both page styles: the shade/size/GSM/weight/tukda style, and the product-description + quantity style. Extracted into one unified register.',
    register: 'production',
    systemPrompt: `You read a handwritten factory production register from a corrugated box factory. There are TWO different page styles used in this register — figure out which one you're looking at and extract accordingly:
STYLE A — shade/size/GSM style: columns are typically SL.NO. (ignore it), a column commonly headed "S/K" (this records paper SHADE, not a party name — decode 'S.K'/'S/K' as shade NS, 'G.Y' as shade GY, normalize to the 2-letter code, leave blank if unfamiliar rather than guessing), then GMS (GSM), Size (no separate BF column exists here), Weight (kg), and Tukda (count of pieces produced for that row). Most rows share one date at the top of the page. This sheet is pre-ruled with a fixed number of numbered SL.NO. slots, and it's completely normal for a real day to only fill in the first several and leave the rest blank — a line that has ONLY a pre-printed SL.NO. with no shade/GSM/size/weight actually written on it is not a row of data; do not output an item for it.
STYLE B — product ledger style: each line has a DATE and a product DESCRIPTION, plus one or two quantity columns.
- Extract the FULL item name exactly as it appears in the factory's own product catalog below — including the weight and pack count (e.g. "Butter Bake 65g x60", "T50 64g x60"). Do NOT shorten it to just the brand/family word (do not output just "T50" or "Butter Bake" alone) — the weight and pack count are part of the item name, not separate data.
- Known handwriting misreads to correct, using the reference catalog below: the letter "g" (grams) is very often misread as the digit "9" — a pattern like "120g x60" is almost always grams, essentially never "1209 x60"; "&" is often misread as "8" or "5"; "Run" is often misread as "Rum". When a line clearly matches one of the catalog items below (allowing for this kind of misread), use the catalog's exact spelling. If it doesn't resemble anything in the catalog, transcribe your best reading rather than forcing a match.
- Reference catalog of known exact item names (case as shown), grouped by customer for your own matching confidence only — still extract just the item name into "description", not the customer name:
{{PRODUCT_CATALOG}}
  (This list isn't exhaustive — other legitimate products exist too. Only use it to correct obvious misreads of these specific items, never to force an unrelated line into matching one of them.)
- Word abbreviations this factory uses that a person has taught the app — expand these to their full form wherever they appear inside "description" (e.g. write "IT 500 Jumbo Container", not "IT 500 J Cont"), so the same item is always written the same way no matter which shorthand a given page happened to use:
{{ABBREVIATIONS}}
- IMPORTANT: a customer name is sometimes written in brackets right next to the item name, either before or after it — e.g. "(Diamond) Cream Burst 30g x140" or "Cream Burst 30g x140 (Diamond)". Pull that bracketed name OUT into its own "customer_hint" field and do NOT leave it inside "description" — "description" should be just the clean item name with no bracket in it.
- Dates are VERY often written once then repeated below with a ditto mark for every following row that shares that same date — this could be a tick, a quote mark ("), the digits 11, two short slashes (//), the word "do", a dash, or any other short repeat-symbol instead of an actual date. Whenever a row's date cell is anything other than a clearly legible date, treat it as a ditto mark and resolve it to the SAME date as the row directly above it — copy that exact date into "date_override" for the row. Never leave a row's date blank or output the ditto symbol itself just because it wasn't written out in full.
- Some pages have a single "Quantity" column — put that value into "pieces". Other pages have TWO numeric columns side by side for the same row (regardless of what they're headed with, e.g. small letters like S/D, two batch tallies, etc.) — both of these are PRODUCTION quantities, never a dispatch quantity. If a row has a number in only one of the two columns, put that number in "pieces". If a row genuinely has a number in BOTH columns, add them together into "pieces" — do not put the second column's number into "dispatch". The "dispatch" field on this register is rare — only use it if the page has an explicit, separately-written note that some of that day's production was sent out immediately (e.g. an actual word like "dispatch"/"bheja" next to a number), never just because a second quantity column exists.
- CROSSED-OUT / CORRECTED VALUES: this is a real working register, so a number is sometimes struck through (a line drawn through it, or scribbled over) with the CORRECTED number written right next to it (before, after, above, or below the crossed-out one, still within that same row). When you see this, use ONLY the number that is NOT crossed out as the real value — completely ignore the struck-through one. This correction always stays within its own row — never borrow a number from the row above or below just because a cell looks messy or hard to read; re-examine that exact row instead of guessing from a neighboring row.
- MISSING/CUT-OFF COLUMNS: if the quantity column (or any column Style A/B normally has) genuinely is not visible for a row — cropped off the edge of the scan, torn page, or the column simply doesn't exist on this particular page — do NOT invent a number by reading some other nearby text instead (e.g. do not use a pack-size number like "120" from "32g x120", and do not use the word "Pkt"/"Pakt"/"Packet" or any part of it as if it were a quantity). Leave "pieces" as 0 for that row and set "flag" to a short reason like "quantity column not visible in this scan" instead. If this affects the WHOLE page (e.g. the entire quantity column got cropped out of the photo), still extract every row's date/description as normal, just flag each one the same way rather than skipping the page.
For EITHER style: do not include page-total or running-total rows (a lone number with no row content, usually at the bottom of a page or column). Extract every real row on the page, in order, exactly once — do not skip rows and do not repeat any row, even if faint printing or ruling lines make it look duplicated. If an actual customer/party name is written somewhere on the page itself (separate from the shade column — e.g. as a page header/title, applying to the WHOLE page, not a per-row bracket), include it as "party" for every item on that page; otherwise omit "party" entirely.
Return ONLY one JSON object: {"date":"shared header date if Style A, DD/MM/YYYY, else omit","items":[{"date_override":"the row's own date if Style A/B and it differs from the header, OR the ditto-resolved date copied from the row above if this row used a ditto mark — include this whenever you're not 100% sure the header date alone is right, else omit","party":"only if found as a whole-page header, else omit","customer_hint":"Style B only — a bracketed customer name found next to this specific item, else omit","shade":"Style A only, else omit","size":"Style A only, else omit","gsm":"Style A only, else omit","weight":"Style A only (number), else omit","description":"Style B only — the FULL exact item name including weight and pack count, with no bracketed customer name in it, else omit","pieces":"number — pieces produced (Tukda in Style A; in Style B, the single Quantity column, or the SUM of both quantity columns if the page has two side by side; the CORRECTED number if this row had a crossed-out value; 0 if the quantity column isn't actually visible for this row)","dispatch":"number — ONLY if the page has an explicit, separately-written dispatch/sent notation for this row (rare) — never just because a second quantity column exists, else 0","flag":"only if you couldn't reliably read this row's quantity — e.g. the column looks cut off/missing, or you're genuinely unsure which of two numbers is the corrected one — a short reason why, else omit"}]}`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      const rows = raw.items.map(it => ({
        id: genId(), date: normalizeDateToDots(it.date_override || raw.date || it.date || ''), party: it.party || '', shade: it.shade || '',
        size: it.size || '', gsm: it.gsm || '', weight: num(it.weight), description: it.description || '',
        customerHint: it.customer_hint || '', pieces: num(it.pieces), dispatch: num(it.dispatch),
        stockConfirmed: false, confirmedCustomer: '',
        flagged: !!(it.flag && String(it.flag).trim()), flagReason: (it.flag || '').trim(),
      }));
      return fillDittoDates(rows);
    },
  },
  {
    key: 'dispatch_bill',
    label: 'Dispatch Bill / Tax Invoice',
    hint: 'A printed tax invoice / dispatch bill sent to any customer (Bindal, Diamond, Anmol, or otherwise) — lands in its own Customer Dispatch Bills tab, separate from the Production Register, and reduces that customer’s stock balance on the Customer Stock tab.',
    register: 'customerDispatch',
    systemPrompt: `You read a printed GST tax invoice / dispatch bill for corrugated boxes sent to a customer.
CRITICAL — you are shown ONE PAGE at a time, in isolation. You have no memory of any other page in this document, so you must decide what to do based only on what's printed on THIS page. Every dispatch bill is printed as a set of near-identical copies, each one labeled directly beside or below the words "Tax Invoice" at the very top: "(ORIGINAL FOR RECIPIENT)", "(DUPLICATE FOR TRANSPORTER)", "(TRIPLICATE FOR SUPPLIER)", "(EXTRA COPY)" — plus separate e-Way Bill pages mixed in between them. These copies are NOT separate invoices and NOT separate deliveries — they are the identical invoice printed several times.
Extract line items ONLY from the page labeled exactly "(ORIGINAL FOR RECIPIENT)". For every other page — one labeled "(DUPLICATE FOR TRANSPORTER)", "(TRIPLICATE FOR SUPPLIER)", "(EXTRA COPY)", or any e-Way Bill page — return the same JSON shape with an empty items array, no matter how clear or legible that particular page is. Do not extract from a Duplicate/Triplicate/Extra copy just because it happens to be easier to read than the Original — a different page in this same upload batch is the Original and will be extracted from instead. If the page has no visible copy label at all, treat it as NOT the Original and return empty items.
If the page you are looking at right now is purely an e-Way Bill, or is any non-Original copy you're deliberately skipping per the rule above, do NOT treat that as an error and do NOT refuse or apologize — just return the same JSON shape with an empty items array: {"invoice_no":"","date":"","party":"","buyer_order_no":"","items":[]}. An empty result is a completely valid, expected answer for most pages in a batch like this.
Extract (only from the Original page):
- "invoice_no": the Invoice No.
- "date": the invoice date
- "party": the Buyer (Bill to) name — the actual company name, e.g. "BINDAL TECHNOPOLYMER PVT. LTD." or "ANMOL INDUSTRIES LTD." — copy it as printed
- "buyer_order_no": the Buyer's Order No. / PO reference if printed, else empty string
- "items": one entry per line under "Description of Goods". Each line typically shows "CORRUGATED BOX" as a generic heading with the real product name in italics underneath (e.g. "IT 500 CONT", "HANDLE LOCK", "N 100 JUMBO XL CONT", "JEERA DHAMAL 35G x 144PKT") — use that italic sub-line as the description, NOT the generic "CORRUGATED BOX" text. Ignore the "(BOX:- N*M = total)" annotation underneath — it's just arithmetic, not part of the item name. "quantity" is the Quantity column value (e.g. 500 from "500.0 nos.").
- Word abbreviations this factory uses that a person has taught the app — expand these to their full form wherever they appear inside "description" (e.g. write "IT 500 Jumbo Container", not "IT 500 J Cont"), so the same item is always written the same way no matter which shorthand a given invoice happened to print:
{{ABBREVIATIONS}}
Return ONLY one JSON object: {"invoice_no":"","date":"","party":"","buyer_order_no":"","items":[{"description":"","quantity":number,"rate":number,"amount":number}]}`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      return raw.items.map(it => ({
        id: genId(), date: normalizeDateToDots(raw.date || ''), party: raw.party || '', description: it.description || '',
        customerHint: raw.party || '', invoice_no: raw.invoice_no || '', buyer_order_no: raw.buyer_order_no || '',
        quantity: num(it.quantity), rate: num(it.rate), amount: num(it.amount),
        stockConfirmed: false, confirmedCustomer: '',
      }));
    },
  },
  {
    key: 'dabur_spec',
    label: 'Dabur Spec Sheet (reference)',
    hint: 'Printed Dabur PM Specification sheet for a box item — replaces the manual diary.',
    register: 'daburSpecs',
    systemPrompt: `You read a printed "Dabur India Limited — PM Specification" sheet for a corrugated box. Return ONLY a JSON array with ONE object:
[{"product_name":"product/item name","box_size":"Length x Width x Height with units, from the Length/Width/Height rows","gsm_combo":"the paper combination and GSM text, e.g. 140(VK)/120(SK)/120(SK)/120(SK)/150(SK)","partition_size":"partition size/count if listed","plate_size":"plate/central-plate size if listed","compression":"compression strength / bursting factor spec","notes":"any other short important spec worth keeping, 1 sentence max"}]`,
    shape: (raw) => (Array.isArray(raw) ? raw : []).map(r => ({ id: genId(), product_name: r.product_name || '', box_size: r.box_size || '', gsm_combo: r.gsm_combo || '', partition_size: r.partition_size || '', plate_size: r.plate_size || '', compression: r.compression || '', notes: r.notes || '' })),
  },
  {
    key: 'dabur_po',
    label: 'Dabur Purchase Order (printed)',
    hint: 'Printed purchase order from Dabur — creates entries on the Pending PO list.',
    register: 'daburPO',
    systemPrompt: `You read a printed Dabur purchase order (may be titled "Draft Purchase Order"). Return ONLY one JSON object:
{"po_number":"P.O. Number","date":"PO date DD/MM/YYYY","items":[{"material_desc":"material/description text","hsn":"HSN/SAC code if present","quantity":number,"rate":number,"delivery_date":"delivery date if present, else empty string"}]}`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      return raw.items.map(it => ({
        id: genId(), po_number: raw.po_number || '', date: normalizeDateToDots(raw.date || ''), material_desc: it.material_desc || '',
        hsn: it.hsn || '', quantity: num(it.quantity), rate: num(it.rate), delivery_date: normalizeDateToDots(it.delivery_date || '')
      }));
    },
  },
  {
    key: 'dabur_dispatch',
    label: 'Dabur Dispatch Bill (updates Pending PO)',
    hint: 'The dispatch bill / invoice sent against a Dabur PO — reduces the pending quantity on that PO.',
    register: 'daburDispatch',
    systemPrompt: `You read a printed tax invoice / dispatch bill for goods dispatched against a Dabur purchase order.
CRITICAL — you are shown ONE PAGE at a time, in isolation. You have no memory of any other page in this document, so you must decide what to do based only on what's printed on THIS page. Every dispatch bill is printed as a set of near-identical copies, each one labeled directly beside or below the words "Tax Invoice" at the very top: "(ORIGINAL FOR RECIPIENT)", "(DUPLICATE FOR TRANSPORTER)", "(TRIPLICATE FOR SUPPLIER)", "(EXTRA COPY)" — plus separate e-Way Bill pages mixed in between them. These copies are NOT separate invoices — they are the identical invoice printed several times.
Extract line items ONLY from the page labeled exactly "(ORIGINAL FOR RECIPIENT)". For every other page — "(DUPLICATE FOR TRANSPORTER)", "(TRIPLICATE FOR SUPPLIER)", "(EXTRA COPY)", or any e-Way Bill page — return the same JSON shape with an empty items array, no matter how clear or legible that page is. Do not extract from a Duplicate/Triplicate/Extra copy just because it's easier to read than the Original — a different page in this same batch is the Original and will be extracted from instead. If the page has no visible copy label at all, treat it as NOT the Original and return empty items.
If the page you are looking at right now is purely an e-Way Bill, or is any non-Original copy you're deliberately skipping per the rule above, do NOT treat that as an error and do NOT refuse or apologize — just return the same JSON shape with an empty items array. An empty result is a completely valid, expected answer for most pages in a batch like this.
Return ONLY one JSON object:
{"invoice_no":"invoice number","date":"invoice date","party":"buyer name, usually Dabur / Anmol Industries etc.","buyer_order_no":"the Buyer's Order No. / PO number referenced on the invoice — this is critical, look carefully, there may be more than one","items":[{"description":"item description","quantity":number,"rate":number,"amount":number}]}
If multiple buyer order numbers are listed, use the first one and mention any others inside the description field of the relevant item.`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      return raw.items.map(it => ({
        id: genId(), date: normalizeDateToDots(raw.date || ''), invoice_no: raw.invoice_no || '', party: raw.party || '',
        buyer_order_no: raw.buyer_order_no || '', description: it.description || '', quantity: num(it.quantity), rate: num(it.rate), amount: num(it.amount)
      }));
    },
  },
];
const COLUMNS = {
  rawMaterialIn: [
    { key: 'date', label: 'Date' }, { key: 'mill', label: 'Mill' }, { key: 'reel_no', label: 'Reel No' },
    { key: 'size', label: 'Size' }, { key: 'unit', label: 'Unit' }, { key: 'gsm', label: 'GSM' }, { key: 'bf', label: 'BF' }, { key: 'shade', label: 'Shade' },
    { key: 'weight_kg', label: 'Weight (kg)', type: 'number' }, { key: 'consumed', label: 'Consumed (date)' },
  ],
  consumption: [
    { key: 'sl_no', label: 'SL. No.' },
    { key: 'date', label: 'Date' }, { key: 'shade', label: 'Shade' }, { key: 'size', label: 'Size' }, { key: 'gsm', label: 'GSM' },
    { key: 'weight_consumed', label: 'Weight Consumed', type: 'number' },
  ],
  production: [
    { key: 'date', label: 'Date' }, { key: 'party', label: 'Party (if noted)' }, { key: 'description', label: 'Description (as written)' },
    { key: 'customerHint', label: 'Bracketed Customer' },
    { key: 'pieces', label: 'Pieces Produced', type: 'number' }, { key: 'dispatch', label: 'Dispatch (handwritten, if noted)', type: 'number' },
  ],
  customerDispatch: [
    { key: 'date', label: 'Date' }, { key: 'invoice_no', label: 'Invoice No' }, { key: 'party', label: 'Party' },
    { key: 'buyer_order_no', label: 'Buyer Order No' }, { key: 'description', label: 'Description' },
    { key: 'quantity', label: 'Quantity Dispatched', type: 'number' }, { key: 'rate', label: 'Rate', type: 'number' }, { key: 'amount', label: 'Amount', type: 'number' },
  ],
  daburSpecs: [
    { key: 'product_name', label: 'Product' }, { key: 'box_size', label: 'Box Size (L x W x H)' },
    { key: 'gsm_combo', label: 'Paper Combination / GSM' }, { key: 'partition_size', label: 'Partition' },
    { key: 'plate_size', label: 'Plate' }, { key: 'compression', label: 'Compression' }, { key: 'notes', label: 'Notes' },
  ],
  daburPO: [
    { key: 'po_number', label: 'PO Number' }, { key: 'date', label: 'Date' }, { key: 'material_desc', label: 'Material' },
    { key: 'hsn', label: 'HSN' }, { key: 'quantity', label: 'Ordered Qty', type: 'number' }, { key: 'rate', label: 'Rate', type: 'number' }, { key: 'delivery_date', label: 'Delivery Date' },
  ],
  daburDispatch: [
    { key: 'date', label: 'Date' }, { key: 'invoice_no', label: 'Invoice No' }, { key: 'party', label: 'Party' },
    { key: 'buyer_order_no', label: 'PO Ref No' }, { key: 'description', label: 'Description' },
    { key: 'quantity', label: 'Quantity', type: 'number' }, { key: 'rate', label: 'Rate', type: 'number' }, { key: 'amount', label: 'Amount', type: 'number' },
  ],
};
// The Raw Material Register view is grouped into one table per size (see rawMaterialBySize in
// FIMSApp), matching how the physical mill-slip register book itself is organized — a size's own
// page. "Size" itself is dropped from these per-row columns since it's now the group heading, not a
// repeated cell; "Reel No" and "Unit" are dropped too since they're not part of what was asked for
// here. "Consumed" is a real, persisted field on the row (see mill_slip's shape() and COLUMNS.rawMaterialIn
// above) — deliberately blank for now, to be filled in once a future feature matches this entry
// against consumption reports.
const RAW_MATERIAL_SIZE_COLUMNS = [
  { key: 'date', label: 'Date' }, { key: 'mill', label: 'Mill' }, { key: 'weight_kg', label: 'Weight (kg)', type: 'number' },
  { key: 'gsm', label: 'GSM' }, { key: 'bf', label: 'BF' }, { key: 'shade', label: 'Shade' }, { key: 'consumed', label: 'Consumed (date)' },
];
const NAV = [
  { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { key: 'upload', label: 'Upload & Extract', icon: Upload },
  { key: 'rawMaterialIn', label: 'Raw Material Register', icon: Archive },
  { key: 'production', label: 'Production Register', icon: ClipboardList },
  { key: 'customerDispatch', label: 'Customer Dispatch Bills', icon: Package },
  { key: 'daburSpecs', label: 'Dabur — Spec Master', icon: FileText },
  { key: 'daburPO', label: 'Dabur — Pending PO', icon: ListChecks },
  { key: 'daburDispatch', label: 'Dabur — Dispatch Log', icon: Truck },
  { key: 'customerStock', label: 'Customer Stock', icon: Boxes },
  { key: 'customerMapping', label: 'Customer Mapping', icon: ListChecks },
  { key: 'aliases', label: 'Aliases', icon: Link2 },
  { key: 'customerSheets', label: 'Customer Sheets', icon: FileSpreadsheet },
  { key: 'settings', label: 'Settings', icon: Trash2 },
];
const GUIDE_STEPS = [
  {
    title: 'Upload a document',
    tab: 'upload',
    what: 'This is where every photo, scan, or PDF enters the system — mill slips, consumption sheets, production register pages, dispatch bills, Dabur specs/POs, all of it.',
    how: 'Pick the document type from the dropdown first (this tells the system what to look for), then upload the file. JPG, PNG, and PDF all work. For iPhone photos, make sure they’re saved as JPG/PNG, not HEIC.',
  },
  {
    title: 'Extract, then review before anything is saved',
    tab: 'upload',
    what: 'Click "Extract data" and it reads the document into a table. Nothing is saved yet — this is a proposal, not a done deal.',
    how: 'Check every row against the actual document. Fix anything wrong by clicking directly into a cell. Delete a row if it shouldn’t be there. Only once it looks right, click "Add rows to register."',
  },
  {
    title: 'Raw Material Register fills itself in from there',
    tab: 'rawMaterialIn',
    what: 'Mill slips become inward entries; daily consumption sheets become consumption entries. The Balance table at the top (In minus Consumed, by size/GSM) is fully automatic — you never edit it directly.',
    how: 'Upload slips and consumption sheets as they come in. If a balance looks wrong, the fix is correcting the underlying inward/consumption row below, not the balance itself.',
  },
  {
    title: 'Production Register handles two different page styles',
    tab: 'production',
    what: 'Some of your handwritten pages track shade/size/GSM (box blanks); others list a product description and quantity (finished goods, e.g. "Butter Bake 65g x60"). Both land in this one table — the system figures out which style a page is automatically.',
    how: 'Upload as usual. If a page has a customer name written on it (as a page header, not a bracket next to one item), it’s worth double-checking the Party column got filled in — that field is searchable from the search bar at the top of the app.',
  },
  {
    title: 'New finished-goods rows wait in Customer Stock for your OK',
    tab: 'customerStock',
    what: 'Any Production Register row with a product description, and any Customer Dispatch Bill row, shows up in "Pending Review" at the top of Customer Stock. Neither counts toward any customer’s balance, and nothing is pushed to their Sheet, until you confirm it.',
    how: 'Check the suggested customer for each row (auto-matched, editable if it’s wrong), then confirm one at a time or use "Push to Sheet" to confirm and push everything resolved in one go. Duplicate rows already in the real Sheet are skipped automatically.',
  },
  {
    title: 'Customer Mapping decides who gets what',
    tab: 'customerMapping',
    what: 'This is the rulebook Customer Stock uses to guess which customer a product belongs to — plus the reference catalog of exact item names that helps correct handwriting misreads during extraction.',
    how: 'Paste a customer’s Google Sheet ID/link on the Customer Sheets tab (new or already-added) to add its products automatically. You can also add, edit, or delete individual rules by hand at any time — rules are checked top to bottom, first match wins.',
  },
  {
    title: 'Dispatch bills record what actually went out',
    tab: 'upload',
    what: 'Upload dispatch bills / tax invoices from the same dropdown as everything else — they land in their own Customer Dispatch Bills tab (kept separate from the Production Register) and, once confirmed, reduce that customer’s balance on Customer Stock.',
    how: 'Same upload-extract-review-confirm pattern as everywhere else. These invoices are usually printed with several duplicate copies (Original/Duplicate/Triplicate/Extra Copy) plus an e-Way Bill page — the extraction already knows to count the transaction once, not once per copy.',
  },
  {
    title: 'The Dabur side runs in parallel: Specs → PO → Dispatch',
    tab: 'daburSpecs',
    what: 'Spec Master replaces the manual diary — one row per box item. Pending PO tracks what Dabur has ordered and what’s still owed, with a 10% tolerance (a PO counts as Fulfilled once you’ve dispatched 90%+ of it). Dabur Dispatch Log is what drives that pending calculation, matched by PO number.',
    how: 'Upload spec sheets once per item. Upload each PO as it arrives. Upload dispatch bills against Dabur POs to keep the pending list current.',
  },
  {
    title: 'Export whenever you need an actual Excel file',
    tab: 'dashboard',
    what: 'Every table has its own "Export" button for just that data. The top-right "Export all to Excel" button bundles everything — every register, plus one sheet per customer’s stock — into one file. Every export opens a window with your data ready to copy-paste (always works), plus a "Download .xlsx" button for a real file (usually works — click Cancel then Export again if a download seems to silently do nothing, since browsers sometimes block a second automatic-feeling download in the same session).',
    how: 'The app itself is the real, permanent record — it doesn’t reset. Exporting just gives you a snapshot copy to hand off or back up; export as often or as rarely as you like.',
  },
];
/* ============================== small components ============================== */
function Pill({ tone = 'neutral', children, title }) {
  return <span className={`pill pill-${tone}`} title={title}>{children}</span>;
}
// A 'fillable' row's Sheet already has a row for this date — just missing production, missing dispatch,
// or missing both — so the label always says exactly which column(s) will actually get written, instead
// of a generic "gap" that could mean either.
function fillableLabel(r) {
  if (r.fillProduction && r.fillDispatch) return 'Fills gap (both)';
  if (r.fillProduction) return 'Fills production';
  if (r.fillDispatch) return 'Fills dispatch';
  return 'Fills gap';
}
function fillableTitle(r) {
  const parts = [];
  if (r.fillProduction) parts.push(`Production (${r.production})`);
  if (r.fillDispatch) parts.push(`Dispatch (${r.dispatch})`);
  const cols = parts.length ? parts.join(' and ') : 'this cell';
  return `This date already has a row in the Sheet, but ${cols} ${parts.length > 1 ? 'were' : 'was'} blank — will fill in just ${parts.length > 1 ? 'those cells' : 'that cell'}, without touching anything else on the row.`;
}
// Bold section separator used on Customer Stock to keep its three very different zones (needs your
// input, nobody's claimed this at all, and the running per-customer ledgers) from blurring into one
// long wall of similar-looking panels.
function SectionDivider({ icon: Icon, label, hint }) {
  return (
    <div className="stock-section-divider">
      {Icon && <Icon size={15} />}
      <span className="stock-section-divider-label">{label}</span>
      {hint && <span className="doc-hint">{hint}</span>}
    </div>
  );
}
// "Suggested Customer" cell on the Pending Production/Dispatch Review tables. Replaces a free-text
// input with a dropdown of customers the app already knows about, plus an explicit "+ New customer"
// option — so a dispatch bill's Party field (or a production row's bracketed name) that doesn't match
// anyone known can never silently spin up a duplicate, disconnected customer the way a free-text field
// pre-filled with a guess could. When the guess IS a real known customer (an alias hit, a mapping-rule
// match, or a fuzzy legal-name match against an existing customer), it's pre-selected as before so the
// already-reliable flows aren't disrupted — only a genuinely unresolved hint forces an explicit pick.
function CustomerSuggestCell({ value, guess, isKnownGuess, knownCustomers, onChange }) {
  const [customMode, setCustomMode] = useState(false);
  const current = value || (isKnownGuess ? guess : '');
  const isCustomValue = !!current && current !== 'Unassigned' && !knownCustomers.includes(current);
  if (customMode || isCustomValue) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
        <input className="cell-input" placeholder="Type new customer name" value={current}
          onChange={e => onChange(e.target.value)} />
        <button type="button" className="icon-btn" title="Pick from existing customers instead"
          onClick={() => { setCustomMode(false); onChange(''); }}><XCircle size={14} /></button>
      </div>
    );
  }
  return (
    <select className="cell-input" value={current} onChange={e => {
      if (e.target.value === '__new__') { setCustomMode(true); onChange(''); return; }
      onChange(e.target.value);
    }}>
      <option value="" disabled>Select customer…</option>
      {knownCustomers.map(c => <option key={c} value={c}>{c}</option>)}
      <option value="Unassigned">Unassigned</option>
      <option value="__new__">+ New customer…</option>
    </select>
  );
}
// "Sheet tab" / "Block" cells on a Pending Production/Dispatch Review row, shown only once a real
// customer is picked/confirmed AND the Product Catalog doesn't already know this item — i.e. exactly
// the case Customer Stock's "not routed to a Sheet tab yet" section exists to catch, just settled here
// instead, in the same step as picking the customer. Confirming the row (see confirmStockRow) reads
// this draft and registers it as a Product Catalog alias, so it's remembered for every future row with
// this same wording — no second pass needed. Both fields are flagged red until EXPLICITLY filled in —
// leaving Block blank still functionally falls back to a new block named after the item on confirm, but
// the row stays flagged until that choice is made on purpose (e.g. picking "+ New block" from its own
// dropdown), not left blank by default.
function TabBlockPickerCells({ rowId, description, sheetGroupOptions, blockOptions, draft, onChange }) {
  const sheetGroup = (draft && draft.sheetGroup) || '';
  const block = (draft && draft.block) || '';
  const tabUnresolved = !sheetGroup.trim();
  const blockUnresolved = !block.trim();
  return (
    <>
      <td>
        <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          <input className="cell-input" style={{ width: 110, fontSize: 12, borderColor: tabUnresolved ? 'var(--ledger-red)' : undefined }}
            list={`pending-sheetgroup-${rowId}`}
            placeholder="Sheet tab" value={sheetGroup} onChange={e => onChange('sheetGroup', e.target.value)} />
          <datalist id={`pending-sheetgroup-${rowId}`}>{sheetGroupOptions.map(s => <option key={s} value={s} />)}</datalist>
          {tabUnresolved && (
            <AlertCircle size={14} color="var(--ledger-red)" style={{ flexShrink: 0 }}
              title="Not yet assigned to a Sheet tab — pick one, or type a new one." />
          )}
        </div>
      </td>
      <td>
        <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          <input className="cell-input" style={{ width: 140, fontSize: 12, borderColor: blockUnresolved ? 'var(--ledger-red)' : undefined }}
            list={`pending-block-${rowId}`}
            placeholder="Block (blank = new block)" value={block} onChange={e => onChange('block', e.target.value)} />
          <datalist id={`pending-block-${rowId}`}>
            {blockOptions.map(b => <option key={b} value={b} />)}
            <option value={description}>{`+ New block: "${description}"`}</option>
          </datalist>
          {blockUnresolved && (
            <AlertCircle size={14} color="var(--ledger-red)" style={{ flexShrink: 0 }}
              title={`Not yet assigned to a block — pick an existing one, or "+ New block: \"${description}\"" to confirm this should be its own new block.`} />
          )}
        </div>
      </td>
    </>
  );
}
// One dropdown per real-world invoice/party, above the itemized Pending Review table, so picking a
// customer for a multi-line bill doesn't mean repeating the same pick on every one of its rows.
// Applying it sets every row in the group at once; "Confirm all N" then confirms just this group
// without waiting for (or being blocked by) anything else still pending elsewhere in the table.
function PendingGroupBar({ group, guess, isKnownGuess, knownCustomers, onBulkChange, onConfirmGroup }) {
  const values = Array.from(new Set(group.rows.map(r => (r.confirmedCustomer || '').trim())));
  const commonValue = values.length === 1 ? values[0] : '';
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px', border: '1px dashed var(--accent)', borderRadius: 6, marginBottom: 6, flexWrap: 'wrap', background: 'var(--accent-soft)' }}>
      <span style={{ fontSize: 12, flex: '1 1 auto', minWidth: 120 }}>{group.rows.length} rows — <strong>{group.key}</strong></span>
      <CustomerSuggestCell value={commonValue} guess={guess} isKnownGuess={isKnownGuess} knownCustomers={knownCustomers}
        onChange={v => onBulkChange(group.rows.map(r => r.id), v)} />
      <button type="button" className="btn btn-ghost" style={{ padding: '4px 8px', fontSize: 12 }} disabled={!commonValue}
        title={commonValue ? `Confirm all ${group.rows.length} rows as ${commonValue}` : 'Pick a customer above first'}
        onClick={onConfirmGroup}>
        <CheckCircle2 size={13} /> Confirm all {group.rows.length}
      </button>
    </div>
  );
}
// Every register table with a "date" column renders in real chronological order (dateSortKey — see
// its own comment — handles D.M.YY correctly, unlike a plain string sort). Rows get appended to a
// register in whatever order they're extracted/confirmed, which is routinely NOT chronological order
// (a page dated last week uploaded today lands after today's already-confirmed rows) — sorting here,
// once, covers every register/search-result table that has a date column, rather than needing every
// call site to remember to sort its own rows before passing them in. A stable sort (guaranteed by the
// spec), so rows sharing the same date keep their existing relative order. sortByDate={false} opts a
// caller out — used only by the pre-confirm extraction review table, which must stay in the exact
// order the model returned it (matching the physical page top-to-bottom) so a row can be cross-checked
// against the original photo by position; sorting THAT one by date would scramble the correspondence.
// showBatchFill (opt-in, only the pre-confirm extraction review table uses it): a row of per-column
// batch controls instead of clicking into every row by hand — up to three actions per column:
// - Fill blanks: sets ONLY the currently-blank cells (e.g. no party/customer name written anywhere on
//   the page, so it's blank on every row) — never overwrites a row that already has a real per-row
//   value, so a page where only some lines had an exception noted never gets a legitimate value
//   clobbered. Disabled when the column has no blank cells at all (nothing to fill).
// - Overwrite all: sets EVERY row's cell to this value, replacing anything already there — for
//   correcting a systematic misread across the whole page (e.g. the model got one column's value
//   wrong for every row). Confirms first whenever it would actually replace an existing value (never
//   for a column that's already all-blank, where it's equivalent to Fill blanks), since this one really
//   can lose real per-row data if clicked on the wrong column.
// - Apply to selected: only shown once at least one row is selected via the gutter (see EditableTable)
//   — sets just the selected rows' cell, ignoring blank/non-blank status entirely, since picking exact
//   rows by hand is itself the deliberate scoping (no confirm needed the way Overwrite all needs one).
function BatchFillRow({ columns, rows, onUpdate, selectedIds, onClearSelection }) {
  const [batchValues, setBatchValues] = useState({});
  const fillBlanks = (colKey) => {
    const value = (batchValues[colKey] ?? '').trim();
    if (!value) return;
    rows.forEach(row => { if (!String(row[colKey] ?? '').trim()) onUpdate(row.id, colKey, value); });
    setBatchValues(prev => ({ ...prev, [colKey]: '' }));
  };
  const overwriteAll = (colKey, label) => {
    const value = (batchValues[colKey] ?? '').trim();
    if (!value) return;
    const nonBlankCount = rows.filter(r => String(r[colKey] ?? '').trim()).length;
    if (nonBlankCount > 0 && !window.confirm(`Set every row's ${label} to "${value}"? This REPLACES ${nonBlankCount} row${nonBlankCount === 1 ? '' : 's'} that already ha${nonBlankCount === 1 ? 's' : 've'} a value here — not just the blank ones.`)) return;
    rows.forEach(row => onUpdate(row.id, colKey, value));
    setBatchValues(prev => ({ ...prev, [colKey]: '' }));
  };
  const applyToSelected = (colKey) => {
    const value = (batchValues[colKey] ?? '').trim();
    if (!value || !selectedIds.size) return;
    rows.forEach(row => { if (selectedIds.has(row.id)) onUpdate(row.id, colKey, value); });
    setBatchValues(prev => ({ ...prev, [colKey]: '' }));
  };
  return (
    <tr className="batch-fill-row">
      <th className="col-select">
        {selectedIds.size > 0 && (
          <button type="button" className="icon-btn" title={`${selectedIds.size} row(s) selected — click to clear`} onClick={onClearSelection} style={{ fontSize: 10, padding: '2px 4px' }}>
            {selectedIds.size}✕
          </button>
        )}
      </th>
      {columns.map(c => {
        const hasBlank = rows.some(r => !String(r[c.key] ?? '').trim());
        return (
          <th key={c.key}>
            <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
              <input
                className="cell-input"
                style={{ fontSize: 11.5 }}
                placeholder="Batch value"
                type={c.type === 'number' ? 'number' : 'text'}
                value={batchValues[c.key] || ''}
                onChange={e => setBatchValues(prev => ({ ...prev, [c.key]: e.target.value }))}
                onKeyDown={e => { if (e.key === 'Enter') fillBlanks(c.key); }}
              />
              <button type="button" className="icon-btn" title={hasBlank ? `Fill every blank ${c.label} cell with this value` : `No blank ${c.label} cells — nothing to fill`} disabled={!hasBlank || !(batchValues[c.key] || '').trim()} onClick={() => fillBlanks(c.key)}>
                <Check size={13} />
              </button>
              <button type="button" className="icon-btn" title={`Set EVERY row's ${c.label} to this value, replacing anything already there`} disabled={!(batchValues[c.key] || '').trim()} onClick={() => overwriteAll(c.key, c.label)}>
                <RefreshCw size={13} />
              </button>
              {selectedIds.size > 0 && (
                <button type="button" className="icon-btn" title={`Set just the ${selectedIds.size} selected row(s)' ${c.label} to this value`} disabled={!(batchValues[c.key] || '').trim()} onClick={() => applyToSelected(c.key)}>
                  <CheckCircle2 size={13} />
                </button>
              )}
            </div>
          </th>
        );
      })}
      <th className="col-action"></th>
    </tr>
  );
}
function EditableTable({ columns, rows, onUpdate, onDelete, emptyLabel = 'No entries yet.', suppressFlags = false, highlightRow, sortByDate = true, showBatchFill = false }) {
  // Row selection (also showBatchFill-only, see BatchFillRow's "Apply to selected"): shift-click for a
  // range, ctrl/cmd-click to toggle one row without disturbing the rest, and click-drag across the
  // gutter for an OS-style rubber-band select — mirrors how selection already works everywhere else on
  // the person's own machine. Declared before the `!rows.length` early return below (rules of hooks:
  // must run in the same order every render) even though selection is meaningless on an empty table.
  // Kept local to this table instance rather than lifted to the caller — intersected against the
  // CURRENT rows on every read rather than reset via an effect keyed on page identity, so switching to
  // a different file's review table silently drops ids that no longer exist, no extra plumbing needed.
  const [rawSelectedIds, setRawSelectedIds] = useState(() => new Set());
  const [dragAnchorId, setDragAnchorId] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  useEffect(() => {
    if (!isDragging) return;
    const stop = () => setIsDragging(false);
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, [isDragging]);
  if (!rows.length) return <div className="empty-state">{emptyLabel}</div>;
  const hasDateColumn = sortByDate && columns.some(c => c.key === 'date');
  // SL. No. sorts ascending whenever the register has that column, in BOTH the persisted register
  // table and the pre-confirm extraction review — unlike date sorting (gated behind sortByDate, since
  // scrambling the review table by date would break cross-checking a row against the original photo
  // by position), SL. No. is the row's own real identity, not something tied to upload/confirm order.
  // Some source sheets (e.g. Consumption) physically lay their columns out in decreasing SL. No. order
  // left-to-right, so preserving "the model's read order" there means preserving a page that's
  // ALREADY backwards — sorting numerically fixes that instead of faithfully reproducing it.
  // Compares numerically when both sides parse as numbers (so "10" sorts after "2", not before).
  const hasSlNoColumn = columns.some(c => c.key === 'sl_no');
  const slNoCompare = (a, b) => {
    const na = parseFloat(a.sl_no), nb = parseFloat(b.sl_no);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
    return String(a.sl_no ?? '').localeCompare(String(b.sl_no ?? ''), undefined, { numeric: true });
  };
  const sortedRows = hasDateColumn
    ? [...rows].sort((a, b) => {
        const dateCmp = dateSortKey(a.date).localeCompare(dateSortKey(b.date));
        if (dateCmp !== 0) return dateCmp;
        return hasSlNoColumn ? slNoCompare(a, b) : 0;
      })
    : (hasSlNoColumn ? [...rows].sort(slNoCompare) : rows);
  const rowIdOrder = sortedRows.map(r => r.id);
  // Re-derived every render rather than trusted as-is: an id that no longer exists in the current
  // `rows` (switched to a different file, or the row got deleted) just silently drops out here.
  const selectedIds = showBatchFill ? new Set(rowIdOrder.filter(id => rawSelectedIds.has(id))) : rawSelectedIds;
  const rangeBetween = (fromId, toId) => {
    const fromIdx = rowIdOrder.indexOf(fromId), toIdx = rowIdOrder.indexOf(toId);
    if (fromIdx === -1 || toIdx === -1) return [toId];
    const [lo, hi] = fromIdx <= toIdx ? [fromIdx, toIdx] : [toIdx, fromIdx];
    return rowIdOrder.slice(lo, hi + 1);
  };
  const handleGutterMouseDown = (rowId, e) => {
    e.preventDefault(); // stop the browser's own text-selection drag from fighting the row-drag below
    if (e.shiftKey && dragAnchorId) {
      setRawSelectedIds(new Set(rangeBetween(dragAnchorId, rowId)));
    } else if (e.ctrlKey || e.metaKey) {
      setRawSelectedIds(prev => {
        const next = new Set(prev);
        if (next.has(rowId)) next.delete(rowId); else next.add(rowId);
        return next;
      });
      setDragAnchorId(rowId);
    } else {
      setRawSelectedIds(new Set([rowId]));
      setDragAnchorId(rowId);
      setIsDragging(true);
    }
  };
  const handleGutterMouseEnter = (rowId) => {
    if (!isDragging || !dragAnchorId) return;
    setRawSelectedIds(new Set(rangeBetween(dragAnchorId, rowId)));
  };
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {showBatchFill && <th className="col-select"></th>}
            {columns.map(c => <th key={c.key}>{c.label}</th>)}
            <th className="col-action"></th>
          </tr>
          {showBatchFill && <BatchFillRow columns={columns} rows={rows} onUpdate={onUpdate} selectedIds={selectedIds} onClearSelection={() => setRawSelectedIds(new Set())} />}
        </thead>
        <tbody>
          {sortedRows.map(row => {
            const isSelected = showBatchFill && selectedIds.has(row.id);
            return (
              <tr key={row.id} className={[
                (!suppressFlags && row.flagged) ? 'flagged-row' : '',
                (highlightRow && highlightRow(row)) ? 'needs-customer-row' : '',
                isSelected ? 'row-selected' : '',
              ].filter(Boolean).join(' ')}>
                {showBatchFill && (
                  <td className="col-select"
                    onMouseDown={(e) => handleGutterMouseDown(row.id, e)}
                    onMouseEnter={() => handleGutterMouseEnter(row.id)}>
                    <input type="checkbox" checked={isSelected} readOnly style={{ pointerEvents: 'none' }} />
                  </td>
                )}
                {columns.map((c, ci) => (
                  <td key={c.key}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                      {ci === 0 && !suppressFlags && row.flagged && (
                        <AlertCircle size={13} color="var(--ledger-red)" style={{ flexShrink: 0 }} title={row.flagReason || 'Flagged during extraction — the model wasn\'t confident about this row. Check it against the original document.'} />
                      )}
                      <input
                        className="cell-input"
                        type={c.type === 'number' ? 'number' : 'text'}
                        value={row[c.key] ?? ''}
                        onChange={(e) => onUpdate(row.id, c.key, c.type === 'number' ? e.target.value : e.target.value)}
                      />
                    </div>
                  </td>
                ))}
                <td className="col-action">
                  <button className="icon-btn danger" title="Delete row" onClick={() => onDelete(row.id)}>
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
function RegisterPanel({ title, subtitle, columns, rows, onUpdate, onDelete, onExport, extra, suppressFlags = false, highlightRow }) {
  return (
    <div className="panel">
      <div className="panel-header">
        <div>
          <h2>{title}</h2>
          {subtitle && <p className="subtitle">{subtitle}</p>}
        </div>
        <button className="btn btn-ghost" onClick={onExport}><Download size={15} /> Export this table</button>
      </div>
      {extra}
      <EditableTable columns={columns} rows={rows} onUpdate={onUpdate} onDelete={onDelete} suppressFlags={suppressFlags} highlightRow={highlightRow} />
    </div>
  );
}
// The export window for every "Export" button in the app. Opens INSTANTLY with the data already laid
// out as copy-paste-ready text (this can never fail — it's just string formatting, no browser API
// involved), and additionally offers a "Download .xlsx" button that builds and downloads a real file
// on click. Building the workbook happens lazily, right here, only when that button is pressed — so if
// anything about a particular table's data ever did trip up the xlsx library, it would show as a small
// inline message in this window with the copy-paste data still sitting right there, not as a silent
// "nothing happened" click.
function CopyExportModal({ data, onClose }) {
  const isMulti = data.sheets.length > 1;
  const [activeIdx, setActiveIdx] = useState(0);
  // 'all' pastes every sheet as one stacked block in a single Ctrl+C — the default for multi-sheet
  // exports (customer stock, export-all) so getting the complete data never requires repeating the
  // copy-paste once per item. 'single' keeps the original one-sheet-at-a-time dropdown view, for anyone
  // who specifically wants to paste one item into its own existing Excel tab.
  const [copyMode, setCopyMode] = useState(isMulti ? 'all' : 'single');
  const [downloadState, setDownloadState] = useState('idle'); // idle | done | error:<msg>
  const textareaRef = useRef(null);
  const sheet = data.sheets[activeIdx] || data.sheets[0];
  const tsv = isMulti && copyMode === 'all' ? toMultiBlockTSV(data.sheets) : toTSV(sheet.rows, sheet.columns);
  useEffect(() => {
    if (textareaRef.current) { textareaRef.current.focus(); textareaRef.current.select(); }
  }, [activeIdx, copyMode, sheet]);
  const selectAll = () => { if (textareaRef.current) { textareaRef.current.focus(); textareaRef.current.select(); } };
  const isMac = typeof navigator !== 'undefined' && /Mac/i.test(navigator.platform || '');
  const downloadXlsx = () => {
    try {
      const wb = buildWorkbookFromSheets(data.sheets);
      downloadWorkbook(wb, `${data.title.replace(/\s+/g, '_')}.xlsx`);
      setDownloadState('done');
    } catch (e) {
      console.error('Download error:', e);
      setDownloadState(`error:${e.message || 'unknown error'}`);
    }
  };
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(35,38,43,0.55)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }} onClick={onClose}>
      <div style={{ background: '#fff', borderRadius: 8, padding: 22, maxWidth: 760, width: '100%', maxHeight: '85vh', display: 'flex', flexDirection: 'column' }} onClick={e => e.stopPropagation()}>
        <h2 style={{ marginBottom: 6 }}>{data.title}</h2>
        <p className="subtitle" style={{ marginBottom: 10 }}>
          Two ways to get this out — the copy-paste method below always works: click "Select all", press {isMac ? 'Cmd+C' : 'Ctrl+C'}, then paste into a blank Excel or Google Sheets cell — it splits into columns automatically. "Download .xlsx" gives you a real file (one sheet per item, matching your usual layout) and usually works too; if a download seems to do nothing, it's likely your browser silently blocking a second automatic-feeling download in this session — the copy-paste data above is unaffected either way.
        </p>
        {isMulti && (
          <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
            <button className={`btn ${copyMode === 'all' ? 'btn-primary' : 'btn-ghost'}`} onClick={() => setCopyMode('all')}>All {data.sheets.length} items (one paste)</button>
            <button className={`btn ${copyMode === 'single' ? 'btn-primary' : 'btn-ghost'}`} onClick={() => setCopyMode('single')}>One item at a time</button>
          </div>
        )}
        {isMulti && copyMode === 'all' && (
          <p className="doc-hint" style={{ marginTop: -4, marginBottom: 10 }}>
            Pastes every item's ledger stacked into one sheet, each with its own name and header row — everything in a single paste. Not split into separate tabs like your original files; use "One item at a time" below, or "Download .xlsx", if you need that exact layout.
          </p>
        )}
        {isMulti && copyMode === 'single' && (
          <select className="doc-select" value={activeIdx} onChange={e => setActiveIdx(Number(e.target.value))} style={{ marginBottom: 10 }}>
            {data.sheets.map((s, i) => <option key={i} value={i}>{s.name}</option>)}
          </select>
        )}
        <textarea ref={textareaRef} readOnly value={tsv}
          style={{ flex: 1, minHeight: 260, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, padding: 10, border: '1px solid var(--rule)', borderRadius: 6, resize: 'vertical' }}
          onFocus={(e) => e.target.select()} />
        {downloadState === 'done' && (
          <div className="doc-hint" style={{ color: 'var(--ok)', marginTop: 8 }}>
            Download triggered — check your browser's downloads (or the tab bar for a new tab). If you don't see it after a few seconds, use copy-paste above instead; it's unaffected.
          </div>
        )}
        {typeof downloadState === 'string' && downloadState.startsWith('error:') && (
          <div className="error-box" style={{ marginTop: 8 }}><AlertCircle size={16} /><span>Couldn't build the file ({downloadState.slice(6)}) — the copy-paste data above is unaffected, use that instead.</span></div>
        )}
        <div className="review-actions" style={{ marginTop: 12 }}>
          <button className="btn btn-primary" onClick={selectAll}>Select all</button>
          <button className="btn btn-ghost" onClick={downloadXlsx}><Download size={15} /> Download .xlsx</button>
          <button className="btn btn-ghost" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
}
// Full-size view of the uploaded document — the inline preview during extraction is already sized
// to roughly match the review table, but small handwriting can still need a closer look. Click
// anywhere on the backdrop (or the × ) to close.
function ImageZoomModal({ src, onClose }) {
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(35,38,43,0.8)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20, cursor: 'zoom-out' }} onClick={onClose}>
      <img src={src} alt="Full-size preview" style={{ maxWidth: '95vw', maxHeight: '92vh', objectFit: 'contain', borderRadius: 6, boxShadow: '0 6px 30px rgba(0,0,0,0.4)' }} />
      <button
        onClick={onClose}
        title="Close"
        style={{ position: 'fixed', top: 16, right: 20, width: 32, height: 32, borderRadius: '50%', border: 'none', background: '#fff', color: '#23262b', fontSize: 16, cursor: 'pointer' }}
      >
        ×
      </button>
    </div>
  );
}
/* ============================== main app ============================== */
function FIMSApp() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [loaded, setLoaded] = useState(false);
  const [rawMaterialIn, setRawMaterialIn] = useState([]);
  const [consumption, setConsumption] = useState([]);
  const [production, setProduction] = useState([]);
  const [customerDispatch, setCustomerDispatch] = useState([]);
  const [daburSpecs, setDaburSpecs] = useState([]);
  const [daburPO, setDaburPO] = useState([]);
  const [daburDispatch, setDaburDispatch] = useState([]);
  const [customerSheetsMirror, setCustomerSheetsMirror] = useState([]); // [{id, customer, sheetTab, block, date, opening, production, dispatch, closing}]
  const [trainingExamples, setTrainingExamples] = useState({});
  const [customerMapping, setCustomerMapping] = useState(DEFAULT_CUSTOMER_MAPPING);
  const [productCatalog, setProductCatalog] = useState(DEFAULT_PRODUCT_CATALOG);
  const [abbreviations, setAbbreviations] = useState(DEFAULT_ABBREVIATIONS); // [{id, short, long}]
  const [customerSheetIds, setCustomerSheetIds] = useState([]); // [{id, customer, sheetId}]
  const [customerNameAliases, setCustomerNameAliases] = useState([]); // [{id, alias, customer}]
  const [pushStatus, setPushStatus] = useState({}); // { [customer]: { state: 'idle'|'pushing'|'done'|'error', message, unmatched } }
  const [serviceAccountEmail, setServiceAccountEmail] = useState('');
  // reviewByCustomer holds the last fetched diff against a customer's REAL Sheet (a dry run — see
  // /api/customer-sheets/preview): existingTabNames/existingBlockTitles feed the Sheet-tab/Block
  // datalist suggestions in Pending Review and Customer Sheets. Pushing itself now happens straight
  // from Pending Review (see pushPendingRows/pushCustomerSheetNow) — no separate review-and-edit UI
  // reads this anymore, it's purely a background data source for autocomplete suggestions.
  const [reviewByCustomer, setReviewByCustomer] = useState({}); // { [customer]: { loading, error, tabs, existingTabNames } }
  // Always empty now — nothing sets it since the review-and-edit UI it backed was removed (pushing
  // happens straight from Pending Review, see pushPendingRows) — kept only because applyReviewEdits/
  // getEditedPayload still read it as a harmless no-op rather than reworking that whole call chain.
  const [reviewEdits, setReviewEdits] = useState({}); // { [customer]: { [variantTitle]: { tabNameOverride, rowEdits: { [rowIndex]: {date,production,dispatch} }, deletedRows: { [rowIndex]: true } } } }
  const registerState = { rawMaterialIn, consumption, production, customerDispatch, daburSpecs, daburPO, daburDispatch };
  const registerSetters = { rawMaterialIn: setRawMaterialIn, consumption: setConsumption, production: setProduction, customerDispatch: setCustomerDispatch, daburSpecs: setDaburSpecs, daburPO: setDaburPO, daburDispatch: setDaburDispatch, customerSheetsMirror: setCustomerSheetsMirror };
  useEffect(() => {
    (async () => {
      const entries = await Promise.all(Object.entries(STORAGE_KEYS).map(async ([k, storageKey]) => [k, await loadRegister(storageKey)]));
      const loadedMap = Object.fromEntries(entries);
      Object.entries(loadedMap).forEach(([k, rows]) => registerSetters[k] && registerSetters[k](rows));
      setTrainingExamples(await loadTraining());
      try {
        const r = await window.storage.get(CUSTOMER_MAPPING_KEY, false);
        if (r && r.value) setCustomerMapping(JSON.parse(r.value));
        else await window.storage.set(CUSTOMER_MAPPING_KEY, JSON.stringify(DEFAULT_CUSTOMER_MAPPING), false);
      } catch (e) { /* keep defaults */ }
      try {
        const r2 = await window.storage.get(CATALOG_KEY, false);
        if (r2 && r2.value) setProductCatalog(JSON.parse(r2.value));
        else await window.storage.set(CATALOG_KEY, JSON.stringify(DEFAULT_PRODUCT_CATALOG), false);
      } catch (e) { /* keep defaults */ }
      try {
        const r3 = await window.storage.get(CUSTOMER_SHEET_IDS_KEY, false);
        if (r3 && r3.value) setCustomerSheetIds(JSON.parse(r3.value));
      } catch (e) { /* keep empty — nothing pushed yet */ }
      try {
        const r5 = await window.storage.get(ABBREVIATIONS_KEY, false);
        if (r5 && r5.value) setAbbreviations(JSON.parse(r5.value));
      } catch (e) { /* keep empty — none taught yet */ }
      try {
        const r6 = await window.storage.get(CUSTOMER_NAME_ALIASES_KEY, false);
        if (r6 && r6.value) setCustomerNameAliases(JSON.parse(r6.value));
      } catch (e) { /* keep empty — none taught yet */ }
      try {
        const r4 = await fetch('/api/service-account-email', { credentials: 'include' });
        if (r4.ok) { const j = await r4.json(); setServiceAccountEmail(j.email || ''); }
      } catch (e) { /* shown as blank below; the per-error messages still work without this */ }
      setLoaded(true);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // Backfills the Raw Material Pivot tab for data that was already sitting in the register before the
  // auto-rebuild-on-save behavior existed (putTab in server/lib/sheets.js handles every SAVE from here
  // on). Fires once per session, right after the initial load, purely fire-and-forget — no UI surface
  // for it (no button, no message), matching how the on-save rebuild is silent too. A failure here just
  // means the pivot tab stays as it was; it isn't disruptive to anything else in the app.
  const pivotBackfillFiredRef = useRef(false);
  useEffect(() => {
    if (!loaded || pivotBackfillFiredRef.current || !rawMaterialIn.length) return;
    pivotBackfillFiredRef.current = true;
    fetch('/api/raw-material/rebuild-pivot', { method: 'POST', credentials: 'include' })
      .then(res => { if (!res.ok) console.warn('Raw Material Pivot backfill failed:', res.status); })
      .catch(e => console.warn('Raw Material Pivot backfill failed:', e));
  }, [loaded, rawMaterialIn.length]);
  // Google Sheets allows only 60 write requests/minute per user (300/minute per project — see
  // https://developers.google.com/workspace/sheets/api/limits). Saving on every single keystroke or
  // row edit blows through that almost immediately — e.g. clicking "Confirm all" on a batch of
  // pending stock rows used to fire one full-register rewrite PER ROW. This debounces every save: an
  // edit schedules a write ~1.5s in the future, and each further edit to the same thing within that
  // window just reschedules it, so a whole burst of edits collapses into exactly one network write
  // of the latest state. flushPendingSaves forces any still-pending writes out immediately when the
  // tab is about to be hidden or closed, so a save scheduled just before that isn't silently lost.
  const SAVE_DEBOUNCE_MS = 1500;
  const saveTimerRef = useRef({});
  const pendingSaveRef = useRef({});
  const scheduleSave = useCallback((key, fn) => {
    pendingSaveRef.current[key] = fn;
    if (saveTimerRef.current[key]) clearTimeout(saveTimerRef.current[key]);
    saveTimerRef.current[key] = setTimeout(() => {
      delete saveTimerRef.current[key];
      const run = pendingSaveRef.current[key];
      delete pendingSaveRef.current[key];
      if (run) run();
    }, SAVE_DEBOUNCE_MS);
  }, []);
  const flushPendingSaves = useCallback(() => {
    Object.values(saveTimerRef.current).forEach(t => clearTimeout(t));
    saveTimerRef.current = {};
    const pending = pendingSaveRef.current;
    pendingSaveRef.current = {};
    Object.values(pending).forEach(run => run());
  }, []);
  useEffect(() => {
    const onVisibility = () => { if (document.visibilityState === 'hidden') flushPendingSaves(); };
    document.addEventListener('visibilitychange', onVisibility);
    window.addEventListener('pagehide', flushPendingSaves);
    return () => {
      document.removeEventListener('visibilitychange', onVisibility);
      window.removeEventListener('pagehide', flushPendingSaves);
    };
  }, [flushPendingSaves]);
  const persist = useCallback((registerKey, rows) => {
    scheduleSave(`register:${registerKey}`, () => saveRegister(STORAGE_KEYS[registerKey], rows));
  }, [scheduleSave]);
  // Full-resync semantics for ONE customer's slice of the Customer Sheets Mirror: replaces every row
  // this customer already had with the fresh set just read from their real Sheet (at import or push
  // time — see confirmSheetImport/pushCustomerSheetNow), so a block/row renamed or removed for real
  // never lingers as a stale, unfindable-in-real-life search result. A functional update — not a plain
  // `setCustomerSheetsMirror([...customerSheetsMirror.filter(...), ...fresh])` reading the outer
  // closure — for the exact same reason registerAliases needs one: guards against two customers'
  // imports/pushes landing in the same render tick and silently clobbering each other.
  const replaceCustomerMirrorRows = (customer, freshRowsForCustomer) => {
    setCustomerSheetsMirror(prev => {
      const next = [...prev.filter(r => r.customer !== customer), ...(freshRowsForCustomer || []).map(r => ({ id: genId(), customer, ...r }))];
      persist('customerSheetsMirror', next);
      return next;
    });
  };
  /* -------- Settings: Clear App Data — a UI version of the direct-API wipe used earlier this
     session, so this doesn't need a one-off script every time a clean retest is wanted. Writes go
     out immediately (not through the debounced `persist`/scheduleSave path) since a clear should be
     final the moment it's confirmed, not sit in a 1.5s window where a tab close could drop it.
     Training examples are deliberately NOT clearable here — they improve extraction accuracy rather
     than being test data, same reasoning as when this was first done manually. -------- */
  const CLEAR_GROUPS = [
    { key: 'rawMaterialIn', label: 'Raw Material Register', registerKeys: ['rawMaterialIn'] },
    { key: 'consumption', label: 'Consumption', registerKeys: ['consumption'] },
    { key: 'production', label: 'Production Register', registerKeys: ['production'] },
    { key: 'customerDispatch', label: 'Customer Dispatch Bills', registerKeys: ['customerDispatch'] },
    { key: 'dabur', label: 'Dabur (Specs, PO, Dispatch)', registerKeys: ['daburSpecs', 'daburPO', 'daburDispatch'] },
    { key: 'catalogMapping', label: 'Product Catalog & Customer Mapping', catalogMapping: true },
    { key: 'sheetIds', label: 'Customer Sheet IDs (tracked links to customer sheets — the sheets themselves are untouched)', sheetIds: true },
    { key: 'abbreviations', label: 'Word Abbreviations', abbreviations: true },
  ];
  const [clearSelected, setClearSelected] = useState({});
  const [clearBusy, setClearBusy] = useState(false);
  const [clearMessage, setClearMessage] = useState('');
  const toggleClearGroup = (key) => setClearSelected(prev => ({ ...prev, [key]: !prev[key] }));
  const runClearSelected = async () => {
    const chosen = CLEAR_GROUPS.filter(g => clearSelected[g.key]);
    if (!chosen.length) return;
    const summary = chosen.map(g => g.label).join(', ');
    if (!window.confirm(`Permanently clear: ${summary}?\n\nThis cannot be undone.`)) return;
    setClearBusy(true); setClearMessage('');
    try {
      for (const g of chosen) {
        if (g.registerKeys) {
          for (const rk of g.registerKeys) {
            registerSetters[rk]([]);
            await saveRegister(STORAGE_KEYS[rk], []);
          }
        } else if (g.catalogMapping) {
          setProductCatalog([]); await window.storage.set(CATALOG_KEY, JSON.stringify([]), false);
          setCustomerMapping([]); await window.storage.set(CUSTOMER_MAPPING_KEY, JSON.stringify([]), false);
        } else if (g.sheetIds) {
          setCustomerSheetIds([]); await window.storage.set(CUSTOMER_SHEET_IDS_KEY, JSON.stringify([]), false);
        } else if (g.abbreviations) {
          setAbbreviations([]); await window.storage.set(ABBREVIATIONS_KEY, JSON.stringify([]), false);
        }
      }
      setClearMessage(`Cleared: ${summary}.`);
      setClearSelected({});
    } catch (e) {
      setClearMessage(`Something went wrong partway through (${e.message || 'unknown error'}) — check which registers above still show data and retry just those.`);
    } finally {
      setClearBusy(false);
    }
  };
  // --- Google Sheet maintenance (Settings tab): scan every tab in the main Shyam Adarsh sheet and
  // delete a person-confirmed list. Only "internal" tabs (bookkeeping the app needs but that never
  // renders as a table anywhere — training examples, the customer-sheets search mirror) get
  // pre-checked; "unrecognized" tabs are shown but left unchecked, since only a person looking at the
  // actual name can tell whether one is stray clutter or something still needed. See
  // KNOWN_APP_TAB_KEYS/INTERNAL_ONLY_TAB_KEYS in server/lib/sheets.js for how a tab gets labeled.
  const [sheetTabs, setSheetTabs] = useState(null); // null = not scanned yet this session
  const [sheetTabsBusy, setSheetTabsBusy] = useState(false);
  const [sheetTabsMessage, setSheetTabsMessage] = useState('');
  const [tabsToDelete, setTabsToDelete] = useState({});
  const scanSheetTabs = async () => {
    setSheetTabsBusy(true); setSheetTabsMessage('');
    try {
      const res = await fetch('/api/maintenance/tabs', { credentials: 'include' });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `HTTP ${res.status}`);
      const data = await res.json();
      setSheetTabs(data.tabs || []);
      setTabsToDelete(Object.fromEntries((data.tabs || []).filter(t => t.kind === 'internal').map(t => [t.title, true])));
    } catch (e) {
      setSheetTabsMessage(`Couldn't scan the sheet (${e.message || 'unknown error'}).`);
    } finally {
      setSheetTabsBusy(false);
    }
  };
  const toggleTabToDelete = (title) => setTabsToDelete(prev => ({ ...prev, [title]: !prev[title] }));
  const deleteSelectedTabs = async () => {
    const chosen = Object.keys(tabsToDelete).filter(t => tabsToDelete[t]);
    if (!chosen.length) return;
    if (!window.confirm(`Permanently delete these tabs from the Shyam Adarsh sheet:\n\n${chosen.join('\n')}\n\nThis cannot be undone.`)) return;
    setSheetTabsBusy(true); setSheetTabsMessage('');
    try {
      const res = await fetch('/api/maintenance/delete-tabs', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'include',
        body: JSON.stringify({ tabNames: chosen }),
      });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `HTTP ${res.status}`);
      const data = await res.json();
      setSheetTabsMessage(`Deleted: ${data.deleted.length ? data.deleted.join(', ') : '(none)'}.${data.notFound.length ? ` Already gone: ${data.notFound.join(', ')}.` : ''}`);
      await scanSheetTabs();
    } catch (e) {
      setSheetTabsMessage(`Something went wrong (${e.message || 'unknown error'}).`);
    } finally {
      setSheetTabsBusy(false);
    }
  };
  // --- Date-format cleanup (Settings tab): normalizeDateToDots (above) runs on every FRESH
  // extraction, but a row that entered the ledger before a normalizer bug was fixed, or was typed/
  // edited by hand, can still be sitting there in the wrong shape (e.g. "23.07.2026" next to "12.4.26"
  // in the same register). One-time, on-demand sweep — not automatic on every load — since it's a
  // retroactive cleanup for already-bad data, not an ongoing need once the root cause is fixed.
  const DATE_FIELD_REGISTERS = {
    rawMaterialIn: ['date'], consumption: ['date'], production: ['date'], customerDispatch: ['date'],
    daburPO: ['date', 'delivery_date'], daburDispatch: ['date'],
  };
  const [dateFixBusy, setDateFixBusy] = useState(false);
  const [dateFixMessage, setDateFixMessage] = useState('');
  const normalizeAllDates = async () => {
    setDateFixBusy(true); setDateFixMessage('');
    try {
      let totalFixed = 0;
      const perRegister = [];
      for (const [registerKey, fields] of Object.entries(DATE_FIELD_REGISTERS)) {
        const rows = registerState[registerKey] || [];
        let changedInThisRegister = 0;
        const next = rows.map(r => {
          const patched = { ...r };
          let rowChanged = false;
          fields.forEach(f => {
            const normalized = normalizeDateToDots(r[f]);
            if (normalized !== (r[f] || '')) { patched[f] = normalized; rowChanged = true; }
          });
          if (rowChanged) changedInThisRegister++;
          return patched;
        });
        if (changedInThisRegister > 0) {
          registerSetters[registerKey](next);
          await saveRegister(STORAGE_KEYS[registerKey], next);
          totalFixed += changedInThisRegister;
          perRegister.push(`${changedInThisRegister} in ${registerKey}`);
        }
      }
      setDateFixMessage(totalFixed ? `Fixed ${totalFixed} row(s): ${perRegister.join(', ')}.` : 'Every date was already consistent — nothing to fix.');
    } catch (e) {
      setDateFixMessage(`Something went wrong partway through (${e.message || 'unknown error'}) — safe to run again, it only touches rows that still need it.`);
    } finally {
      setDateFixBusy(false);
    }
  };
  // The Raw Material Pivot tab on the Shyam Adarsh sheet rebuilds itself automatically whenever raw
  // material data saves (see putTab in server/lib/sheets.js) — no button/trigger needed here.
  // --- Raw Material Register view: the app-side equivalent of the Sheet pivot's filter — narrows the
  // per-size tables down to one size and/or one GSM at a time, without summing anything.
  const [rmSizeFilter, setRmSizeFilter] = useState('');
  const [rmGsmFilter, setRmGsmFilter] = useState('');
  const updateRow = (registerKey) => (id, field, value) => {
    registerSetters[registerKey](prev => {
      const next = prev.map(r => r.id === id ? { ...r, [field]: value } : r);
      persist(registerKey, next);
      return next;
    });
  };
  const deleteRow = (registerKey) => (id) => {
    registerSetters[registerKey](prev => {
      const next = prev.filter(r => r.id !== id);
      persist(registerKey, next);
      return next;
    });
  };
  // Every extraction-driven register gets de-duplicated on add — any of them can suffer the same
  // re-photographed-page or overlapping-scan mistake, not just Production/Dispatch. Checked against the
  // CURRENT register via registerState (read from the outer render closure, not a functional-updater
  // `prev` — safe here specifically because addRows is only ever called once per user action, never in
  // a loop the way some other multi-row confirms elsewhere in this app are, so there's no risk of the
  // staleness that pattern would otherwise need guarding against) plus within the new batch itself, so
  // two copies of the same page queued together in one "Confirm all" don't both land either. Returns how
  // many were skipped so the caller can tell the user.
  const DEDUP_REGISTERS = new Set(['rawMaterialIn', 'consumption', 'production', 'customerDispatch', 'daburSpecs', 'daburPO', 'daburDispatch']);
  const addRows = (registerKey, rows) => {
    let toAdd = rows;
    let skipped = 0;
    if (DEDUP_REGISTERS.has(registerKey)) {
      const existingKeys = new Set((registerState[registerKey] || []).map(rowDedupKey));
      const seenInBatch = new Set();
      toAdd = rows.filter(r => {
        const key = rowDedupKey(r);
        if (existingKeys.has(key) || seenInBatch.has(key)) { skipped++; return false; }
        seenInBatch.add(key);
        return true;
      });
    }
    if (toAdd.length) {
      registerSetters[registerKey](prev => {
        const next = [...prev, ...toAdd];
        persist(registerKey, next);
        return next;
      });
    }
    return skipped;
  };
  // Collapsible sidebar — mainly so the upload preview (and the review table next to it) has more
  // room to breathe on a laptop-width screen. A pure UI preference, not app data, so it's kept in
  // plain localStorage (not the Google-Sheets-backed window.storage shim used for real business data)
  // and just falls back to expanded if that read ever fails.
  const [sidebarCollapsed, setSidebarCollapsed] = useState(() => {
    try { return localStorage.getItem('fims_sidebar_collapsed') === '1'; } catch (e) { return false; }
  });
  useEffect(() => {
    try { localStorage.setItem('fims_sidebar_collapsed', sidebarCollapsed ? '1' : '0'); } catch (e) { /* noop */ }
  }, [sidebarCollapsed]);
  // Topbar search — replaced the old dedicated "Order Availability Check" tab. Instead of a single
  // party/keyword form scoped to just Production + Dispatch, this is a plain substring search across
  // every flat register PLUS Customer Stock (produced/dispatched/balance by customer+item), so the same
  // box that finds "was this order dispatched" also finds a mill slip, a Dabur PO, anything else.
  const [globalQuery, setGlobalQuery] = useState('');
  /* -------- upload & extract state -------- */
  const [docType, setDocType] = useState(DOCUMENT_TYPES[0].key);
  const [preview, setPreview] = useState(null);
  const [base64Img, setBase64Img] = useState(null);
  const [zoomedImage, setZoomedImage] = useState(null); // full-size lightbox for the upload preview, so tiny handwriting is actually legible
  const [queuedPages, setQueuedPages] = useState([]); // [{id, dataUrl, base64, label}] — one entry per image file, or per PDF page
  const [skippedPages, setSkippedPages] = useState([]); // pages auto-detected as a non-original copy / e-Way Bill — never sent to the API
  const [queuedIndex, setQueuedIndex] = useState(0);
  const [pdfLoading, setPdfLoading] = useState(false);
  const [extracting, setExtracting] = useState(false);
  const [cancelRequested, setCancelRequested] = useState(false);
  const abortControllerRef = useRef(null);
  const [errorMsg, setErrorMsg] = useState('');
  const [fileResults, setFileResults] = useState([]); // [{id, label, status: pending|extracting|done|error, rows, originalRows, error}]
  const [activeResultIndex, setActiveResultIndex] = useState(0);
  const activeConfig = DOCUMENT_TYPES.find(d => d.key === docType);
  const isDispatchDocType = docType === 'dispatch_bill' || docType === 'dabur_dispatch';
  const handleFiles = async (fileList) => {
    setErrorMsg(''); setFileResults([]); setActiveResultIndex(0); setQueuedPages([]); setSkippedPages([]); setQueuedIndex(0);
    const files = Array.from(fileList || []);
    if (!files.length) return;
    setPdfLoading(true);
    try {
      const allPages = [];
      const failedFiles = [];
      for (const file of files) {
        const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
        try {
          if (isPdf) {
            const pages = await pdfFileToPages(file);
            if (!pages.length) throw new Error('NO_PAGES');
            pages.forEach((p, idx) => allPages.push({
              id: genId(), dataUrl: p.dataUrl, base64: p.base64, copyLabel: p.copyLabel,
              label: pages.length > 1 ? `${file.name} — page ${idx + 1}` : file.name,
            }));
          } else {
            const { dataUrl, base64 } = await resizeImageToBase64(file);
            allPages.push({ id: genId(), dataUrl, base64, copyLabel: 'unknown', label: file.name });
          }
        } catch (e) {
          failedFiles.push(file.name);
        }
      }
      if (!allPages.length) throw new Error('NO_READABLE_FILES');
      // For dispatch-bill uploads, a single bill's PDF usually contains 5-6 near-identical copies
      // (Original, Duplicate, Triplicate, Extra, plus e-Way Bill pages) that we only ever want the
      // Original from anyway (see the extraction prompt). Sending every copy to the vision API just
      // to get an empty result back was burning through the rate limit on a single bill — since the
      // PDF has a real text layer, we can tell copies apart for free and skip the API call entirely
      // for ones we're always going to discard. Pages we can't confidently classify (e.g. photographed
      // JPGs with no text layer) stay queued as normal, nothing is silently dropped.
      let toQueue = allPages;
      let toSkip = [];
      if (isDispatchDocType) {
        const original = allPages.filter(p => p.copyLabel === 'original');
        const undetermined = allPages.filter(p => p.copyLabel === 'unknown');
        const nonOriginal = allPages.filter(p => p.copyLabel && !['original', 'unknown'].includes(p.copyLabel));
        // if nothing was confidently detected as the Original, don't skip anything — safer to let
        // extraction run on every page (as before) than to risk silently discarding the real bill
        if (original.length > 0) { toQueue = [...original, ...undetermined]; toSkip = nonOriginal; }
      }
      setQueuedPages(toQueue);
      setSkippedPages(toSkip);
      setQueuedIndex(0);
      setPreview(toQueue[0].dataUrl);
      setBase64Img(toQueue[0].base64);
      const notes = [];
      if (toSkip.length) notes.push(`Detected and skipped ${toSkip.length} duplicate/e-Way Bill page(s) automatically — no API calls used for those. ${toQueue.length} original page(s) left to extract.`);
      if (failedFiles.length) notes.push(`Couldn't read: ${failedFiles.join(', ')}. If any is a HEIC photo, re-save as JPG/PNG first. The rest loaded fine below.`);
      if (notes.length) setErrorMsg(notes.join(' '));
    } catch (e) {
      setErrorMsg('Could not read any of these files. If they’re HEIC photos from an iPhone, re-save as JPG/PNG first (Settings → Camera → Formats → Most Compatible), or take a screenshot and upload that instead.');
      setPreview(null); setBase64Img(null);
    } finally {
      setPdfLoading(false);
    }
  };
  const restoreSkippedPage = (id) => {
    setSkippedPages(prev => {
      const page = prev.find(p => p.id === id);
      if (page) setQueuedPages(q => [...q, page]);
      return prev.filter(p => p.id !== id);
    });
  };
  const selectQueuedPage = (idx) => {
    setQueuedIndex(idx);
    setPreview(queuedPages[idx].dataUrl);
    setBase64Img(queuedPages[idx].base64);
    setActiveResultIndex(idx);
  };
  // Detects both textual "rate limit" wording in an error body and a raw 429 status code —
  // some responses phrase it as "rate_limit_error" (underscore, no space) which the old text-only
  // check could miss, which is why dispatch-bill batches (many requests back-to-back) were slipping
  // through as generic failures instead of being treated as rate limits.
  const isRateLimitError = (e) => e?.status === 429 || /rate limit/i.test(e?.message || '') || /EXTRACT_HTTP_429/.test(e?.message || '');
  // See SERVER_ERROR_BACKOFFS_MS above for what these actually are (Render's proxy, not Anthropic).
  const isTransientServerError = (e) => [502, 503, 504].includes(e?.status) || /EXTRACT_HTTP_50[234]/.test(e?.message || '');
  // attempts=3 (was 2): EXTRACT_EMPTY ("no text block in response") shows up occasionally and is a
  // transient API hiccup, not a property of the image — a plain retry usually succeeds. One extra
  // attempt costs little (it's only reached on a genuine failure) and cuts down on files landing in
  // the error state that otherwise just needed a second try, which the person had to trigger by hand.
  const extractWithRetry = async (prompt, base64, signal, attempts = 3) => {
    let lastErr;
    for (let i = 0; i < attempts; i++) {
      try {
        return await callClaudeExtract(prompt, base64, signal);
      } catch (e) {
        lastErr = e;
        // let the caller handle rate limits/gateway blips with a longer backoff, cancellation with a
        // clean stop — not this quick retry (600ms is nowhere near enough for either of those to clear)
        if (isCancelled(e) || isRateLimitError(e) || isTransientServerError(e)) break;
        if (i < attempts - 1) await sleep(600, signal); // brief pause before retry
      }
    }
    throw lastErr;
  };
  // Runs one extraction, and if it's specifically a rate-limit or transient-server error, waits and
  // retries automatically (growing backoff, on whichever schedule matches the error) before giving up
  // — both are usually gone within seconds to a couple minutes, so a wait-and-retry recovers most of
  // the time without the person having to notice the failure and click retry themselves.
  // onTick fires every second during a wait so the UI can show a live countdown instead of a frozen
  // message, and the whole wait aborts immediately if `signal` is cancelled (the person hit Cancel).
  const extractWithRateLimitBackoff = async (prompt, base64, signal, onWaiting, onTick) => {
    const maxAttempts = Math.max(RATE_LIMIT_BACKOFFS_MS.length, SERVER_ERROR_BACKOFFS_MS.length);
    for (let i = 0; i <= maxAttempts; i++) {
      try {
        return await extractWithRetry(prompt, base64, signal);
      } catch (e) {
        const kind = isRateLimitError(e) ? 'rate_limit' : isTransientServerError(e) ? 'server_error' : null;
        const backoffs = kind === 'rate_limit' ? RATE_LIMIT_BACKOFFS_MS : kind === 'server_error' ? SERVER_ERROR_BACKOFFS_MS : null;
        if (backoffs && i < backoffs.length) {
          const waitMs = e.retryAfterMs && e.retryAfterMs > 0 ? e.retryAfterMs : backoffs[i];
          if (onWaiting) onWaiting(waitMs, i + 1, backoffs.length, kind);
          let remaining = Math.ceil(waitMs / 1000);
          if (onTick) onTick(remaining);
          const tickTimer = setInterval(() => { remaining -= 1; if (onTick) onTick(Math.max(remaining, 0)); }, 1000);
          try {
            await sleep(waitMs, signal);
          } finally {
            clearInterval(tickTimer);
          }
          continue;
        }
        throw e;
      }
    }
  };
  const cancelExtraction = () => {
    setCancelRequested(true);
    if (abortControllerRef.current) abortControllerRef.current.abort();
  };
  const removeQueuedPage = (id) => {
    setQueuedPages(prev => {
      const next = prev.filter(p => p.id !== id);
      const clampedIdx = Math.max(0, Math.min(queuedIndex, next.length - 1));
      if (next.length) { setPreview(next[clampedIdx].dataUrl); setBase64Img(next[clampedIdx].base64); }
      else { setPreview(null); setBase64Img(null); }
      setQueuedIndex(clampedIdx);
      return next;
    });
    setFileResults(prev => prev.filter(r => r.id !== id));
  };
  // The friendly explanation appends the raw server response (what e.message actually contains, e.g.
  // "EXTRACT_HTTP_429: {...}") so the real diagnostic text is right there to copy over if needed,
  // instead of requiring browser DevTools.
  const retryExhaustedExplainer = (e) => {
    const friendly = isTransientServerError(e)
      ? 'The server hosting this app (Render, not Anthropic) was still unreachable even after automatic retries — usually a brief restart, redeploy, or the service waking back up from being idle. This almost always clears within well under a minute; wait a bit and try again.'
      : 'Still hitting a rate limit even after automatic retries — unusual at normal usage (this app\'s API key gets 1,000 requests/minute). Wait a minute and try again, or check the Anthropic Console for the account this key belongs to.';
    const detail = e && e.message ? String(e.message).slice(0, 300) : '';
    return detail ? `${friendly} (Server said: ${detail})` : friendly;
  };
  const runExtraction = async () => {
    // Defensive guard against a double-fire (e.g. a fast double-click landing before React re-renders
    // the button's disabled state) — without this, two overlapping calls could each open their own
    // AbortController and fire their own API request for what was visually a single click.
    if (extracting || !base64Img || !queuedPages.length) return;
    const current = queuedPages[queuedIndex];
    setExtracting(true); setErrorMsg(''); setCancelRequested(false);
    abortControllerRef.current = new AbortController();
    setFileResults(prev => {
      const existing = prev.find(r => r.id === current.id);
      const base = existing || { id: current.id, label: current.label, status: 'pending', rows: [], originalRows: [], error: '' };
      const next = prev.some(r => r.id === current.id) ? prev.map(r => r.id === current.id ? { ...base, status: 'extracting' } : r) : [...prev, { ...base, status: 'extracting' }];
      return next;
    });
    setActiveResultIndex(queuedIndex);
    try {
      const basePrompt = activeConfig.systemPrompt.replace('{{PRODUCT_CATALOG}}', buildCatalogText(productCatalog)).replace('{{ABBREVIATIONS}}', buildAbbreviationsText(abbreviations));
      const promptWithTraining = buildPromptWithTraining(basePrompt, trainingExamples[docType]);
      const { raw, truncated } = await extractWithRateLimitBackoff(promptWithTraining, base64Img, abortControllerRef.current.signal, (waitMs, attempt, total, kind) => {
        const label = kind === 'server_error' ? 'Server temporarily unreachable' : 'Rate limit hit';
        setErrorMsg(`${label} — waiting ${Math.round(waitMs / 1000)}s and retrying automatically (attempt ${attempt} of ${total})… Click Cancel below if you'd rather stop.`);
      }, (secondsLeft) => {
        setErrorMsg(prevMsg => prevMsg.replace(/waiting \d+s/, `waiting ${secondsLeft}s`));
      });
      const rows = activeConfig.shape(raw);
      if (!rows.length) setErrorMsg('No line items found on this page. If this is a duplicate copy or an e-Way Bill page, that’s expected — just move to the next one. Otherwise, try a clearer photo or crop closer to the table.');
      else if (truncated) setErrorMsg(`Claude's response was cut off before it finished this page — it may have more rows than the ${rows.length} shown below. Check against the original, and use "Re-extract this file" if anything's missing.`);
      else setErrorMsg('');
      setFileResults(prev => prev.map(r => r.id === current.id ? { ...r, status: 'done', rows, originalRows: rows.map(x => ({ ...x })), truncated } : r));
    } catch (e) {
      console.error('Extraction error:', e);
      if (isCancelled(e)) {
        setErrorMsg('Cancelled.');
        setFileResults(prev => prev.map(r => r.id === current.id ? { ...r, status: 'pending', error: '' } : r));
      } else if (isRateLimitError(e) || isTransientServerError(e)) {
        setErrorMsg(retryExhaustedExplainer(e));
        setFileResults(prev => prev.map(r => r.id === current.id ? { ...r, status: 'pending', error: '' } : r));
      } else {
        const msg = e.message || 'unknown error';
        setErrorMsg(`Extraction failed (${msg}). Retried automatically already — try a clearer photo, better lighting, or make sure the whole document fits in frame.`);
        setFileResults(prev => prev.map(r => r.id === current.id ? { ...r, status: 'error', error: msg } : r));
      }
    } finally {
      setExtracting(false);
      abortControllerRef.current = null;
    }
  };
  const extractQueue = async (targets) => {
    setExtracting(true); setErrorMsg(''); setCancelRequested(false);
    abortControllerRef.current = new AbortController();
    const signal = abortControllerRef.current.signal;
    const basePrompt = activeConfig.systemPrompt.replace('{{PRODUCT_CATALOG}}', buildCatalogText(productCatalog)).replace('{{ABBREVIATIONS}}', buildAbbreviationsText(abbreviations));
    const promptWithTraining = buildPromptWithTraining(basePrompt, trainingExamples[docType]);
    const gapMs = BATCH_REQUEST_GAP_MS;
    let anySucceeded = false;
    let anyTruncated = false;
    let haltedForRetry = false;
    let lastHaltError = null;
    let cancelled = false;
    for (let t = 0; t < targets.length; t++) {
      const p = targets[t];
      // skip files that were removed from the queue mid-batch (e.g. via the × on a thumbnail)
      if (!queuedPages.some(q => q.id === p.id)) continue;
      try {
        // small pacing gap between requests (skip before the very first one) so a big multi-copy
        // dispatch-bill upload doesn't fire a literal burst of requests in the same instant.
        if (t > 0) await sleep(gapMs, signal);
      } catch (e) { cancelled = true; break; }
      setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'extracting' } : r));
      try {
        const { raw, truncated } = await extractWithRateLimitBackoff(promptWithTraining, p.base64, signal, (waitMs, attempt, total, kind) => {
          const label = kind === 'server_error' ? `Server temporarily unreachable while extracting "${p.label}"` : `Rate limit hit on "${p.label}"`;
          setErrorMsg(`${label} — waiting ${Math.round(waitMs / 1000)}s and retrying automatically (attempt ${attempt} of ${total})… Click Cancel below if you'd rather stop.`);
        }, (secondsLeft) => {
          setErrorMsg(prevMsg => prevMsg.replace(/waiting \d+s/, `waiting ${secondsLeft}s`));
        });
        const rows = activeConfig.shape(raw);
        anySucceeded = true;
        if (truncated) anyTruncated = true;
        setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'done', rows, originalRows: rows.map(x => ({ ...x })), truncated } : r));
      } catch (e) {
        console.error('Extraction error:', e);
        if (isCancelled(e)) {
          cancelled = true;
          setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'pending', error: '' } : r));
          break;
        }
        if (isRateLimitError(e) || isTransientServerError(e)) {
          haltedForRetry = true;
          lastHaltError = e;
          setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'pending', error: '' } : r));
          break; // stop hammering the same wall even after retries — leftover files stay pending, resumable
        }
        setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'error', error: e.message || 'unknown error' } : r));
      }
    }
    if (cancelled) {
      setErrorMsg('Cancelled — files not yet extracted are still queued below, untouched. Remove any you don\'t want with the × on its thumbnail, or hit "Retry remaining" to pick back up.');
    } else if (haltedForRetry) {
      setErrorMsg(`${retryExhaustedExplainer(lastHaltError)} Anything already extracted in this batch is safe and untouched — use "Retry remaining" below once you're ready to continue.`);
    } else if (!anySucceeded && targets.length) {
      setErrorMsg('Every file in this batch failed to extract. Check your connection and try again, or extract one file at a time to isolate the problem.');
    } else if (anyTruncated) {
      setErrorMsg('One or more files in this batch had a response cut off before it finished the page — they\'re marked below. Check each against the original and use "Re-extract this file" for any that look incomplete.');
    } else {
      setErrorMsg('');
    }
    setExtracting(false);
    abortControllerRef.current = null;
  };
  const runExtractionAllQueued = async () => {
    if (extracting || !queuedPages.length) return;
    const initial = queuedPages.map(p => ({ id: p.id, label: p.label, status: 'pending', rows: [], originalRows: [], error: '' }));
    setFileResults(initial);
    setActiveResultIndex(0);
    await extractQueue(queuedPages);
  };
  const runExtractionRemaining = async () => {
    if (extracting) return;
    const targets = queuedPages.filter(p => {
      const r = fileResults.find(fr => fr.id === p.id);
      return r && (r.status === 'pending' || r.status === 'error');
    });
    if (!targets.length) return;
    await extractQueue(targets);
  };
  // Re-runs extraction for ONE already-"done" file — the response to being told a page's response was
  // cut off (page.truncated) is to just try again, since a fresh request isn't guaranteed to hit the
  // same wall a second time. Resets that file back to 'pending' first (clearing its old rows/truncated
  // flag) rather than appending to what's already there, so a second cut-off attempt can't leave stale
  // rows from the first mixed in with a fresh partial result.
  const reextractFile = async (id) => {
    if (extracting) return;
    const page = queuedPages.find(p => p.id === id);
    if (!page) return;
    setFileResults(prev => prev.map(r => r.id === id ? { ...r, status: 'pending', rows: [], originalRows: [], error: '', truncated: false } : r));
    await extractQueue([page]);
  };
  // Walks the ORIGINAL rows (not the final ones) specifically so a row deleted entirely during review
  // — e.g. a blank pre-ruled line the model hallucinated a row for — still gets picked up as a lesson
  // (recorded as {after: null}, see buildPromptWithTraining), not just rows that survived with an
  // edited field. Previously only field-level edits were recorded, so "delete the junk rows and
  // confirm" taught the model nothing and the same blank rows kept coming back on the next upload.
  const recordCorrections = (originalRows, finalRows) => {
    if (!originalRows || !originalRows.length) return;
    const corrections = [];
    const finalById = new Map(finalRows.map(r => [r.id, r]));
    originalRows.forEach(orig => {
      const { id: _oi, ...beforeClean } = orig;
      const row = finalById.get(orig.id);
      if (!row) { corrections.push({ before: beforeClean, after: null }); return; }
      const changed = Object.keys(row).some(k => k !== 'id' && String(orig[k] ?? '') !== String(row[k] ?? ''));
      if (changed) {
        const { id: _i2, ...afterClean } = row;
        corrections.push({ before: beforeClean, after: afterClean });
      }
    });
    if (!corrections.length) return;
    setTrainingExamples(prev => {
      const list = [...(prev[docType] || []), ...corrections].slice(-MAX_EXAMPLES_STORED);
      const next = { ...prev, [docType]: list };
      saveTraining(next);
      return next;
    });
  };
  const clearTrainingForType = () => {
    setTrainingExamples(prev => {
      const next = { ...prev };
      delete next[docType];
      saveTraining(next);
      return next;
    });
  };
  const confirmPage = (idx) => {
    const page = fileResults[idx];
    if (!page || !page.rows.length) return;
    recordCorrections(page.originalRows, page.rows);
    const skipped = addRows(activeConfig.register, page.rows);
    setFileResults(prev => prev.filter((_, i) => i !== idx));
    setActiveResultIndex(prev => Math.max(0, Math.min(prev, fileResults.length - 2)));
    if (skipped) window.alert(`${skipped} row${skipped === 1 ? '' : 's'} exactly matched one already in the register and ${skipped === 1 ? "wasn't" : "weren't"} added again — looks like this page (or part of it) was uploaded before.`);
  };
  const confirmAllPages = () => {
    const donePages = fileResults.filter(r => r.status === 'done' && r.rows.length);
    donePages.forEach(page => recordCorrections(page.originalRows, page.rows));
    const allRows = donePages.flatMap(p => p.rows);
    const skipped = allRows.length ? addRows(activeConfig.register, allRows) : 0;
    setFileResults([]); setActiveResultIndex(0); setPreview(null); setBase64Img(null); setQueuedPages([]);
    if (skipped) window.alert(`${skipped} row${skipped === 1 ? '' : 's'} exactly matched one already in the register and ${skipped === 1 ? "wasn't" : "weren't"} added again — looks like a page (or part of it) was uploaded before.`);
  };
  const discardPage = (idx) => {
    setFileResults(prev => prev.filter((_, i) => i !== idx));
    setActiveResultIndex(prev => Math.max(0, Math.min(prev, fileResults.length - 2)));
  };
  const discardAllPages = () => { setFileResults([]); setActiveResultIndex(0); };
  const updateReviewCell = (pageIdx, rowId, field, value) => {
    setFileResults(prev => prev.map((page, i) => i !== pageIdx ? page : { ...page, rows: page.rows.map(r => r.id === rowId ? { ...r, [field]: value } : r) }));
  };
  const deleteReviewRow = (pageIdx, rowId) => {
    setFileResults(prev => prev.map((page, i) => i !== pageIdx ? page : { ...page, rows: page.rows.filter(r => r.id !== rowId) }));
  };
  /* -------- raw material balance -------- */
  const balanceRows = (() => {
    const map = {};
    rawMaterialIn.forEach(r => {
      const key = `${(r.size || '').trim()}|${(r.gsm || '').trim()}`;
      if (!map[key]) map[key] = { id: key, size: r.size, gsm: r.gsm, weight_in: 0, weight_consumed: 0 };
      map[key].weight_in += num(r.weight_kg);
    });
    consumption.forEach(c => {
      const key = `${(c.size || '').trim()}|${(c.gsm || '').trim()}`;
      if (!map[key]) map[key] = { id: key, size: c.size, gsm: c.gsm, weight_in: 0, weight_consumed: 0 };
      map[key].weight_consumed += num(c.weight_consumed);
    });
    return Object.values(map).map(v => ({ ...v, balance: v.weight_in - v.weight_consumed }));
  })();
  // One group per distinct size, exactly as the physical register book itself is organized (each
  // size gets its own page). Deliberately NOT summed the way balanceRows above is — every individual
  // mill-slip line item stays its own row; grouping only changes which table it's displayed in, never
  // the data itself. Row order within a group is left to EditableTable's own date sort (its default),
  // rather than sorted here too, to avoid sorting the same rows twice.
  const rawMaterialBySize = (() => {
    const groups = {};
    rawMaterialIn.forEach(r => {
      const raw = (r.size || '').trim();
      const n = parseFloat(raw);
      // Collapse an all-zero decimal tail ("62.00" -> "62") so it groups with plain "62" instead of
      // splitting into its own table — String(Number(...)) drops trailing zeros but keeps a real
      // fraction intact ("56.50" -> "56.5"), so nothing here ever changes the actual size value.
      const size = raw ? (Number.isNaN(n) ? raw : String(n)) : '(no size)';
      (groups[size] = groups[size] || []).push(r);
    });
    return Object.entries(groups)
      .map(([size, rows]) => ({ size, rows }))
      .sort((a, b) => {
        const na = parseFloat(a.size), nb = parseFloat(b.size);
        if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
        return a.size.localeCompare(b.size, undefined, { numeric: true });
      });
  })();
  // Distinct GSM values present, for the filter dropdown below — sorted numerically the same way
  // sizes are, so "100" doesn't land before "62" the way a plain string sort would.
  const rawMaterialGsmOptions = Array.from(new Set(rawMaterialIn.map(r => (r.gsm || '').trim()).filter(Boolean)))
    .sort((a, b) => {
      const na = parseFloat(a), nb = parseFloat(b);
      if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
      return a.localeCompare(b, undefined, { numeric: true });
    });
  // The app-side "pivot" — narrows the per-size tables down to just the selected size and/or GSM,
  // same as the Sheet pivot's filter dropdowns, without summing or hiding anything beyond that.
  const rawMaterialGroupsFiltered = rawMaterialBySize
    .filter(g => !rmSizeFilter || g.size === rmSizeFilter)
    .map(g => ({ ...g, rows: rmGsmFilter ? g.rows.filter(r => (r.gsm || '').trim() === rmGsmFilter) : g.rows }))
    .filter(g => g.rows.length > 0);
  /* -------- dabur PO pending calc -------- */
  const PO_TOLERANCE = 0.10; // ±10% — a PO counts as fulfilled once dispatched qty reaches 90% of ordered qty
  const daburPOWithPending = daburPO.map(po => {
    const matched = daburDispatch.filter(d => d.buyer_order_no && po.po_number && d.buyer_order_no.replace(/\s/g, '') === String(po.po_number).replace(/\s/g, ''));
    const dispatchedQty = matched.reduce((s, d) => s + num(d.quantity), 0);
    const orderedQty = num(po.quantity);
    const fulfilled = orderedQty > 0 && dispatchedQty >= orderedQty * (1 - PO_TOLERANCE);
    const shortfall = orderedQty - dispatchedQty;
    return { ...po, dispatched_qty: dispatchedQty, pending_qty: fulfilled ? 0 : Math.max(0, shortfall), fulfilled };
  });
  /* -------- customer mapping (editable) -------- */
  const persistCustomerMapping = (next) => {
    setCustomerMapping(next);
    scheduleSave('customerMapping', () => window.storage.set(CUSTOMER_MAPPING_KEY, JSON.stringify(next), false).catch(() => {}));
  };
  const addMappingRow = () => persistCustomerMapping([...customerMapping, { id: genId(), keyword: '', customer: '' }]);
  const updateMappingRow = (id, field, value) => persistCustomerMapping(customerMapping.map(r => r.id === id ? { ...r, [field]: value } : r));
  const deleteMappingRow = (id) => persistCustomerMapping(customerMapping.filter(r => r.id !== id));
  // "Merge duplicate customer" form state — was a per-customer inline dropdown on the (now removed)
  // Customer Ledgers review panel; moved here since it's a genuine recovery tool worth keeping, not
  // something that only made sense next to a review UI that no longer exists.
  const [mergeFromCustomer, setMergeFromCustomer] = useState('');
  const [mergeToCustomer, setMergeToCustomer] = useState('');
  /* -------- word abbreviations (editable; "J" -> "Jumbo", "CONT" -> "Container", etc.) -------- */
  const persistAbbreviations = (next) => {
    setAbbreviations(next);
    scheduleSave('abbreviations', () => window.storage.set(ABBREVIATIONS_KEY, JSON.stringify(next), false).catch(() => {}));
  };
  const addAbbreviationRow = () => persistAbbreviations([...abbreviations, { id: genId(), short: '', long: '' }]);
  const updateAbbreviationRow = (id, field, value) => persistAbbreviations(abbreviations.map(a => a.id === id ? { ...a, [field]: value } : a));
  const deleteAbbreviationRow = (id) => persistAbbreviations(abbreviations.filter(a => a.id !== id));
  /* -------- customer name aliases (a dispatch bill's/register's raw "as written" customer text ->
     the one real known customer it means, e.g. "Bindal technopolymer pvt. ltd." -> "BINDAL STOCK
     1.08.26") -------- */
  // Functional setState updater (never a plain `setCustomerNameAliases([...customerNameAliases, entry])`
  // reading the outer closure) — required because pushPendingRows calls this once per row via
  // confirmStockRow in a tight synchronous loop, and a closure-based read would have every call but
  // the last one silently clobber the ones before it.
  const registerCustomerNameAlias = (alias, customer) => {
    const a = (alias || '').trim();
    const c = (customer || '').trim();
    if (!a || !c || a.toLowerCase() === c.toLowerCase()) return;
    const key = a.toLowerCase();
    setCustomerNameAliases(prev => {
      if (prev.some(x => (x.alias || '').trim().toLowerCase() === key)) return prev;
      const next = [...prev, { id: genId(), alias: a, customer: c }];
      scheduleSave('customerNameAliases', () => window.storage.set(CUSTOMER_NAME_ALIASES_KEY, JSON.stringify(next), false).catch(() => {}));
      return next;
    });
  };
  // Recovery tool for exactly the "phantom customer" situation the dropdown above now mostly prevents
  // going forward: an existing bucket of confirmed Production/Dispatch rows sitting under a name that
  // turned out to be a duplicate/misspelling of a real customer (e.g. "Bindal technopolymer pvt. ltd."
  // instead of "BINDAL STOCK 1.08.26"). Reassigns every confirmed row, catalog entry, mapping rule, and
  // Sheet ID from one name to the other, then teaches the alias so the same raw text never creates a
  // new duplicate again. Nothing in the underlying Production Register or Dispatch Bills is deleted —
  // only which customer each row is attributed to changes.
  const mergeCustomerInto = (fromCustomer, toCustomer) => {
    const from = (fromCustomer || '').trim();
    const to = (toCustomer || '').trim();
    if (!from || !to || from.toLowerCase() === to.toLowerCase()) return;
    if (!window.confirm(`Merge everything under "${from}" into "${to}"?\n\nThis moves all their confirmed Production Register and Dispatch Bill rows, Product Catalog entries, Customer Mapping rules, and Sheet ID (if any) over to "${to}", and remembers "${from}" as an alias for "${to}" going forward. The underlying register/dispatch rows themselves are never deleted — only the customer they're attributed to changes.`)) return;
    ['production', 'customerDispatch'].forEach(key => {
      registerSetters[key](prev => {
        const next = prev.map(r => r.confirmedCustomer === from ? { ...r, confirmedCustomer: to } : r);
        persist(key, next);
        return next;
      });
    });
    if (productCatalog.some(c => c.customer === from)) persistCatalog(productCatalog.map(c => c.customer === from ? { ...c, customer: to } : c));
    if (customerMapping.some(r => r.customer === from)) persistCustomerMapping(customerMapping.map(r => r.customer === from ? { ...r, customer: to } : r));
    if (customerSheetIds.some(c => c.customer === from) && !getCustomerSheetId(to).trim()) {
      persistCustomerSheetIds(customerSheetIds.map(c => c.customer === from ? { ...c, customer: to } : c));
    } else if (customerSheetIds.some(c => c.customer === from)) {
      persistCustomerSheetIds(customerSheetIds.filter(c => c.customer !== from));
    }
    registerCustomerNameAlias(from, to);
  };
  // Whole-word, case-insensitive replacement of every taught abbreviation — used both to normalize
  // descriptions for catalog/variant matching (see normalizeForCatalogMatch/normalizeVariant below)
  // and available for anything else that wants a "read as the person would say it" version of a raw
  // description. Word-boundary matching so "J" only matches the standalone letter, never part of a
  // longer word.
  const applyAbbreviations = (s) => {
    let out = s || '';
    abbreviations.forEach(a => {
      const short = (a.short || '').trim();
      const long = (a.long || '').trim();
      if (!short || !long) return;
      const escaped = short.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      let pattern = `\\b${escaped}\\b`;
      // Most real abbreviations are the long form's own opening word ("Jeera" -> "Jeera Dhamal",
      // "CONT" -> "Container") — without this guard, text that was ALREADY written out in full gets
      // re-expanded on top of itself ("Jeera Dhamal 35g" -> "Jeera Dhamal Dhamal 35g"), which then fails
      // to match anything. So if the long form starts with the short form, skip an occurrence that's
      // already immediately followed by the rest of the long form — it's already expanded.
      if (long.length > short.length && long.toLowerCase().startsWith(short.toLowerCase())) {
        const rest = long.slice(short.length).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        pattern += `(?!${rest})`;
      }
      out = out.replace(new RegExp(pattern, 'gi'), long);
    });
    return out;
  };
  /* -------- product catalog (editable; populated via Customer Sheets tab's Sheet-ID import) -------- */
  const persistCatalog = (next) => {
    setProductCatalog(next);
    scheduleSave('catalog', () => window.storage.set(CATALOG_KEY, JSON.stringify(next), false).catch(() => {}));
  };
  // Shared dedup keys for the Sheet-ID import flow below — trims BOTH sides before comparing.
  // Catalog items imported from a Sheet are always already trimmed at extraction time, but a row
  // added by hand through the UI might have stray leading/trailing whitespace; without trimming both
  // sides symmetrically, a manually-entered "Hit & Run 120g " with a trailing space wouldn't match an
  // incoming "Hit & Run 120g" and would get silently re-added as a duplicate. Same reasoning for
  // mapping keywords.
  const catalogDedupKey = (customer, item) => `${(customer || '').trim().toLowerCase()}||${(item || '').trim().toLowerCase()}`;
  const mappingDedupKey = (keyword) => (keyword || '').trim().toLowerCase();
  // Deleting or renaming a catalog item ("alias") also keeps its paired exact-match Customer Mapping
  // keyword rule in sync (if one exists) — same idea everywhere in the app that creates one of these
  // pairs (see registerAliases below): a name that routes to a Sheet tab and a name that routes to a
  // customer should never be able to silently drift apart just because they're stored in two arrays.
  // Used by both the Known Product Catalog table (Customer Mapping tab) and the dedicated Aliases tab
  // — one implementation, so editing/deleting behaves identically no matter which screen you're on.
  const deleteCatalogItem = (id) => {
    const entry = productCatalog.find(c => c.id === id);
    if (!entry) return;
    persistCatalog(productCatalog.filter(c => c.id !== id));
    const exact = mappingDedupKey(entry.item);
    const rule = customerMapping.find(r => mappingDedupKey(r.keyword) === exact && r.customer === entry.customer);
    if (rule) persistCustomerMapping(customerMapping.filter(r => r.id !== rule.id));
  };
  const updateCatalogItem = (id, field, value) => {
    const entry = productCatalog.find(c => c.id === id);
    if (!entry) return;
    persistCatalog(productCatalog.map(c => c.id === id ? { ...c, [field]: value } : c));
    if (field === 'item' || field === 'customer') {
      const oldExact = mappingDedupKey(entry.item);
      const rule = customerMapping.find(r => mappingDedupKey(r.keyword) === oldExact && r.customer === entry.customer);
      if (rule) {
        persistCustomerMapping(customerMapping.map(r => r.id === rule.id
          ? { ...r, keyword: field === 'item' ? mappingDedupKey(value) : r.keyword, customer: field === 'customer' ? value : r.customer }
          : r));
      }
    }
  };
  // Registers one or more alias/name spellings for a customer + sheet tab as a single batch — adds a
  // Product Catalog entry (routes to the right Sheet tab) and an exact-match Customer Mapping keyword
  // rule (routes to the right customer) for whichever names aren't already known.
  // Uses FUNCTIONAL setState updaters (reading `prev`, never the outer `productCatalog`/`customerMapping`
  // closure) — required because pushPendingRows now calls this once per row in a tight synchronous
  // loop (via confirmStockRow) to register each row's picked block. With a
  // plain `persistCatalog([...productCatalog, ...newEntries])` call, EVERY call in that loop would read
  // the SAME pre-loop snapshot of productCatalog (React doesn't re-render between them), so each call's
  // write would silently clobber the previous one's — only the LAST row's alias would ever actually
  // persist. The functional form correctly threads the real, up-to-date array through every call.
  // `block` (optional): the REAL block title in the customer's actual Sheet this name should land
  // under, when it differs from the name itself (e.g. alias "HANDLE LOCK" really means block "Tijori
  // Handle") — stored on the entry only when it differs, same "blank = same as item name" convention
  // the Sheet Tab column already uses. Read back at push time by buildCustomerSheetPayload, which
  // applies it as an automatic blockTitleOverride — so once an alias's real block is known here, every
  // future row with that same wording routes straight into the right block, no manual re-picking.
  const registerAliases = (customer, names, sheetGroup, block) => {
    const c = (customer || '').trim();
    const sg = (sheetGroup || '').trim();
    const blk = (block || '').trim();
    const list = Array.from(new Set((names || []).map(n => (n || '').trim()).filter(Boolean)));
    if (!c || !sg || !list.length) return;
    setProductCatalog(prev => {
      const existingItems = new Set(prev.map(x => catalogDedupKey(x.customer, x.item)));
      const seenCatalogKeys = new Set();
      const newCatalogEntries = [];
      list.forEach(name => {
        const key = catalogDedupKey(c, name);
        if (existingItems.has(key) || seenCatalogKeys.has(key)) return;
        seenCatalogKeys.add(key);
        newCatalogEntries.push({ id: genId(), customer: c, item: name, sheetGroup: sg, ...(blk && blk !== name ? { block: blk } : {}) });
      });
      if (!newCatalogEntries.length) return prev;
      const next = [...prev, ...newCatalogEntries];
      scheduleSave('catalog', () => window.storage.set(CATALOG_KEY, JSON.stringify(next), false).catch(() => {}));
      return next;
    });
    setCustomerMapping(prev => {
      const existingKeywords = new Set(prev.map(r => mappingDedupKey(r.keyword)));
      const seenKeywords = new Set();
      const newMappingRules = [];
      list.forEach(name => {
        const exact = mappingDedupKey(name);
        if (!exact || existingKeywords.has(exact) || seenKeywords.has(exact)) return;
        seenKeywords.add(exact);
        newMappingRules.push({ id: genId(), keyword: exact, customer: c });
      });
      if (!newMappingRules.length) return prev;
      const next = [...newMappingRules, ...prev];
      scheduleSave('customerMapping', () => window.storage.set(CUSTOMER_MAPPING_KEY, JSON.stringify(next), false).catch(() => {}));
      return next;
    });
  };
  const matchCustomer = (row) => {
    const hint = (row.customerHint || '').trim();
    const d = (row.description || '').toLowerCase();
    if (hint) {
      // A dispatch bill's Party field often carries a full legal name ("Anmol Industries
      // Ltd.") rather than the short name we already track ("Anmol"). Before treating this
      // hint as a brand-new customer, check it against customers we already know — anyone
      // with a tracked Sheet, or anyone already used in Customer Mapping. If the hint
      // contains (or is contained by) a known customer's name, route there instead of
      // spinning up a disconnected duplicate. This is what makes it sustainable: a new
      // legal-name variant of an EXISTING customer resolves automatically from now on, no
      // code change needed — only a genuinely new customer still lands here as "new," and
      // that just needs a Customer Mapping keyword rule (added from the UI, not code).
      const hintLower = hint.toLowerCase();
      // Explicit, exact alias someone has already taught (see Customer Name Aliases) — checked before
      // the fuzzy guess below since it's unambiguous by construction.
      const aliasHit = customerNameAliases.find(a => (a.alias || '').trim().toLowerCase() === hintLower);
      if (aliasHit) return aliasHit.customer;
      const knownCustomers = new Set([
        ...customerSheetIds.map(c => c.customer),
        ...customerMapping.map(r => r.customer),
      ].filter(Boolean));
      for (const known of knownCustomers) {
        const knownLower = known.toLowerCase();
        if (hintLower.includes(knownLower) || knownLower.includes(hintLower)) return known;
      }
      return hint.charAt(0).toUpperCase() + hint.slice(1).toLowerCase();
    }
    for (const rule of customerMapping) {
      if (rule.keyword && d.includes(rule.keyword.toLowerCase())) return rule.customer || 'Unassigned';
    }
    return 'Unassigned';
  };
  // Strips a trailing "x<digits>"/"×<digits>" pack-count — e.g. "Kaju Bake 65g x60" -> "Kaju Bake
  // 65g" — the SAME pattern normalizeForCatalogMatch strips below. That day's carton count isn't part
  // of the item's identity: the same physical item shows up with a different count on every
  // production day, and without stripping it here too, "Kaju Bake 65g x60" and "Kaju Bake 65g x40"
  // (or a day with no count at all, just "Kaju Bake 65g") would each become a SEPARATE stock group —
  // fragmenting one item's ledger into several, each looking like its own new, disconnected variant.
  const stripPackCount = (s) => (s || '').replace(/\s*[x×]\s*\d+\s*$/i, '').trim();
  // NOTE: this used to also strip "container"/"lid"/"jumbo"/"j" as noise words — found while testing
  // the fix above that this silently merges GENUINELY DIFFERENT real items into one stock group,
  // because both words normalize to nothing: "IT 500 Lid" and "IT 500 Container" both became "it500";
  // "IT 500 Container" and "IT 500 Jumbo Container" (two separate real rows in Bindal's own summary
  // sheet) both became "it500" too. That's a silent-data-merging bug — production/dispatch quantities
  // for two different physical items would get added into one blended ledger. Fixed the same way
  // normalizeForCatalogMatch already handles this below: FOLD the "cont" abbreviation into "container"
  // instead of erasing it, and only strip words that are genuinely never part of an item's identity
  // (pure packaging/count units).
  const normalizeVariant = (description) => {
    // Expand taught abbreviations FIRST — "IT 500 J Cont" and "IT 500 Jumbo Cont" are the same real
    // item written two different ways, and should land in the same group; without this they normalize
    // to different keys and silently fragment into two "different" items.
    let s = stripPackCount(applyAbbreviations(description)).toLowerCase();
    s = s.replace(/\(diamond\)/g, '');
    s = s.replace(/\bcorrugated box\b/g, '');
    s = s.replace(/\bcont\.?\b/g, 'container');
    s = s.replace(/\b(pkt|pkts|pcs|nos)\b/g, '');
    s = s.replace(/[^a-z0-9]/g, '');
    return s.trim();
  };
  /* -------- customer stock: combine confirmed Production Register rows (stock in) with confirmed
     Customer Dispatch Bill rows (stock out), grouped by customer + variant -------- */
  const pendingProductionRows = production.filter(row => (row.description || '').trim() && !row.stockConfirmed);
  const pendingDispatchRows = customerDispatch.filter(row => (row.description || '').trim() && !row.stockConfirmed);
  const confirmedProductionRows = production.filter(row => (row.description || '').trim() && row.stockConfirmed);
  const confirmedDispatchRows = customerDispatch.filter(row => (row.description || '').trim() && row.stockConfirmed);
  const confirmStockRow = (registerKey, row, chosenCustomer) => {
    // NOTE: must be `||`, not `??` — every fresh row starts with confirmedCustomer as '' (empty
    // string, not null/undefined), so `??` would treat that '' as "already chosen" and never fall
    // through to the matched suggestion, silently confirming everything as Unassigned.
    const customer = (chosenCustomer || row.confirmedCustomer || matchCustomer(row)).trim() || 'Unassigned';
    // A dispatch bill's Party field (or a production row's bracketed name) is "as written" real-world
    // text — a legal-name variant, a typo, whatever. Whenever confirming lands this row on a REAL
    // customer that's spelled differently from that raw text, remember the mapping so the next bill
    // that says the exact same thing routes straight there next time, no re-picking needed. No-ops
    // automatically when the hint already IS the customer name, or when there's nothing to learn.
    const rawHint = (row.customerHint || '').trim();
    if (rawHint && customer !== 'Unassigned') registerCustomerNameAlias(rawHint, customer);
    // If a Sheet tab (and optionally a block) was picked right here on the Pending Review row for an
    // item the Product Catalog didn't already know, register it now — same alias mechanism as Customer
    // Stock's "not routed" flow, just settled in this same confirm step instead of a second pass later.
    const tabBlockDraft = pendingTabBlockForms[row.id];
    const draftSheetGroup = (tabBlockDraft && tabBlockDraft.sheetGroup || '').trim();
    if (customer !== 'Unassigned' && draftSheetGroup) {
      const description = (row.description || '').trim();
      const block = (tabBlockDraft.block || description).trim();
      registerAliases(customer, [description], draftSheetGroup, block);
    }
    registerSetters[registerKey](prev => {
      const next = prev.map(r => r.id === row.id ? { ...r, confirmedCustomer: customer, stockConfirmed: true } : r);
      persist(registerKey, next);
      return next;
    });
    if (tabBlockDraft) setPendingTabBlockForms(prev => { const next = { ...prev }; delete next[row.id]; return next; });
  };
  // "Push to Sheet" (Pending Production/Dispatch Review) must never silently confirm a row onto a
  // fabricated new-customer guess — only rows that are either already explicitly picked
  // (row.confirmedCustomer set via the dropdown) or whose automatic guess is a real known customer get
  // bulk-confirmed and pushed. Anything genuinely unresolved is left pending and reported, so it gets a
  // deliberate pick instead of a phantom customer.
  //
  // Confirming and pushing can't happen in the same synchronous pass: confirmStockRow's setState only
  // takes effect on the NEXT render, so buildCustomerSheetPayload (which reads live production/
  // customerDispatch state) would still see the pre-confirm data if called right after. Instead this
  // queues the touched customers in pendingPushCustomers; the effect below fires once the confirmed
  // rows have actually landed in state and does the real push then. Duplicate detection happens
  // entirely server-side during that push (see classifyIncomingRows in server/lib/sheets.js) — a row
  // that's already correctly in the real Sheet is silently skipped, never re-written.
  const [pendingPushCustomers, setPendingPushCustomers] = useState(null);
  const pushPendingRows = (registerKey, rows) => {
    const resolved = rows.filter(r => r.confirmedCustomer || isKnownCustomerGuess(r));
    const skipped = rows.length - resolved.length;
    const customersTouched = new Set();
    resolved.forEach(r => {
      const customer = (r.confirmedCustomer || matchCustomer(r)).trim() || 'Unassigned';
      if (customer !== 'Unassigned') customersTouched.add(customer);
      confirmStockRow(registerKey, r);
    });
    if (skipped > 0) window.alert(`Confirmed and queued ${resolved.length} row${resolved.length === 1 ? '' : 's'} to push. Skipped ${skipped} row${skipped === 1 ? '' : 's'} whose customer couldn't be matched automatically — pick a customer in the dropdown for ${skipped === 1 ? 'it' : 'them'}, then push again.`);
    if (customersTouched.size) setPendingPushCustomers(prev => new Set([...(prev || []), ...customersTouched]));
  };
  useEffect(() => {
    if (!pendingPushCustomers || !pendingPushCustomers.size) return;
    const toPush = pendingPushCustomers;
    setPendingPushCustomers(null);
    toPush.forEach(customer => {
      if (getCustomerSheetId(customer).trim()) pushCustomerSheetNow(customer);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingPushCustomers, production, customerDispatch]);
  const updatePendingCustomer = (registerKey) => (id, value) => {
    registerSetters[registerKey](prev => {
      const next = prev.map(r => r.id === id ? { ...r, confirmedCustomer: value } : r);
      persist(registerKey, next);
      return next;
    });
  };
  // Same real-world invoice/party almost always means several pending rows in a row — picking a
  // customer separately for each line item of the same bill is exactly the busywork this avoids. Rows
  // are grouped by the raw text that produced them (customerHint for production, Party/customerHint
  // for dispatch; falls back to Invoice No if that's blank), so one dropdown pick applies to every
  // line of that same bill at once.
  const groupKeyForRow = (row) => (row.customerHint || '').trim() || (row.invoice_no || '').trim();
  const buildPendingGroups = (rows) => {
    const map = {};
    rows.forEach(row => {
      const key = groupKeyForRow(row);
      if (!key) return;
      if (!map[key]) map[key] = [];
      map[key].push(row);
    });
    // Only surface groups that actually need the help: more than one row, sharing a raw hint, with at
    // least one row still lacking both an explicit pick and a resolvable automatic guess — otherwise
    // "Push to Sheet" already handles it without any extra UI.
    return Object.entries(map)
      .map(([key, groupRows]) => ({ key, rows: groupRows }))
      .filter(g => g.rows.length > 1 && g.rows.some(r => !r.confirmedCustomer && !isKnownCustomerGuess(r)));
  };
  const updatePendingCustomerBulk = (registerKey) => (ids, value) => {
    registerSetters[registerKey](prev => {
      const idSet = new Set(ids);
      const next = prev.map(r => idSet.has(r.id) ? { ...r, confirmedCustomer: value } : r);
      persist(registerKey, next);
      return next;
    });
  };
  const customerStockGroups = (() => {
    const groups = {};
    // Merges by DATE within each customer+item group — a Production Register row and a Customer
    // Dispatch Bill row landing on the same date are two separate approvals (confirmed independently,
    // on their own timelines, in their own registers) but describe ONE real row in the real Sheet, so
    // they combine into one ledger entry here rather than becoming two competing half-rows that each
    // show 0 in the column the other one owns. Production Register rows can ALSO carry their own
    // handwritten dispatch figure (noted on the same page) — that's a second, independent source of
    // dispatch for the same date, summed in alongside any dispatch bill's quantity, not overwritten by
    // it. `productionIds`/`dispatchIds` keep every contributing row's real id (there can be more than
    // one dispatch bill for the same item on the same day), so a delete on the merged row can still
    // reach back to the exact underlying register row(s) it came from.
    const addEntry = (row, source) => {
      const customer = row.confirmedCustomer || matchCustomer(row);
      const variantKey = normalizeVariant(row.description);
      const key = `${customer}||${variantKey}`;
      // Display description is the pack-count-stripped, clean item name — not whichever day's raw
      // register wording happened to create this group first. Without this, the group's permanent
      // label (and the Sheet tab this pushes to) would be whatever pack count was on THAT entry,
      // e.g. "Kaju Bake 65g x60" forever, even on a day production ran a batch of 40 instead.
      if (!groups[key]) groups[key] = { id: key, customer, description: stripPackCount(row.description) || row.description, byDate: {} };
      const dateKey = dateSortKey(row.date);
      if (!groups[key].byDate[dateKey]) {
        groups[key].byDate[dateKey] = { date: row.date, pieces: 0, dispatch: 0, productionIds: [], dispatchIds: [], hasDispatch: false };
      }
      const entry = groups[key].byDate[dateKey];
      if (source === 'production') {
        entry.pieces += num(row.pieces);
        entry.productionIds.push(row.id);
        // The Production Register's own handwritten dispatch column, if noted — a second, independent
        // source of dispatch for the same date. Only counts as a real dispatch opinion when it's
        // actually non-zero; a blank/0 handwritten field isn't "we know dispatch was zero," it's just
        // that column going unused, same as everywhere else in this app.
        const handDispatch = num(row.dispatch);
        if (handDispatch) { entry.dispatch += handDispatch; entry.hasDispatch = true; }
      } else {
        entry.dispatch += num(row.quantity);
        entry.dispatchIds.push(row.id);
        entry.hasDispatch = true;
      }
    };
    confirmedProductionRows.forEach(row => addEntry(row, 'production'));
    confirmedDispatchRows.forEach(row => addEntry(row, 'customerDispatch'));
    return Object.values(groups).map(g => {
      const sorted = Object.values(g.byDate).sort((a, b) => dateSortKey(a.date).localeCompare(dateSortKey(b.date)));
      let running = 0;
      const ledger = sorted.map(e => {
        const opening = running;
        running = opening + num(e.pieces) - num(e.dispatch);
        return { ...e, opening, closing: running };
      });
      const totalProduction = sorted.reduce((s, e) => s + num(e.pieces), 0);
      const totalDispatch = sorted.reduce((s, e) => s + num(e.dispatch), 0);
      return { id: g.id, customer: g.customer, description: g.description, ledger, totalProduction, totalDispatch, closingBalance: running };
    });
  })();
  const customerNames = Array.from(new Set(customerStockGroups.map(g => g.customer))).sort((a, b) => (a === 'Unassigned') - (b === 'Unassigned') || a.localeCompare(b));
  // Every flat register the topbar search box can look through — one entry per register, carrying its
  // own rows/columns so the results panel can render each with the exact same EditableTable used on
  // that register's own tab (fix a row right from the search results, no need to go find it by hand).
  const SEARCHABLE_REGISTERS = [
    { key: 'rawMaterialIn', label: 'Raw Material Register', rows: rawMaterialIn, columns: COLUMNS.rawMaterialIn },
    { key: 'consumption', label: 'Consumption', rows: consumption, columns: COLUMNS.consumption },
    { key: 'production', label: 'Production Register', rows: production, columns: COLUMNS.production },
    { key: 'customerDispatch', label: 'Customer Dispatch Bills', rows: customerDispatch, columns: COLUMNS.customerDispatch },
    { key: 'daburSpecs', label: 'Dabur — Spec Master', rows: daburSpecs, columns: COLUMNS.daburSpecs },
    { key: 'daburPO', label: 'Dabur — Pending PO', rows: daburPO, columns: COLUMNS.daburPO },
    { key: 'daburDispatch', label: 'Dabur — Dispatch Log', rows: daburDispatch, columns: COLUMNS.daburDispatch },
  ];
  // Plain case-insensitive substring match against every column's own displayed value — deliberately
  // simple (no fuzzy matching, no per-register special-casing) so the same box that used to only search
  // Production's Party field for Order Availability Check now finds a party name, an item description,
  // an invoice number, a PO number, anything, in whichever register it actually lives in. The Customer
  // Stock match is kept separate from the register matches (not just another "register") since it's the
  // one search result showing the produced/dispatched/balance-so-far numbers the old dedicated check
  // used to show, matched against the customer name and item description together. The Customer Sheets
  // Mirror match is ALSO kept separate — it's read-only (reflects each customer's real, external Sheet;
  // editing a search result row there would silently diverge from the real Sheet instead of fixing
  // anything), and can genuinely be large across many customers/blocks/dates, so it's capped rather than
  // ever rendering an unbounded table.
  const MIRROR_SEARCH_RESULT_CAP = 300;
  const globalSearchResults = (() => {
    const q = globalQuery.trim().toLowerCase();
    if (!q) return null;
    const registerMatches = SEARCHABLE_REGISTERS
      .map(reg => ({ ...reg, rows: reg.rows.filter(row => reg.columns.some(c => String(row[c.key] ?? '').toLowerCase().includes(q))) }))
      .filter(reg => reg.rows.length > 0);
    const stockMatches = customerStockGroups.filter(g => g.customer.toLowerCase().includes(q) || g.description.toLowerCase().includes(q));
    const allMirrorMatches = customerSheetsMirror.filter(r =>
      r.customer.toLowerCase().includes(q) || r.sheetTab.toLowerCase().includes(q) || r.block.toLowerCase().includes(q) || String(r.date || '').toLowerCase().includes(q)
    );
    const mirrorMatches = allMirrorMatches.slice(0, MIRROR_SEARCH_RESULT_CAP);
    const mirrorMatchesTruncated = allMirrorMatches.length > MIRROR_SEARCH_RESULT_CAP;
    return { registerMatches, stockMatches, mirrorMatches, mirrorMatchesTotal: allMirrorMatches.length, mirrorMatchesTruncated };
  })();
  /* -------- Customer Sheets: push a customer's stock ledger out to THEIR OWN separate Google
     Sheet (not a tab on this app's main sheet) — matching the real BINDAL STOCK.xlsx / DIAMOND.xlsx /
     anmol stock dec 22.xlsx layout exactly: one tab per base item ("IT 500", "Digestive", "CREAM"),
     with every variant of that item as its own side-by-side table within that tab. This is a
     write-only computed mirror (the Production Register and Customer Dispatch Bills stay the real
     source of truth) and is pushed ONLY when you click "Push to Sheet" on the Customer Stock tab,
     right under that customer's own preview — never automatically. The item/variant tabs are MERGED,
     never overwritten: the backend reads what's
     already in each tab first and only appends rows for dates it doesn't already have there, using
     the SAME formula convention already used throughout the real file (Opening references the
     previous row's Closing cell, Closing = Opening+Production-Dispatch) — anything already in the
     sheet, pushed by the app before or typed in by hand, is never touched, not even to re-write it
     unchanged. The "summary" tab is never read or written at all: it's a handful of formulas that
     already point at each item's block's own last row, so it keeps calculating itself correctly for
     as long as a block only ever grows downward. */
  // "Unassigned" is deliberately excluded here — it's not a real customer with a Sheet to push to, so
  // it should never get a Sheet-ID field, a Push button, or a preview/review panel. It's just the
  // holding pen for register/dispatch rows that didn't match any customer at all (a DIFFERENT,
  // earlier problem than an item not matching a customer's catalog) — those get resolved by fixing
  // Customer Mapping and re-syncing, not by pushing anything for "Unassigned" itself.
  const allCustomerTabNames = Array.from(new Set([
    ...productCatalog.map(c => c.customer),
    ...customerNames,
  ])).filter(c => c && c !== 'Unassigned');
  // A guess is "known" when matchCustomer resolved the row to a real, already-tracked customer (an
  // alias hit, a Customer Mapping rule, or a fuzzy legal-name match) or explicitly to 'Unassigned'.
  // It's "unknown" only in the one fallback case matchCustomer has left: a hint that didn't match
  // anything, title-cased and returned as-is — a fabricated new-customer guess that should never be
  // silently confirmed. Used to decide whether the Pending Review dropdown can safely pre-select the
  // guess, or must force an explicit pick instead.
  const isKnownCustomerGuess = (row) => {
    const guess = (row.confirmedCustomer || matchCustomer(row)).trim();
    return guess === 'Unassigned' || allCustomerTabNames.includes(guess);
  };
  // Row-level highlight for the Production Register and Customer Dispatch Bills tabs themselves — so
  // a row that's confirmed-but-Unassigned, or still pending with no resolvable customer, is easy to
  // spot right there (and delete, if that's the right call) without having to go hunt for it on the
  // Customer Stock tab. Clears itself the instant the row gets a real customer — nothing to "undo" by
  // hand. Rows with no description are the shade/size/GSM style entries that never feed Customer Stock
  // at all, so they're never in scope for this.
  const needsCustomerHighlight = (row) => {
    if (!(row.description || '').trim()) return false;
    if (row.stockConfirmed) return (row.confirmedCustomer || '').trim() === 'Unassigned';
    return !isKnownCustomerGuess(row);
  };
  // Groups this customer's variant ledgers under the tab name (sheetGroup) their catalog entry
  // says they belong to. A variant whose description doesn't exactly match any catalog item for
  // this customer falls back to using its own description as the tab name AND is flagged in
  // `unmatched` — surfaced as a warning before pushing, since an unmapped/misspelled item (e.g. a
  // dispatch bill saying "HANDLE LOCK" when the catalog only knows "Tijori Handle") would otherwise
  // silently land in the wrong tab of a sheet that looks like an official customer record.
  // Dispatch bills and handwritten ledgers routinely abbreviate ("IT 500 CONT" for what the catalog
  // calls "IT 500 Container") — a plain case-insensitive string match would flag nearly everything
  // as unmapped. This folds that one specific abbreviation and strips spacing/punctuation noise, but
  // deliberately does NOT strip words like "Lid"/"Container" themselves (unlike normalizeVariant
  // above) — those are exactly what tells two real catalog entries apart, so erasing them would
  // silently merge "IT 500 Lid" and "IT 500 Container" into looking like the same item. Verified
  // against all 4 real dispatch bills: correctly matches every real abbreviation variant, and still
  // correctly flags a genuinely different name ("HANDLE LOCK" vs. catalog's "Tijori Handle") as
  // unmatched rather than guessing.
  //
  // The Production Register has its own systematic abbreviation, found by tracing a real scanned
  // page (Anmol, Dec 2025) through this exact matching logic: every entry is written as "<item> x<N>"
  // — e.g. "Hit & Run 128g x30", "Digestive 120g x60" — where "x<N>" is that day's carton/batch count,
  // NOT part of the item's own identity (the catalog just calls it "hit & run 128g" / "Digestive
  // 120g"). Without stripping it, EVERY register entry for an item whose catalog name has no such
  // suffix fails to match — which is exactly what was flooding customer sheets with one tab per
  // register line instead of grouping into the item's real tab. This strips a trailing "x<digits>"
  // (or "×<digits>") ONLY at the end of the string, so it's safe against catalog items whose own name
  // legitimately ends in a pack count written the OTHER way round ("64gm * 60pkt", "32g*144pkt" —
  // neither has a literal "x" immediately before the trailing digits, so neither is touched).
  const normalizeForCatalogMatch = (s) => applyAbbreviations(s || '')
    .toLowerCase()
    // "pkt"/"pkt." is pure packaging noise (pack-count units), not part of an item's identity — strip
    // the word itself first (no word-boundary requirement, since it's routinely glued straight onto a
    // digit with no space: "40pkt", "70pkt"). Stripping it here — rather than trying to match the whole
    // "x<N> pkt" phrase as one pattern — is what lets the trailing x<N> strip below actually reach the
    // end of the string in cases like "65g x 70 pkt", which previously survived untouched because the
    // string ended in "pkt", not a digit.
    .replace(/pkt\.?/g, '')
    // Trailing pack-count separator — Sheet block titles write this as "x", "×", or a bare "*"
    // ("35g*144", as opposed to the register's "35g x144"). All three mean the same thing here.
    .replace(/\s*[x×*]\s*\d+\s*$/i, '')
    // "gm" and "g" both mean grams — real Sheet block titles and the handwritten register don't agree
    // on which one to use for the same item ("64gm" vs "64g"), so fold "gm" down to "g" wherever it's
    // written immediately after a digit, before the final alphanumeric-only strip below would otherwise
    // let that one-letter difference silently keep them apart.
    .replace(/(\d)\s*gm\b/g, '$1g')
    .replace(/\bcont\.?\b/g, 'container')
    .replace(/[^a-z0-9]/g, '');
  // The Product Catalog entry (if any) a description already resolves to for this customer — i.e.
  // where it would land on push without needing a Sheet tab/Block picked for it anywhere. Used both to
  // decide whether the Pending Review tables need to show the tab/block picker for a row at confirm
  // time (before it's even reached Customer Stock's "not routed" fallback section), and, when it IS
  // already known, to actually show what it's routed to instead of just a blank dash.
  const getCatalogEntryForItem = (customer, description) => {
    if (!customer || customer === 'Unassigned') return null;
    const key = normalizeForCatalogMatch(description || '');
    return productCatalog.find(c => c.customer === customer && normalizeForCatalogMatch(c.item) === key) || null;
  };
  // Every real customer Sheet writes dates as dot-separated D.M.YY ("5.1.24", "26.1.24") — never with
  // Rows land in the ledger already dot-formatted now (normalizeDateToDots runs at extraction time,
  // see DOCUMENT_TYPES above) — this second call right before push is a safety net for anything that
  // reached the ledger some other way (a manually typed/edited row), not the primary defense anymore.
  const buildCustomerSheetPayload = (customer) => {
    const groups = customerStockGroups.filter(g => g.customer === customer);
    const sheetGroupByItem = {};
    const blockByItem = {};
    productCatalog.filter(c => c.customer === customer).forEach(c => {
      const key = normalizeForCatalogMatch(c.item);
      sheetGroupByItem[key] = (c.sheetGroup || c.item || '').trim();
      if (c.block && c.block.trim()) blockByItem[key] = c.block.trim();
    });
    const unmatched = [];
    const tabsMap = {};
    groups.forEach(g => {
      const key = normalizeForCatalogMatch(g.description);
      const sheetGroup = sheetGroupByItem[key];
      // An item with no catalog match must NEVER get pushed under its own improvised tab — that's
      // exactly what fragmented real customer sheets into one tab per pack-size variant instead of
      // one tab per product category. It's reported in `unmatched` so it's visible and fixable, but
      // it's excluded from itemGroups entirely: nothing gets written to the Sheet for it until it's
      // deliberately mapped in the Known Product Catalog with the exact wording that customer's own
      // Sheet uses for that variant's shared tab.
      if (!sheetGroup) { unmatched.push(g.description || '(blank description)'); return; }
      if (!tabsMap[sheetGroup]) tabsMap[sheetGroup] = [];
      // blockTitleOverride from the catalog's stored alias->block mapping (see registerAliases) — set
      // automatically here so a wording once routed by hand through the block picker never needs to be
      // re-picked on a later push; an explicit reviewEdits override (picked fresh in the review UI)
      // still wins over this, see applyReviewEdits.
      const catalogBlock = blockByItem[key];
      tabsMap[sheetGroup].push({
        title: g.description || 'Item',
        header: ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'],
        // null (not 0) whenever this date has NO real production/dispatch entry behind it at all — a
        // dispatch bill has no opinion on production, and vice versa. Sending a literal 0 in that case
        // would read as "confirmed zero," which the server would then compare against whatever the real
        // Sheet already has for that column — exactly the false "mismatch" a pure dispatch-only upload
        // must never trigger just because it has nothing to say about production.
        rows: g.ledger.map(e => [
          normalizeDateToDots(e.date), e.opening,
          (e.productionIds && e.productionIds.length) ? (e.pieces || 0) : null,
          e.hasDispatch ? (e.dispatch || 0) : null,
          e.closing,
        ]),
        ...(catalogBlock ? { blockTitleOverride: catalogBlock } : {}),
      });
    });
    const itemGroups = Object.entries(tabsMap).map(([tabName, variants]) => ({ tabName, variants }));
    // No summary payload here on purpose — the real "summary" tab in every customer's own Sheet is a
    // handful of live formulas (one per item, pointing at that item's block's own last row), which
    // already keeps itself correct as long as the app only ever appends rows to a block and never
    // rewrites it. The app never reads or writes that tab at all now, for any customer.
    return { itemGroups, unmatched: Array.from(new Set(unmatched)) };
  };
  // reviewEdits[customer] is always empty now (see its declaration) — kept as a harmless no-op pass-
  // through rather than reworking every call site that still asks for the "edited" payload.
  const applyReviewEdits = (itemGroups, editsForCustomer) => {
    if (!editsForCustomer || !Object.keys(editsForCustomer).length) return itemGroups;
    const regrouped = {};
    itemGroups.forEach(g => (g.variants || []).forEach(v => {
      const e = editsForCustomer[v.title];
      const finalTab = (e && e.tabNameOverride && e.tabNameOverride.trim()) ? e.tabNameOverride.trim() : g.tabName;
      const deletedRows = (e && e.deletedRows) || {};
      const rows = (v.rows || [])
        .map((r, i) => {
          if (deletedRows[i]) return null;
          const re = e && e.rowEdits && e.rowEdits[i];
          if (!re) return r;
          return [
            re.date !== undefined ? re.date : r[0],
            r[1],
            (re.production !== undefined && re.production !== '') ? Number(re.production) : r[2],
            (re.dispatch !== undefined && re.dispatch !== '') ? Number(re.dispatch) : r[3],
            r[4],
          ];
        })
        .filter(Boolean);
      if (!regrouped[finalTab]) regrouped[finalTab] = [];
      // blockTitleOverride: which of the tab's REAL existing blocks this item was explicitly assigned
      // to — v.title itself (the item's real name) is never rewritten, so it keeps working as the
      // stable key reviewEdits is keyed by. Only sent when actually set, so an untouched item's
      // behavior is unchanged. reviewEdits is always empty now (see its declaration), so this whole
      // function is a no-op in practice — left in place rather than reworking every caller.
      const blockTitleOverride = (e && e.blockTitleOverride && e.blockTitleOverride.trim()) ? e.blockTitleOverride.trim() : undefined;
      regrouped[finalTab].push(blockTitleOverride ? { ...v, rows, blockTitleOverride } : { ...v, rows });
    }));
    return Object.entries(regrouped).map(([tabName, variants]) => ({ tabName, variants }));
  };
  const getEditedPayload = (customer) => {
    const { itemGroups, unmatched } = buildCustomerSheetPayload(customer);
    return { itemGroups: applyReviewEdits(itemGroups, reviewEdits[customer]), unmatched };
  };
  // Declared up here (rather than down by the rest of their assign-form logic, further below) because
  // refreshReview/reviewRefreshKey read them and reviewRefreshKey's value is computed synchronously on
  // every render, not deferred inside a callback — referencing a later `const [x] = useState()` from
  // here would be a genuine temporal-dead-zone crash, not just a stale-closure issue.
  const [assignForms, setAssignForms] = useState({});
  const [unmatchedAssignForms, setUnmatchedAssignForms] = useState({}); // keyed by `${customer}::${description}`
  // Sheet tab / Block picked directly on a Pending Production/Dispatch Review row, for an item that
  // isn't in the Product Catalog yet — keyed by the register row's own id. Applied at confirm time (see
  // confirmStockRow) so an unknown item's routing is settled in the SAME step as picking its customer,
  // instead of needing a second pass through Customer Stock's "not routed" section afterward.
  const [pendingTabBlockForms, setPendingTabBlockForms] = useState({});
  const updatePendingTabBlockForm = (id, field, value) => setPendingTabBlockForms(prev => ({ ...prev, [id]: { ...prev[id], [field]: value } }));
  // Dry-runs the CURRENT (edited) payload against the customer's real Sheet so the review area always
  // shows an accurate diff — old row (from the real Sheet, read fresh every time) next to the new
  // row(s) about to be appended — rather than a locally-guessed one. Called automatically whenever the
  // computed payload changes (a new register entry gets confirmed, a catalog mapping changes, or a
  // review edit is made), debounced below, and manually never needs to be triggered by hand.
  const refreshReview = async (customer) => {
    const sheetId = getCustomerSheetId(customer).trim();
    const { itemGroups } = getEditedPayload(customer);
    // Bare tab-name probes: whatever's currently typed into this customer's "Sheet tab" field on the
    // unassigned / not-yet-routed assign forms, even before anything is confirmed to land there. Sent
    // as tabName-only groups (no variants) so readPushState/previewCustomerSheet still read that tab's
    // REAL current blocks — otherwise a tab with zero currently-routed items never gets read at all,
    // and the Block picker for a customer whose items are ALL still unmatched would always show empty.
    const pendingRowsForCustomer = [...pendingProductionRows, ...pendingDispatchRows].filter(row => {
      const guess = row.confirmedCustomer || (isKnownCustomerGuess(row) ? matchCustomer(row) : '');
      return (guess || '').trim().toLowerCase() === customer.toLowerCase();
    });
    const probeTabNames = Array.from(new Set([
      ...Object.entries(unmatchedAssignForms)
        .filter(([key]) => key.startsWith(`${customer}::`))
        .map(([, f]) => (f.sheetGroup || '').trim()),
      ...Object.values(assignForms)
        .filter(f => (f.customer || '').trim().toLowerCase() === customer.toLowerCase())
        .map(f => (f.sheetGroup || '').trim()),
      ...pendingRowsForCustomer.map(row => ((pendingTabBlockForms[row.id] && pendingTabBlockForms[row.id].sheetGroup) || '').trim()),
    ].filter(Boolean)));
    // Deliberately NOT excluding a probe name just because it looks like it already matches an
    // itemGroups tab: real customer sheets routinely have trailing spaces / inconsistent casing in tab
    // names (see normalizeTabKey's comments), so a plain client-side string comparison here could
    // wrongly skip probing a tab whose REAL resolved name (only the backend actually knows it) differs
    // by exactly that kind of formatting — leaving the Block picker with nothing and no fallback. A
    // probe that turns out to duplicate an already-included tab just costs one harmless extra read.
    const newProbeNames = probeTabNames;
    const itemGroupsForPreview = [...itemGroups, ...newProbeNames.map(t => ({ tabName: t, variants: [] }))];
    // Still fire the preview call with an EMPTY itemGroups array when there's nothing routed/probed yet
    // — the backend reads the real Sheet's tab list regardless of itemGroups content (see readPushState),
    // so this is what makes existingTabNames available for the Sheet tab pickers even for a customer
    // with zero confirmed items so far, instead of only ever showing catalog-guessed tab names.
    if (!sheetId) { setReviewByCustomer(prev => ({ ...prev, [customer]: null })); return; }
    setReviewByCustomer(prev => ({ ...prev, [customer]: { ...(prev[customer] || {}), loading: true, error: '' } }));
    try {
      const res = await fetch('/api/customer-sheets/preview', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'include',
        body: JSON.stringify({ spreadsheetId: sheetId, itemGroups: itemGroupsForPreview }),
      });
      if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return; }
      const data = await res.json().catch(() => ({}));
      if (res.ok && data.ok) {
        // Probe-only entries come back at the same positions they were appended at, in order — pull
        // their real block titles into a name -> titles map so the picker can look a typed Sheet tab
        // name up directly (exact string, same as what was sent), no tab-name re-normalization needed.
        const probedBlocksByTabName = {};
        newProbeNames.forEach((name, i) => {
          const t = (data.tabs || [])[itemGroups.length + i];
          probedBlocksByTabName[name] = (t && t.existingBlockTitles) || [];
        });
        setReviewByCustomer(prev => ({ ...prev, [customer]: { loading: false, error: '', tabs: data.tabs, existingTabNames: data.existingTabNames || [], probedBlocksByTabName } }));
      } else {
        setReviewByCustomer(prev => ({ ...prev, [customer]: { loading: false, error: data.error || `Preview failed (HTTP ${res.status}).`, tabs: [] } }));
      }
    } catch (e) {
      setReviewByCustomer(prev => ({ ...prev, [customer]: { loading: false, error: e.message || 'Network error — could not reach the server.', tabs: [] } }));
    }
  };
  // Real block titles for a customer's Sheet tab, sourced from whatever refreshReview last fetched —
  // checks an already-routed tab (`tabs`, matched by name) first, then a bare probe-only tab
  // (`probedBlocksByTabName`, keyed by the exact string sent). Shared by every Block picker in this app
  // (Customer Stock's unassigned/not-routed forms, Pending Review) so a typed Sheet tab name resolves to
  // the same real blocks no matter which form it's typed into.
  const getRealBlocksForTab = (customer, sheetGroup) => {
    const sg = (sheetGroup || '').trim();
    if (!sg) return [];
    const rev = reviewByCustomer[customer];
    if (!rev) return [];
    const matchedTab = (rev.tabs || []).find(t => t.tabName.toLowerCase() === sg.toLowerCase());
    if (matchedTab && matchedTab.existingBlockTitles) return matchedTab.existingBlockTitles;
    return (rev.probedBlocksByTabName && rev.probedBlocksByTabName[sg]) || [];
  };
  // Real Sheet tab names for a customer, straight from the customer's actual spreadsheet (see
  // refreshReview, which now fires even with nothing routed/probed yet purely to populate this) — used
  // ALONGSIDE whatever tab names the Product Catalog already knows, never instead of them, so a Sheet
  // tab picker always offers the real tabs even for a customer with zero catalog entries so far.
  const getRealTabNamesForCustomer = (customer) => (reviewByCustomer[customer] && reviewByCustomer[customer].existingTabNames) || [];
  // Block names already assigned to OTHER items under this same customer+Sheet tab in the Product
  // Catalog — e.g. if "Jeera Dhamal 35gm" already routes to block "jeera 35gm x 144" under tab "jeera",
  // that block should be offered as a pick for any OTHER item being routed to "jeera" too, even before
  // it's ever actually been pushed to the real Sheet (getRealBlocksForTab alone only knows about blocks
  // that already physically exist there, which a freshly-aliased-but-not-yet-pushed block never is).
  const getCatalogBlocksForTab = (customer, sheetGroup) => {
    const sg = (sheetGroup || '').trim().toLowerCase();
    if (!customer || !sg) return [];
    return Array.from(new Set(
      productCatalog
        .filter(c => c.customer === customer && (c.sheetGroup || c.item || '').trim().toLowerCase() === sg)
        .map(c => (c.block || c.item || '').trim())
    )).filter(Boolean);
  };
  // Keeps the underlying per-customer real-Sheet diff (existingTabNames/existingBlockTitles) fresh in
  // the background — the visible "Customer Ledgers" review UI this used to feed was removed (pushing
  // now happens straight from Pending Review, see pushPendingRows), but getRealTabNamesForCustomer/
  // getRealBlocksForTab still read this to populate the Sheet-tab/Block datalist suggestions there and
  // on Customer Sheets. Debounced so a burst of confirms (e.g. "Push to Sheet") triggers one preview
  // call per customer, not one per row.
  const reviewRefreshKey = JSON.stringify({
    groups: customerStockGroups.map(g => ({ c: g.customer, d: g.description, ledger: g.ledger })),
    catalog: productCatalog, sheetIds: customerSheetIds, edits: reviewEdits,
    // Sheet tab names typed into the not-yet-routed assign forms — included so pausing on a newly
    // typed tab name re-probes the real Sheet for that tab's blocks (see refreshReview).
    pendingTabs: [
      ...Object.entries(unmatchedAssignForms).map(([k, f]) => [k, f.sheetGroup || '']),
      ...Object.entries(assignForms).map(([k, f]) => [k, f.customer || '', f.sheetGroup || '']),
      ...Object.entries(pendingTabBlockForms).map(([k, f]) => [k, f.sheetGroup || '', f.block || '']),
      ...[...pendingProductionRows, ...pendingDispatchRows].map(r => [r.id, r.confirmedCustomer || '']),
    ],
  });
  useEffect(() => {
    const timer = setTimeout(() => {
      allCustomerTabNames.forEach(customer => {
        if (getCustomerSheetId(customer).trim()) refreshReview(customer);
      });
    }, 700);
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reviewRefreshKey]);
  const persistCustomerSheetIds = (next) => {
    setCustomerSheetIds(next);
    scheduleSave('customer-sheet-ids', () => window.storage.set(CUSTOMER_SHEET_IDS_KEY, JSON.stringify(next), false).catch(() => {}));
  };
  // Lets every Sheet ID field accept either the raw ID or the full URL you'd copy straight out of the
  // browser's address bar (https://docs.google.com/spreadsheets/d/<ID>/edit#gid=0) — pulls just the ID
  // out of the URL so what's actually stored/sent to the backend is always the plain ID either way.
  // If it doesn't look like a Sheets URL, the input is trusted as-is (already a raw ID).
  const extractSheetIdFromInput = (raw) => {
    const s = (raw || '').trim();
    const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return m ? m[1] : s;
  };
  const updateCustomerSheetId = (customer, rawInput) => {
    const sheetId = extractSheetIdFromInput(rawInput);
    const exists = customerSheetIds.some(c => c.customer === customer);
    const next = exists
      ? customerSheetIds.map(c => c.customer === customer ? { ...c, sheetId } : c)
      : [...customerSheetIds, { id: genId(), customer, sheetId }];
    persistCustomerSheetIds(next);
  };
  const getCustomerSheetId = (customer) => (customerSheetIds.find(c => c.customer === customer) || {}).sheetId || '';
  // Merges any fields (sheetId, lastImportedAt, lastPushedAt, itemCount) into a customer's tracked-sheet
  // registry entry, creating it if it doesn't exist yet. This registry entry IS the "already imported"
  // record the CRUD workflow below needs — it's the same array that already backed the manual Sheet ID
  // field for Push, just extended with tracking metadata.
  const patchCustomerSheetEntry = (customer, patch) => {
    const exists = customerSheetIds.some(c => c.customer === customer);
    const next = exists
      ? customerSheetIds.map(c => c.customer === customer ? { ...c, ...patch } : c)
      : [...customerSheetIds, { id: genId(), customer, sheetId: '', ...patch }];
    persistCustomerSheetIds(next);
  };
  // Delete: stops tracking a customer's Sheet ID (so it's no longer counted as "already imported" and
  // the manual Push field for it clears). Deliberately does NOT touch productCatalog/customerMapping —
  // those items were confirmed by a human at import time and stay until removed one-by-one from the
  // Known Product Catalog table, exactly like any other catalog edit.
  const removeCustomerSheetEntry = (customer) => {
    persistCustomerSheetIds(customerSheetIds.filter(c => c.customer !== customer));
  };
  /* -------- Customer Sheets: IMPORT direction (Create + Update/re-sync of the CRUD workflow) --------
     Reads an arbitrary customer's own Google Sheet by ID (no hardcoded per-customer identity — any
     Sheet ID can be pasted in) via the backend's /api/customer-sheets/import, and turns its item list
     into a review screen before anything is actually added — same "review before confirm" pattern as
     the existing .xlsx upload import above, just sourced from a live Sheet instead of a file.
     mode: 'create' treats every returned item as a candidate to add for a NEW (or existing) customer,
     whose name is suggested from the sheet's own title. mode: 'resync' is the "Update" half of CRUD:
     it's scoped to one already-tracked customer and keeps ONLY items that aren't already in that
     customer's catalog — additive-only, so it can never silently overwrite a sheetGroup someone edited
     by hand in the Known Product Catalog table. */
  const [sheetImportBusy, setSheetImportBusy] = useState(false);
  const [sheetImportError, setSheetImportError] = useState('');
  const [newSheetId, setNewSheetId] = useState('');
  const [sheetReview, setSheetReview] = useState(null); // { spreadsheetId, spreadsheetTitle, customerName, mode, resyncCustomer, items, skippedExisting }
  const [importResultMessage, setImportResultMessage] = useState('');
  const importSheetById = async (spreadsheetId, { mode = 'create', resyncCustomer = '' } = {}) => {
    const id = (spreadsheetId || '').trim();
    if (!id) { setSheetImportError('Paste a Google Sheet ID first.'); return; }
    setSheetImportError(''); setImportResultMessage(''); setSheetReview(null); setSheetImportBusy(true);
    try {
      const res = await fetch('/api/customer-sheets/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ spreadsheetId: id }),
      });
      if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return; }
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) {
        setSheetImportError(data.error || `Could not read that sheet (HTTP ${res.status}). Make sure it's been shared with the service account's email, same as your main FIMS sheet.`);
        return;
      }
      if (!data.items || !data.items.length) {
        setSheetImportError(`No items found (checked ${data.tabCount || 0} tab${data.tabCount === 1 ? '' : 's'}, excluding "summary"). Nothing to import.`);
        return;
      }
      const customerName = mode === 'resync' ? resyncCustomer : (data.spreadsheetTitle || '').trim();
      const existingItemNames = new Set(productCatalog.filter(c => c.customer === customerName).map(c => c.item.toLowerCase().trim()));
      let items = data.items.map(it => ({
        id: genId(), item: it.item, sheetGroup: it.sheetGroup || it.item, include: true,
        isNew: !existingItemNames.has(String(it.item || '').toLowerCase().trim()),
      }));
      // Full re-sync now: show every item currently in the Sheet, not just new ones — confirming below
      // REPLACES this customer's whole catalog + auto-generated mapping with exactly this list, so
      // anything stale (a block that got renamed or removed in the Sheet) drops out too, instead of
      // piling up forever. isNew is kept per item just so the review can show what's actually changed.
      const skippedExisting = mode === 'resync' ? items.filter(it => !it.isNew).length : 0;
      // Carried through to confirmSheetImport, which is where the customer name is actually finalized
      // (still editable right up to confirm) — the Customer Sheets Mirror gets tagged with whatever
      // name the person confirms, not this suggested one.
      setSheetReview({ spreadsheetId: id, spreadsheetTitle: data.spreadsheetTitle || '', customerName, mode, resyncCustomer, items, skippedExisting, mirrorRows: data.mirrorRows || [] });
    } catch (e) {
      setSheetImportError(e.message || 'Network error — could not reach the server.');
    } finally {
      setSheetImportBusy(false);
    }
  };
  const updateSheetReviewItem = (id, field, value) => {
    setSheetReview(prev => ({ ...prev, items: prev.items.map(it => it.id === id ? { ...it, [field]: value } : it) }));
  };
  // Same catalog-then-mapping confirm logic as confirmCatalogImport (exact-name rule per item, checked
  // first, plus a broader fallback keyword), then registers/updates this Sheet ID in the tracked-sheet
  // registry, then re-checks every confirmed-but-Unassigned production/dispatch row against the newly
  // widened mapping — the "if you find anything that matches in unassigned, add it to the respective
  // customer sheet" half of the request.
  // Purely additive for BOTH the catalog and Customer Mapping, resync or not — an item already in the
  // catalog (same customer + item, see catalogDedupKey) is left completely untouched: not regenerated
  // with a new id, not duplicated, not removed even if it's no longer in the Sheet. Only a genuinely
  // new item gets appended. Resync used to treat the Sheet as the source of truth and replace/drop
  // whatever wasn't currently in `included` — which meant an item temporarily missing from the Sheet (a
  // bad push undone, a row someone deleted by mistake) silently lost its catalog entry on the next
  // resync, the same class of data loss the Customer Mapping side was already fixed for. Resync and
  // import now share this identical additive logic; the only remaining difference between them is which
  // items even reach the review screen (resync shows everything currently in the Sheet, import already
  // pre-filtered to just the new ones) and the wording in the result message.
  const confirmSheetImport = () => {
    if (!sheetReview || !sheetReview.customerName.trim()) return;
    const customer = sheetReview.customerName.trim();
    const included = sheetReview.items.filter(it => it.include && it.item.trim());
    const isFullResync = sheetReview.mode === 'resync';
    const existingItems = new Set(productCatalog.map(c => catalogDedupKey(c.customer, c.item)));
    const newCatalogEntries = included
      .filter(it => !existingItems.has(catalogDedupKey(customer, it.item)))
      .map(it => ({ id: genId(), customer, item: it.item.trim(), sheetGroup: (it.sheetGroup || it.item).trim() }));
    const addedCount = newCatalogEntries.length;
    const nextCatalog = [...productCatalog, ...newCatalogEntries];
    const existingKeywords = new Set(customerMapping.map(r => mappingDedupKey(r.keyword)));
    const exactRules = [];
    const fallbackRules = [];
    included.forEach(it => {
      const exact = mappingDedupKey(it.item);
      if (exact && !existingKeywords.has(exact)) { exactRules.push({ id: genId(), keyword: exact, customer }); existingKeywords.add(exact); }
      const fallback = exact.replace(/[\d].*$/, '').trim();
      if (fallback && fallback.length > 2 && !existingKeywords.has(fallback)) { fallbackRules.push({ id: genId(), keyword: fallback, customer }); existingKeywords.add(fallback); }
    });
    const nextMapping = [...exactRules, ...customerMapping, ...fallbackRules];
    persistCatalog(nextCatalog);
    persistCustomerMapping(nextMapping);
    patchCustomerSheetEntry(customer, {
      sheetId: sheetReview.spreadsheetId,
      lastImportedAt: new Date().toISOString(),
      itemCount: nextCatalog.filter(c => c.customer === customer).length,
    });
    // Seeds/refreshes this customer's slice of the Customer Sheets Mirror with every real ledger row
    // read at import time, so global search can find it immediately — no separate sync step needed.
    replaceCustomerMirrorRows(customer, sheetReview.mirrorRows);
    const reassignedCount = reassignUnassignedRows();
    const verb = isFullResync ? 'Re-synced' : sheetReview.mode === 'generate' ? 'Generated' : 'Imported';
    setImportResultMessage(
      `${verb} ${customer}: added ${addedCount} new catalog item${addedCount === 1 ? '' : 's'}` +
      (sheetReview.skippedExisting ? ` (${sheetReview.skippedExisting} already known, skipped)` : '') +
      (reassignedCount ? `, reassigned ${reassignedCount} previously-Unassigned row${reassignedCount === 1 ? '' : 's'} to ${customer}` : '') + '.'
    );
    setSheetReview(null);
    setNewSheetId('');
  };
  /* -------- Customer Sheets: GENERATE a brand-new customer's tabs/blocks/summary from scratch, into a
     blank Sheet the person already created and shared with the service account (the service account
     can't create+own a new file itself — verified directly, plain service accounts have no Drive
     storage of their own). Category name -> tab, each line under it -> one item block within that tab.
     Reuses the exact same sheetReview/confirmSheetImport review-then-confirm flow as the Sheet-ID
     import above (mode: 'generate'), so a generated structure gets the same "review before it touches
     your catalog" treatment as anything else. -------- */
  const [genCustomerName, setGenCustomerName] = useState('');
  const [genSheetId, setGenSheetId] = useState('');
  const [genCategories, setGenCategories] = useState([{ id: genId(), name: '', itemsText: '' }]);
  const [genBusy, setGenBusy] = useState(false);
  const [genError, setGenError] = useState('');
  const addGenCategory = () => setGenCategories(prev => [...prev, { id: genId(), name: '', itemsText: '' }]);
  const removeGenCategory = (id) => setGenCategories(prev => prev.filter(c => c.id !== id));
  const updateGenCategory = (id, field, value) => setGenCategories(prev => prev.map(c => c.id === id ? { ...c, [field]: value } : c));
  const buildGeneratedStructure = async () => {
    const customer = genCustomerName.trim();
    const sheetId = extractSheetIdFromInput(genSheetId);
    if (!customer) { setGenError('Enter a customer name.'); return; }
    if (!sheetId) { setGenError("Paste the blank Google Sheet's link or ID (create it and share it with the service account email above first)."); return; }
    const categories = genCategories
      .map(c => ({ name: c.name.trim(), items: c.itemsText.split('\n').map(s => s.trim()).filter(Boolean) }))
      .filter(c => c.name && c.items.length);
    if (!categories.length) { setGenError('Add at least one category name with at least one item listed under it.'); return; }
    setGenError(''); setImportResultMessage(''); setGenBusy(true);
    try {
      const res = await fetch('/api/customer-sheets/generate-structure', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ spreadsheetId: sheetId, categories }),
      });
      if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return; }
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) {
        setGenError(data.error || `Could not build the structure (HTTP ${res.status}).`);
        return;
      }
      const items = (data.items || []).map(it => ({ id: genId(), item: it.item, sheetGroup: it.sheetGroup || it.item, include: true, isNew: true }));
      setSheetReview({ spreadsheetId: sheetId, spreadsheetTitle: customer, customerName: customer, mode: 'generate', resyncCustomer: '', items, skippedExisting: 0, mirrorRows: [] });
      setGenCustomerName(''); setGenSheetId(''); setGenCategories([{ id: genId(), name: '', itemsText: '' }]);
    } catch (e) {
      setGenError(e.message || 'Network error — could not reach the server.');
    } finally {
      setGenBusy(false);
    }
  };
  // Re-check every already-confirmed production/dispatch row still sitting as Unassigned against the
  // (possibly just-widened) customer mapping. Only PENDING rows re-evaluate matchCustomer live on every
  // render — a CONFIRMED row's confirmedCustomer is locked in at confirm-time, so a new customer's
  // rules added after that row was confirmed would never reach it without this explicit re-pass.
  // Returns the number of rows it moved off Unassigned, for the confirmation message above.
  const reassignUnassignedRows = () => {
    const reassignInRegister = (rows) => {
      let count = 0;
      const next = rows.map(r => {
        if (!r.stockConfirmed) return r;
        const current = (r.confirmedCustomer || '').trim();
        if (current && current !== 'Unassigned') return r;
        const match = matchCustomer(r);
        if (match && match !== 'Unassigned') { count++; return { ...r, confirmedCustomer: match }; }
        return r;
      });
      return { next, count };
    };
    const p = reassignInRegister(production);
    const d = reassignInRegister(customerDispatch);
    if (p.count) { setProduction(p.next); persist('production', p.next); }
    if (d.count) { setCustomerDispatch(d.next); persist('customerDispatch', d.next); }
    return p.count + d.count;
  };
  // Manual assign form for the Unassigned section (Customer Stock tab) — lets a person route a
  // stuck item straight to a customer + sheet tab + item block by hand instead of only being able
  // to fix it indirectly via a Customer Mapping keyword rule. One form per unassigned group, keyed
  // by group id, so several can be filled in without clobbering each other. (State declared earlier,
  // above refreshReview — see the comment there.)
  const updateAssignForm = (groupId, field, value) => setAssignForms(prev => ({ ...prev, [groupId]: { ...prev[groupId], [field]: value } }));
  // Same two-step fix an unassigned row already gets from a Customer Mapping rule: (1) make sure a
  // Product Catalog entry exists mapping this item to the chosen customer + sheet tab, so it lands
  // in the right block when the customer's sheet is next generated/pushed, and (2) add an exact-match
  // Customer Mapping keyword so this and any future row with the same description routes here
  // automatically. Then re-run the same reassignUnassignedRows() pass confirmSheetImport already
  // relies on to move this (and anything else now matching) off Unassigned immediately.
  //
  // Alias handling: `description` is the raw wording this row actually scanned as (e.g. "HANDLE
  // LOCK"); `item` is whatever the person typed/picked in the Block field, which may be a DIFFERENT,
  // already-known block name they're merging this row into (e.g. "Tijori Handle"). If those two
  // differ, this is exactly the "an item got scanned under an alias" case — registering only the
  // canonical `item` name would fix THIS row but leave the next dispatch bill that also says "HANDLE
  // LOCK" falling right back into Unassigned. So both spellings get registered against the same
  // customer + sheet tab: the raw alias text (so future rows with this exact wording route straight
  // through) and the canonical block name (so it's tracked too, in case it wasn't already).
  const assignUnassignedGroup = (groupId, description) => {
    const form = assignForms[groupId] || {};
    const customer = (form.customer || '').trim();
    const sheetGroup = (form.sheetGroup || '').trim();
    const rawAlias = (description || '').trim();
    const item = (form.item || rawAlias).trim();
    if (!customer || !sheetGroup || !item) return;
    registerAliases(customer, [item, rawAlias], sheetGroup, item);
    reassignUnassignedRows();
    setAssignForms(prev => { const next = { ...prev }; delete next[groupId]; return next; });
  };
  // Assign form for "unmatched" items — a confirmed row already routed to the right CUSTOMER (it
  // matched a Customer Mapping rule), but its exact wording doesn't match any Product Catalog item,
  // so buildCustomerSheetPayload has nowhere to route it and silently excludes it from what gets
  // pushed. Same underlying fix as the Unassigned assign form (register the wording as an alias), just
  // with the customer field already fixed instead of picked — there's nothing to guess there. (State
  // declared earlier, above refreshReview — see the comment there.)
  const updateUnmatchedAssignForm = (key, field, value) => setUnmatchedAssignForms(prev => ({ ...prev, [key]: { ...prev[key], [field]: value } }));
  const assignUnmatchedItem = (customer, description) => {
    const key = `${customer}::${description}`;
    const form = unmatchedAssignForms[key] || {};
    const sheetGroup = (form.sheetGroup || '').trim();
    const rawAlias = (description || '').trim();
    const item = (form.item || rawAlias).trim();
    if (!customer || !sheetGroup || !item) return;
    registerAliases(customer, [item, rawAlias], sheetGroup, item);
    setUnmatchedAssignForms(prev => { const next = { ...prev }; delete next[key]; return next; });
  };
  // Add-new-alias form for the dedicated Aliases tab.
  const [newAliasForm, setNewAliasForm] = useState({ customer: '', item: '', sheetGroup: '', block: '' });
  const updateNewAliasForm = (field, value) => setNewAliasForm(prev => ({ ...prev, [field]: value }));
  const addAliasFromForm = () => {
    const { customer, item, sheetGroup, block } = newAliasForm;
    if (!customer.trim() || !item.trim() || !sheetGroup.trim()) return;
    registerAliases(customer, [item], sheetGroup, block);
    setNewAliasForm({ customer: '', item: '', sheetGroup: '', block: '' });
  };
  const pushCustomerSheetNow = async (customer) => {
    const sheetId = getCustomerSheetId(customer).trim();
    if (!sheetId) {
      setPushStatus(prev => ({ ...prev, [customer]: { state: 'error', message: 'Add a Google Sheet ID for this customer first (see field above).' } }));
      return;
    }
    // Always the CURRENT edited payload (review edits — value changes, deletions, moves, added rows —
    // applied on top of the freshly computed ledger), so what gets written always matches exactly what
    // was last shown on screen. No separate "Approve" step anymore — just a plain confirm, same as any
    // other real write in this app (Clear Data, Discard changes, etc.).
    const { itemGroups, unmatched } = getEditedPayload(customer);
    if (!itemGroups.some(g => (g.variants || []).some(v => (v.rows || []).length > 0))) {
      setPushStatus(prev => ({ ...prev, [customer]: { state: 'error', message: 'Nothing new to push right now.' } }));
      return;
    }
    if (!window.confirm(`Push ${itemGroups.length} item tab${itemGroups.length === 1 ? '' : 's'} to ${customer}'s Google Sheet now?`)) return;
    setPushStatus(prev => ({ ...prev, [customer]: { state: 'pushing', message: '', unmatched } }));
    try {
      const res = await fetch('/api/customer-sheets/push', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ spreadsheetId: sheetId, itemGroups }),
      });
      if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return; }
      const data = await res.json().catch(() => ({}));
      if (res.ok && data.ok) {
        // A push can succeed (every genuinely new row written fine) while ALSO leaving some dates
        // unresolved — a date already had a row in the sheet, but with different numbers than what we
        // sent, so nothing was written for it (existing rows are never touched). That's not a failure,
        // but it must never look like a silent, complete success either.
        const mismatchCount = (data.results || []).reduce((s, r) => s + (r.mismatches || []).reduce((s2, m) => s2 + ((m.mismatches || []).length), 0), 0);
        const mismatchNote = mismatchCount ? ` ${mismatchCount} date${mismatchCount === 1 ? '' : 's'} already had a row with different numbers and ${mismatchCount === 1 ? "wasn't" : "weren't"} touched — see the Preview below for details.` : '';
        setPushStatus(prev => ({ ...prev, [customer]: { state: 'done', message: `Pushed ${itemGroups.length} item tab${itemGroups.length === 1 ? '' : 's'} just now.${mismatchNote}`, unmatched } }));
        patchCustomerSheetEntry(customer, { sheetId, lastPushedAt: new Date().toISOString() });
        // Refreshes this customer's slice of the Customer Sheets Mirror with what's really in the Sheet
        // post-push (the server re-reads it fresh — see pushCustomerSheetHandler) so search reflects the
        // just-written rows immediately, not just whatever the mirror last had at import time.
        replaceCustomerMirrorRows(customer, data.mirrorRows || []);
        // The staged edits were for THIS specific diff — once it's actually written, their values are
        // now baked into the real rows the app just appended, so clear them and pull a fresh diff.
        setReviewEdits(prev => { const next = { ...prev }; delete next[customer]; return next; });
        refreshReview(customer);
      } else {
        const failedTabs = (data.results || []).filter(r => !r.ok).map(r => `${r.tab}: ${r.error}`).join(' · ');
        setPushStatus(prev => ({ ...prev, [customer]: { state: 'error', message: failedTabs || data.error || `Push failed (HTTP ${res.status}).`, unmatched } }));
      }
    } catch (e) {
      setPushStatus(prev => ({ ...prev, [customer]: { state: 'error', message: e.message || 'Network error — could not reach the server.', unmatched } }));
    }
  };
  /* -------- export --------
     Every export button below just packages the relevant rows/columns into a { title, sheets } object
     and hands it to the CopyExportModal — it does NOT build an .xlsx workbook or trigger a download
     itself. That's deliberate: packaging plain JS data can't fail, so the modal (with the
     always-works copy-paste view) is guaranteed to open. The actual .xlsx file is only built, on
     demand, if the person clicks "Download .xlsx" inside that modal. */
  const [copyModal, setCopyModal] = useState(null); // { title, sheets: [{name, rows, columns, kind, title}] }
  const exportSheet = (name, rows, columns) => {
    setCopyModal({ title: name, sheets: [{ name, rows, columns, kind: 'table' }] });
  };
  const exportCustomerStock = (customer) => {
    const groupsForCustomer = customerStockGroups.filter(g => g.customer === customer);
    const stockColumns = [
      { key: 'date', label: 'Date' }, { key: 'opening', label: 'Opening' }, { key: 'pieces', label: 'Production' },
      { key: 'dispatch', label: 'Dispatch' }, { key: 'closing', label: 'Closing' },
    ];
    const sheets = groupsForCustomer.length
      ? groupsForCustomer.map(g => ({ name: g.description || 'variant', title: g.description, rows: g.ledger, columns: stockColumns, kind: 'titled' }))
      : [{ name: 'No data yet', rows: [], columns: stockColumns, kind: 'table' }];
    setCopyModal({ title: `${customer} Stock`, sheets });
  };
  const exportAll = () => {
    const sheets = [];
    sheets.push({ name: 'Raw Material In', rows: rawMaterialIn, columns: COLUMNS.rawMaterialIn, kind: 'table' });
    sheets.push({ name: 'Consumption', rows: consumption, columns: COLUMNS.consumption, kind: 'table' });
    sheets.push({
      name: 'RM Balance', kind: 'table', rows: balanceRows, columns: [
        { key: 'size', label: 'Size' }, { key: 'gsm', label: 'GSM' },
        { key: 'weight_in', label: 'Total In (kg)' }, { key: 'weight_consumed', label: 'Total Consumed (kg)' }, { key: 'balance', label: 'Balance Left (kg)' },
      ],
    });
    sheets.push({ name: 'Production Register', rows: production, columns: COLUMNS.production, kind: 'table' });
    sheets.push({ name: 'Customer Dispatch Bills', rows: customerDispatch, columns: COLUMNS.customerDispatch, kind: 'table' });
    sheets.push({ name: 'Dabur Spec Master', rows: daburSpecs, columns: COLUMNS.daburSpecs, kind: 'table' });
    sheets.push({
      name: 'Dabur Pending PO', kind: 'table', rows: daburPOWithPending.map(r => ({ ...r, status: r.fulfilled ? 'Fulfilled' : 'Pending' })), columns: [
        ...COLUMNS.daburPO, { key: 'dispatched_qty', label: 'Dispatched Qty' }, { key: 'pending_qty', label: 'Pending Qty' }, { key: 'status', label: 'Status' },
      ],
    });
    sheets.push({ name: 'Dabur Dispatch', rows: daburDispatch, columns: COLUMNS.daburDispatch, kind: 'table' });
    customerNames.forEach(customer => {
      const rows = customerStockGroups.filter(g => g.customer === customer).flatMap(g => g.ledger.map(e => ({
        description: g.description, date: e.date, opening: e.opening, production: e.pieces, dispatch: e.dispatch, closing: e.closing,
      })));
      sheets.push({
        name: `Stock - ${customer}`, kind: 'table', rows, columns: [
          { key: 'description', label: 'Description' }, { key: 'date', label: 'Date' }, { key: 'opening', label: 'Opening' },
          { key: 'production', label: 'Production' }, { key: 'dispatch', label: 'Dispatch' }, { key: 'closing', label: 'Closing' },
        ],
      });
    });
    setCopyModal({ title: 'All Registers', sheets });
  };
  /* ============================== render ============================== */
  const counts = {
    rawMaterialIn: rawMaterialIn.length, consumption: consumption.length, production: production.length,
    customerDispatch: customerDispatch.length, daburSpecs: daburSpecs.length, daburPO: daburPO.length, daburDispatch: daburDispatch.length,
  };
  return (
    <>
    <div className="fims-root">
      <style>{`
        .fims-root {
          --paper: #ece8df;
          --paper-raised: #f6f3ec;
          --ink: #23262b;
          --ink-soft: #5b5e63;
          --rule: #cfc9ba;
          --ledger-red: #a23b2e;
          --accent: #a97a2f;
          --accent-soft: #f0e2c4;
          --ok: #2f6f5e;
          --ok-soft: #dceae4;
          --warn-soft: #f3ddd6;
          font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif;
          background: var(--paper);
          color: var(--ink);
          min-height: 100%;
          display: flex;
          border-radius: 8px;
          overflow: hidden;
          border: 1px solid var(--rule);
        }
        .fims-root * { box-sizing: border-box; }
        .fims-root h1, .fims-root h2, .fims-root h3 {
          font-family: 'Fraunces', Georgia, serif;
          margin: 0;
          color: var(--ink);
        }
        .sidebar {
          width: 232px;
          flex-shrink: 0;
          background: var(--ink);
          color: #e8e5dc;
          display: flex;
          flex-direction: column;
          padding: 20px 0;
          transition: width 0.15s ease;
        }
        .sidebar.collapsed { width: 60px; }
        .sidebar.collapsed .nav-item { justify-content: center; padding: 10px 0; }
        .sidebar.collapsed .nav-item span { display: none; }
        .brand {
          padding: 0 18px 18px 18px;
          border-bottom: 1px solid rgba(255,255,255,0.12);
          margin-bottom: 10px;
        }
        .sidebar.collapsed .brand { padding: 0 0 14px 0; text-align: center; }
        .brand h1 { font-size: 17px; line-height: 1.3; color: #f6f3ec; }
        .brand p { font-size: 11px; color: #a7a396; margin-top: 4px; letter-spacing: 0.04em; text-transform: uppercase; }
        .sidebar-toggle {
          margin: 10px auto 0 auto;
          background: transparent; border: 1px solid rgba(255,255,255,0.18); color: #cfcabd;
          border-radius: 6px; width: 26px; height: 26px; display: flex; align-items: center; justify-content: center;
          cursor: pointer; flex-shrink: 0;
        }
        .sidebar-toggle:hover { background: rgba(255,255,255,0.08); color: #f6f3ec; }
        .nav-item {
          display: flex; align-items: center; gap: 10px;
          padding: 10px 18px;
          font-size: 13.5px;
          cursor: pointer;
          color: #cfcabd;
          border-left: 3px solid transparent;
          transition: background 0.12s, color 0.12s;
        }
        .nav-item:hover { background: rgba(255,255,255,0.05); color: #f6f3ec; }
        .nav-item.active { background: rgba(169,122,47,0.18); color: #f6f3ec; border-left-color: var(--accent); }
        .nav-count { margin-left: auto; font-size: 11px; background: rgba(255,255,255,0.12); padding: 1px 7px; border-radius: 20px; color: #e8e5dc; }
        .main {
          flex: 1;
          min-width: 0;
          display: flex;
          flex-direction: column;
          max-height: 88vh;
          overflow-y: auto;
        }
        .topbar {
          display: flex; align-items: center; justify-content: space-between;
          padding: 16px 26px;
          background: var(--paper-raised);
          border-bottom: 1px solid var(--rule);
          position: sticky; top: 0; z-index: 5;
        }
        .topbar h2 { font-size: 20px; }
        .content { padding: 24px 26px 40px 26px; }
        .panel { background: var(--paper-raised); border: 1px solid var(--rule); border-radius: 6px; padding: 20px; margin-bottom: 20px; position: relative; }
        .panel::before {
          content: ''; position: absolute; left: 14px; top: 0; bottom: 0; width: 2px; background: var(--ledger-red); opacity: 0.55;
        }
        .panel { padding-left: 30px; }
        .panel-header { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 14px; gap: 12px; }
        .subtitle { font-size: 12.5px; color: var(--ink-soft); margin-top: 4px; max-width: 560px; }
        .btn {
          font-family: inherit; font-size: 13px; font-weight: 600; cursor: pointer;
          border-radius: 5px; padding: 8px 14px; display: inline-flex; align-items: center; gap: 6px;
          border: 1px solid transparent; transition: transform 0.08s, background 0.12s;
        }
        .btn:active { transform: scale(0.98); }
        .btn-primary { background: var(--accent); color: #fff; }
        .btn-primary:hover { background: #96692a; }
        .btn-primary:disabled { background: #c9bda0; cursor: not-allowed; }
        .btn-ghost { background: transparent; border-color: var(--rule); color: var(--ink); }
        .btn-ghost:hover { background: var(--accent-soft); }
        .btn-danger { background: var(--ledger-red); color: #fff; }
        .icon-btn { background: transparent; border: none; cursor: pointer; color: var(--ink-soft); padding: 4px; border-radius: 4px; }
        .icon-btn.danger:hover { color: var(--ledger-red); background: var(--warn-soft); }
        .table-wrap { overflow-x: auto; border: 1px solid var(--rule); border-radius: 4px; }
        table { border-collapse: collapse; width: 100%; font-family: 'IBM Plex Mono', monospace; font-size: 12.5px; }
        thead th {
          text-align: left; background: var(--accent-soft); color: #5c4419; font-family: 'IBM Plex Sans', sans-serif;
          font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; padding: 8px 10px; border-bottom: 2px solid var(--rule); white-space: nowrap;
        }
        tbody td { border-bottom: 1px solid var(--rule); padding: 2px 4px; }
        tbody tr:hover { background: rgba(169,122,47,0.06); }
        .cell-input {
          width: 100%; border: 1px solid transparent; background: transparent; font: inherit; color: var(--ink);
          padding: 6px 6px; border-radius: 3px; min-width: 70px;
        }
        .cell-input:focus { outline: none; border-color: var(--accent); background: #fff; }
        .col-action { width: 34px; text-align: center; }
        .empty-state { padding: 30px; text-align: center; color: var(--ink-soft); font-size: 13px; border: 1px dashed var(--rule); border-radius: 6px; }
        .pill { font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; display: inline-block; }
        .pill-ok { background: var(--ok-soft); color: var(--ok); }
        .pill-warn { background: var(--warn-soft); color: var(--ledger-red); }
        .pill-neutral { background: var(--accent-soft); color: #5c4419; }
        .dash-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 14px; margin-bottom: 8px; }
        .dash-card {
          background: var(--paper-raised); border: 1px solid var(--rule); border-radius: 6px; padding: 16px 18px; cursor: pointer;
          transition: transform 0.1s, box-shadow 0.1s;
        }
        .dash-card:hover { transform: translateY(-2px); box-shadow: 0 4px 10px rgba(0,0,0,0.08); }
        .dash-card .num { font-family: 'Fraunces', serif; font-size: 30px; }
        .dash-card .lbl { font-size: 12px; color: var(--ink-soft); margin-top: 2px; }
        .guide-step { display: flex; gap: 14px; padding: 16px 0; border-bottom: 1px solid var(--rule); cursor: pointer; }
        .guide-step:last-child { border-bottom: none; }
        .guide-step:hover .guide-title { color: var(--accent); }
        .guide-num {
          flex-shrink: 0; width: 30px; height: 30px; border-radius: 50%; background: var(--accent); color: #fff;
          font-family: 'Fraunces', serif; font-size: 14px; display: flex; align-items: center; justify-content: center;
        }
        .guide-title { font-size: 14.5px; font-weight: 700; margin-bottom: 4px; transition: color 0.12s; }
        .guide-what { font-size: 12.5px; color: var(--ink); margin-bottom: 4px; line-height: 1.5; }
        .guide-how { font-size: 12.5px; color: var(--ink-soft); line-height: 1.5; }
        .guide-how b { color: var(--ink); font-weight: 600; }
        .upload-grid { display: grid; grid-template-columns: 1.1fr 1fr; gap: 20px; align-items: start; }
        @media (max-width: 880px) { .upload-grid { grid-template-columns: 1fr; } }
        .doc-select { width: 100%; padding: 10px 12px; border-radius: 5px; border: 1px solid var(--rule); font: inherit; font-size: 13.5px; background: #fff; margin-bottom: 6px; }
        .doc-hint { font-size: 12px; color: var(--ink-soft); margin-bottom: 16px; }
        .dropzone {
          display: flex; flex-direction: column; align-items: center; justify-content: center;
          width: 100%; box-sizing: border-box;
          border: 2px dashed var(--rule); border-radius: 8px; padding: 26px; text-align: center;
          background: #fff; cursor: pointer; transition: border-color 0.12s, background 0.12s;
        }
        .dropzone:hover { background: var(--accent-soft); border-color: var(--accent); }
        .preview-img { max-width: 100%; max-height: 70vh; object-fit: contain; border-radius: 6px; border: 1px solid var(--rule); margin-top: 12px; cursor: zoom-in; background: #fff; }
        .error-box { display: flex; gap: 8px; align-items: flex-start; background: var(--warn-soft); border: 1px solid var(--ledger-red); color: #6b241a; padding: 10px 12px; border-radius: 5px; font-size: 12.5px; margin: 10px 0; }
        .flagged-row { background: var(--warn-soft) !important; }
        /* A different tone than .flagged-row on purpose — flagged-row means "extraction wasn't
           confident," this means "confirmed, but no real customer attached yet" (still Unassigned, or
           a suggestion that didn't match anyone known). Disappears automatically the moment the row
           gets a real customer, so it's always safe to leave a row highlighted until you get to it. */
        .needs-customer-row { background: var(--accent-soft) !important; }
        /* A row on Pending Production/Dispatch Review whose customer IS known but whose item isn't in
           the Product Catalog yet — i.e. it needs a Sheet tab/Block picked by hand right here. Same
           warning tone as .flagged-row (deliberately, since it's the same "needs your attention before
           this counts" meaning) but its own class so it's never confused with extraction confidence. */
        .needs-tabblock-row { background: var(--warn-soft) !important; }
        .batch-fill-row th { background: var(--accent-soft); font-weight: 400; padding: 4px 6px; }
        .col-select { width: 26px; text-align: center; cursor: pointer; user-select: none; }
        /* A distinct tone from flagged-row/needs-customer-row's amber/red so a selected row still
           reads clearly even when it also happens to be flagged. */
        .row-selected { background: var(--ok-soft) !important; }
        .info-box { display: flex; gap: 8px; align-items: flex-start; background: var(--accent-soft); border: 1px solid var(--rule); color: var(--ink); padding: 10px 12px; border-radius: 5px; font-size: 12.5px; margin: 10px 0; }
        .field-row { display: flex; gap: 10px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }
        .text-input { padding: 8px 10px; border-radius: 5px; border: 1px solid var(--rule); font: inherit; font-size: 13px; }
        .review-box { border: 1px solid var(--accent); border-radius: 6px; padding: 16px; background: #fff; margin-top: 18px; }
        .review-actions { display: flex; gap: 10px; margin-top: 14px; }
        .section-label { font-size: 13px; font-weight: 700; margin: 18px 0 8px 0; color: var(--ink); }
        /* Customer Stock's per-customer panels are collapsible <details> so a customer with nothing
           to look at collapses to one line instead of taking up the same space as one that needs a
           decision — the page used to be a long wall of identical-looking tables no matter how much
           (or little) actually needed attention; this is the fix for that. */
        .customer-review-summary { list-style: none; cursor: pointer; display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
        .customer-review-summary::-webkit-details-marker { display: none; }
        .customer-review-summary-main { display: flex; align-items: center; gap: 8px; }
        .customer-review-summary-main h2 { margin: 0; }
        .customer-review-summary-badges { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
        .details-chevron { flex-shrink: 0; color: var(--ink-soft); transition: transform 0.12s; }
        .customer-review-panel[open] > .customer-review-summary .details-chevron { transform: rotate(90deg); }
        .customer-review-panel:not([open]) { padding-bottom: 20px; }
        /* Bold, high-contrast dividers between the three very different kinds of things Customer
           Stock shows (things waiting on a decision, things nobody's claimed at all, and the running
           per-customer ledgers) — distinct enough in weight that the eye can jump straight to the
           right zone instead of parsing every panel in order to find it. */
        .stock-section-divider { display: flex; align-items: baseline; gap: 10px; margin: 28px 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid var(--ink); }
        .stock-section-divider:first-child { margin-top: 0; }
        .stock-section-divider-label { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink); }
        .spin { animation: fims-spin 1s linear infinite; }
        @keyframes fims-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      `}</style>
      <aside className={`sidebar ${sidebarCollapsed ? 'collapsed' : ''}`}>
        <div className="brand">
          {sidebarCollapsed
            ? <div style={{ fontSize: 13, fontWeight: 700, color: '#f6f3ec' }} title="Shyam Adarsh Pack — Inventory & Production">SA</div>
            : <><h1>Shyam Adarsh Pack</h1><p>Inventory &amp; Production</p></>}
        </div>
        {NAV.map(item => {
          const Icon = item.icon;
          const count = counts[item.key];
          return (
            <div key={item.key} className={`nav-item ${activeTab === item.key ? 'active' : ''}`} onClick={() => setActiveTab(item.key)} title={sidebarCollapsed ? item.label : undefined}>
              <Icon size={16} />
              <span>{item.label}</span>
              {!sidebarCollapsed && typeof count === 'number' && <span className="nav-count">{count}</span>}
            </div>
          );
        })}
        <button className="sidebar-toggle" onClick={() => setSidebarCollapsed(v => !v)} title={sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}>
          {sidebarCollapsed ? <ChevronRight size={15} /> : <ChevronLeft size={15} />}
        </button>
      </aside>
      <div className="main">
        <div className="topbar">
          <h2>{activeTab === 'search' ? `Search: "${globalQuery}"` : NAV.find(n => n.key === activeTab)?.label}</h2>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <input
              className="text-input"
              placeholder="Search every register — party, item, invoice, PO no…"
              value={globalQuery}
              onChange={e => setGlobalQuery(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter' && globalQuery.trim()) setActiveTab('search'); }}
              style={{ width: 280 }}
            />
            <button className="btn" onClick={() => { if (globalQuery.trim()) setActiveTab('search'); }} title="Search across every register"><Search size={15} /></button>
            <button className="btn btn-primary" onClick={exportAll}><FileSpreadsheet size={15} /> Export all to Excel</button>
          </div>
        </div>
        <div className="content">
          {!loaded && <div className="empty-state">Loading your registers…</div>}
          {loaded && activeTab === 'dashboard' && (
            <div>
              <div className="panel">
                <h2 style={{ marginBottom: 4 }}>How this system works</h2>
                <p className="subtitle" style={{ marginBottom: 6 }}>The order things usually happen in, start to finish. Click any step to jump straight to that tab.</p>
                {GUIDE_STEPS.map((step, i) => (
                  <div className="guide-step" key={i} onClick={() => setActiveTab(step.tab)}>
                    <div className="guide-num">{i + 1}</div>
                    <div>
                      <div className="guide-title">{step.title}</div>
                      <div className="guide-what">{step.what}</div>
                      <div className="guide-how"><b>Do this:</b> {step.how}</div>
                    </div>
                  </div>
                ))}
              </div>
              <p className="subtitle" style={{ marginBottom: 18 }}>Everything below is stored on this device for your account and stays put between visits. Click any card to jump to that register.</p>
              <div className="dash-grid">
                {NAV.filter(n => n.key !== 'dashboard' && n.key !== 'upload' && n.key !== 'customerStock' && n.key !== 'customerMapping' && n.key !== 'aliases' && n.key !== 'settings').map(n => (
                  <div className="dash-card" key={n.key} onClick={() => setActiveTab(n.key)}>
                    <div className="num">{counts[n.key]}</div>
                    <div className="lbl">{n.label}</div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {loaded && activeTab === 'upload' && (
            <div className="upload-grid">
              <div className="panel" style={{ paddingLeft: 30 }}>
                <h2 style={{ marginBottom: 14 }}>1. Choose document &amp; upload</h2>
                <select className="doc-select" value={docType} onChange={(e) => { setDocType(e.target.value); setFileResults([]); setActiveResultIndex(0); setPreview(null); setBase64Img(null); setQueuedPages([]); setSkippedPages([]); }}>
                  {DOCUMENT_TYPES.map(d => <option key={d.key} value={d.key}>{d.label}</option>)}
                </select>
                <div className="doc-hint">{activeConfig.hint}</div>
                <div className="doc-hint" style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: -8 }}>
                  {(trainingExamples[docType]?.length || 0) > 0
                    ? <>
                        <span>📚 {trainingExamples[docType].length} correction{trainingExamples[docType].length === 1 ? '' : 's'} learned for this document type.</span>
                        <button className="icon-btn" style={{ textDecoration: 'underline', fontSize: 11.5 }} onClick={clearTrainingForType}>reset</button>
                      </>
                    : <span>No corrections learned yet for this document type — edits you make in the review step below will be remembered for next time.</span>}
                </div>
                <label className="dropzone" htmlFor="fileInput">
                  <Upload size={22} style={{ marginBottom: 6 }} />
                  <div style={{ fontSize: 13.5, fontWeight: 600 }}>Click to upload photos, scans, or PDFs</div>
                  <div style={{ fontSize: 11.5, color: 'var(--ink-soft)', marginTop: 4 }}>JPG, PNG, or PDF — select multiple files at once if you like. Emailed or hard-copy docs — scan, photograph, or save the PDF, then upload here.</div>
                </label>
                <input id="fileInput" type="file" accept="image/*,application/pdf" multiple style={{ display: 'none' }} onChange={(e) => handleFiles(e.target.files)} />
                {pdfLoading && <div className="doc-hint" style={{ marginTop: 10 }}><Loader2 size={13} className="spin" style={{ verticalAlign: 'middle', marginRight: 6 }} />Reading file(s)…</div>}
                {preview && <img src={preview} alt="preview" className="preview-img" title="Click to enlarge" onClick={() => setZoomedImage(preview)} />}
                {queuedPages.length > 1 && (
                  <div style={{ marginTop: 10 }}>
                    <div className="doc-hint">{queuedIndex + 1} of {queuedPages.length}: {queuedPages[queuedIndex]?.label} — click a thumbnail to preview it, click × to drop it from the queue (works anytime, even mid-extraction), or extract everything at once.</div>
                    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 6 }}>
                      {queuedPages.map((p, idx) => (
                        <div key={p.id} style={{ position: 'relative' }}>
                          <img src={p.dataUrl} alt={p.label} title={p.label}
                            onClick={() => selectQueuedPage(idx)}
                            style={{ width: 54, height: 54, objectFit: 'cover', borderRadius: 4, cursor: 'pointer', border: idx === queuedIndex ? '2px solid var(--accent)' : '1px solid var(--rule)' }} />
                          <button
                            title="Remove from queue"
                            onClick={(e) => { e.stopPropagation(); removeQueuedPage(p.id); }}
                            style={{ position: 'absolute', top: -6, right: -6, width: 18, height: 18, borderRadius: '50%', border: 'none', background: 'var(--ledger-red)', color: '#fff', fontSize: 11, lineHeight: '18px', textAlign: 'center', padding: 0, cursor: 'pointer' }}>
                            ×
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {skippedPages.length > 0 && (
                  <div style={{ marginTop: 10 }}>
                    <div className="doc-hint" style={{ marginBottom: 4 }}>Auto-skipped as duplicate/e-Way Bill pages (no API call used) — click + to queue one anyway if you actually need it:</div>
                    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                      {skippedPages.map(p => (
                        <div key={p.id} style={{ position: 'relative' }}>
                          <img src={p.dataUrl} alt={p.label} title={`${p.label} (${p.copyLabel})`}
                            style={{ width: 54, height: 54, objectFit: 'cover', borderRadius: 4, opacity: 0.5, border: '1px dashed var(--rule)' }} />
                          <button
                            title="Queue this page anyway"
                            onClick={() => restoreSkippedPage(p.id)}
                            style={{ position: 'absolute', top: -6, right: -6, width: 18, height: 18, borderRadius: '50%', border: 'none', background: 'var(--accent)', color: '#fff', fontSize: 12, lineHeight: '18px', textAlign: 'center', padding: 0, cursor: 'pointer' }}>
                            +
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {errorMsg && <div className="error-box"><AlertCircle size={16} /><span>{errorMsg}</span></div>}
                <div style={{ marginTop: 14, display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                  <button className="btn btn-primary" disabled={!base64Img || extracting} onClick={runExtraction}>
                    {extracting ? <Loader2 size={15} className="spin" /> : <ImageIcon size={15} />}
                    {extracting ? 'Reading document…' : queuedPages.length > 1 ? `Extract this one only` : 'Extract data'}
                  </button>
                  {queuedPages.length > 1 && (
                    <button className="btn btn-ghost" disabled={extracting} onClick={runExtractionAllQueued}>
                      <FileText size={15} /> Extract all {queuedPages.length}
                    </button>
                  )}
                  {extracting && (
                    <button className="btn btn-danger" onClick={cancelExtraction}>
                      <XCircle size={15} /> Cancel
                    </button>
                  )}
                </div>
              </div>
              <div className="panel" style={{ paddingLeft: 30 }}>
                <h2 style={{ marginBottom: 14 }}>2. Review &amp; confirm</h2>
                {!fileResults.length && <div className="empty-state">Extracted rows will show up here for you to check and correct before they're added to the register. Nothing is saved automatically.</div>}
                {fileResults.length > 0 && (() => {
                  const idx = Math.min(activeResultIndex, fileResults.length - 1);
                  const page = fileResults[idx];
                  const doneCount = fileResults.filter(r => r.status === 'done').length;
                  const remainingCount = fileResults.filter(r => r.status === 'pending' || r.status === 'error').length;
                  return (
                    <div>
                      {fileResults.length > 1 && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, flexWrap: 'wrap' }}>
                          <button className="icon-btn" disabled={idx === 0} onClick={() => setActiveResultIndex(idx - 1)}>◀</button>
                          <span style={{ fontSize: 12.5 }}>File {idx + 1} of {fileResults.length} — {doneCount} done so far</span>
                          <button className="icon-btn" disabled={idx === fileResults.length - 1} onClick={() => setActiveResultIndex(idx + 1)}>▶</button>
                          <div style={{ display: 'flex', gap: 5, marginLeft: 6, flexWrap: 'wrap' }}>
                            {fileResults.map((r, i) => (
                              <span key={r.id} onClick={() => setActiveResultIndex(i)} title={r.label}
                                style={{
                                  width: 22, height: 22, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                                  fontSize: 10, cursor: 'pointer', border: i === idx ? '2px solid var(--accent)' : '1px solid var(--rule)',
                                  background: r.status === 'done' ? 'var(--ok-soft)' : r.status === 'error' ? 'var(--warn-soft)' : r.status === 'extracting' ? 'var(--accent-soft)' : '#fff',
                                }}>
                                {r.status === 'extracting' ? <Loader2 size={11} className="spin" /> : i + 1}
                              </span>
                            ))}
                          </div>
                        </div>
                      )}
                      <p className="subtitle" style={{ marginBottom: 10 }}>{page.label}</p>
                      {page.status === 'extracting' && <div className="doc-hint"><Loader2 size={13} className="spin" style={{ verticalAlign: 'middle', marginRight: 6 }} />Reading this one…</div>}
                      {page.status === 'pending' && <div className="doc-hint">Waiting its turn…</div>}
                      {page.status === 'error' && (
                        <div>
                          <div className="error-box"><AlertCircle size={16} /><span>Failed: {page.error}</span></div>
                          <button className="btn btn-ghost" onClick={() => selectQueuedPage(idx)}>Select this file, then hit "Extract this one only" to retry</button>
                        </div>
                      )}
                      {page.status === 'done' && (
                        <div>
                          {page.truncated && (
                            <div className="error-box" style={{ marginBottom: 10 }}>
                              <AlertCircle size={16} />
                              <span>Claude's response was cut off before it finished this page — there may be more rows than the {page.rows.length} shown below. Check against the original document, then re-extract if anything's missing (a fresh attempt isn't guaranteed to hit the same cutoff).</span>
                            </div>
                          )}
                          {page.rows.some(r => r.flagged) && (
                            <div className="error-box" style={{ marginBottom: 10 }}>
                              <AlertCircle size={16} />
                              <span>{page.rows.filter(r => r.flagged).length} row(s) below are flagged (highlighted, with a warning icon on the first cell) — the model wasn't confident about a value, usually because a column looked cut off/missing or a crossed-out number was ambiguous. Check those against the original document and fix by hand; nothing was guessed for them.</span>
                            </div>
                          )}
                          <p className="subtitle" style={{ marginBottom: 10 }}>{page.rows.length} row(s) found for <strong>{activeConfig.label}</strong>. Fix anything that looks wrong, then confirm.</p>
                          <EditableTable columns={COLUMNS[activeConfig.register]} rows={page.rows}
                            onUpdate={(rowId, field, value) => updateReviewCell(idx, rowId, field, value)}
                            onDelete={(rowId) => deleteReviewRow(idx, rowId)}
                            emptyLabel="All rows removed — nothing to add."
                            sortByDate={false} showBatchFill />
                          <div className="review-actions">
                            <button className="btn btn-primary" onClick={() => confirmPage(idx)} disabled={!page.rows.length}><CheckCircle2 size={15} /> Add {page.rows.length} row(s) to register</button>
                            {page.truncated && (
                              <button className="btn btn-ghost" disabled={extracting} onClick={() => reextractFile(page.id)}><RefreshCw size={15} /> Re-extract this file</button>
                            )}
                            <button className="btn btn-ghost" onClick={() => discardPage(idx)}><XCircle size={15} /> Discard this file</button>
                          </div>
                        </div>
                      )}
                      {fileResults.length > 1 && (doneCount > 0 || remainingCount > 0) && (
                        <div className="review-actions" style={{ borderTop: '1px solid var(--rule)', paddingTop: 14, marginTop: 14, flexWrap: 'wrap' }}>
                          {doneCount > 0 && <button className="btn btn-primary" onClick={confirmAllPages}><CheckCircle2 size={15} /> Confirm all {doneCount} completed file(s) at once</button>}
                          {extracting && <button className="btn btn-danger" onClick={cancelExtraction}><XCircle size={15} /> Cancel</button>}
                          {!extracting && remainingCount > 0 && <button className="btn btn-ghost" onClick={runExtractionRemaining}><Upload size={15} /> Retry remaining {remainingCount}</button>}
                          <button className="btn btn-ghost" disabled={extracting} onClick={discardAllPages}><XCircle size={15} /> Discard everything</button>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>
            </div>
          )}
          {loaded && activeTab === 'rawMaterialIn' && (
            <div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Inward Entries (from mill slips)</h2><p className="subtitle">One table per size, same as the physical register book. Nothing is summed here — every mill-slip line stays its own row. "Consumed" is left blank for now; it'll be filled in once entries are matched against consumption reports.</p></div>
                  <button className="btn btn-ghost" onClick={() => exportSheet('Raw_Material_In', rawMaterialIn, COLUMNS.rawMaterialIn)}><Download size={15} /> Export</button>
                </div>
                {!!rawMaterialIn.length && (
                  <div className="field-row" style={{ marginBottom: 14 }}>
                    <select className="doc-select" style={{ width: 160, marginBottom: 0 }} value={rmSizeFilter} onChange={e => setRmSizeFilter(e.target.value)}>
                      <option value="">All sizes</option>
                      {rawMaterialBySize.map(g => <option value={g.size} key={g.size}>Size {g.size}</option>)}
                    </select>
                    <select className="doc-select" style={{ width: 160, marginBottom: 0 }} value={rmGsmFilter} onChange={e => setRmGsmFilter(e.target.value)}>
                      <option value="">All GSM</option>
                      {rawMaterialGsmOptions.map(g => <option value={g} key={g}>GSM {g}</option>)}
                    </select>
                    {(rmSizeFilter || rmGsmFilter) && (
                      <button className="btn btn-ghost" onClick={() => { setRmSizeFilter(''); setRmGsmFilter(''); }}>Clear filters</button>
                    )}
                  </div>
                )}
                {!rawMaterialIn.length && <div className="empty-state">Upload some mill slips to see inward entries here.</div>}
                {!!rawMaterialIn.length && !rawMaterialGroupsFiltered.length && <div className="empty-state">No entries match that filter.</div>}
                {rawMaterialGroupsFiltered.map(group => (
                  <div key={group.size} style={{ marginBottom: 20 }}>
                    <div style={{ marginBottom: 6 }}><strong>Size {group.size}</strong></div>
                    <EditableTable columns={RAW_MATERIAL_SIZE_COLUMNS} rows={group.rows}
                      onUpdate={updateRow('rawMaterialIn')} onDelete={deleteRow('rawMaterialIn')} />
                  </div>
                ))}
              </div>
              <RegisterPanel title="Consumption Entries (from daily reports)" columns={COLUMNS.consumption} rows={consumption}
                onUpdate={updateRow('consumption')} onDelete={deleteRow('consumption')} onExport={() => exportSheet('Consumption', consumption, COLUMNS.consumption)} />
            </div>
          )}
          {loaded && activeTab === 'production' && (
            <RegisterPanel title="Production Register" subtitle="From handwritten daily production sheets — covers both the shade/size/GSM style and the product-description style. Rows from the description style also feed the Customer Stock tab. Dispatch bills do NOT land here — see the Customer Dispatch Bills tab." columns={COLUMNS.production} rows={production}
              onUpdate={updateRow('production')} onDelete={deleteRow('production')} onExport={() => exportSheet('Production_Register', production, COLUMNS.production)} suppressFlags highlightRow={needsCustomerHighlight} />
          )}
          {loaded && activeTab === 'customerDispatch' && (
            <RegisterPanel title="Customer Dispatch Bills" subtitle="From dispatch bills / tax invoices sent to customers (Bindal, Diamond, Anmol, or otherwise) — kept separate from the Production Register. Once confirmed on the Customer Stock tab, these reduce that customer's balance there." columns={COLUMNS.customerDispatch} rows={customerDispatch}
              onUpdate={updateRow('customerDispatch')} onDelete={deleteRow('customerDispatch')} onExport={() => exportSheet('Customer_Dispatch_Bills', customerDispatch, COLUMNS.customerDispatch)} highlightRow={needsCustomerHighlight} />
          )}
          {loaded && activeTab === 'search' && (
            <div>
              <div className="panel">
                <h2 style={{ marginBottom: 6 }}>Search Results</h2>
                <p className="subtitle" style={{ marginBottom: !globalSearchResults || (globalSearchResults.stockMatches.length === 0 && globalSearchResults.registerMatches.length === 0 && globalSearchResults.mirrorMatches.length === 0) ? 0 : 16 }}>
                  Matches for “{globalQuery}” across every register, Customer Stock, and every customer's real Google Sheet — plain substring match, case-insensitive. Register rows are the real row — fix or delete right here, same as on their own tab; Customer Sheet rows are read-only (they reflect the real Sheet — fix them there, on the Customer Sheets tab).
                </p>
                {(!globalSearchResults || (globalSearchResults.stockMatches.length === 0 && globalSearchResults.registerMatches.length === 0 && globalSearchResults.mirrorMatches.length === 0)) && (
                  <p className="subtitle">No matches found.</p>
                )}
              </div>
              {globalSearchResults && globalSearchResults.stockMatches.length > 0 && (
                <div className="panel">
                  <div className="section-label">Customer Stock ({globalSearchResults.stockMatches.length})</div>
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Customer</th><th>Item</th><th>Total Produced</th><th>Total Dispatched</th><th>Balance</th></tr></thead>
                      <tbody>
                        {globalSearchResults.stockMatches.map(g => (
                          <tr key={g.id}>
                            <td>{g.customer}</td><td>{g.description}</td>
                            <td>{g.totalProduction}</td><td>{g.totalDispatch}</td>
                            <td style={{ color: g.closingBalance > 0 ? 'var(--ok)' : 'var(--ledger-red)' }}>{g.closingBalance}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              {globalSearchResults && globalSearchResults.mirrorMatches.length > 0 && (
                <div className="panel">
                  <div className="section-label">Customer Sheets — real data ({globalSearchResults.mirrorMatchesTotal})</div>
                  {globalSearchResults.mirrorMatchesTruncated && (
                    <p className="doc-hint" style={{ marginBottom: 8 }}>Showing the first {MIRROR_SEARCH_RESULT_CAP} of {globalSearchResults.mirrorMatchesTotal} matches — narrow your search to see the rest.</p>
                  )}
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Customer</th><th>Sheet Tab</th><th>Block</th><th>Date</th><th>Opening</th><th>Production</th><th>Dispatch</th><th>Closing</th></tr></thead>
                      <tbody>
                        {globalSearchResults.mirrorMatches.map(r => (
                          <tr key={r.id}>
                            <td>{r.customer}</td><td>{r.sheetTab}</td><td>{r.block}</td><td>{r.date}</td>
                            <td>{r.opening}</td><td>{r.production}</td><td>{r.dispatch}</td><td>{r.closing}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              {globalSearchResults && globalSearchResults.registerMatches.map(reg => (
                <div className="panel" key={reg.key}>
                  <div className="section-label">{reg.label} ({reg.rows.length})</div>
                  <EditableTable columns={reg.columns} rows={reg.rows} onUpdate={updateRow(reg.key)} onDelete={deleteRow(reg.key)} emptyLabel="No matches." />
                </div>
              ))}
            </div>
          )}
          {loaded && activeTab === 'daburSpecs' && (
            <RegisterPanel title="Dabur — Spec Master" subtitle="Replaces the manual diary — one row per box item, from printed PM Specification sheets." columns={COLUMNS.daburSpecs} rows={daburSpecs}
              onUpdate={updateRow('daburSpecs')} onDelete={deleteRow('daburSpecs')} onExport={() => exportSheet('Dabur_Spec_Master', daburSpecs, COLUMNS.daburSpecs)} />
          )}
          {loaded && activeTab === 'daburPO' && (
            <div className="panel">
              <div className="panel-header">
                <div><h2>Dabur — Pending PO</h2><p className="subtitle">Pending quantity is ordered qty minus dispatches matched by PO number, with a ±10% tolerance — a PO is marked Fulfilled once dispatches reach 90% of the ordered quantity. If one PO covers several different materials, double-check against the Dabur Dispatch Log tab.</p></div>
                <button className="btn btn-ghost" onClick={() => exportSheet('Dabur_Pending_PO', daburPOWithPending.map(r => ({ ...r, status: r.fulfilled ? 'Fulfilled' : 'Pending' })), [...COLUMNS.daburPO, { key: 'dispatched_qty', label: 'Dispatched Qty' }, { key: 'pending_qty', label: 'Pending Qty' }, { key: 'status', label: 'Status' }])}><Download size={15} /> Export</button>
              </div>
              {!daburPOWithPending.length && <div className="empty-state">Upload a Dabur PO to start tracking pending quantities.</div>}
              {!!daburPOWithPending.length && (
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>{COLUMNS.daburPO.map(c => <th key={c.key}>{c.label}</th>)}<th>Dispatched</th><th>Pending</th><th className="col-action"></th></tr>
                    </thead>
                    <tbody>
                      {daburPOWithPending.map(row => (
                        <tr key={row.id}>
                          {COLUMNS.daburPO.map(c => (
                            <td key={c.key}><input className="cell-input" value={row[c.key] ?? ''} onChange={e => updateRow('daburPO')(row.id, c.key, e.target.value)} /></td>
                          ))}
                          <td style={{ padding: '6px 10px' }}>{row.dispatched_qty}</td>
                          <td style={{ padding: '6px 10px' }}>
                            <Pill tone={row.fulfilled ? 'ok' : 'warn'}>{row.fulfilled ? 'Fulfilled' : row.pending_qty}</Pill>
                          </td>
                          <td className="col-action"><button className="icon-btn danger" onClick={() => deleteRow('daburPO')(row.id)}><Trash2 size={15} /></button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
          {loaded && activeTab === 'daburDispatch' && (
            <RegisterPanel title="Dabur — Dispatch Log" subtitle="From dispatch bills against Dabur POs. The PO Ref No column drives the pending-quantity calculation on the Pending PO tab." columns={COLUMNS.daburDispatch} rows={daburDispatch}
              onUpdate={updateRow('daburDispatch')} onDelete={deleteRow('daburDispatch')} onExport={() => exportSheet('Dabur_Dispatch', daburDispatch, COLUMNS.daburDispatch)} />
          )}
          {loaded && activeTab === 'customerStock' && (
            <div>
              <div className="panel">
                <h2 style={{ marginBottom: 6 }}>Customer Stock</h2>
                <p className="subtitle">Review new Production Register and Customer Dispatch Bill entries here and confirm which customer they belong to. Matching is done by the Customer Mapping tab, plus anything literally bracketed next to the item name. "Push to Sheet" below confirms the customer AND pushes straight to their real Google Sheet in one step — duplicate rows already there are skipped automatically, server-side. The Customer Sheets tab is just for adding/syncing a Sheet ID.</p>
              </div>
              {Object.entries(pushStatus).filter(([, s]) => s && s.state && s.state !== 'idle').length > 0 && (
                <div className="panel">
                  <h2 style={{ marginBottom: 6 }}>Push Status</h2>
                  {Object.entries(pushStatus).filter(([, s]) => s && s.state && s.state !== 'idle').map(([customer, s]) => (
                    <div key={customer} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                      <strong>{customer}:</strong>
                      {s.state === 'pushing' && <span className="doc-hint"><Loader2 size={12} className="spin" style={{ verticalAlign: 'middle', marginRight: 4 }} />pushing…</span>}
                      {s.state === 'done' && <span className="doc-hint" style={{ color: 'var(--ok)' }}>✓ {s.message}</span>}
                      {s.state === 'error' && <span style={{ color: 'var(--ledger-red)', fontSize: 12.5 }}>{s.message}</span>}
                    </div>
                  ))}
                </div>
              )}
              {(pendingProductionRows.length > 0 || pendingDispatchRows.length > 0) && (
                <SectionDivider icon={ListChecks} label="Needs your input" hint="New entries whose customer isn't confirmed yet — nothing here counts toward any balance until you confirm it." />
              )}
              {pendingProductionRows.length > 0 && (
                <div className="panel" style={{ borderColor: 'var(--accent)' }}>
                  <div className="panel-header">
                    <div><h2>Pending Production Review ({pendingProductionRows.length})</h2><p className="subtitle">New Production Register entries not yet pushed to any customer's Sheet. Check or fix the suggested customer, then push — nothing here counts toward balances or reaches the real Sheet until you do.</p></div>
                    <button className="btn btn-primary" onClick={() => pushPendingRows('production', pendingProductionRows)}><FileSpreadsheet size={15} /> Push to Sheet</button>
                  </div>
                  {buildPendingGroups(pendingProductionRows).map(group => (
                    <PendingGroupBar key={group.key} group={group} guess={matchCustomer(group.rows[0])}
                      isKnownGuess={isKnownCustomerGuess(group.rows[0])} knownCustomers={allCustomerTabNames}
                      onBulkChange={updatePendingCustomerBulk('production')}
                      onConfirmGroup={() => group.rows.forEach(r => confirmStockRow('production', r))} />
                  ))}
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Date</th><th>Description</th><th>Pieces</th><th>Suggested Customer</th><th>Sheet tab</th><th>Block</th><th className="col-action"></th></tr></thead>
                      <tbody>
                        {pendingProductionRows.map(row => {
                          const effectiveCustomer = row.confirmedCustomer || (isKnownCustomerGuess(row) ? matchCustomer(row) : '');
                          const catalogEntry = getCatalogEntryForItem(effectiveCustomer, row.description);
                          const needsTabBlock = effectiveCustomer && !catalogEntry;
                          const draft = pendingTabBlockForms[row.id];
                          const tabBlockUnresolved = needsTabBlock && (!((draft && draft.sheetGroup) || '').trim() || !((draft && draft.block) || '').trim());
                          return (
                          <tr key={row.id} className={tabBlockUnresolved ? 'needs-tabblock-row' : ''}>
                            <td style={{ padding: '6px 10px' }}>{row.date}</td>
                            <td style={{ padding: '6px 10px' }}>{row.description}</td>
                            <td style={{ padding: '6px 10px' }}>{row.pieces || ''}</td>
                            <td>
                              <CustomerSuggestCell value={row.confirmedCustomer} guess={matchCustomer(row)}
                                isKnownGuess={isKnownCustomerGuess(row)} knownCustomers={allCustomerTabNames}
                                onChange={v => updatePendingCustomer('production')(row.id, v)} />
                            </td>
                            {needsTabBlock ? (
                              <TabBlockPickerCells rowId={row.id} description={row.description}
                                sheetGroupOptions={Array.from(new Set([
                                  ...productCatalog.filter(c => c.customer.toLowerCase() === effectiveCustomer.toLowerCase()).map(c => c.sheetGroup),
                                  ...getRealTabNamesForCustomer(effectiveCustomer),
                                ])).filter(Boolean)}
                                blockOptions={Array.from(new Set([
                                  ...getRealBlocksForTab(effectiveCustomer, (draft && draft.sheetGroup) || ''),
                                  ...getCatalogBlocksForTab(effectiveCustomer, (draft && draft.sheetGroup) || ''),
                                ])).filter(Boolean)}
                                draft={draft} onChange={(field, value) => updatePendingTabBlockForm(row.id, field, value)} />
                            ) : (
                              <>
                                <td className="doc-hint">{catalogEntry ? (catalogEntry.sheetGroup || catalogEntry.item) : ''}</td>
                                <td className="doc-hint">{catalogEntry ? (catalogEntry.block || catalogEntry.item) : ''}</td>
                              </>
                            )}
                            <td className="col-action">
                              <button className="icon-btn" style={{ color: 'var(--ok)' }} title="Confirm this row" onClick={() => confirmStockRow('production', row)}><CheckCircle2 size={16} /></button>
                            </td>
                          </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              {pendingDispatchRows.length > 0 && (
                <div className="panel" style={{ borderColor: 'var(--ledger-red)' }}>
                  <div className="panel-header">
                    <div><h2>Pending Dispatch Bill Review ({pendingDispatchRows.length})</h2><p className="subtitle">New Customer Dispatch Bill entries not yet pushed to any customer's Sheet. Check or fix the suggested customer, then push — nothing here counts toward balances or reaches the real Sheet until you do.</p></div>
                    <button className="btn btn-primary" onClick={() => pushPendingRows('customerDispatch', pendingDispatchRows)}><FileSpreadsheet size={15} /> Push to Sheet</button>
                  </div>
                  {buildPendingGroups(pendingDispatchRows).map(group => (
                    <PendingGroupBar key={group.key} group={group} guess={matchCustomer(group.rows[0])}
                      isKnownGuess={isKnownCustomerGuess(group.rows[0])} knownCustomers={allCustomerTabNames}
                      onBulkChange={updatePendingCustomerBulk('customerDispatch')}
                      onConfirmGroup={() => group.rows.forEach(r => confirmStockRow('customerDispatch', r))} />
                  ))}
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Date</th><th>Invoice No</th><th>Description</th><th>Quantity</th><th>Suggested Customer</th><th>Sheet tab</th><th>Block</th><th className="col-action"></th></tr></thead>
                      <tbody>
                        {pendingDispatchRows.map(row => {
                          const effectiveCustomer = row.confirmedCustomer || (isKnownCustomerGuess(row) ? matchCustomer(row) : '');
                          const catalogEntry = getCatalogEntryForItem(effectiveCustomer, row.description);
                          const needsTabBlock = effectiveCustomer && !catalogEntry;
                          const draft = pendingTabBlockForms[row.id];
                          const tabBlockUnresolved = needsTabBlock && (!((draft && draft.sheetGroup) || '').trim() || !((draft && draft.block) || '').trim());
                          return (
                          <tr key={row.id} className={tabBlockUnresolved ? 'needs-tabblock-row' : ''}>
                            <td style={{ padding: '6px 10px' }}>{row.date}</td>
                            <td style={{ padding: '6px 10px' }}>{row.invoice_no}</td>
                            <td style={{ padding: '6px 10px' }}>{row.description}</td>
                            <td style={{ padding: '6px 10px' }}>{row.quantity || ''}</td>
                            <td>
                              <CustomerSuggestCell value={row.confirmedCustomer} guess={matchCustomer(row)}
                                isKnownGuess={isKnownCustomerGuess(row)} knownCustomers={allCustomerTabNames}
                                onChange={v => updatePendingCustomer('customerDispatch')(row.id, v)} />
                            </td>
                            {needsTabBlock ? (
                              <TabBlockPickerCells rowId={row.id} description={row.description}
                                sheetGroupOptions={Array.from(new Set([
                                  ...productCatalog.filter(c => c.customer.toLowerCase() === effectiveCustomer.toLowerCase()).map(c => c.sheetGroup),
                                  ...getRealTabNamesForCustomer(effectiveCustomer),
                                ])).filter(Boolean)}
                                blockOptions={Array.from(new Set([
                                  ...getRealBlocksForTab(effectiveCustomer, (draft && draft.sheetGroup) || ''),
                                  ...getCatalogBlocksForTab(effectiveCustomer, (draft && draft.sheetGroup) || ''),
                                ])).filter(Boolean)}
                                draft={draft} onChange={(field, value) => updatePendingTabBlockForm(row.id, field, value)} />
                            ) : (
                              <>
                                <td className="doc-hint">{catalogEntry ? (catalogEntry.sheetGroup || catalogEntry.item) : ''}</td>
                                <td className="doc-hint">{catalogEntry ? (catalogEntry.block || catalogEntry.item) : ''}</td>
                              </>
                            )}
                            <td className="col-action">
                              <button className="icon-btn" style={{ color: 'var(--ok)' }} title="Confirm this row" onClick={() => confirmStockRow('customerDispatch', row)}><CheckCircle2 size={16} /></button>
                            </td>
                          </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              {(() => {
                // Restores the per-item ledger table this page originally had for Unassigned entries
                // (it got collapsed to a one-line description list when the real per-customer ledgers
                // moved to the Customer Sheets tab — but Unassigned isn't a real customer with a Sheet,
                // so it never got a home over there, and lost its table in the move). Each item shown
                // here is a confirmed Production/Dispatch row that still doesn't match any Customer
                // Mapping rule, with its full production+dispatch ledger so you can see exactly what's
                // sitting unrouted before writing the keyword rule that will pick it up.
                const unassignedGroups = customerStockGroups.filter(g => g.customer === 'Unassigned');
                if (!unassignedGroups.length) return null;
                return (
                  <>
                  <SectionDivider icon={AlertCircle} label="Unassigned" hint="Confirmed, but no customer at all — nobody's claimed these yet." />
                  <div className="panel" style={{ borderColor: 'var(--ledger-red)' }}>
                    <div className="panel-header">
                      <div>
                        <h2>Unassigned ({unassignedGroups.length})</h2>
                        <p className="subtitle">Confirmed rows that didn't match any rule in Customer Mapping, so they're not attributed to any customer yet. Add a keyword rule for them in Customer Mapping, then they'll route to the right customer automatically.</p>
                      </div>
                    </div>
                    {unassignedGroups.map(g => (
                      <div key={g.id} style={{ marginBottom: 18 }}>
                        <div className="section-label">{g.description} — closing balance: <Pill tone={g.closingBalance >= 0 ? 'ok' : 'warn'}>{g.closingBalance}</Pill></div>
                        <div className="table-wrap">
                          <table>
                            <thead><tr><th>Date</th><th>Opening</th><th>Production</th><th>Dispatch</th><th>Closing</th><th className="col-action"></th></tr></thead>
                            <tbody>
                              {g.ledger.map((e, i) => (
                                <tr key={i}>
                                  <td style={{ padding: '6px 10px' }}>{e.date}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.opening}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.pieces || ''}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.dispatch || ''}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.closing}</td>
                                  <td className="col-action">
                                    <div style={{ display: 'flex', gap: 4, justifyContent: 'flex-end' }}>
                                      {(e.productionIds || []).map(id => (
                                        <button key={`p-${id}`} className="icon-btn danger" title="Delete this date's Production Register row" onClick={() => deleteRow('production')(id)}><Trash2 size={14} /></button>
                                      ))}
                                      {(e.dispatchIds || []).map(id => (
                                        <button key={`d-${id}`} className="icon-btn danger" title="Delete this date's Customer Dispatch Bill row" onClick={() => deleteRow('customerDispatch')(id)}><Trash2 size={14} /></button>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                        {(() => {
                          const form = assignForms[g.id] || {};
                          const chosenCustomer = (form.customer || '').trim();
                          const customerSheetGroups = Array.from(new Set([
                            ...productCatalog.filter(c => c.customer.toLowerCase() === chosenCustomer.toLowerCase()).map(c => c.sheetGroup),
                            ...getRealTabNamesForCustomer(chosenCustomer),
                          ])).filter(Boolean);
                          const chosenSheetGroup = (form.sheetGroup || '').trim();
                          // Both what physically already exists in the real Sheet tab, AND whatever
                          // block name other Product Catalog entries under this same tab already use
                          // (which may not be pushed to the real Sheet yet — see getCatalogBlocksForTab).
                          const existingBlocks = Array.from(new Set([
                            ...getRealBlocksForTab(chosenCustomer, chosenSheetGroup),
                            ...getCatalogBlocksForTab(chosenCustomer, chosenSheetGroup),
                          ])).filter(Boolean);
                          const canAssign = chosenCustomer && chosenSheetGroup && (form.item || g.description || '').trim();
                          return (
                            <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 8, flexWrap: 'wrap' }}>
                              <input
                                list={`assign-customer-${g.id}`}
                                placeholder="Customer"
                                value={form.customer || ''}
                                onChange={e => updateAssignForm(g.id, 'customer', e.target.value)}
                                style={{ width: 150 }}
                              />
                              <datalist id={`assign-customer-${g.id}`}>
                                {allCustomerTabNames.map(c => <option key={c} value={c} />)}
                              </datalist>
                              <input
                                list={`assign-sheetgroup-${g.id}`}
                                placeholder="Sheet tab"
                                value={form.sheetGroup || ''}
                                onChange={e => updateAssignForm(g.id, 'sheetGroup', e.target.value)}
                                style={{ width: 150 }}
                              />
                              <datalist id={`assign-sheetgroup-${g.id}`}>
                                {customerSheetGroups.map(s => <option key={s} value={s} />)}
                              </datalist>
                              <input
                                list={`assign-item-${g.id}`}
                                placeholder="Block (blank = new block)"
                                value={form.item || ''}
                                onChange={e => updateAssignForm(g.id, 'item', e.target.value)}
                                style={{ width: 180 }}
                              />
                              <datalist id={`assign-item-${g.id}`}>
                                {existingBlocks.map(i => <option key={i} value={i} />)}
                                <option value={g.description}>{`+ New block: "${g.description}"`}</option>
                              </datalist>
                              <button className="btn btn-primary" disabled={!canAssign} onClick={() => assignUnassignedGroup(g.id, g.description)}>
                                <Check size={15} /> Assign
                              </button>
                            </div>
                          );
                        })()}
                      </div>
                    ))}
                  </div>
                  </>
                );
              })()}
            </div>
          )}
          {loaded && activeTab === 'customerMapping' && (
            <div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Merge Duplicate Customer</h2><p className="subtitle">If a name here turns out to be a duplicate or misspelling of another customer (e.g. "Bindal technopolymer pvt. ltd." vs "BINDAL STOCK 1.08.26"), merge everything under it into the real one — moves all confirmed Production/Dispatch rows, Product Catalog entries, Mapping rules, and the Sheet ID over, then remembers the old name as an alias so it routes correctly going forward. Nothing in the underlying registers is ever deleted, only who each row is attributed to.</p></div>
                </div>
                <div className="field-row">
                  <select className="cell-input" style={{ width: 220 }} value={mergeFromCustomer} onChange={e => setMergeFromCustomer(e.target.value)}>
                    <option value="">Merge this customer…</option>
                    {allCustomerTabNames.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <span className="doc-hint">into</span>
                  <select className="cell-input" style={{ width: 220 }} value={mergeToCustomer} onChange={e => setMergeToCustomer(e.target.value)}>
                    <option value="">this real customer…</option>
                    {allCustomerTabNames.filter(c => c !== mergeFromCustomer).map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <button
                    className="btn btn-primary"
                    disabled={!mergeFromCustomer || !mergeToCustomer}
                    onClick={() => { mergeCustomerInto(mergeFromCustomer, mergeToCustomer); setMergeFromCustomer(''); setMergeToCustomer(''); }}
                  >
                    <RefreshCw size={15} /> Merge
                  </button>
                </div>
              </div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Customer Mapping</h2><p className="subtitle">Which keyword in a production-ledger item name routes to which customer. Checked top to bottom — the first match wins, so more specific keywords should sit above general ones. A bracketed customer name found right next to an item always overrides this list.</p></div>
                  <button className="btn btn-primary" onClick={addMappingRow}><Plus size={15} /> Add rule</button>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead><tr><th>Keyword (matches anywhere in item name, case-insensitive)</th><th>Customer</th><th className="col-action"></th></tr></thead>
                    <tbody>
                      {customerMapping.map(rule => (
                        <tr key={rule.id}>
                          <td><input className="cell-input" value={rule.keyword} onChange={e => updateMappingRow(rule.id, 'keyword', e.target.value)} /></td>
                          <td><input className="cell-input" value={rule.customer} onChange={e => updateMappingRow(rule.id, 'customer', e.target.value)} /></td>
                          <td className="col-action"><button className="icon-btn danger" onClick={() => deleteMappingRow(rule.id)}><Trash2 size={15} /></button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <div className="panel">
                <h2 style={{ marginBottom: 6 }}>Known Product Catalog</h2>
                <p className="subtitle" style={{ marginBottom: 14 }}>Exact item names used to correct handwriting misreads during extraction (e.g. "g" vs "9", "&" vs "8"). "Sheet Tab" is which tab this item lands under in that customer's own Google Sheet (Customer Sheets tab) — several variants of one base item (e.g. "IT 500 Lid" and "IT 500 Container") should share the same Sheet Tab name, exactly like your original files. Built from whatever you've imported via the Customer Sheets tab — edit or delete anything that's wrong.</p>
                {customerNames.length === 0 && !productCatalog.length && <div className="empty-state">No catalog items yet.</div>}
                {Array.from(new Set(productCatalog.map(c => c.customer))).map(customer => (
                  <div key={customer} style={{ marginBottom: 16 }}>
                    <div className="section-label">{customer} ({productCatalog.filter(c => c.customer === customer).length} items)</div>
                    <div className="table-wrap">
                      <table>
                        <thead><tr><th>Item name</th><th>Sheet Tab (item group)</th><th>Block (blank = same as item name)</th><th className="col-action"></th></tr></thead>
                        <tbody>
                          {productCatalog.filter(c => c.customer === customer).map(c => (
                            <tr key={c.id}>
                              <td><input className="cell-input" value={c.item} onChange={e => updateCatalogItem(c.id, 'item', e.target.value)} /></td>
                              <td><input className="cell-input" value={c.sheetGroup || ''} placeholder={c.item} onChange={e => updateCatalogItem(c.id, 'sheetGroup', e.target.value)} /></td>
                              <td><input className="cell-input" value={c.block || ''} placeholder={c.item} onChange={e => updateCatalogItem(c.id, 'block', e.target.value)} /></td>
                              <td className="col-action"><button className="icon-btn danger" onClick={() => deleteCatalogItem(c.id)}><Trash2 size={15} /></button></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {loaded && activeTab === 'aliases' && (
            <div>
              <div className="panel">
                <div className="panel-header">
                  <div>
                    <h2>Word Abbreviations</h2>
                    <p className="subtitle">Shorthand this factory writes on real pages — "J" for "Jumbo", "CONT" for "Container", and so on. Teach the app once here and it applies everywhere from then on: extraction writes the item name out in full going forward, and matching (so an item routes to the right Sheet tab / customer stock group) treats the short and long forms as the same thing, even for entries already extracted before you added the rule. No code changes needed for a new abbreviation — just add it below.</p>
                  </div>
                  <button className="btn btn-primary" onClick={addAbbreviationRow}><Plus size={15} /> Add abbreviation</button>
                </div>
                {!abbreviations.length && <div className="empty-state">None taught yet — add one above, e.g. short "J", full "Jumbo".</div>}
                {!!abbreviations.length && (
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Short form (as written)</th><th>Full word</th><th className="col-action"></th></tr></thead>
                      <tbody>
                        {abbreviations.map(a => (
                          <tr key={a.id}>
                            <td><input className="cell-input" placeholder="J" value={a.short} onChange={e => updateAbbreviationRow(a.id, 'short', e.target.value)} /></td>
                            <td><input className="cell-input" placeholder="Jumbo" value={a.long} onChange={e => updateAbbreviationRow(a.id, 'long', e.target.value)} /></td>
                            <td className="col-action"><button className="icon-btn danger" onClick={() => deleteAbbreviationRow(a.id)}><Trash2 size={15} /></button></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
              <div className="panel">
                <h2 style={{ marginBottom: 6 }}>Aliases</h2>
                <p className="subtitle" style={{ marginBottom: 14 }}>An alias is an alternate spelling/wording for an item that should route to the same customer and the same Sheet tab as the "real" name — e.g. a dispatch bill written as "HANDLE LOCK" that should really land under Diamond's "Tijori Handle" block. Adding one here does two things together, kept in sync: it adds a Product Catalog entry (so it lands in the right Sheet tab) and an exact-match Customer Mapping keyword rule (so it routes to the right customer). Editing or deleting an alias updates both. This is the same data shown under Customer Mapping → Known Product Catalog — just a dedicated, add-capable view of it.</p>
                <div className="field-row">
                  <input className="text-input" list="alias-customers" placeholder="Customer" style={{ width: 160 }} value={newAliasForm.customer} onChange={e => updateNewAliasForm('customer', e.target.value)} />
                  <datalist id="alias-customers">{allCustomerTabNames.map(c => <option value={c} key={c} />)}</datalist>
                  <input className="text-input" placeholder="Alias / item name (as it's actually written)" style={{ minWidth: 240, flex: 1 }} value={newAliasForm.item} onChange={e => updateNewAliasForm('item', e.target.value)} />
                  <input
                    className="text-input"
                    list={`alias-sheetgroups-${(newAliasForm.customer || '').trim().toLowerCase()}`}
                    placeholder="Sheet tab it should land under"
                    style={{ width: 200 }}
                    value={newAliasForm.sheetGroup}
                    onChange={e => updateNewAliasForm('sheetGroup', e.target.value)}
                  />
                  <datalist id={`alias-sheetgroups-${(newAliasForm.customer || '').trim().toLowerCase()}`}>
                    {Array.from(new Set(productCatalog.filter(c => c.customer.toLowerCase() === (newAliasForm.customer || '').trim().toLowerCase()).map(c => c.sheetGroup))).filter(Boolean).map(s => <option value={s} key={s} />)}
                  </datalist>
                  <input
                    className="text-input"
                    list={`alias-blocks-${(newAliasForm.customer || '').trim().toLowerCase()}`}
                    placeholder="Block it should land under (blank = same as item name)"
                    style={{ width: 220 }}
                    value={newAliasForm.block}
                    onChange={e => updateNewAliasForm('block', e.target.value)}
                  />
                  <datalist id={`alias-blocks-${(newAliasForm.customer || '').trim().toLowerCase()}`}>
                    {Array.from(new Set(productCatalog.filter(c => c.customer.toLowerCase() === (newAliasForm.customer || '').trim().toLowerCase()).map(c => c.block || c.item))).filter(Boolean).map(b => <option value={b} key={b} />)}
                  </datalist>
                  <button className="btn btn-primary" disabled={!newAliasForm.customer.trim() || !newAliasForm.item.trim() || !newAliasForm.sheetGroup.trim()} onClick={addAliasFromForm}>
                    <Plus size={15} /> Add alias
                  </button>
                </div>
              </div>
              <div className="panel">
                {customerNames.length === 0 && !productCatalog.length && <div className="empty-state">No aliases yet — add one above.</div>}
                {Array.from(new Set(productCatalog.map(c => c.customer))).map(customer => (
                  <div key={customer} style={{ marginBottom: 16 }}>
                    <div className="section-label">{customer} ({productCatalog.filter(c => c.customer === customer).length})</div>
                    <div className="table-wrap">
                      <table>
                        <thead><tr><th>Alias / item name</th><th>Sheet tab</th><th>Block (blank = same as item name)</th><th className="col-action"></th></tr></thead>
                        <tbody>
                          {productCatalog.filter(c => c.customer === customer).map(c => (
                            <tr key={c.id}>
                              <td><input className="cell-input" value={c.item} onChange={e => updateCatalogItem(c.id, 'item', e.target.value)} /></td>
                              <td><input className="cell-input" value={c.sheetGroup || ''} placeholder={c.item} onChange={e => updateCatalogItem(c.id, 'sheetGroup', e.target.value)} /></td>
                              <td><input className="cell-input" value={c.block || ''} placeholder={c.item} onChange={e => updateCatalogItem(c.id, 'block', e.target.value)} /></td>
                              <td className="col-action"><button className="icon-btn danger" onClick={() => deleteCatalogItem(c.id)}><Trash2 size={15} /></button></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
          {loaded && activeTab === 'customerSheets' && (
            <div>
              <div className="panel">
                <h2 style={{ marginBottom: 6 }}>Customer Sheets</h2>
                <p className="subtitle" style={{ marginBottom: 4 }}>Add and sync a customer's Sheet ID here — their own separate Google Sheet, not a tab on this app's main sheet — laid out the same way as your original BINDAL STOCK.xlsx / DIAMOND.xlsx / anmol stock files: one tab per base item, every variant of that item as its own side-by-side table within that tab.</p>
                <p className="subtitle" style={{ marginBottom: 4 }}>Reviewing and pushing a customer's ledger out to their Sheet now happens on the Customer Stock tab, right under their per-item preview — nothing pushes from here. Pushing only ever appends: it never touches or duplicates a row that's already in the sheet, whether it got there from a previous push or someone typed it in by hand, and new rows use the same Opening/Closing formula convention already in your sheet instead of dumping in flat numbers. The "summary" tab is never touched at all — its existing formulas already point at each item's latest row, so they keep calculating themselves.</p>
                <p className="subtitle">Setup per customer, once: create or open their Google Sheet, share it with the service account email below as an Editor, then paste that sheet's link below — you can paste the whole browser address bar URL, no need to dig out just the ID. Skipping the share step is the #1 cause of "the caller does not have permission" errors.</p>
                <div className="field-row" style={{ marginTop: 10 }}>
                  <span style={{ fontSize: 13, fontWeight: 600 }}>Share every customer Sheet with:</span>
                  {serviceAccountEmail
                    ? <code style={{ background: 'var(--accent-soft)', padding: '4px 8px', borderRadius: 4, fontSize: 12.5 }}>{serviceAccountEmail}</code>
                    : <span className="doc-hint">(couldn't read it — check GOOGLE_SERVICE_ACCOUNT_JSON on the server)</span>}
                  {serviceAccountEmail && (
                    <button className="btn btn-ghost" onClick={() => navigator.clipboard?.writeText(serviceAccountEmail)}><FileSpreadsheet size={14} /> Copy</button>
                  )}
                </div>
              </div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Add a Customer Sheet</h2><p className="subtitle">Works for any customer — paste any Google Sheet ID (shared with the service account first) and it's read back to fill in the item catalog and mapping, same as before, without needing to know the customer's name ahead of time. Also tracks which Sheet IDs you've already imported.</p></div>
                </div>
                <div className="field-row">
                  <input className="text-input" style={{ minWidth: 340, flex: 1 }} placeholder="Paste the Google Sheet link or ID" value={newSheetId} onChange={e => setNewSheetId(extractSheetIdFromInput(e.target.value))} autoComplete="off" spellCheck={false} />
                  <button className="btn btn-primary" onClick={() => importSheetById(newSheetId, { mode: 'create' })} disabled={sheetImportBusy || !newSheetId.trim()}>
                    {sheetImportBusy ? <Loader2 size={15} className="spin" /> : <Link2 size={15} />} Fetch &amp; Import
                  </button>
                </div>
                {sheetImportError && <div className="error-box" style={{ marginTop: 10 }}><AlertCircle size={16} /><span>{sheetImportError}</span></div>}
                {importResultMessage && !sheetReview && <div className="doc-hint" style={{ color: 'var(--ok)', marginTop: 10 }}>✓ {importResultMessage}</div>}
                {sheetReview && (
                  <div className="review-box">
                    <p className="subtitle" style={{ marginBottom: 6 }}>
                      {sheetReview.mode === 'resync' ? `Re-syncing ${sheetReview.customerName} — this is EVERY item currently in the Sheet. Confirming replaces this customer's whole catalog with exactly this list, so anything removed or renamed there drops out too.`
                        : sheetReview.mode === 'generate' ? `Structure built in the Sheet for "${sheetReview.customerName}" — every tab, item block, and the summary tab are already written and live.`
                        : `Imported from "${sheetReview.spreadsheetTitle || sheetReview.spreadsheetId}".`}
                    </p>
                    <div className="field-row">
                      <span style={{ fontSize: 13, fontWeight: 600 }}>Customer name:</span>
                      <input className="text-input" value={sheetReview.customerName} disabled={sheetReview.mode === 'resync'} onChange={e => setSheetReview(prev => ({ ...prev, customerName: e.target.value }))} />
                    </div>
                    <p className="subtitle" style={{ marginBottom: 10 }}>
                      {sheetReview.mode === 'resync'
                        ? `${sheetReview.items.length} item${sheetReview.items.length === 1 ? '' : 's'} found in the Sheet${sheetReview.skippedExisting ? ` (${sheetReview.skippedExisting} already known)` : ''}. Uncheck anything that isn't really a product, fix any typos, adjust which Sheet Tab each lands under, then confirm the replace.`
                        : `${sheetReview.items.length} new item${sheetReview.items.length === 1 ? '' : 's'} found${sheetReview.skippedExisting ? ` (${sheetReview.skippedExisting} already known, not shown)` : ''}. Uncheck anything that isn't really a product, fix any typos, adjust which Sheet Tab each lands under, then confirm.`}
                    </p>
                    <div className="table-wrap">
                      <table>
                        <thead><tr><th style={{ width: 34 }}></th><th>Item name</th><th>Sheet Tab (item group)</th></tr></thead>
                        <tbody>
                          {sheetReview.items.map(it => (
                            <tr key={it.id}>
                              <td style={{ textAlign: 'center' }}><input type="checkbox" checked={it.include} onChange={e => updateSheetReviewItem(it.id, 'include', e.target.checked)} /></td>
                              <td><input className="cell-input" value={it.item} onChange={e => updateSheetReviewItem(it.id, 'item', e.target.value)} /></td>
                              <td><input className="cell-input" value={it.sheetGroup} onChange={e => updateSheetReviewItem(it.id, 'sheetGroup', e.target.value)} /></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="review-actions">
                      <button className="btn btn-primary" onClick={confirmSheetImport} disabled={!sheetReview.customerName.trim()}><CheckCircle2 size={15} /> Add new items to catalog & mapping</button>
                      <button className="btn btn-ghost" onClick={() => setSheetReview(null)}><XCircle size={15} /> Discard</button>
                    </div>
                  </div>
                )}
              </div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Generate a Brand-New Customer Sheet</h2><p className="subtitle">For a customer with no Sheet yet: create one blank Google Sheet, share it with the service account email above as an Editor, paste its link below, then list the customer's product categories and the items under each. Every tab, item block (with live Opening/Closing formulas), and a summary tab get written into that Sheet for you — same structure and formulas as every other customer here, just built from scratch instead of pushed to over time.</p></div>
                </div>
                <div className="field-row">
                  <input className="text-input" style={{ minWidth: 220 }} placeholder="Customer name" value={genCustomerName} onChange={e => setGenCustomerName(e.target.value)} autoComplete="off" spellCheck={false} />
                  <input className="text-input" style={{ minWidth: 300, flex: 1 }} placeholder="Paste the blank Google Sheet's link or ID" value={genSheetId} onChange={e => setGenSheetId(e.target.value)} autoComplete="off" spellCheck={false} />
                </div>
                <div style={{ marginTop: 10 }}>
                  {genCategories.map((cat, idx) => (
                    <div key={cat.id} style={{ display: 'flex', gap: 8, marginBottom: 8, alignItems: 'flex-start' }}>
                      <input className="text-input" style={{ minWidth: 180 }} placeholder="Category name (becomes a tab)" value={cat.name} onChange={e => updateGenCategory(cat.id, 'name', e.target.value)} autoComplete="off" spellCheck={false} />
                      <textarea className="text-input" style={{ flex: 1, minHeight: 60, resize: 'vertical', fontFamily: 'inherit' }} placeholder={'One item per line, e.g.\nIT 500 LID\nIT 500 CONTAINER'} value={cat.itemsText} onChange={e => updateGenCategory(cat.id, 'itemsText', e.target.value)} spellCheck={false} />
                      {genCategories.length > 1 && (
                        <button className="icon-btn danger" onClick={() => removeGenCategory(cat.id)} title="Remove category"><Trash2 size={15} /></button>
                      )}
                    </div>
                  ))}
                  <button className="btn btn-ghost" onClick={addGenCategory}><Plus size={15} /> Add category</button>
                </div>
                <div className="review-actions" style={{ marginTop: 10 }}>
                  <button className="btn btn-primary" onClick={buildGeneratedStructure} disabled={genBusy}>
                    {genBusy ? <Loader2 size={15} className="spin" /> : <FileSpreadsheet size={15} />} Build structure
                  </button>
                </div>
                {genError && <div className="error-box" style={{ marginTop: 10 }}><AlertCircle size={16} /><span>{genError}</span></div>}
              </div>
              {(() => {
                const unassignedGroups = customerStockGroups.filter(g => g.customer === 'Unassigned');
                if (!unassignedGroups.length) return null;
                return (
                  <div className="info-box">
                    <Info size={16} />
                    <span>{unassignedGroups.length} item{unassignedGroups.length === 1 ? '' : 's'} didn't match any customer at all, so they're not shown here: {unassignedGroups.map(g => g.description).join(', ')}. Fix this in Customer Mapping (add a keyword rule), then it'll route to the right customer automatically — nothing to push here for "Unassigned" itself.</span>
                  </div>
                );
              })()}
              {allCustomerTabNames.length === 0 && <div className="panel"><div className="empty-state">No customers yet — add one above.</div></div>}
              {allCustomerTabNames.map(customer => {
                const { itemGroups } = buildCustomerSheetPayload(customer);
                const variantCount = itemGroups.reduce((s, g) => s + (g.variants || []).length, 0);
                const sheetId = getCustomerSheetId(customer);
                const registryEntry = customerSheetIds.find(c => c.customer === customer);
                return (
                  <div className="panel" key={customer}>
                    <div className="panel-header">
                      <div>
                        <h2>{customer}</h2>
                        <p className="subtitle">{itemGroups.length} item tab{itemGroups.length === 1 ? '' : 's'} · {variantCount} variant{variantCount === 1 ? '' : 's'} in the catalog
                          {registryEntry?.lastImportedAt && <> · last imported {new Date(registryEntry.lastImportedAt).toLocaleString()}</>}
                          {registryEntry?.lastPushedAt && <> · last pushed {new Date(registryEntry.lastPushedAt).toLocaleString()}</>}
                        </p>
                      </div>
                      <button className="btn btn-ghost" onClick={() => importSheetById(sheetId, { mode: 'resync', resyncCustomer: customer })} disabled={sheetImportBusy || !sheetId.trim()} title="Re-read this customer's Sheet fresh and replace their catalog with exactly what's in it now (removes anything renamed or deleted there) — you'll see a review + confirm before anything changes">
                        <RefreshCw size={15} /> Re-sync
                      </button>
                    </div>
                    <div className="field-row">
                      <input className="text-input" style={{ minWidth: 340, flex: 1 }} placeholder="Paste this customer's Google Sheet link or ID" value={sheetId} onChange={e => updateCustomerSheetId(customer, e.target.value)} autoComplete="off" spellCheck={false} />
                      {registryEntry && (
                        <button className="btn btn-ghost" onClick={() => { if (window.confirm(`Stop tracking ${customer}'s Sheet ID? This only removes the tracking record — nothing is deleted from the Known Product Catalog or Customer Mapping.`)) removeCustomerSheetEntry(customer); }} title="Stop tracking this Sheet ID (catalog items stay)">
                          <Trash2 size={15} /> Remove tracking
                        </button>
                      )}
                    </div>
                    {!sheetId.trim() && <div className="doc-hint" style={{ marginTop: 10 }}>Add a Sheet ID above to enable syncing and pushing (pushing itself now happens on the Customer Stock tab).</div>}
                  </div>
                );
              })}
            </div>
          )}
          {loaded && activeTab === 'settings' && (
            <div>
              <div className="panel">
                <h2 style={{ marginBottom: 6 }}>Clear App Data</h2>
                <p className="subtitle" style={{ marginBottom: 14 }}>Pick exactly what to wipe, then confirm — nothing happens until you click "Clear selected," and there's a confirmation prompt listing exactly what you picked before anything is actually deleted. This only clears data inside this app's own registers; it never touches any customer's real Google Sheet.</p>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10, paddingBottom: 10, borderBottom: '1px solid var(--border)', cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={CLEAR_GROUPS.every(g => !!clearSelected[g.key])}
                    ref={el => { if (el) el.indeterminate = CLEAR_GROUPS.some(g => !!clearSelected[g.key]) && !CLEAR_GROUPS.every(g => !!clearSelected[g.key]); }}
                    onChange={e => setClearSelected(CLEAR_GROUPS.every(g => !!clearSelected[g.key]) ? {} : Object.fromEntries(CLEAR_GROUPS.map(g => [g.key, true])))}
                  />
                  <span style={{ fontSize: 13.5, fontWeight: 600 }}>Select all</span>
                </label>
                {CLEAR_GROUPS.map(g => (
                  <label key={g.key} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, cursor: 'pointer' }}>
                    <input type="checkbox" checked={!!clearSelected[g.key]} onChange={() => toggleClearGroup(g.key)} />
                    <span style={{ fontSize: 13.5 }}>{g.label}</span>
                  </label>
                ))}
                <div className="review-actions" style={{ marginTop: 10 }}>
                  <button className="btn btn-danger" disabled={clearBusy || !Object.values(clearSelected).some(Boolean)} onClick={runClearSelected}>
                    {clearBusy ? <Loader2 size={15} className="spin" /> : <Trash2 size={15} />} Clear selected
                  </button>
                </div>
                {clearMessage && <div className="doc-hint" style={{ marginTop: 10 }}>{clearMessage}</div>}
              </div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Date Formatting</h2><p className="subtitle">Every date should read D.M.YY (e.g. "16.7.26") throughout — a bug let some rows through with a leading zero and/or a 4-digit year instead (e.g. "23.07.2026"), which is now fixed for anything freshly extracted. This is a one-time sweep for rows that entered the ledger before that fix, or were typed/edited by hand — it rewrites the date field(s) in place across every register, nothing else. Safe to run more than once; it only touches rows that still need it.</p></div>
                  <button className="btn btn-primary" disabled={dateFixBusy} onClick={normalizeAllDates}>
                    {dateFixBusy ? <Loader2 size={15} className="spin" /> : <RefreshCw size={15} />} Fix date formatting
                  </button>
                </div>
                {dateFixMessage && <div className="doc-hint">{dateFixMessage}</div>}
              </div>
              <div className="panel">
                <div className="panel-header">
                  <div><h2>Shyam Adarsh Sheet — Tabs</h2><p className="subtitle">Scans the main Google Sheet and shows every tab it finds. "Internal" tabs are bookkeeping the app needs but that never shows up as a table anywhere in the app (training corrections, the customer-sheets search mirror) — checked by default, safe to delete. "Unrecognized" tabs aren't anything this app manages — look at the name before deleting one. This never touches any customer's own external Sheet.</p></div>
                  <button className="btn btn-ghost" disabled={sheetTabsBusy} onClick={scanSheetTabs}>
                    {sheetTabsBusy ? <Loader2 size={15} className="spin" /> : <Search size={15} />} Scan tabs
                  </button>
                </div>
                {sheetTabs === null && <div className="empty-state">Click "Scan tabs" to see what's currently in the Shyam Adarsh sheet.</div>}
                {sheetTabs !== null && !sheetTabs.length && <div className="empty-state">No tabs found.</div>}
                {sheetTabs !== null && !!sheetTabs.length && (
                  <>
                    {sheetTabs.map(t => (
                      <label key={t.title} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, cursor: 'pointer' }}>
                        <input type="checkbox" checked={!!tabsToDelete[t.title]} onChange={() => toggleTabToDelete(t.title)} />
                        <span style={{ fontSize: 13.5 }}>{t.title}</span>
                        <span className={`pill ${t.kind === 'internal' ? 'pill-warn' : t.kind === 'unrecognized' ? 'pill-neutral' : 'pill-ok'}`}>{t.kind}</span>
                      </label>
                    ))}
                    <div className="review-actions" style={{ marginTop: 10 }}>
                      <button className="btn btn-danger" disabled={sheetTabsBusy || !Object.values(tabsToDelete).some(Boolean)} onClick={deleteSelectedTabs}>
                        {sheetTabsBusy ? <Loader2 size={15} className="spin" /> : <Trash2 size={15} />} Delete selected tabs
                      </button>
                    </div>
                  </>
                )}
                {sheetTabsMessage && <div className="doc-hint" style={{ marginTop: 10 }}>{sheetTabsMessage}</div>}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
    {copyModal && <CopyExportModal data={copyModal} onClose={() => setCopyModal(null)} />}
    {zoomedImage && <ImageZoomModal src={zoomedImage} onClose={() => setZoomedImage(null)} />}
    </>
  );
}
/* ============================== login gate ============================== */
// The whole app sits behind one shared team password (see server/lib/auth.js). This wrapper checks
// /api/session on load, shows a plain login form if there's no valid session cookie yet, and renders
// the real app once logged in. It also listens for the 'fims-unauthorized' event — dispatched by
// callClaudeExtract/window.storage above whenever the backend returns a 401 (e.g. the session cookie
// expired after 30 days, or the server restarted with a new SESSION_SECRET) — so an expired session
// mid-use drops back to the login screen instead of the app silently failing every request.
const LOGIN_PAGE_STYLE = {
  minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center',
  fontFamily: "'IBM Plex Sans', 'Segoe UI', sans-serif", background: '#ece8df', color: '#23262b',
};
function LoginScreen({ onLoggedIn }) {
  const [password, setPassword] = useState('');
  const [loggingIn, setLoggingIn] = useState(false);
  const [error, setError] = useState('');
  const submit = async (e) => {
    e.preventDefault();
    if (!password) return;
    setLoggingIn(true); setError('');
    try {
      const res = await fetch('/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ password }),
      });
      if (res.ok) { onLoggedIn(); return; }
      const data = await res.json().catch(() => ({}));
      setError(data.error || `Login failed (HTTP ${res.status}).`);
    } catch (e) {
      setError(e.message || 'Could not reach the server — check your connection and try again.');
    } finally {
      setLoggingIn(false);
    }
  };
  return (
    <div style={LOGIN_PAGE_STYLE}>
      <form onSubmit={submit} style={{ background: '#f6f3ec', border: '1px solid #cfc9ba', borderRadius: 8, padding: 28, width: 320, boxShadow: '0 2px 10px rgba(0,0,0,0.08)' }}>
        <h1 style={{ fontSize: 18, marginBottom: 4 }}>FIMS</h1>
        <p style={{ fontSize: 12.5, color: '#5b5e63', marginBottom: 16 }}>Enter the shared team password to continue.</p>
        <input
          type="password" value={password} onChange={e => setPassword(e.target.value)}
          placeholder="Password" autoFocus autoComplete="current-password"
          style={{ width: '100%', padding: '9px 10px', borderRadius: 6, border: '1px solid #cfc9ba', fontSize: 14, marginBottom: 12, boxSizing: 'border-box' }}
        />
        <button type="submit" disabled={loggingIn || !password} style={{ width: '100%', padding: '9px 10px', borderRadius: 6, border: 'none', background: '#a97a2f', color: '#fff', fontSize: 14, fontWeight: 600, cursor: loggingIn ? 'default' : 'pointer', opacity: loggingIn ? 0.7 : 1 }}>
          {loggingIn ? 'Checking…' : 'Log in'}
        </button>
        {error && <p style={{ fontSize: 12.5, color: '#a23b2e', marginTop: 12 }}>{error}</p>}
      </form>
    </div>
  );
}
function LogoutButton({ onLoggedOut }) {
  const doLogout = async () => {
    try { await fetch('/api/logout', { method: 'POST', credentials: 'include' }); } catch (e) { /* noop */ }
    onLoggedOut();
  };
  return (
    <button
      onClick={doLogout}
      title="Log out"
      style={{ position: 'fixed', top: 10, right: 12, zIndex: 1001, padding: '5px 10px', borderRadius: 6, border: '1px solid #cfc9ba', background: '#f6f3ec', color: '#5b5e63', fontSize: 11.5, cursor: 'pointer' }}
    >
      Log out
    </button>
  );
}
export default function App() {
  const [authed, setAuthed] = useState(null); // null = still checking, then true/false
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch('/api/session', { credentials: 'include' });
        const data = await res.json();
        setAuthed(!!data.authed);
      } catch (e) {
        setAuthed(false);
      }
    })();
    const onUnauthorized = () => setAuthed(false);
    window.addEventListener('fims-unauthorized', onUnauthorized);
    return () => window.removeEventListener('fims-unauthorized', onUnauthorized);
  }, []);
  if (authed === null) {
    return <div style={LOGIN_PAGE_STYLE}>Loading…</div>;
  }
  if (!authed) {
    return <LoginScreen onLoggedIn={() => setAuthed(true)} />;
  }
  return (
    <>
      <LogoutButton onLoggedOut={() => setAuthed(false)} />
      <FIMSApp />
    </>
  );
}
