import React, { useState, useEffect, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload, Image as ImageIcon, Package, Boxes, Search, Truck, ClipboardList,
  FileSpreadsheet, Download, CheckCircle2, XCircle, Trash2, Loader2,
  AlertCircle, LayoutDashboard, FileText, Archive, ListChecks, Plus, RefreshCw, Link2, Info, Check
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
// Lower-priority fallback keywords — only used when a ledger entry doesn't exactly match one of the
// catalog's known item names above (e.g. a pack-size variant that isn't in the catalog yet). Starts
// empty along with the catalog above; importing a customer sheet adds both an exact-name rule per
// item and a broader fallback rule automatically (see confirmSheetImport).
const FALLBACK_CUSTOMER_MAPPING = [];
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
function buildPromptWithTraining(basePrompt, examples) {
  if (!examples || !examples.length) return basePrompt;
  const recent = examples.slice(-MAX_EXAMPLES_IN_PROMPT);
  const examplesText = recent.map((ex, i) =>
    `Example ${i + 1}:\nExtracted (this had a mistake): ${JSON.stringify(ex.before)}\nHuman-corrected (this is right): ${JSON.stringify(ex.after)}`
  ).join('\n\n');
  return `${basePrompt}
LEARNED CORRECTIONS — a person has corrected real extraction mistakes on this exact document type before. Each pair below shows a row as it was first extracted (with a mistake) and how a human corrected it. Study the pattern behind each correction — what kind of value was misread, which field it belongs in, how it should be formatted — and apply that same fix logic to this new document. Do not copy these exact values into unrelated rows; only apply the underlying pattern.
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
  // Already dot-separated ("5.1.24", "13.08.24") — leave as-is, that's the target format.
  if (/^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(s)) return s;
  // Slash-separated DD/MM/YYYY or D/M/YY (what the Production Register's OCR tends to produce).
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
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
      return raw.items.map(it => ({ id: genId(), date: normalizeDateToDots(raw.date || ''), mill: raw.mill || '', reel_no: it.reel_no || '', size: it.size || '', unit: it.unit || '', gsm: it.gsm || '', bf: it.bf || '', shade: it.shade || '', weight_kg: num(it.weight_kg) }));
    },
  },
  {
    key: 'consumption_sheet',
    label: 'Daily Consumption Report (handwritten)',
    hint: 'The handwritten daily sheet workers use to record raw material consumed — columns are usually Shade, GSM, Size, Weight, and Balance Left.',
    register: 'consumption',
    systemPrompt: `You read a handwritten daily raw-material consumption report from a corrugated box factory. The sheet has Hindi column headers. Most rows share one date written once at the top. Return ONLY one JSON object:
{"date":"as written at the top, DD/MM/YYYY","items":[{"shade":"shade code — see rules below","gsm":"the ग्रा / GMS column value","size":"the साइज़ / Size column value","weight_consumed":"the वजन / Weight column value, as a number","balance_left":"the टुकड़ा / Tukda column value (this is the running balance left, NOT a piece count) as a number, or empty string if that row has none","date_override":"only include this if a specific row has a different date than the header, else omit"}]}
IMPORTANT:
- The column that looks like it's labeled "S/K" is actually the SHADE column, not a party name or code. Decode its handwritten values: "S.K" or "SK" means shade NS (Sada Kraft / natural shade). "G.Y" or "GY" means shade GY (Golden Yellow). If you see a different value you don't recognize, copy it as written rather than forcing it into NS or GY.
- There is no BF field on this document — do not invent one. What might look like a stray extra column is the Size column.
- Do not skip the last column (टुकड़ा / Tukda) — it is the balance left, and must be captured for every row that has a value in it.
- Interpret unclear handwriting as best you can; if a value is genuinely illegible leave it as an empty string rather than guessing wildly.`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      return raw.items.map(it => ({ id: genId(), date: normalizeDateToDots(it.date_override || raw.date || ''), shade: it.shade || '', size: it.size || '', gsm: it.gsm || '', weight_consumed: num(it.weight_consumed), balance_left: it.balance_left === '' || it.balance_left == null ? '' : num(it.balance_left) }));
    },
  },
  {
    key: 'production_sheet',
    label: 'Production Register (handwritten)',
    hint: 'The handwritten daily production register — covers both page styles: the shade/size/GSM/weight/tukda style, and the product-description + quantity style. Extracted into one unified register.',
    register: 'production',
    systemPrompt: `You read a handwritten factory production register from a corrugated box factory. There are TWO different page styles used in this register — figure out which one you're looking at and extract accordingly:
STYLE A — shade/size/GSM style: columns are typically SL.NO. (ignore it), a column commonly headed "S/K" (this records paper SHADE, not a party name — decode 'S.K'/'S/K' as shade NS, 'G.Y' as shade GY, normalize to the 2-letter code, leave blank if unfamiliar rather than guessing), then GMS (GSM), Size (no separate BF column exists here), Weight (kg), and Tukda (count of pieces produced for that row). Most rows share one date at the top of the page.
STYLE B — product ledger style: each line has a DATE and a product DESCRIPTION, plus one or two quantity columns.
- Extract the FULL item name exactly as it appears in the factory's own product catalog below — including the weight and pack count (e.g. "Butter Bake 65g x60", "T50 64g x60"). Do NOT shorten it to just the brand/family word (do not output just "T50" or "Butter Bake" alone) — the weight and pack count are part of the item name, not separate data.
- Known handwriting misreads to correct, using the reference catalog below: the letter "g" (grams) is very often misread as the digit "9" — a pattern like "120g x60" is almost always grams, essentially never "1209 x60"; "&" is often misread as "8" or "5"; "Run" is often misread as "Rum". When a line clearly matches one of the catalog items below (allowing for this kind of misread), use the catalog's exact spelling. If it doesn't resemble anything in the catalog, transcribe your best reading rather than forcing a match.
- Reference catalog of known exact item names (case as shown), grouped by customer for your own matching confidence only — still extract just the item name into "description", not the customer name:
{{PRODUCT_CATALOG}}
  (This list isn't exhaustive — other legitimate products exist too. Only use it to correct obvious misreads of these specific items, never to force an unrelated line into matching one of them.)
- IMPORTANT: a customer name is sometimes written in brackets right next to the item name, either before or after it — e.g. "(Diamond) Cream Burst 30g x140" or "Cream Burst 30g x140 (Diamond)". Pull that bracketed name OUT into its own "customer_hint" field and do NOT leave it inside "description" — "description" should be just the clean item name with no bracket in it.
- Dates are often written once then repeated below with a ditto mark (a tick, quote mark, or short symbol like " or 11 or //) — resolve a ditto mark to the same date as the row above, never leave date blank because of one.
- Some pages have a single "Quantity" column — put that value into "pieces". Other pages have two columns headed roughly S and D side by side — S goes into "pieces", D goes into "dispatch". If a row only has one of the two, leave the other as 0.
For EITHER style: do not include page-total or running-total rows (a lone number with no row content, usually at the bottom of a page or column). Extract every real row on the page, in order, exactly once — do not skip rows and do not repeat any row, even if faint printing or ruling lines make it look duplicated. If an actual customer/party name is written somewhere on the page itself (separate from the shade column — e.g. as a page header/title, applying to the WHOLE page, not a per-row bracket), include it as "party" for every item on that page; otherwise omit "party" entirely.
Return ONLY one JSON object: {"date":"shared header date if Style A, DD/MM/YYYY, else omit","items":[{"date_override":"only if a specific row's date differs from the header or ditto-resolves to something else, else omit","party":"only if found as a whole-page header, else omit","customer_hint":"Style B only — a bracketed customer name found next to this specific item, else omit","shade":"Style A only, else omit","size":"Style A only, else omit","gsm":"Style A only, else omit","weight":"Style A only (number), else omit","description":"Style B only — the FULL exact item name including weight and pack count, with no bracketed customer name in it, else omit","pieces":"number — pieces produced (Tukda in Style A, the single Quantity or the S column in Style B)","dispatch":"number — Style B's D column only, else 0"}]}`,
    shape: (raw) => {
      if (!raw || !Array.isArray(raw.items)) return [];
      return raw.items.map(it => ({
        id: genId(), date: normalizeDateToDots(it.date_override || raw.date || it.date || ''), party: it.party || '', shade: it.shade || '',
        size: it.size || '', gsm: it.gsm || '', weight: num(it.weight), description: it.description || '',
        customerHint: it.customer_hint || '', pieces: num(it.pieces), dispatch: num(it.dispatch),
        stockConfirmed: false, confirmedCustomer: '',
      }));
    },
  },
  {
    key: 'dispatch_bill',
    label: 'Dispatch Bill / Tax Invoice',
    hint: 'A printed tax invoice / dispatch bill sent to any customer (Bindal, Diamond, Anmol, or otherwise) — lands in its own Customer Dispatch Bills tab, separate from the Production Register, and reduces that customer’s stock balance on the Customer Stock tab and feeds the Order Availability Check.',
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
    { key: 'weight_kg', label: 'Weight (kg)', type: 'number' },
  ],
  consumption: [
    { key: 'date', label: 'Date' }, { key: 'shade', label: 'Shade' }, { key: 'size', label: 'Size' }, { key: 'gsm', label: 'GSM' },
    { key: 'weight_consumed', label: 'Weight Consumed', type: 'number' }, { key: 'balance_left', label: 'Balance Left (Tukda)', type: 'number' },
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
const NAV = [
  { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { key: 'upload', label: 'Upload & Extract', icon: Upload },
  { key: 'rawMaterialIn', label: 'Raw Material Register', icon: Archive },
  { key: 'production', label: 'Production Register', icon: ClipboardList },
  { key: 'customerDispatch', label: 'Customer Dispatch Bills', icon: Package },
  { key: 'orderCheck', label: 'Order Availability Check', icon: Search },
  { key: 'daburSpecs', label: 'Dabur — Spec Master', icon: FileText },
  { key: 'daburPO', label: 'Dabur — Pending PO', icon: ListChecks },
  { key: 'daburDispatch', label: 'Dabur — Dispatch Log', icon: Truck },
  { key: 'customerStock', label: 'Customer Stock', icon: Boxes },
  { key: 'customerMapping', label: 'Customer Mapping', icon: ListChecks },
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
    how: 'Upload as usual. If a page has a customer name written on it (as a page header, not a bracket next to one item), it’s worth double-checking the Party column got filled in — that field is what the Order Availability Check searches by.',
  },
  {
    title: 'New finished-goods rows wait in Customer Stock for your OK',
    tab: 'customerStock',
    what: 'Any Production Register row with a product description, and any Customer Dispatch Bill row, shows up in "Pending Review" at the top of Customer Stock. Neither counts toward any customer’s balance until you confirm it.',
    how: 'Check the suggested customer for each row (auto-matched, editable if it’s wrong), then confirm one at a time or use "Confirm all suggested." Once confirmed, production rows add to that customer’s stock and dispatch bill rows subtract from it, in the running balance below.',
  },
  {
    title: 'Customer Mapping decides who gets what',
    tab: 'customerMapping',
    what: 'This is the rulebook Customer Stock uses to guess which customer a product belongs to — plus the reference catalog of exact item names that helps correct handwriting misreads during extraction.',
    how: 'Paste a customer’s Google Sheet ID/link on the Customer Sheets tab (new or already-added) to add its products automatically. You can also add, edit, or delete individual rules by hand at any time — rules are checked top to bottom, first match wins.',
  },
  {
    title: 'Check what’s available before promising a customer an order',
    tab: 'orderCheck',
    what: 'Type a party name (and optionally a size/GSM keyword) to see pieces produced vs. pieces dispatched, and what’s left. It shows you the exact matching rows too, not just a number, so you can double-check it.',
    how: 'This searches the Production Register’s Party field, so it only works well once party names are actually filled in on those rows.',
  },
  {
    title: 'Dispatch bills record what actually went out',
    tab: 'upload',
    what: 'Upload dispatch bills / tax invoices from the same dropdown as everything else — they land in their own Customer Dispatch Bills tab (kept separate from the Production Register) and, once confirmed, reduce that customer’s balance on both the Order Availability Check and Customer Stock.',
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
function Pill({ tone = 'neutral', children }) {
  return <span className={`pill pill-${tone}`}>{children}</span>;
}
function EditableTable({ columns, rows, onUpdate, onDelete, emptyLabel = 'No entries yet.' }) {
  if (!rows.length) return <div className="empty-state">{emptyLabel}</div>;
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>{columns.map(c => <th key={c.key}>{c.label}</th>)}<th className="col-action"></th></tr>
        </thead>
        <tbody>
          {rows.map(row => (
            <tr key={row.id}>
              {columns.map(c => (
                <td key={c.key}>
                  <input
                    className="cell-input"
                    type={c.type === 'number' ? 'number' : 'text'}
                    value={row[c.key] ?? ''}
                    onChange={(e) => onUpdate(row.id, c.key, c.type === 'number' ? e.target.value : e.target.value)}
                  />
                </td>
              ))}
              <td className="col-action">
                <button className="icon-btn danger" title="Delete row" onClick={() => onDelete(row.id)}>
                  <Trash2 size={15} />
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
function RegisterPanel({ title, subtitle, columns, rows, onUpdate, onDelete, onExport, extra }) {
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
      <EditableTable columns={columns} rows={rows} onUpdate={onUpdate} onDelete={onDelete} />
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
  const [trainingExamples, setTrainingExamples] = useState({});
  const [customerMapping, setCustomerMapping] = useState(DEFAULT_CUSTOMER_MAPPING);
  const [productCatalog, setProductCatalog] = useState(DEFAULT_PRODUCT_CATALOG);
  const [customerSheetIds, setCustomerSheetIds] = useState([]); // [{id, customer, sheetId}]
  const [pushStatus, setPushStatus] = useState({}); // { [customer]: { state: 'idle'|'pushing'|'done'|'error', message, unmatched } }
  const [serviceAccountEmail, setServiceAccountEmail] = useState('');
  // Review-before-push: reviewByCustomer holds the last fetched diff (a dry run against the customer's
  // REAL Sheet — see /api/customer-sheets/preview), reviewEdits holds only what a person has changed
  // in that preview (date/production/dispatch per row, or which tab an item is routed to) — never
  // written back into production/customerDispatch, only ever applied on top of the computed payload at
  // preview-refresh and push time. approvedByCustomer stores a snapshot of the edited payload at the
  // moment "Approve" was clicked; Push to Sheet is only enabled while the CURRENT edited payload still
  // matches that snapshot exactly — any further edit or newly confirmed register row invalidates it,
  // requiring a fresh look before anything real gets written.
  const [reviewByCustomer, setReviewByCustomer] = useState({}); // { [customer]: { loading, error, tabs, existingTabNames } }
  const [reviewEdits, setReviewEdits] = useState({}); // { [customer]: { [variantTitle]: { tabNameOverride, rowEdits: { [rowIndex]: {date,production,dispatch} } } } }
  const [approvedByCustomer, setApprovedByCustomer] = useState({}); // { [customer]: JSON string of the approved itemGroups }
  const registerState = { rawMaterialIn, consumption, production, customerDispatch, daburSpecs, daburPO, daburDispatch };
  const registerSetters = { rawMaterialIn: setRawMaterialIn, consumption: setConsumption, production: setProduction, customerDispatch: setCustomerDispatch, daburSpecs: setDaburSpecs, daburPO: setDaburPO, daburDispatch: setDaburDispatch };
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
        const r4 = await fetch('/api/service-account-email', { credentials: 'include' });
        if (r4.ok) { const j = await r4.json(); setServiceAccountEmail(j.email || ''); }
      } catch (e) { /* shown as blank below; the per-error messages still work without this */ }
      setLoaded(true);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
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
  const addRows = (registerKey, rows) => {
    registerSetters[registerKey](prev => {
      const next = [...prev, ...rows];
      persist(registerKey, next);
      return next;
    });
  };
  /* -------- upload & extract state -------- */
  const [docType, setDocType] = useState(DOCUMENT_TYPES[0].key);
  const [preview, setPreview] = useState(null);
  const [base64Img, setBase64Img] = useState(null);
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
  const extractWithRetry = async (prompt, base64, signal, attempts = 2) => {
    let lastErr;
    for (let i = 0; i < attempts; i++) {
      try {
        return await callClaudeExtract(prompt, base64, signal);
      } catch (e) {
        lastErr = e;
        if (isCancelled(e) || isRateLimitError(e)) break; // let the caller handle rate limits with a longer backoff, cancellation with a clean stop — not this quick retry
        if (i < attempts - 1) await sleep(600, signal); // brief pause before retry
      }
    }
    throw lastErr;
  };
  // Runs one extraction, and if it's specifically a rate-limit error, waits and retries automatically
  // (growing backoff) before giving up — rate limits are usually a per-minute window that clears on
  // its own, so a short wait-and-retry recovers most of the time without the person having to do anything.
  // onTick fires every second during a wait so the UI can show a live countdown instead of a frozen
  // message, and the whole wait aborts immediately if `signal` is cancelled (the person hit Cancel).
  const extractWithRateLimitBackoff = async (prompt, base64, signal, onWaiting, onTick) => {
    for (let i = 0; i <= RATE_LIMIT_BACKOFFS_MS.length; i++) {
      try {
        return await extractWithRetry(prompt, base64, signal);
      } catch (e) {
        if (isRateLimitError(e) && i < RATE_LIMIT_BACKOFFS_MS.length) {
          const waitMs = e.retryAfterMs && e.retryAfterMs > 0 ? e.retryAfterMs : RATE_LIMIT_BACKOFFS_MS[i];
          if (onWaiting) onWaiting(waitMs, i + 1, RATE_LIMIT_BACKOFFS_MS.length);
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
  const rateLimitExplainer = (e) => {
    const friendly = 'Still hitting a rate limit even after automatic retries — unusual at normal usage (this app\'s API key gets 1,000 requests/minute). Wait a minute and try again, or check the Anthropic Console for the account this key belongs to.';
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
      const basePrompt = activeConfig.systemPrompt.replace('{{PRODUCT_CATALOG}}', buildCatalogText(productCatalog));
      const promptWithTraining = buildPromptWithTraining(basePrompt, trainingExamples[docType]);
      const { raw, truncated } = await extractWithRateLimitBackoff(promptWithTraining, base64Img, abortControllerRef.current.signal, (waitMs, attempt, total) => {
        setErrorMsg(`Rate limit hit — waiting ${Math.round(waitMs / 1000)}s and retrying automatically (attempt ${attempt} of ${total})… Click Cancel below if you'd rather stop.`);
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
      } else if (isRateLimitError(e)) {
        setErrorMsg(rateLimitExplainer(e));
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
    const basePrompt = activeConfig.systemPrompt.replace('{{PRODUCT_CATALOG}}', buildCatalogText(productCatalog));
    const promptWithTraining = buildPromptWithTraining(basePrompt, trainingExamples[docType]);
    const gapMs = BATCH_REQUEST_GAP_MS;
    let anySucceeded = false;
    let anyTruncated = false;
    let rateLimited = false;
    let lastRateLimitError = null;
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
        const { raw, truncated } = await extractWithRateLimitBackoff(promptWithTraining, p.base64, signal, (waitMs, attempt, total) => {
          setErrorMsg(`Rate limit hit on "${p.label}" — waiting ${Math.round(waitMs / 1000)}s and retrying automatically (attempt ${attempt} of ${total})… Click Cancel below if you'd rather stop.`);
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
        if (isRateLimitError(e)) {
          rateLimited = true;
          lastRateLimitError = e;
          setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'pending', error: '' } : r));
          break; // stop hammering the same wall even after retries — leftover files stay pending, resumable
        }
        setFileResults(prev => prev.map(r => r.id === p.id ? { ...r, status: 'error', error: e.message || 'unknown error' } : r));
      }
    }
    if (cancelled) {
      setErrorMsg('Cancelled — files not yet extracted are still queued below, untouched. Remove any you don\'t want with the × on its thumbnail, or hit "Retry remaining" to pick back up.');
    } else if (rateLimited) {
      setErrorMsg(`${rateLimitExplainer(lastRateLimitError)} Anything already extracted in this batch is safe and untouched — use "Retry remaining" below once you're ready to continue.`);
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
  const recordCorrections = (originalRows, finalRows) => {
    if (!originalRows || !finalRows.length) return;
    const corrections = [];
    finalRows.forEach(row => {
      const orig = originalRows.find(o => o.id === row.id);
      if (!orig) return;
      const changed = Object.keys(row).some(k => k !== 'id' && String(orig[k] ?? '') !== String(row[k] ?? ''));
      if (changed) {
        const { id: _i1, ...beforeClean } = orig;
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
    addRows(activeConfig.register, page.rows);
    setFileResults(prev => prev.filter((_, i) => i !== idx));
    setActiveResultIndex(prev => Math.max(0, Math.min(prev, fileResults.length - 2)));
  };
  const confirmAllPages = () => {
    const donePages = fileResults.filter(r => r.status === 'done' && r.rows.length);
    donePages.forEach(page => recordCorrections(page.originalRows, page.rows));
    const allRows = donePages.flatMap(p => p.rows);
    if (allRows.length) addRows(activeConfig.register, allRows);
    setFileResults([]); setActiveResultIndex(0); setPreview(null); setBase64Img(null); setQueuedPages([]);
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
  /* -------- order availability check -------- */
  const [checkParty, setCheckParty] = useState('');
  const [checkKeyword, setCheckKeyword] = useState('');
  const [checkResult, setCheckResult] = useState(null);
  const runCheck = () => {
    const p = checkParty.trim().toLowerCase();
    const k = checkKeyword.trim().toLowerCase();
    if (!p) { setCheckResult(null); return; }
    const partyMatch = r => (r.party || '').toLowerCase().includes(p) || (r.confirmedCustomer || '').toLowerCase().includes(p) || (r.customerHint || '').toLowerCase().includes(p);
    const keywordMatch = r => !k || (r.size || '').toLowerCase().includes(k) || (r.gsm || '').toLowerCase().includes(k) || (r.description || '').toLowerCase().includes(k);
    const matchProd = production.filter(r => partyMatch(r) && keywordMatch(r) && num(r.pieces) > 0);
    const matchDisp = customerDispatch.filter(r => partyMatch(r) && keywordMatch(r) && num(r.quantity) > 0);
    const produced = matchProd.reduce((s, r) => s + num(r.pieces), 0);
    const dispatched = matchDisp.reduce((s, r) => s + num(r.quantity), 0);
    setCheckResult({ produced, dispatched, available: produced - dispatched, matchProd, matchDisp });
  };
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
  /* -------- product catalog (editable; populated via Customer Sheets tab's Sheet-ID import) -------- */
  const persistCatalog = (next) => {
    setProductCatalog(next);
    scheduleSave('catalog', () => window.storage.set(CATALOG_KEY, JSON.stringify(next), false).catch(() => {}));
  };
  const deleteCatalogItem = (id) => persistCatalog(productCatalog.filter(c => c.id !== id));
  const updateCatalogItem = (id, field, value) => persistCatalog(productCatalog.map(c => c.id === id ? { ...c, [field]: value } : c));
  // Shared dedup keys for the Sheet-ID import flow below — trims BOTH sides before comparing.
  // Catalog items imported from a Sheet are always already trimmed at extraction time, but a row
  // added by hand through the UI might have stray leading/trailing whitespace; without trimming both
  // sides symmetrically, a manually-entered "Hit & Run 120g " with a trailing space wouldn't match an
  // incoming "Hit & Run 120g" and would get silently re-added as a duplicate. Same reasoning for
  // mapping keywords.
  const catalogDedupKey = (customer, item) => `${(customer || '').trim().toLowerCase()}||${(item || '').trim().toLowerCase()}`;
  const mappingDedupKey = (keyword) => (keyword || '').trim().toLowerCase();
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
    let s = stripPackCount(description).toLowerCase();
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
    registerSetters[registerKey](prev => {
      const next = prev.map(r => r.id === row.id ? { ...r, confirmedCustomer: customer, stockConfirmed: true } : r);
      persist(registerKey, next);
      return next;
    });
  };
  const confirmAllPendingProduction = () => { pendingProductionRows.forEach(r => confirmStockRow('production', r)); };
  const confirmAllPendingDispatch = () => { pendingDispatchRows.forEach(r => confirmStockRow('customerDispatch', r)); };
  const updatePendingCustomer = (registerKey) => (id, value) => {
    registerSetters[registerKey](prev => {
      const next = prev.map(r => r.id === id ? { ...r, confirmedCustomer: value } : r);
      persist(registerKey, next);
      return next;
    });
  };
  const customerStockGroups = (() => {
    const groups = {};
    const addEntry = (row, pieces, dispatchQty) => {
      const customer = row.confirmedCustomer || matchCustomer(row);
      const variantKey = normalizeVariant(row.description);
      const key = `${customer}||${variantKey}`;
      // Display description is the pack-count-stripped, clean item name — not whichever day's raw
      // register wording happened to create this group first. Without this, the group's permanent
      // label (and the Sheet tab this pushes to) would be whatever pack count was on THAT entry,
      // e.g. "Kaju Bake 65g x60" forever, even on a day production ran a batch of 40 instead.
      if (!groups[key]) groups[key] = { id: key, customer, description: stripPackCount(row.description) || row.description, entries: [] };
      groups[key].entries.push({ id: row.id, date: row.date, pieces, dispatch: dispatchQty });
    };
    confirmedProductionRows.forEach(row => addEntry(row, num(row.pieces), num(row.dispatch)));
    confirmedDispatchRows.forEach(row => addEntry(row, 0, num(row.quantity)));
    return Object.values(groups).map(g => {
      const sorted = [...g.entries].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
      let running = 0;
      const ledger = sorted.map(e => {
        const opening = running;
        running = opening + num(e.pieces) - num(e.dispatch);
        return { ...e, opening, closing: running };
      });
      const totalProduction = sorted.reduce((s, e) => s + num(e.pieces), 0);
      const totalDispatch = sorted.reduce((s, e) => s + num(e.dispatch), 0);
      return { ...g, ledger, totalProduction, totalDispatch, closingBalance: running };
    });
  })();
  const customerNames = Array.from(new Set(customerStockGroups.map(g => g.customer))).sort((a, b) => (a === 'Unassigned') - (b === 'Unassigned') || a.localeCompare(b));
  /* -------- Customer Sheets: push a customer's stock ledger out to THEIR OWN separate Google
     Sheet (not a tab on this app's main sheet) — matching the real BINDAL STOCK.xlsx / DIAMOND.xlsx /
     anmol stock dec 22.xlsx layout exactly: one tab per base item ("IT 500", "Digestive", "CREAM"),
     with every variant of that item as its own side-by-side table within that tab. This is a
     write-only computed mirror (the Production Register and Customer Dispatch Bills stay the real
     source of truth) and is pushed ONLY when you click "Push to Sheet" for that customer below —
     never automatically. The item/variant tabs are MERGED, never overwritten: the backend reads what's
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
  const normalizeForCatalogMatch = (s) => (s || '')
    .toLowerCase()
    .replace(/\s*[x×]\s*\d+\s*$/i, '')
    .replace(/\bcont\.?\b/g, 'container')
    .replace(/[^a-z0-9]/g, '');
  // Every real customer Sheet writes dates as dot-separated D.M.YY ("5.1.24", "26.1.24") — never with
  // Rows land in the ledger already dot-formatted now (normalizeDateToDots runs at extraction time,
  // see DOCUMENT_TYPES above) — this second call right before push is a safety net for anything that
  // reached the ledger some other way (a manually typed/edited row), not the primary defense anymore.
  const buildCustomerSheetPayload = (customer) => {
    const groups = customerStockGroups.filter(g => g.customer === customer);
    const sheetGroupByItem = {};
    productCatalog.filter(c => c.customer === customer).forEach(c => {
      sheetGroupByItem[normalizeForCatalogMatch(c.item)] = (c.sheetGroup || c.item || '').trim();
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
      tabsMap[sheetGroup].push({
        title: g.description || 'Item',
        header: ['Date', 'Opening', 'Production', 'Dispatch', 'Closing'],
        rows: g.ledger.map(e => [normalizeDateToDots(e.date), e.opening, e.pieces || 0, e.dispatch || 0, e.closing]),
      });
    });
    const itemGroups = Object.entries(tabsMap).map(([tabName, variants]) => ({ tabName, variants }));
    // No summary payload here on purpose — the real "summary" tab in every customer's own Sheet is a
    // handful of live formulas (one per item, pointing at that item's block's own last row), which
    // already keeps itself correct as long as the app only ever appends rows to a block and never
    // rewrites it. The app never reads or writes that tab at all now, for any customer.
    return { itemGroups, unmatched: Array.from(new Set(unmatched)) };
  };
  // Applies whatever's staged in reviewEdits[customer] on top of the freshly computed itemGroups —
  // used both to regenerate the diff shown on screen and to build the exact payload Push actually
  // sends, so there's no way for what got approved to differ from what gets written. Edits live ONLY
  // here — a changed date/production/dispatch value, or a reassigned tab, is never written back into
  // `production`/`customerDispatch`, so the original register stays exactly as scanned or entered.
  const applyReviewEdits = (itemGroups, editsForCustomer) => {
    if (!editsForCustomer || !Object.keys(editsForCustomer).length) return itemGroups;
    const regrouped = {};
    itemGroups.forEach(g => (g.variants || []).forEach(v => {
      const e = editsForCustomer[v.title];
      const finalTab = (e && e.tabNameOverride && e.tabNameOverride.trim()) ? e.tabNameOverride.trim() : g.tabName;
      const rows = (v.rows || []).map((r, i) => {
        const re = e && e.rowEdits && e.rowEdits[i];
        if (!re) return r;
        return [
          re.date !== undefined ? re.date : r[0],
          r[1],
          (re.production !== undefined && re.production !== '') ? Number(re.production) : r[2],
          (re.dispatch !== undefined && re.dispatch !== '') ? Number(re.dispatch) : r[3],
          r[4],
        ];
      });
      if (!regrouped[finalTab]) regrouped[finalTab] = [];
      regrouped[finalTab].push({ ...v, rows });
    }));
    return Object.entries(regrouped).map(([tabName, variants]) => ({ tabName, variants }));
  };
  const getEditedPayload = (customer) => {
    const { itemGroups, unmatched } = buildCustomerSheetPayload(customer);
    return { itemGroups: applyReviewEdits(itemGroups, reviewEdits[customer]), unmatched };
  };
  // Dry-runs the CURRENT (edited) payload against the customer's real Sheet so the review area always
  // shows an accurate diff — old row (from the real Sheet, read fresh every time) next to the new
  // row(s) about to be appended — rather than a locally-guessed one. Called automatically whenever the
  // computed payload changes (a new register entry gets confirmed, a catalog mapping changes, or a
  // review edit is made), debounced below, and manually never needs to be triggered by hand.
  const refreshReview = async (customer) => {
    const sheetId = getCustomerSheetId(customer).trim();
    const { itemGroups } = getEditedPayload(customer);
    if (!sheetId || !itemGroups.length) { setReviewByCustomer(prev => ({ ...prev, [customer]: null })); return; }
    setReviewByCustomer(prev => ({ ...prev, [customer]: { ...(prev[customer] || {}), loading: true, error: '' } }));
    try {
      const res = await fetch('/api/customer-sheets/preview', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, credentials: 'include',
        body: JSON.stringify({ spreadsheetId: sheetId, itemGroups }),
      });
      if (res.status === 401) { window.dispatchEvent(new Event('fims-unauthorized')); return; }
      const data = await res.json().catch(() => ({}));
      if (res.ok && data.ok) {
        setReviewByCustomer(prev => ({ ...prev, [customer]: { loading: false, error: '', tabs: data.tabs, existingTabNames: data.existingTabNames || [] } }));
      } else {
        setReviewByCustomer(prev => ({ ...prev, [customer]: { loading: false, error: data.error || `Preview failed (HTTP ${res.status}).`, tabs: [] } }));
      }
    } catch (e) {
      setReviewByCustomer(prev => ({ ...prev, [customer]: { loading: false, error: e.message || 'Network error — could not reach the server.', tabs: [] } }));
    }
  };
  const setRowEdit = (customer, variantTitle, rowIndex, field, value) => {
    setReviewEdits(prev => {
      const forCustomer = { ...(prev[customer] || {}) };
      const forVariant = { ...(forCustomer[variantTitle] || {}) };
      const rowEdits = { ...(forVariant.rowEdits || {}) };
      rowEdits[rowIndex] = { ...(rowEdits[rowIndex] || {}), [field]: value };
      forVariant.rowEdits = rowEdits;
      forCustomer[variantTitle] = forVariant;
      return { ...prev, [customer]: forCustomer };
    });
  };
  const setTabOverride = (customer, variantTitle, tabName) => {
    setReviewEdits(prev => {
      const forCustomer = { ...(prev[customer] || {}) };
      forCustomer[variantTitle] = { ...(forCustomer[variantTitle] || {}), tabNameOverride: tabName };
      return { ...prev, [customer]: forCustomer };
    });
  };
  const approveReview = (customer) => {
    const { itemGroups } = getEditedPayload(customer);
    setApprovedByCustomer(prev => ({ ...prev, [customer]: JSON.stringify(itemGroups) }));
  };
  const isReviewApproved = (customer) => {
    const { itemGroups } = getEditedPayload(customer);
    return !!approvedByCustomer[customer] && approvedByCustomer[customer] === JSON.stringify(itemGroups);
  };
  // Auto-refreshes every customer's review whenever what would actually get pushed changes — a new
  // production/dispatch entry gets confirmed, a catalog mapping is edited, or a review edit is made.
  // Debounced so a burst of confirms (e.g. "Confirm all suggested") triggers one preview call per
  // customer, not one per row.
  const reviewRefreshKey = JSON.stringify({
    groups: customerStockGroups.map(g => ({ c: g.customer, d: g.description, ledger: g.ledger })),
    catalog: productCatalog, sheetIds: customerSheetIds, edits: reviewEdits,
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
      let skippedExisting = 0;
      if (mode === 'resync') {
        skippedExisting = items.filter(it => !it.isNew).length;
        items = items.filter(it => it.isNew); // additive-only re-sync — never re-offer/overwrite items already in the catalog
      }
      if (mode === 'resync' && !items.length) {
        setImportResultMessage(`${customerName}: re-synced — no new items found (${skippedExisting} already in the catalog).`);
        patchCustomerSheetEntry(customerName, { sheetId: id, lastImportedAt: new Date().toISOString() });
        return;
      }
      setSheetReview({ spreadsheetId: id, spreadsheetTitle: data.spreadsheetTitle || '', customerName, mode, resyncCustomer, items, skippedExisting });
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
  const confirmSheetImport = () => {
    if (!sheetReview || !sheetReview.customerName.trim()) return;
    const customer = sheetReview.customerName.trim();
    const included = sheetReview.items.filter(it => it.include && it.item.trim());
    const existingItems = new Set(productCatalog.map(c => catalogDedupKey(c.customer, c.item)));
    const newCatalogEntries = included
      .filter(it => !existingItems.has(catalogDedupKey(customer, it.item)))
      .map(it => ({ id: genId(), customer, item: it.item.trim(), sheetGroup: (it.sheetGroup || it.item).trim() }));
    const nextCatalog = [...productCatalog, ...newCatalogEntries];
    persistCatalog(nextCatalog);
    const existingKeywords = new Set(customerMapping.map(r => mappingDedupKey(r.keyword)));
    const exactRules = [];
    const fallbackRules = [];
    included.forEach(it => {
      const exact = it.item.trim().toLowerCase();
      if (exact && !existingKeywords.has(exact)) { exactRules.push({ id: genId(), keyword: exact, customer }); existingKeywords.add(exact); }
      const fallback = exact.replace(/[\d].*$/, '').trim();
      if (fallback && fallback.length > 2 && !existingKeywords.has(fallback)) { fallbackRules.push({ id: genId(), keyword: fallback, customer }); existingKeywords.add(fallback); }
    });
    persistCustomerMapping([...exactRules, ...customerMapping, ...fallbackRules]);
    patchCustomerSheetEntry(customer, {
      sheetId: sheetReview.spreadsheetId,
      lastImportedAt: new Date().toISOString(),
      itemCount: nextCatalog.filter(c => c.customer === customer).length,
    });
    const reassignedCount = reassignUnassignedRows();
    const verb = sheetReview.mode === 'resync' ? 'Re-synced' : sheetReview.mode === 'generate' ? 'Generated' : 'Imported';
    setImportResultMessage(
      `${verb} ${customer}: added ${newCatalogEntries.length} new catalog item${newCatalogEntries.length === 1 ? '' : 's'}` +
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
      setSheetReview({ spreadsheetId: sheetId, spreadsheetTitle: customer, customerName: customer, mode: 'generate', resyncCustomer: '', items, skippedExisting: 0 });
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
  // by group id, so several can be filled in without clobbering each other.
  const [assignForms, setAssignForms] = useState({});
  const updateAssignForm = (groupId, field, value) => setAssignForms(prev => ({ ...prev, [groupId]: { ...prev[groupId], [field]: value } }));
  // Same two-step fix an unassigned row already gets from a Customer Mapping rule: (1) make sure a
  // Product Catalog entry exists mapping this item to the chosen customer + sheet tab, so it lands
  // in the right block when the customer's sheet is next generated/pushed, and (2) add an exact-match
  // Customer Mapping keyword so this and any future row with the same description routes here
  // automatically. Then re-run the same reassignUnassignedRows() pass confirmSheetImport already
  // relies on to move this (and anything else now matching) off Unassigned immediately.
  const assignUnassignedGroup = (groupId, description) => {
    const form = assignForms[groupId] || {};
    const customer = (form.customer || '').trim();
    const sheetGroup = (form.sheetGroup || '').trim();
    const item = (form.item || description || '').trim();
    if (!customer || !sheetGroup || !item) return;
    const existingItems = new Set(productCatalog.map(c => catalogDedupKey(c.customer, c.item)));
    if (!existingItems.has(catalogDedupKey(customer, item))) {
      persistCatalog([...productCatalog, { id: genId(), customer, item, sheetGroup }]);
    }
    const existingKeywords = new Set(customerMapping.map(r => mappingDedupKey(r.keyword)));
    const exact = item.toLowerCase();
    if (exact && !existingKeywords.has(exact)) {
      persistCustomerMapping([{ id: genId(), keyword: exact, customer }, ...customerMapping]);
    }
    reassignUnassignedRows();
    setAssignForms(prev => { const next = { ...prev }; delete next[groupId]; return next; });
  };
  const pushCustomerSheetNow = async (customer) => {
    const sheetId = getCustomerSheetId(customer).trim();
    if (!sheetId) {
      setPushStatus(prev => ({ ...prev, [customer]: { state: 'error', message: 'Add a Google Sheet ID for this customer first (see field above).' } }));
      return;
    }
    if (!isReviewApproved(customer)) {
      setPushStatus(prev => ({ ...prev, [customer]: { state: 'error', message: 'Review the changes below and click Approve first — the diff has to be approved (and match exactly what you\'re about to push) before Push to Sheet will do anything.' } }));
      return;
    }
    // Approved and current — use the EDITED payload (review edits applied), never the raw computed
    // one, so what gets written matches exactly what was reviewed and approved.
    const { itemGroups, unmatched } = getEditedPayload(customer);
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
        setPushStatus(prev => ({ ...prev, [customer]: { state: 'done', message: `Pushed ${itemGroups.length} item tab${itemGroups.length === 1 ? '' : 's'} just now.`, unmatched } }));
        patchCustomerSheetEntry(customer, { sheetId, lastPushedAt: new Date().toISOString() });
        // The approval was for THIS specific diff — once it's actually written, clear both the
        // approval and the row-level edits (their values are now baked into the real rows the app
        // just appended) and pull a fresh diff so the screen immediately shows "nothing new" instead
        // of a stale approved state that happens to still match.
        setApprovedByCustomer(prev => { const next = { ...prev }; delete next[customer]; return next; });
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
        }
        .brand {
          padding: 0 18px 18px 18px;
          border-bottom: 1px solid rgba(255,255,255,0.12);
          margin-bottom: 10px;
        }
        .brand h1 { font-size: 17px; line-height: 1.3; color: #f6f3ec; }
        .brand p { font-size: 11px; color: #a7a396; margin-top: 4px; letter-spacing: 0.04em; text-transform: uppercase; }
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
        .preview-img { max-width: 100%; max-height: 260px; border-radius: 6px; border: 1px solid var(--rule); margin-top: 12px; }
        .error-box { display: flex; gap: 8px; align-items: flex-start; background: var(--warn-soft); border: 1px solid var(--ledger-red); color: #6b241a; padding: 10px 12px; border-radius: 5px; font-size: 12.5px; margin: 10px 0; }
        .info-box { display: flex; gap: 8px; align-items: flex-start; background: var(--accent-soft); border: 1px solid var(--rule); color: var(--ink); padding: 10px 12px; border-radius: 5px; font-size: 12.5px; margin: 10px 0; }
        .field-row { display: flex; gap: 10px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }
        .text-input { padding: 8px 10px; border-radius: 5px; border: 1px solid var(--rule); font: inherit; font-size: 13px; }
        .check-result { margin-top: 16px; }
        .check-stats { display: flex; gap: 14px; margin-bottom: 14px; flex-wrap: wrap; }
        .stat-box { background: var(--paper); border: 1px solid var(--rule); border-radius: 6px; padding: 12px 18px; min-width: 130px; }
        .stat-box .val { font-family: 'Fraunces', serif; font-size: 24px; }
        .stat-box .lbl { font-size: 11px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; }
        .review-box { border: 1px solid var(--accent); border-radius: 6px; padding: 16px; background: #fff; margin-top: 18px; }
        .review-actions { display: flex; gap: 10px; margin-top: 14px; }
        .section-label { font-size: 13px; font-weight: 700; margin: 18px 0 8px 0; color: var(--ink); }
        .spin { animation: fims-spin 1s linear infinite; }
        @keyframes fims-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      `}</style>
      <aside className="sidebar">
        <div className="brand">
          <h1>Shyam Adarsh Pack</h1>
          <p>Inventory &amp; Production</p>
        </div>
        {NAV.map(item => {
          const Icon = item.icon;
          const count = counts[item.key];
          return (
            <div key={item.key} className={`nav-item ${activeTab === item.key ? 'active' : ''}`} onClick={() => setActiveTab(item.key)}>
              <Icon size={16} />
              <span>{item.label}</span>
              {typeof count === 'number' && <span className="nav-count">{count}</span>}
            </div>
          );
        })}
      </aside>
      <div className="main">
        <div className="topbar">
          <h2>{NAV.find(n => n.key === activeTab)?.label}</h2>
          <button className="btn btn-primary" onClick={exportAll}><FileSpreadsheet size={15} /> Export all to Excel</button>
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
                {NAV.filter(n => n.key !== 'dashboard' && n.key !== 'upload' && n.key !== 'orderCheck' && n.key !== 'customerStock' && n.key !== 'customerMapping' && n.key !== 'settings').map(n => (
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
                {preview && <img src={preview} alt="preview" className="preview-img" />}
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
                          <p className="subtitle" style={{ marginBottom: 10 }}>{page.rows.length} row(s) found for <strong>{activeConfig.label}</strong>. Fix anything that looks wrong, then confirm.</p>
                          <EditableTable columns={COLUMNS[activeConfig.register]} rows={page.rows}
                            onUpdate={(rowId, field, value) => updateReviewCell(idx, rowId, field, value)}
                            onDelete={(rowId) => deleteReviewRow(idx, rowId)}
                            emptyLabel="All rows removed — nothing to add." />
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
                  <div><h2>Raw Material Balance</h2><p className="subtitle">Computed automatically from inward slips minus consumption entries, grouped by size / GSM. (BF isn't part of the match — the handwritten consumption sheets don't record it.)</p></div>
                  <button className="btn btn-ghost" onClick={() => exportSheet('RM_Balance', balanceRows, [
                    { key: 'size', label: 'Size' }, { key: 'gsm', label: 'GSM' },
                    { key: 'weight_in', label: 'Total In (kg)' }, { key: 'weight_consumed', label: 'Total Consumed (kg)' }, { key: 'balance', label: 'Balance Left (kg)' },
                  ])}><Download size={15} /> Export</button>
                </div>
                {!balanceRows.length && <div className="empty-state">Upload some mill slips and consumption reports to see balances here.</div>}
                {!!balanceRows.length && (
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Size</th><th>GSM</th><th>Total In (kg)</th><th>Total Consumed (kg)</th><th>Balance Left (kg)</th></tr></thead>
                      <tbody>
                        {balanceRows.map(b => (
                          <tr key={b.id}>
                            <td style={{ padding: '6px 10px' }}>{b.size}</td><td style={{ padding: '6px 10px' }}>{b.gsm}</td>
                            <td style={{ padding: '6px 10px' }}>{b.weight_in.toFixed(1)}</td><td style={{ padding: '6px 10px' }}>{b.weight_consumed.toFixed(1)}</td>
                            <td style={{ padding: '6px 10px' }}><Pill tone={b.balance <= 0 ? 'warn' : 'ok'}>{b.balance.toFixed(1)} kg</Pill></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
              <RegisterPanel title="Inward Entries (from mill slips)" columns={COLUMNS.rawMaterialIn} rows={rawMaterialIn}
                onUpdate={updateRow('rawMaterialIn')} onDelete={deleteRow('rawMaterialIn')} onExport={() => exportSheet('Raw_Material_In', rawMaterialIn, COLUMNS.rawMaterialIn)} />
              <RegisterPanel title="Consumption Entries (from daily reports)" columns={COLUMNS.consumption} rows={consumption}
                onUpdate={updateRow('consumption')} onDelete={deleteRow('consumption')} onExport={() => exportSheet('Consumption', consumption, COLUMNS.consumption)} />
            </div>
          )}
          {loaded && activeTab === 'production' && (
            <RegisterPanel title="Production Register" subtitle="From handwritten daily production sheets — covers both the shade/size/GSM style and the product-description style. Rows from the description style also feed the Customer Stock tab. Dispatch bills do NOT land here — see the Customer Dispatch Bills tab." columns={COLUMNS.production} rows={production}
              onUpdate={updateRow('production')} onDelete={deleteRow('production')} onExport={() => exportSheet('Production_Register', production, COLUMNS.production)} />
          )}
          {loaded && activeTab === 'customerDispatch' && (
            <RegisterPanel title="Customer Dispatch Bills" subtitle="From dispatch bills / tax invoices sent to customers (Bindal, Diamond, Anmol, or otherwise) — kept separate from the Production Register. Once confirmed on the Customer Stock tab, these reduce that customer's balance there and in the Order Availability Check." columns={COLUMNS.customerDispatch} rows={customerDispatch}
              onUpdate={updateRow('customerDispatch')} onDelete={deleteRow('customerDispatch')} onExport={() => exportSheet('Customer_Dispatch_Bills', customerDispatch, COLUMNS.customerDispatch)} />
          )}
          {loaded && activeTab === 'orderCheck' && (
            <div className="panel">
              <h2 style={{ marginBottom: 6 }}>Order Availability Check</h2>
              <p className="subtitle" style={{ marginBottom: 16 }}>Enter a party name (and optionally a size/GSM keyword) to see pieces produced (from the Production Register) vs. dispatched (from confirmed Customer Dispatch Bills), and what's still available. Note: the handwritten Production Register usually doesn't name a customer on the sheet itself — if "Party" is blank on those rows, click into the cell on the Production Register tab and fill it in manually so this check can find it.</p>
              <div className="field-row">
                <input className="text-input" placeholder="Party name (e.g. Anmol, Dabur, Ashoka)" value={checkParty} onChange={e => setCheckParty(e.target.value)} />
                <input className="text-input" placeholder="Size / GSM keyword (optional)" value={checkKeyword} onChange={e => setCheckKeyword(e.target.value)} />
                <button className="btn btn-primary" onClick={runCheck}><Search size={15} /> Check availability</button>
              </div>
              {checkResult && (
                <div className="check-result">
                  <div className="check-stats">
                    <div className="stat-box"><div className="val">{checkResult.produced}</div><div className="lbl">Produced (pcs)</div></div>
                    <div className="stat-box"><div className="val">{checkResult.dispatched}</div><div className="lbl">Dispatched (pcs)</div></div>
                    <div className="stat-box" style={{ borderColor: checkResult.available > 0 ? 'var(--ok)' : 'var(--ledger-red)' }}>
                      <div className="val" style={{ color: checkResult.available > 0 ? 'var(--ok)' : 'var(--ledger-red)' }}>{checkResult.available}</div>
                      <div className="lbl">Available (pcs)</div>
                    </div>
                  </div>
                  <div className="section-label">Matched production rows ({checkResult.matchProd.length})</div>
                  <EditableTable columns={COLUMNS.production} rows={checkResult.matchProd} onUpdate={updateRow('production')} onDelete={deleteRow('production')} emptyLabel="No matching production rows." />
                  <div className="section-label">Matched dispatch bill rows ({checkResult.matchDisp.length})</div>
                  <EditableTable columns={COLUMNS.customerDispatch} rows={checkResult.matchDisp} onUpdate={updateRow('customerDispatch')} onDelete={deleteRow('customerDispatch')} emptyLabel="No matching dispatch bill rows." />
                </div>
              )}
              {!checkResult && <p className="subtitle">No search run yet — results aren't guessed, only shown once you check.</p>}
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
                <p className="subtitle">Review new Production Register and Customer Dispatch Bill entries here and confirm which customer they belong to — that's the only thing that still lives on this tab. Once confirmed, matching is done by the Customer Mapping tab, plus anything literally bracketed next to the item name. The actual per-customer ledger, the diff against each customer's real Sheet, and pushing now all live in the Customer Sheets tab.</p>
              </div>
              {pendingProductionRows.length > 0 && (
                <div className="panel" style={{ borderColor: 'var(--accent)' }}>
                  <div className="panel-header">
                    <div><h2>Pending Production Review ({pendingProductionRows.length})</h2><p className="subtitle">New Production Register entries not yet reflected in the customer stock totals below. Check or fix the suggested customer, then confirm — nothing here counts toward balances until you do.</p></div>
                    <button className="btn btn-primary" onClick={confirmAllPendingProduction}><CheckCircle2 size={15} /> Confirm all suggested</button>
                  </div>
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Date</th><th>Description</th><th>Pieces</th><th>Suggested Customer</th><th className="col-action"></th></tr></thead>
                      <tbody>
                        {pendingProductionRows.map(row => (
                          <tr key={row.id}>
                            <td style={{ padding: '6px 10px' }}>{row.date}</td>
                            <td style={{ padding: '6px 10px' }}>{row.description}</td>
                            <td style={{ padding: '6px 10px' }}>{row.pieces || ''}</td>
                            <td>
                              <input className="cell-input" value={row.confirmedCustomer || matchCustomer(row)}
                                onChange={e => updatePendingCustomer('production')(row.id, e.target.value)} />
                            </td>
                            <td className="col-action">
                              <button className="icon-btn" style={{ color: 'var(--ok)' }} title="Confirm this row" onClick={() => confirmStockRow('production', row)}><CheckCircle2 size={16} /></button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              {pendingDispatchRows.length > 0 && (
                <div className="panel" style={{ borderColor: 'var(--ledger-red)' }}>
                  <div className="panel-header">
                    <div><h2>Pending Dispatch Bill Review ({pendingDispatchRows.length})</h2><p className="subtitle">New Customer Dispatch Bill entries not yet reflected in the customer stock totals below. Check or fix the suggested customer, then confirm — nothing here counts toward balances until you do.</p></div>
                    <button className="btn btn-primary" onClick={confirmAllPendingDispatch}><CheckCircle2 size={15} /> Confirm all suggested</button>
                  </div>
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>Date</th><th>Invoice No</th><th>Description</th><th>Quantity</th><th>Suggested Customer</th><th className="col-action"></th></tr></thead>
                      <tbody>
                        {pendingDispatchRows.map(row => (
                          <tr key={row.id}>
                            <td style={{ padding: '6px 10px' }}>{row.date}</td>
                            <td style={{ padding: '6px 10px' }}>{row.invoice_no}</td>
                            <td style={{ padding: '6px 10px' }}>{row.description}</td>
                            <td style={{ padding: '6px 10px' }}>{row.quantity || ''}</td>
                            <td>
                              <input className="cell-input" value={row.confirmedCustomer || matchCustomer(row)}
                                onChange={e => updatePendingCustomer('customerDispatch')(row.id, e.target.value)} />
                            </td>
                            <td className="col-action">
                              <button className="icon-btn" style={{ color: 'var(--ok)' }} title="Confirm this row" onClick={() => confirmStockRow('customerDispatch', row)}><CheckCircle2 size={16} /></button>
                            </td>
                          </tr>
                        ))}
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
                            <thead><tr><th>Date</th><th>Opening</th><th>Production</th><th>Dispatch</th><th>Closing</th></tr></thead>
                            <tbody>
                              {g.ledger.map((e, i) => (
                                <tr key={i}>
                                  <td style={{ padding: '6px 10px' }}>{e.date}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.opening}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.pieces || ''}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.dispatch || ''}</td>
                                  <td style={{ padding: '6px 10px' }}>{e.closing}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                        {(() => {
                          const form = assignForms[g.id] || {};
                          const chosenCustomer = (form.customer || '').trim();
                          const customerSheetGroups = Array.from(new Set(
                            productCatalog.filter(c => c.customer.toLowerCase() === chosenCustomer.toLowerCase()).map(c => c.sheetGroup)
                          )).filter(Boolean);
                          const chosenSheetGroup = (form.sheetGroup || '').trim();
                          const sheetGroupItems = Array.from(new Set(
                            productCatalog
                              .filter(c => c.customer.toLowerCase() === chosenCustomer.toLowerCase() && c.sheetGroup.toLowerCase() === chosenSheetGroup.toLowerCase())
                              .map(c => c.item)
                          )).filter(Boolean);
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
                                placeholder="Block (item name)"
                                value={form.item !== undefined ? form.item : g.description}
                                onChange={e => updateAssignForm(g.id, 'item', e.target.value)}
                                style={{ width: 180 }}
                              />
                              <datalist id={`assign-item-${g.id}`}>
                                {sheetGroupItems.map(i => <option key={i} value={i} />)}
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
                );
              })()}
              {!pendingProductionRows.length && !pendingDispatchRows.length && !customerStockGroups.filter(g => g.customer === 'Unassigned').length && (
                <div className="empty-state">Nothing pending review right now. Per-customer stock and the diff before pushing now live in the Customer Sheets tab.</div>
              )}
            </div>
          )}
          {loaded && activeTab === 'customerMapping' && (
            <div>
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
                        <thead><tr><th>Item name</th><th>Sheet Tab (item group)</th><th className="col-action"></th></tr></thead>
                        <tbody>
                          {productCatalog.filter(c => c.customer === customer).map(c => (
                            <tr key={c.id}>
                              <td><input className="cell-input" value={c.item} onChange={e => updateCatalogItem(c.id, 'item', e.target.value)} /></td>
                              <td><input className="cell-input" value={c.sheetGroup || ''} placeholder={c.item} onChange={e => updateCatalogItem(c.id, 'sheetGroup', e.target.value)} /></td>
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
                <p className="subtitle" style={{ marginBottom: 4 }}>Push a customer's stock ledger out to THEIR OWN separate Google Sheet — not a tab on this app's main sheet — laid out the same way as your original BINDAL STOCK.xlsx / DIAMOND.xlsx / anmol stock files: one tab per base item, every variant of that item as its own side-by-side table within that tab.</p>
                <p className="subtitle" style={{ marginBottom: 4 }}>Pushing only happens when you click "Push to Sheet" below — nothing is pushed automatically. Pushing only appends: it never touches or duplicates a row that's already in the sheet, whether it got there from a previous push or someone typed it in by hand, and new rows use the same Opening/Closing formula convention already in your sheet instead of dumping in flat numbers. The "summary" tab is never touched at all — its existing formulas already point at each item's latest row, so they keep calculating themselves.</p>
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
                      {sheetReview.mode === 'resync' ? `Re-syncing ${sheetReview.customerName} — showing only items not already in the catalog.`
                        : sheetReview.mode === 'generate' ? `Structure built in the Sheet for "${sheetReview.customerName}" — every tab, item block, and the summary tab are already written and live.`
                        : `Imported from "${sheetReview.spreadsheetTitle || sheetReview.spreadsheetId}".`}
                    </p>
                    <div className="field-row">
                      <span style={{ fontSize: 13, fontWeight: 600 }}>Customer name:</span>
                      <input className="text-input" value={sheetReview.customerName} disabled={sheetReview.mode === 'resync'} onChange={e => setSheetReview(prev => ({ ...prev, customerName: e.target.value }))} />
                    </div>
                    <p className="subtitle" style={{ marginBottom: 10 }}>{sheetReview.items.length} new item{sheetReview.items.length === 1 ? '' : 's'} found{sheetReview.skippedExisting ? ` (${sheetReview.skippedExisting} already known, not shown)` : ''}. Uncheck anything that isn't really a product, fix any typos, adjust which Sheet Tab each lands under, then confirm.</p>
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
                      <button className="btn btn-primary" onClick={confirmSheetImport} disabled={!sheetReview.customerName.trim()}><CheckCircle2 size={15} /> Add to catalog &amp; mapping</button>
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
                const status = pushStatus[customer] || {};
                const { itemGroups, unmatched } = buildCustomerSheetPayload(customer);
                const variantCount = itemGroups.reduce((s, g) => s + (g.variants || []).length, 0);
                const sheetId = getCustomerSheetId(customer);
                const registryEntry = customerSheetIds.find(c => c.customer === customer);
                const review = reviewByCustomer[customer];
                const reviewTabs = (review && review.tabs) || [];
                const hasAnyNewRows = reviewTabs.some(t => (t.variants || []).some(v => v.rows.length > 0));
                const approved = isReviewApproved(customer);
                return (
                  <div className="panel" key={customer}>
                    <div className="panel-header">
                      <div>
                        <h2>{customer}</h2>
                        <p className="subtitle">{itemGroups.length} item tab{itemGroups.length === 1 ? '' : 's'} · {variantCount} variant{variantCount === 1 ? '' : 's'} ready to push
                          {registryEntry?.lastImportedAt && <> · last imported {new Date(registryEntry.lastImportedAt).toLocaleString()}</>}
                          {registryEntry?.lastPushedAt && <> · last pushed {new Date(registryEntry.lastPushedAt).toLocaleString()}</>}
                        </p>
                      </div>
                      <div style={{ display: 'flex', gap: 8 }}>
                        <button className="btn btn-ghost" onClick={() => importSheetById(sheetId, { mode: 'resync', resyncCustomer: customer })} disabled={sheetImportBusy || !sheetId.trim()} title="Re-check this customer's Sheet for newly added items and add them to the catalog">
                          <RefreshCw size={15} /> Re-sync
                        </button>
                        <button className="btn btn-primary" onClick={() => pushCustomerSheetNow(customer)} disabled={status.state === 'pushing' || !sheetId.trim() || !approved} title={!approved ? 'Approve the changes in the review section below first' : ''}>
                          {status.state === 'pushing' ? <Loader2 size={15} className="spin" /> : <FileSpreadsheet size={15} />} Push to Sheet
                        </button>
                      </div>
                    </div>
                    <div className="field-row">
                      <input className="text-input" style={{ minWidth: 340, flex: 1 }} placeholder="Paste this customer's Google Sheet link or ID" value={sheetId} onChange={e => updateCustomerSheetId(customer, e.target.value)} autoComplete="off" spellCheck={false} />
                      {registryEntry && (
                        <button className="btn btn-ghost" onClick={() => { if (window.confirm(`Stop tracking ${customer}'s Sheet ID? This only removes the tracking record — nothing is deleted from the Known Product Catalog or Customer Mapping.`)) removeCustomerSheetEntry(customer); }} title="Stop tracking this Sheet ID (catalog items stay)">
                          <Trash2 size={15} /> Remove tracking
                        </button>
                      )}
                    </div>
                    {(unmatched.length > 0 && sheetId.trim()) && (
                      <div className="info-box" style={{ marginTop: 10 }}>
                        <Info size={16} />
                        <span>{unmatched.length} item{unmatched.length === 1 ? '' : 's'} in your Production/Dispatch records don't exactly match a known item name for {customer}, so they will NOT be pushed to their Sheet at all (nothing gets written for them, and no new tab gets created): {unmatched.join(', ')}. If these are real products, add them to the Known Product Catalog (Customer Mapping tab) with the exact wording that customer's Sheet uses, then push again.</span>
                      </div>
                    )}
                    {status.state === 'done' && <div className="doc-hint" style={{ color: 'var(--ok)', marginTop: 10 }}>✓ {status.message}</div>}
                    {status.state === 'error' && <div className="error-box" style={{ marginTop: 10 }}><AlertCircle size={16} /><span>{status.message}</span></div>}
                    {!sheetId.trim() && <div className="doc-hint" style={{ marginTop: 10 }}>Add a Sheet ID above to enable pushing or re-syncing.</div>}
                    {sheetId.trim() && review && (
                      <div style={{ marginTop: 14, borderTop: '1px solid var(--border)', paddingTop: 12 }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                          <h3 style={{ margin: 0, fontSize: 14 }}>
                            Review before push
                            {review.loading && <span className="doc-hint" style={{ marginLeft: 8, fontWeight: 400 }}><Loader2 size={12} className="spin" style={{ verticalAlign: 'middle' }} /> checking against the real Sheet…</span>}
                          </h3>
                          {hasAnyNewRows && (
                            approved
                              ? <span style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                                  <span className="doc-hint" style={{ color: 'var(--ok)' }}>✓ Approved — matches what Push will send</span>
                                  <button className="btn btn-ghost" onClick={() => setApprovedByCustomer(prev => { const next = { ...prev }; delete next[customer]; return next; })}>Edit again</button>
                                </span>
                              : <button className="btn btn-primary" onClick={() => approveReview(customer)}><CheckCircle2 size={15} /> Approve these changes</button>
                          )}
                        </div>
                        {review.error && <div className="error-box"><AlertCircle size={16} /><span>{review.error}</span></div>}
                        {!review.error && !hasAnyNewRows && !review.loading && <div className="doc-hint">Nothing new to push right now — every confirmed entry is already reflected in the real Sheet.</div>}
                        {reviewTabs.map(tab => (tab.variants || []).filter(v => v.rows.length > 0).map(v => (
                          <div key={`${tab.tabName}::${v.title}`} style={{ border: '1px solid var(--border)', borderRadius: 8, padding: 10, marginBottom: 8 }}>
                            <div style={{ display: 'flex', alignItems: 'baseline', marginBottom: 6, flexWrap: 'wrap', gap: 6 }}>
                              {/* alignItems:'baseline' (not 'center') — the title is 16px bold and the
                                  label/input/new-tag are all 12px, so centering by height put the smaller
                                  text visibly higher than the title instead of sitting on the same line of
                                  text. Baseline alignment lines up their actual text baselines, which is
                                  what "sitting on the same line" means for mixed font sizes. */}
                              <strong>{v.title}</strong>
                              <span className="doc-hint" style={{ whiteSpace: 'nowrap' }}>— Sheet tab:</span>
                              {(() => {
                                const tabNameValue = (reviewEdits[customer] && reviewEdits[customer][v.title] && reviewEdits[customer][v.title].tabNameOverride) ?? tab.tabName;
                                // Sits directly after the item name now (no more floating it to the far
                                // right of the row, which is what caused the label and box to separate
                                // onto different lines). Width is one flat, generous fixed value rather
                                // than a per-character estimate — character-counting approaches (px/char,
                                // then `ch` units) both undersized real lowercase tab names because
                                // neither actually measures the font's real glyph widths. 200px comfortably
                                // fits the longest real tab name across all three customers ("coconut
                                // premium cookies") with room to spare, so nothing truncates — no ellipsis
                                // needed.
                                return (
                                  <input
                                    className="cell-input"
                                    style={{ width: 200, fontSize: 12 }}
                                    list={`review-tabs-${customer}`}
                                    value={tabNameValue}
                                    onChange={e => setTabOverride(customer, v.title, e.target.value)}
                                    disabled={approved}
                                  />
                                );
                              })()}
                              <datalist id={`review-tabs-${customer}`}>
                                {(review.existingTabNames || []).map(t => <option value={t} key={t} />)}
                              </datalist>
                              {tab.isNewTab
                                ? <span className="doc-hint">(new — this tab doesn't exist yet)</span>
                                : v.isNewBlock && <span className="doc-hint">(new — this block doesn't exist yet)</span>}
                            </div>
                            <table style={{ width: '100%', fontSize: 13, borderCollapse: 'collapse' }}>
                              <thead>
                                <tr style={{ textAlign: 'left', color: 'var(--muted)' }}>
                                  <th style={{ padding: '2px 6px', fontWeight: 500 }}>Date</th>
                                  <th style={{ padding: '2px 6px', fontWeight: 500 }}>Opening</th>
                                  <th style={{ padding: '2px 6px', fontWeight: 500 }}>Production</th>
                                  <th style={{ padding: '2px 6px', fontWeight: 500 }}>Dispatch</th>
                                  <th style={{ padding: '2px 6px', fontWeight: 500 }}>Closing</th>
                                </tr>
                              </thead>
                              <tbody>
                                {v.lastExisting && (
                                  <tr style={{ opacity: 0.55 }} title="Already in the real Sheet — shown for reference, never touched">
                                    <td style={{ padding: '2px 6px' }}>{v.lastExisting.date}</td>
                                    <td style={{ padding: '2px 6px' }} colSpan={3}>(already in the Sheet)</td>
                                    <td style={{ padding: '2px 6px' }}>{v.lastExisting.closing}</td>
                                  </tr>
                                )}
                                {v.rows.map((r, i) => {
                                  const edit = (reviewEdits[customer] && reviewEdits[customer][v.title] && reviewEdits[customer][v.title].rowEdits && reviewEdits[customer][v.title].rowEdits[i]) || {};
                                  return (
                                    <tr key={i} style={{ background: 'rgba(214,163,80,0.14)' }}>
                                      <td style={{ padding: '2px 6px' }}>
                                        <input className="cell-input" style={{ width: 100 }} value={edit.date ?? r.date} onChange={e => setRowEdit(customer, v.title, i, 'date', e.target.value)} disabled={approved} />
                                      </td>
                                      <td style={{ padding: '2px 6px' }}>{r.opening}</td>
                                      <td style={{ padding: '2px 6px' }}>
                                        <input className="cell-input" style={{ width: 80 }} type="number" value={edit.production ?? r.production} onChange={e => setRowEdit(customer, v.title, i, 'production', e.target.value)} disabled={approved} />
                                      </td>
                                      <td style={{ padding: '2px 6px' }}>
                                        <input className="cell-input" style={{ width: 80 }} type="number" value={edit.dispatch ?? r.dispatch} onChange={e => setRowEdit(customer, v.title, i, 'dispatch', e.target.value)} disabled={approved} />
                                      </td>
                                      <td style={{ padding: '2px 6px' }}>{r.closing}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        )))}
                      </div>
                    )}
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
            </div>
          )}
        </div>
      </div>
    </div>
    {copyModal && <CopyExportModal data={copyModal} onClose={() => setCopyModal(null)} />}
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
