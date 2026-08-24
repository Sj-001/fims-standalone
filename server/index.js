const express = require('express');
const cookieSession = require('cookie-session');
const cors = require('cors');
const path = require('path');
const { requireEnv, login, logout, requireAuth } = require('./lib/auth');
const { extract } = require('./lib/anthropic');
const { getTab, putTab, putBlocksTab, pushCustomerSheetHandler, previewCustomerSheetHandler, importCustomerSheetHandler, getServiceAccountEmailHandler, generateCustomerSheetStructureHandler, listTabsHandler, deleteTabsHandler } = require('./lib/sheets');

// Fail fast and loud if required secrets are missing, instead of the app half-working with
// confusing downstream errors — matches the "no guesses" standard this project was built to.
requireEnv('APP_PASSWORD');
requireEnv('SESSION_SECRET');
requireEnv('ANTHROPIC_API_KEY');
requireEnv('GOOGLE_SHEET_ID');
if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON && !process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64) {
  console.error('FATAL: set either GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_JSON_BASE64.');
  process.exit(1);
}

const app = express();
// Render (like Heroku and most PaaS hosts) terminates HTTPS at its edge and forwards requests to
// this app over plain HTTP internally. Without this line, Express has no way to know the original
// request was actually HTTPS, so it reports every request as insecure. That silently breaks the
// session cookie below: with secure:true (forced by NODE_ENV=production), the cookie library
// refuses to set a "secure" cookie on what it believes is an unencrypted connection, throws
// internally, and cookie-session swallows that error — so login appears to succeed (200 OK) but no
// session cookie is ever actually stored, and every request right after gets bounced back out as
// unauthenticated. Trusting the first proxy hop makes Express read Render's X-Forwarded-Proto
// header instead, so it correctly sees "https" and the cookie gets set as intended.
app.set('trust proxy', 1);
app.use(cors());
app.use(express.json({ limit: '15mb' })); // extraction requests include a base64 JPEG
app.use(cookieSession({
  name: 'fims_session',
  secret: process.env.SESSION_SECRET,
  maxAge: 30 * 24 * 60 * 60 * 1000, // 30 days
  httpOnly: true,
  sameSite: 'lax',
  secure: process.env.NODE_ENV === 'production',
}));

// --- auth ---
app.post('/api/login', login);
app.post('/api/logout', logout);
app.get('/api/session', (req, res) => res.json({ authed: !!(req.session && req.session.authed) }));

// everything below requires a valid session
app.use('/api', requireAuth);

// --- extraction proxy ---
app.post('/api/extract', extract);

// --- sheets-backed storage ---
app.get('/api/sheets/:tab', getTab);
app.post('/api/sheets/:tab', putTab);
// Multi-item "titled table" tabs — generic stacked-blocks writer, kept for any future use.
app.post('/api/sheets/:tab/blocks', putBlocksTab);
// Customer Sheets: pushes a customer's computed stock ledger to THEIR OWN external Google Sheet
// (spreadsheetId comes from the request body, not GOOGLE_SHEET_ID) — one tab per base item with
// variants laid out side by side, plus a summary tab. Manually triggered from the Customer Sheets
// tab in the app, never automatic.
app.post('/api/customer-sheets/push', pushCustomerSheetHandler);
// Dry-run of the above — same request shape, returns the diff (old row vs. new rows about to be
// appended) without writing anything. Backs the review-before-push screen in the Customer Sheets tab.
app.post('/api/customer-sheets/preview', previewCustomerSheetHandler);
// Reads an existing customer's own Google Sheet (any spreadsheet ID) and returns its item list with
// which tab each item belongs to, for review before adding to the Product Catalog / Customer Mapping.
app.post('/api/customer-sheets/import', importCustomerSheetHandler);
// The service account's own email — the frontend shows this directly in the Customer Sheets tab so
// there's a one-click answer to "what do I share the customer's Sheet with," instead of making
// someone dig it out of a Render env var or a downloaded JSON key file.
app.get('/api/service-account-email', getServiceAccountEmailHandler);
// Builds the full tab/item-block/formula structure (and a generated summary tab) for a brand-new
// customer inside a spreadsheet the person already created and shared with the service account
// themselves — the service account can't create+own a new file itself (confirmed: plain service
// accounts have no Drive storage of their own), but writing into an already-shared blank sheet needs
// nothing beyond the existing Sheets permission. See generateCustomerSheetStructure in lib/sheets.js.
app.post('/api/customer-sheets/generate-structure', generateCustomerSheetStructureHandler);
// Main-sheet maintenance: list every tab (labeled known/internal/unrecognized, see lib/sheets.js) and
// delete a person-confirmed list of them — surfaced in the app's Settings screen.
app.get('/api/maintenance/tabs', listTabsHandler);
app.post('/api/maintenance/delete-tabs', deleteTabsHandler);

// --- serve the built frontend (client/dist), if present, so this is a single deployable service ---
const clientDist = path.join(__dirname, '..', 'client', 'dist');
app.use(express.static(clientDist));
app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api/')) return next();
  res.sendFile(path.join(clientDist, 'index.html'), (err) => { if (err) next(); });
});

app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error.' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`FIMS server listening on port ${PORT}`));