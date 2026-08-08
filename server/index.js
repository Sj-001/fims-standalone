const express = require('express');
const cookieSession = require('cookie-session');
const cors = require('cors');
const path = require('path');
const { requireEnv, login, logout, requireAuth } = require('./lib/auth');
const { extract } = require('./lib/anthropic');
const { getTab, putTab } = require('./lib/sheets');

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
