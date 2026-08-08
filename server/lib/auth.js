// Simple shared-password gate for the whole app. No per-user accounts — everyone on the team
// uses the same password (set via the APP_PASSWORD environment variable on Render). A successful
// login sets a signed, httpOnly cookie (via cookie-session) so the browser stays logged in;
// nothing about the session is stored server-side, which keeps this safe to run on Render's
// free/ephemeral instances (a restart doesn't lose a session store because there isn't one).

function requireEnv(name) {
  const v = process.env[name];
  if (!v) {
    console.error(`FATAL: missing required environment variable ${name}. Set it in Render's Environment tab.`);
    process.exit(1);
  }
  return v;
}

function login(req, res) {
  const { password } = req.body || {};
  const expected = process.env.APP_PASSWORD;
  if (!expected) {
    return res.status(500).json({ error: 'Server misconfigured: APP_PASSWORD is not set.' });
  }
  if (typeof password !== 'string' || password !== expected) {
    return res.status(401).json({ error: 'Wrong password.' });
  }
  req.session.authed = true;
  res.json({ ok: true });
}

function logout(req, res) {
  req.session = null;
  res.json({ ok: true });
}

// Applied to every /api/* route except /api/login. Returns 401 if the session cookie isn't
// present/valid, which the frontend treats as "show the login screen."
function requireAuth(req, res, next) {
  if (req.session && req.session.authed) return next();
  return res.status(401).json({ error: 'Not logged in.' });
}

module.exports = { requireEnv, login, logout, requireAuth };
