# FIMS — standalone app

A password-gated web app: React frontend + a small Node/Express backend that holds your Anthropic
API key and Google service-account credentials, proxying document extraction to Anthropic and
reading/writing your Google Sheet. Deploys as one Render web service.

This has been built and tested locally (backend auth/session/proxy flow, and a full production
build of the frontend served by the backend) but **not yet tested with your real Anthropic key or
real Google Sheet** — that only becomes possible once it's deployed with real credentials, since
this environment has no outbound internet access to Google or Anthropic.

## What you already have

- GitHub account
- Google Cloud project with the Sheets API enabled, a service account, its downloaded JSON key,
  and your spreadsheet shared with the service account's email as an Editor
- Anthropic Console workspace "FIMS Factory App" with a $10 spend limit and a scoped API key
- Render account (free Hobby tier)

## 1. Push this code to GitHub

From a terminal, inside this `fims-standalone` folder:

```
git init
git add .
git commit -m "FIMS standalone app"
```

Then create a new empty repository on github.com (no README/gitignore — you already have one),
and follow GitHub's "push an existing repository" instructions it shows you, e.g.:

```
git remote add origin https://github.com/YOUR-USERNAME/fims-standalone.git
git branch -M main
git push -u origin main
```

Do this part yourself — pasting GitHub credentials or tokens isn't something I can do for you.

## 2. Create the Render web service

1. On [render.com](https://render.com), click **New +** → **Web Service**.
2. Connect your GitHub account if you haven't, and select the `fims-standalone` repo.
3. Fill in:
   - **Name**: anything, e.g. `fims`
   - **Region**: closest to you
   - **Branch**: `main`
   - **Root Directory**: leave blank (the repo root)
   - **Build Command**: `npm run build`
   - **Start Command**: `npm start`
   - **Instance Type**: Free (Hobby) — or Starter ($7/mo) if you decide you don't want the sleep
     delay after all

Don't click "Create Web Service" yet — set the environment variables first (next step), so the
first deploy doesn't crash on a missing variable.

## 3. Set environment variables

In the same Render setup screen, scroll to **Environment Variables** and add each of these. You
paste these values yourself directly into Render's form — that's a security boundary I follow, not
a limitation of the app.

| Key | Value |
|---|---|
| `APP_PASSWORD` | The shared password you and your friend will use to log in. Pick something memorable but not guessable. |
| `SESSION_SECRET` | Any long random string. Generate one by running `node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"` on your own computer, then paste the output here. |
| `ANTHROPIC_API_KEY` | Your `sk-ant-...` key from the "FIMS Factory App" workspace. |
| `GOOGLE_SHEET_ID` | The ID from your Sheet's URL — the part between `/d/` and `/edit`. |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The full contents of your downloaded service-account JSON key file, pasted as-is. |
| `NODE_ENV` | `production` |

If Render's text box mangles the multi-line JSON key (line breaks or quotes getting stripped), use
`GOOGLE_SERVICE_ACCOUNT_JSON_BASE64` instead: base64-encode the JSON file first
(`base64 -i your-key.json | tr -d '\n'` on Mac/Linux) and paste that single-line result instead —
the backend accepts either one.

Now click **Create Web Service**. The first deploy takes a few minutes (installing dependencies,
building the frontend).

## 4. Test it

Once the deploy finishes, Render gives you a URL like `https://fims.onrender.com`. Open it:

1. You should see a password screen — enter the `APP_PASSWORD` you set.
2. Upload a real document and extract it — this is the first real test against your actual
   Anthropic key.
3. Check your Google Sheet — a new tab should appear (e.g. `fims_production`) with your data as
   real rows and columns.

If anything fails, check Render's **Logs** tab for the specific error — the backend is written to
return clear error messages (e.g. "Could not read from Google Sheets: ...") rather than crash
silently.

## 5. Share with your friend

Just send them the Render URL and the `APP_PASSWORD`. No separate account needed — it's the same
shared login for everyone. On the free Hobby tier, whoever opens the link after 15 minutes of no
use waits ~30-60 seconds for the server to wake up; that's the tradeoff you already accepted.

## Notes

- **Changing the password later**: update `APP_PASSWORD` in Render's Environment tab and redeploy
  (Render does this automatically when an env var changes). Anyone already logged in stays logged
  in until their session cookie expires (30 days) or they log out.
- **Local development**: `npm run build` at the repo root builds everything; `npm start` runs the
  server, which serves the built frontend on the same port. For live-reloading frontend development,
  run `npm run dev` inside `client/` (proxies `/api/*` to a backend you run separately with
  `npm start` inside `server/`) — you'll need all the same environment variables set locally too
  (copy `.env.example` to `server/.env`, or export them in your shell).
- **Costs**: see the cost summary already reviewed — Render Hobby is $0/mo, Google Sheets API is
  $0 (far under the free quota at this usage), Anthropic extraction is roughly 1-2 cents/page from
  your $10-capped workspace.
