// Proxies document-extraction requests to Anthropic using the server's own API key — this is the
// whole reason a backend exists instead of a static site: the key must never reach the browser.
// Deliberately forwards Anthropic's status code and response body through almost unchanged, so the
// frontend's existing error-parsing logic (which reads the HTTP status and body text to distinguish
// a rate limit from other failures) keeps working with zero changes after the migration.

const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';
const DEFAULT_MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-5';

async function extract(req, res) {
  const { systemPrompt, base64Image } = req.body || {};
  if (!systemPrompt || !base64Image) {
    return res.status(400).json({ error: 'systemPrompt and base64Image are required.' });
  }
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'Server misconfigured: ANTHROPIC_API_KEY is not set.' });
  }
  try {
    const upstream = await fetch(ANTHROPIC_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: DEFAULT_MODEL,
        // Was 4096, then 8192 — both confirmed too low directly: a ~33-row production-register page
        // came back with only 1 row (cut off after row 1), a 31-row page came back with 23, and a
        // 26+-row page still got cut off at 8192. Denser pages, or ones with harder handwriting the
        // model has to work through (more reasoning per row before it commits to a value), can
        // genuinely need more room than that. This is a hard generation ceiling per request — money/
        // credits are billed on tokens actually generated, not on this ceiling, so raising it doesn't
        // cost anything extra on requests that were already finishing well under it; it only helps the
        // dense ones that were hitting the wall. See the stop_reason/truncated handling in
        // callClaudeExtract (client/src/App.jsx) for how a response that STILL hits this ceiling gets
        // surfaced instead of silently accepted.
        max_tokens: 16000,
        system: systemPrompt,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: base64Image } },
            { type: 'text', text: 'Extract the data from this document image now. Return ONLY the JSON described in your instructions — no markdown fences, no commentary, no explanation. Be concise: do not repeat header-level values inside each item unless asked to. CRITICAL: extract each physical row exactly once, in the same top-to-bottom order it appears — do not output the same row twice, even if faint printing, ruling lines, or a neighboring similar-looking row makes it seem repeated. The number of items you return should equal the number of actual handwritten/printed lines on the page, no more, no less. CRITICAL — READ ROW BY ROW, NOT COLUMN BY COLUMN: work through the page one complete physical row at a time — read every field for that row, left to right, output that item, THEN move down to the next row. Do NOT read all the way down one column first (e.g. every GSM value top to bottom) and then move to the next column — that reading pattern is exactly what causes a single field to silently drift into the row above or below it, especially when a row has a blank/missing cell in one column and the columns fall out of sync with each other. Every field inside one item MUST come from that exact same physical row — never carry, shift, or borrow a single value from a neighboring row, even if a neighboring row\'s value seems like it would fit better or look more consistent. Before finalizing your answer, re-check each row against the image one more time and confirm every value in it — not just the row count — genuinely belongs to that physical line.' },
          ],
        }],
      }),
    });
    // Pass Retry-After through explicitly since the frontend's backoff logic reads it.
    const retryAfter = upstream.headers.get('retry-after');
    if (retryAfter) res.setHeader('retry-after', retryAfter);
    const bodyText = await upstream.text();
    res.status(upstream.status);
    res.setHeader('Content-Type', 'application/json');
    res.send(bodyText);
  } catch (e) {
    console.error('Anthropic proxy error:', e);
    res.status(502).json({ error: `Could not reach Anthropic: ${e.message || 'unknown error'}` });
  }
}

module.exports = { extract };
