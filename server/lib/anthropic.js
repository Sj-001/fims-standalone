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
        max_tokens: 4096,
        system: systemPrompt,
        messages: [{
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: base64Image } },
            { type: 'text', text: 'Extract the data from this document image now. Return ONLY the JSON described in your instructions — no markdown fences, no commentary, no explanation. Be concise: do not repeat header-level values inside each item unless asked to. CRITICAL: extract each physical row exactly once, in the same top-to-bottom order it appears — do not output the same row twice, even if faint printing, ruling lines, or a neighboring similar-looking row makes it seem repeated. The number of items you return should equal the number of actual handwritten/printed lines on the page, no more, no less.' },
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
