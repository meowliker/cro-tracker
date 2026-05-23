const SCRIPT_URL = process.env.GOOGLE_APPS_SCRIPT_URL;

function sendJson(res, status, data) {
  res.setHeader('Content-Type', 'application/json');
  res.status(status).send(JSON.stringify(data));
}

async function readBody(req) {
  if (req.body) {
    return typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
  }

  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
    });
    req.on('end', () => resolve(body));
    req.on('error', reject);
  });
}

module.exports = async function handler(req, res) {
  res.setHeader('Content-Type', 'application/json');

  if (!SCRIPT_URL) {
    sendJson(res, 500, {
      ok: false,
      error: 'Missing GOOGLE_APPS_SCRIPT_URL environment variable.',
    });
    return;
  }

  try {
    if (req.method === 'GET') {
      const upstream = await fetch(`${SCRIPT_URL}?action=getData`);
      const text = await upstream.text();
      res.status(upstream.status).send(text);
      return;
    }

    if (req.method === 'POST') {
      const body = await readBody(req);
      const upstream = await fetch(SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body,
      });
      const text = await upstream.text();
      res.status(upstream.status).send(text);
      return;
    }

    res.setHeader('Allow', 'GET, POST');
    sendJson(res, 405, { ok: false, error: 'Method not allowed.' });
  } catch (error) {
    sendJson(res, 502, {
      ok: false,
      error: error && error.message ? error.message : String(error),
    });
  }
};
