// Netlify serverless proxy — Nova Empresarial
// La app NUNCA lee hojas crudas. Solo lee pestañas DASHBOARD-* calculadas.
// Requiere variables de entorno:
//   SHEETS_SERVICE_ACCOUNT_EMAIL
//   SHEETS_SERVICE_ACCOUNT_KEY   (private key PEM, \n literal)
//   SHEET_ID_DUENO               (ID del spreadsheet Centro de Comando)
//   SHEET_ID_GESTOR              (ID del spreadsheet Gestor, opcional)

const https = require('https');
const crypto = require('crypto');

// Tabs que la app tiene permitido leer (whitelist estricta)
const ALLOWED_TABS = [
  'DASHBOARD-DUENO',
  'DASHBOARD-GESTOR',
  'DASHBOARD-ADMIN',
  'ENTREGA-DUENO-DIARIO',
  'ENTREGA-DUENO-MENSUAL',
];

exports.handler = async (event) => {
  const origin = event.headers['origin'] || '';
  const corsHeaders = {
    'Access-Control-Allow-Origin': origin.includes('netlify.app') || origin.includes('localhost')
      ? origin : 'null',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Cache-Control': 'no-store',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: corsHeaders, body: '' };
  }
  if (event.httpMethod !== 'GET') {
    return { statusCode: 405, headers: corsHeaders, body: 'Method Not Allowed' };
  }

  const { tab, sheet } = event.queryStringParameters || {};

  if (!tab || !ALLOWED_TABS.includes(tab)) {
    return {
      statusCode: 400,
      headers: corsHeaders,
      body: JSON.stringify({ error: `Tab no permitida. Permitidas: ${ALLOWED_TABS.join(', ')}` }),
    };
  }

  const sheetId = sheet === 'gestor'
    ? process.env.SHEET_ID_GESTOR
    : process.env.SHEET_ID_DUENO;

  if (!sheetId) {
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'SHEET_ID no configurado en variables de entorno' }),
    };
  }

  try {
    const token = await getAccessToken();
    const data = await readSheet(sheetId, tab, token);
    return {
      statusCode: 200,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({ tab, rows: data }),
    };
  } catch (err) {
    console.error('sheets proxy error:', err.message);
    return {
      statusCode: 502,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Error al leer Google Sheets', detail: err.message }),
    };
  }
};

// ── Google Sheets API ────────────────────────────────────────────────────────

async function readSheet(sheetId, tabName, token) {
  const range = encodeURIComponent(`${tabName}!A1:ZZ`);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}`;
  const raw = await httpsGet(url, { Authorization: `Bearer ${token}` });
  const json = JSON.parse(raw);
  return json.values || [];
}

// ── JWT / OAuth2 service account ─────────────────────────────────────────────

async function getAccessToken() {
  const email = process.env.SHEETS_SERVICE_ACCOUNT_EMAIL;
  const rawKey = (process.env.SHEETS_SERVICE_ACCOUNT_KEY || '').replace(/\\n/g, '\n');

  if (!email || !rawKey) {
    throw new Error('SHEETS_SERVICE_ACCOUNT_EMAIL o SHEETS_SERVICE_ACCOUNT_KEY no configurados');
  }

  const now = Math.floor(Date.now() / 1000);
  const header = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = b64url(JSON.stringify({
    iss: email,
    scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now,
  }));

  const unsigned = `${header}.${claim}`;
  const key = crypto.createPrivateKey(rawKey);
  const sig = b64url(crypto.sign('sha256', Buffer.from(unsigned), key));
  const jwt = `${unsigned}.${sig}`;

  const body = `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`;
  const raw = await httpsPost('https://oauth2.googleapis.com/token', body, {
    'Content-Type': 'application/x-www-form-urlencoded',
  });
  const json = JSON.parse(raw);
  if (!json.access_token) throw new Error(`OAuth error: ${JSON.stringify(json)}`);
  return json.access_token;
}

function b64url(str) {
  const buf = Buffer.isBuffer(str) ? str : Buffer.from(str);
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

// ── HTTP helpers ─────────────────────────────────────────────────────────────

function httpsGet(url, headers = {}) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers }, (res) => {
      let body = '';
      res.on('data', (c) => (body += c));
      res.on('end', () => resolve(body));
    });
    req.on('error', reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error('timeout')); });
  });
}

function httpsPost(url, body, headers = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const opts = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: 'POST',
      headers: { ...headers, 'Content-Length': Buffer.byteLength(body) },
    };
    const req = https.request(opts, (res) => {
      let data = '';
      res.on('data', (c) => (data += c));
      res.on('end', () => resolve(data));
    });
    req.on('error', reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error('timeout')); });
    req.write(body);
    req.end();
  });
}
