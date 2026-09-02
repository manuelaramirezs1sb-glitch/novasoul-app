/**
 * Netlify serverless proxy — Nova Empresarial
 *
 * RUTAS:
 *   GET  /?tab=DASHBOARD-DUENO   → lee tab normalizada (solo DASHBOARD-*)
 *   POST /write                  → escribe acción de panel (estado, nota, etc.)
 *   POST /import                 → procesa archivo subido y escribe en hoja normalizada
 *
 * Variables de entorno requeridas:
 *   SHEETS_SERVICE_ACCOUNT_EMAIL
 *   SHEETS_SERVICE_ACCOUNT_KEY   (private key PEM, \n literal)
 *   SHEET_ID_DUENO               (ID del spreadsheet Centro de Comando)
 *   SHEET_ID_GESTOR              (ID del spreadsheet Gestor, si aplica)
 *
 * Seguridad:
 *   - La app nunca lee hojas crudas. Solo DASHBOARD-* y tabs de lectura permitidas.
 *   - Escrituras van solo a columnas permitidas (no sobreescriben columnas de import).
 *   - Cada escritura registra en Movimientos: usuario, campo, valor anterior/nuevo.
 */

const https  = require('https');
const crypto = require('crypto');

// Tabs que la app puede LEER (whitelist estricta — nunca hojas crudas)
const READABLE_TABS = [
  // Nova Empresarial
  'DASHBOARD-DUENO', 'DASHBOARD-ADMIN', 'DASHBOARD-GESTORA',
  // Nova Central
  'DASHBOARD-CENTRAL',
  // NovaSoul
  'DASHBOARD-SOUL',
  // novAcademy
  'DASHBOARD-ACADEMY-PROFESOR', 'DASHBOARD-ACADEMY-ESTUDIANTE',
];

// Columnas que la app puede ESCRIBIR por hoja (no sobreescriben datos importados)
const WRITABLE_COLS = {
  // Nova Empresarial
  Pedidos:    ['estado','gestora_asignada','guia','direccion','fecha_promesa','actualizado_en','actualizado_por'],
  Novedades:  ['estado','gestora','nota','intentos','resuelta_en','actualizado_en'],
  Equipo:     ['casos_asignados','casos_resueltos','nota_auditoria','ultima_conexion'],
  Inventario: ['stock','costo_unitario','precio','ultimo_conteo'],
  Cartera:    ['gestora_asignada','notas','dias_sin_contacto','ultimo_contacto'],
  Parametros: ['valor','actualizado_en','actualizado_por'],
  // Nova Central
  Clientes:   ['plan','tarifa','estado','ultimo_pago','fecha_corte'],
  Solicitudes:['estado','resuelta_en','resuelta_por'],
  Candidatas: ['estado','nota'],
  // NovaSoul — toda la data es escritura de la plataforma
  Pendientes: ['texto','tipo','origen','fecha','hecho','hecho_en','plataforma_id'],
  Dias:       ['comidas_marcadas','movimiento_hecho','puntos','cerrado','cerrado_en','perdonado'],
  Recompensas:['canjeada','canjeada_en'],
  // novAcademy
  Estudiantes:['perfil','estado','acceso_vence','primer_ingreso'],
  Permisos:   ['ver_dinero','descargar_plantillas','ver_avance_equipo','emitir_certificado','entrar_practica'],
  Progreso:   ['estado','completada_en','nota_quiz','intentos','ultima_conexion'],
  Ejercicios: ['comentario_profesor','comentado_en'],
  Certificados:['emitido_en','vence','url_pdf'],
};

// Sheets normalizados donde Netlify puede escribir (importación)
const IMPORTABLE_SHEETS = ['Pedidos','Novedades','Pauta','Inventario','Equipo'];

// ── Handler principal ─────────────────────────────────────────────────────────

exports.handler = async (event) => {
  const origin = event.headers['origin'] || '';
  const corsHeaders = {
    'Access-Control-Allow-Origin': allowedOrigin(origin),
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Nova-User',
    'Cache-Control': 'no-store',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: corsHeaders, body: '' };

  const path = (event.path || '').replace(/^\/\.netlify\/functions\/sheets/, '');

  try {
    if (event.httpMethod === 'GET') return await handleRead(event, corsHeaders);
    if (event.httpMethod === 'POST' && path === '/write') return await handleWrite(event, corsHeaders);
    if (event.httpMethod === 'POST' && path === '/import') return await handleImport(event, corsHeaders);
    return json(405, { error: 'Método no permitido' }, corsHeaders);
  } catch (err) {
    console.error('sheets handler error:', err.message);
    return json(502, { error: 'Error interno', detail: err.message }, corsHeaders);
  }
};

// ── READ: leer tab DASHBOARD-* ────────────────────────────────────────────────

async function handleRead(event, corsHeaders) {
  const { tab, sheet } = event.queryStringParameters || {};
  if (!tab || !READABLE_TABS.includes(tab)) {
    return json(400, { error: 'Tab no permitida. Permitidas: ' + READABLE_TABS.join(', ') }, corsHeaders);
  }
  const sheetId = getSheetId(sheet);
  const token   = await getAccessToken();
  const rows    = await readSheetTab(sheetId, tab, token);
  return json(200, { tab, rows }, corsHeaders);
}

// ── WRITE: acción de panel (cambio de estado, nota, asignación) ───────────────

async function handleWrite(event, corsHeaders) {
  let body;
  try { body = JSON.parse(event.body || '{}'); } catch { return json(400, { error: 'JSON inválido' }, corsHeaders); }

  const { hoja, id_col, id_val, campos, usuario, sheet } = body;

  if (!hoja || !WRITABLE_COLS[hoja]) {
    return json(400, { error: 'Hoja no permitida para escritura: ' + hoja }, corsHeaders);
  }
  const colsPermitidas = WRITABLE_COLS[hoja];
  const camposValidos  = Object.keys(campos || {}).filter(c => colsPermitidas.includes(c));
  if (camposValidos.length === 0) {
    return json(400, { error: 'Sin columnas válidas para escribir' }, corsHeaders);
  }

  const sheetId = getSheetId(sheet);
  const token   = await getAccessToken();

  // Leer la hoja para encontrar la fila por id_col/id_val
  const rows = await readSheetTab(sheetId, hoja, token);
  if (!rows || rows.length < 2) return json(404, { error: 'Hoja vacía o sin datos' }, corsHeaders);

  const enc      = rows[0];
  const idColIdx = enc.indexOf(id_col);
  if (idColIdx < 0) return json(400, { error: 'Columna ' + id_col + ' no encontrada' }, corsHeaders);

  const rowIdx = rows.slice(1).findIndex(r => String(r[idColIdx]) === String(id_val));
  if (rowIdx < 0) return json(404, { error: 'Fila no encontrada: ' + id_col + '=' + id_val }, corsHeaders);

  const sheetRowNum = rowIdx + 2; // +1 header, +1 base-1
  const ts = new Date().toISOString().replace('T',' ').substring(0,19);

  // Construir las actualizaciones (solo columnas permitidas)
  const updates = [];
  const movimientos = [];

  camposValidos.forEach(col => {
    const colIdx = enc.indexOf(col);
    if (colIdx < 0) return;
    const valAnterior = rows[sheetRowNum - 1][colIdx];
    const valNuevo    = campos[col];
    updates.push({
      range: `${hoja}!${colLetter(colIdx + 1)}${sheetRowNum}`,
      values: [[valNuevo]],
    });
    movimientos.push({ campo: col, valAnterior, valNuevo });
  });

  // Agregar actualizado_en y actualizado_por si la hoja los tiene
  if (enc.includes('actualizado_en')) {
    const ci = enc.indexOf('actualizado_en');
    updates.push({ range: `${hoja}!${colLetter(ci+1)}${sheetRowNum}`, values: [[ts]] });
  }
  if (enc.includes('actualizado_por') && usuario) {
    const ci = enc.indexOf('actualizado_por');
    updates.push({ range: `${hoja}!${colLetter(ci+1)}${sheetRowNum}`, values: [[usuario]] });
  }

  await batchUpdate(sheetId, updates, token);

  // Registrar en Movimientos
  await registrarMovimientos(sheetId, id_val, hoja, movimientos, usuario, ts, token);

  return json(200, { ok: true, fila: sheetRowNum, cambios: camposValidos.length }, corsHeaders);
}

// ── IMPORT: procesar archivo subido y escribir en hoja normalizada ─────────────

async function handleImport(event, corsHeaders) {
  let body;
  try { body = JSON.parse(event.body || '{}'); } catch { return json(400, { error: 'JSON inválido' }, corsHeaders); }

  const { hoja, plataforma, filas, tienda, usuario, sheet } = body;
  // filas: array de objetos ya normalizados por el cliente antes de enviar

  if (!hoja || !IMPORTABLE_SHEETS.includes(hoja)) {
    return json(400, { error: 'Hoja no permitida para importación: ' + hoja }, corsHeaders);
  }
  if (!Array.isArray(filas) || filas.length === 0) {
    return json(400, { error: 'Sin filas para importar' }, corsHeaders);
  }

  const sheetId = getSheetId(sheet);
  const token   = await getAccessToken();

  // Leer headers de la hoja destino
  const rows = await readSheetTab(sheetId, hoja, token);
  const enc  = rows && rows[0] ? rows[0] : [];

  const ts = new Date().toISOString().replace('T',' ').substring(0,19);

  // Convertir cada objeto a fila usando el orden de headers de la hoja
  const filasEscritas = filas.map(obj => {
    return enc.map(col => {
      if (col === 'tienda' && !obj[col]) return tienda || '';
      if (col === 'actualizado_en')       return ts;
      if (col === 'actualizado_por')      return usuario || 'import';
      return obj[col] !== undefined ? obj[col] : '';
    });
  });

  // Encontrar la primera fila vacía
  const primeraVacia = rows.length + 1; // +1 porque rows incluye header
  await writeRows(sheetId, hoja, primeraVacia, filasEscritas, token);

  return json(200, {
    ok: true,
    hoja,
    plataforma,
    filas_importadas: filasEscritas.length,
    primera_fila: primeraVacia,
  }, corsHeaders);
}

// ── Google Sheets API ─────────────────────────────────────────────────────────

async function readSheetTab(sheetId, tabName, token) {
  const range = encodeURIComponent(`${tabName}!A1:ZZ`);
  const url   = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}`;
  const raw   = await httpsGet(url, { Authorization: `Bearer ${token}` });
  const json_ = JSON.parse(raw);
  return json_.values || [];
}

async function batchUpdate(sheetId, updates, token) {
  const url  = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`;
  const body = JSON.stringify({
    valueInputOption: 'USER_ENTERED',
    data: updates,
  });
  await httpsPost(url, body, {
    Authorization:  `Bearer ${token}`,
    'Content-Type': 'application/json',
  });
}

async function writeRows(sheetId, tabName, startRow, rows, token) {
  const enc = rows[0] ? rows[0].map((_,i) => colLetter(i+1)) : [];
  const lastCol = enc[enc.length - 1] || 'Z';
  const range = encodeURIComponent(`${tabName}!A${startRow}:${lastCol}${startRow + rows.length - 1}`);
  const url   = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?valueInputOption=USER_ENTERED`;
  const body  = JSON.stringify({ values: rows });
  await httpsPost(url, body, {
    Authorization:  `Bearer ${token}`,
    'Content-Type': 'application/json',
  }, 'PUT');
}

async function registrarMovimientos(sheetId, entidadId, hoja, movimientos, usuario, ts, token) {
  if (!movimientos || movimientos.length === 0) return;
  const filas = movimientos.map(m => [ts, usuario||'', hoja, entidadId, m.campo, m.valAnterior, m.valNuevo]);
  try {
    const rows = await readSheetTab(sheetId, 'Movimientos', token);
    const nextRow = rows.length + 1;
    await writeRows(sheetId, 'Movimientos', nextRow, filas, token);
  } catch (e) {
    console.error('registrarMovimientos error (no crítico):', e.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function getSheetId(sheet) {
  return sheet === 'gestor' ? process.env.SHEET_ID_GESTOR : process.env.SHEET_ID_DUENO;
}

function colLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(65 + (n-1) % 26) + s; n = Math.floor((n-1) / 26); }
  return s;
}

function allowedOrigin(origin) {
  if (!origin) return 'null';
  if (origin.includes('netlify.app') || origin.includes('localhost') ||
      origin.includes('nova') || origin.includes('127.0.0.1')) return origin;
  return 'null';
}

function json(status, data, headers) {
  return { statusCode: status, headers: { ...headers, 'Content-Type': 'application/json' },
           body: JSON.stringify(data) };
}

// ── JWT / OAuth2 ──────────────────────────────────────────────────────────────

async function getAccessToken() {
  const email  = process.env.SHEETS_SERVICE_ACCOUNT_EMAIL;
  const rawKey = (process.env.SHEETS_SERVICE_ACCOUNT_KEY || '').replace(/\\n/g, '\n');
  if (!email || !rawKey) throw new Error('Service account no configurado en env vars');

  const now  = Math.floor(Date.now() / 1000);
  const hdr  = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const clm  = b64url(JSON.stringify({
    iss: email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600, iat: now,
  }));
  const unsigned = `${hdr}.${clm}`;
  const key  = crypto.createPrivateKey(rawKey);
  const sig  = b64url(crypto.sign('sha256', Buffer.from(unsigned), key));
  const jwt  = `${unsigned}.${sig}`;

  const body = `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`;
  const raw  = await httpsPost('https://oauth2.googleapis.com/token', body,
    { 'Content-Type': 'application/x-www-form-urlencoded' });
  const res  = JSON.parse(raw);
  if (!res.access_token) throw new Error('OAuth error: ' + JSON.stringify(res));
  return res.access_token;
}

function b64url(str) {
  const buf = Buffer.isBuffer(str) ? str : Buffer.from(str);
  return buf.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
}

// ── HTTP helpers ──────────────────────────────────────────────────────────────

function httpsGet(url, headers = {}) {
  return new Promise((res, rej) => {
    const req = https.get(url, { headers }, r => {
      let b = ''; r.on('data', c => b += c); r.on('end', () => res(b));
    });
    req.on('error', rej);
    req.setTimeout(12000, () => { req.destroy(); rej(new Error('timeout')); });
  });
}

function httpsPost(url, body, headers = {}, method = 'POST') {
  return new Promise((res, rej) => {
    const u    = new URL(url);
    const opts = {
      hostname: u.hostname, path: u.pathname + u.search, method,
      headers: { ...headers, 'Content-Length': Buffer.byteLength(body) },
    };
    const req = https.request(opts, r => {
      let d = ''; r.on('data', c => d += c); r.on('end', () => res(d));
    });
    req.on('error', rej);
    req.setTimeout(12000, () => { req.destroy(); rej(new Error('timeout')); });
    req.write(body); req.end();
  });
}
