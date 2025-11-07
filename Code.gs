/* ===========================================================================
   CALCULADORA HIPOTECARIA – Bayteca (Apps Script)
   Versión 2025-08 · GS con Euríbor desde EuriborDiario.es
   - getEuriborActual(): lee el valor (indicativo) desde https://www.euribordiario.es/
     y, si no hay dato, cae al oficial del BdE (media mensual) como respaldo.
   - Mantiene: limpieza de bancos vacíos, logging, envío Zapier.
   ========================================================================== */

/* ───────────── CONFIGURACIÓN BÁSICA ───────────── */
const SHEET_ID = '1gjqW5zQ_6HtpdRZKmHxIqC0sodydXlUKemZYF4-mkJs';
const ZAP_URL  = 'https://hooks.zapier.com/hooks/catch/15130242/u30cbdk/';
const EXTRAER_HOOK_URL = 'https://hook.eu2.make.com/8griym530n3hd6v5f3bivv8vxyeg5irl';

const CONSULTOR_SHEET_NAME = 'Consultores';
const CONSULTOR_NOMBRES_ESTATICOS = {
  '17835825': 'Javier Domínguez',
  '18267388': 'Ana Negrete',
  '18753280': 'Gabriel Mendez',
  '19799325': 'Alejandro Martinez',
  '21416380': 'Fernando Bermúdez',
  '21416391': 'Jose Garcia',
  '22003857': 'Soraya',
  '22224550': 'Diego Sanz',
  '22321834': 'Jose Ortega',
  '22125000': 'Ismael Gomez',
  '22125011': 'Sara Garcia',
  '22592599': 'Susana de Armas',
  '22592588': 'Carlos Hidalgo',
  '23116573': 'Kristian Zlatkov',
  '23222866': 'Borja Señor',
  '23275941': 'Javier Bartolome',
  '23341457': 'Mario Esquer',
  '23612651': 'Marilena Cabrera',
  '23665011': 'Maria Jose Bita',
  '23665022': 'Veronica Diaz',
  '23750074': 'Gaston Murray',
  '23750063': 'Raul Fuentes',
  '23953750': 'Daniel Fagil',
  '23953761': 'Adriana Gaitan',
  '24070889': 'Macarena Lauro',
  '25300634': 'Enrique Wazzan',
  '25266842': 'David Blazquez',
  '16884622': 'Paco del Río',
  '26212017': 'Jaime Piñar',
  '26709987': 'Osmel Montiel',
  '26709976': 'Bruno Estaun',
  '26850314': 'Jennifer Vera',
  '27199399': 'Eva Gómez',
  '27081677': 'Laura',
  '27338340': 'Matias Burgueño',
  '26955342': 'Andrea Garzón',
  '27081666': 'Cristina Castellar',
};

// Serie BdE (media mensual oficial) usada sólo como Fallback
const SERIE_BDE_EURIBOR12M = 'D_1NBAF472';

/* ───────────── FUNCIÓN doGet – sirve el HTML ───────────── */
function doGet(e) {
  const tpl     = HtmlService.createTemplateFromFile('index');
  tpl.consultor = (e && e.parameter && e.parameter.consultor) ? e.parameter.consultor : 'Unknown';
  tpl.consultorNombre = lookupConsultorNombre_(tpl.consultor);
  return tpl
    .evaluate()
    .setTitle('Calculadora Hipotecaria · Bayteca')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ───────────── Hoja de registro / crea si falta ───────────── */
function logSheet_() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const name = 'Envios';
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* ───────────── Util: sanea payload antes de loguear/enviar ───────────── */
function sanitizePayload_(payload){
  const out = Object.assign({}, payload);
  ['bancoCliente','entidadesEnviar1','entidadesEnviar2','entidadesEnviar3']
    .forEach(k => {
      if (out[k] == null) return;
      const v = String(out[k]).trim();
      if (!v || v === '-' || /no\s*bank\s*selected/i.test(v)) delete out[k];
    });
  Object.keys(out).forEach(k => { if (typeof out[k] === 'string') out[k] = out[k].trim(); });
  return out;
}

/* ───────────── Recibe payload, valida, registra y reenvía ───────────── */
function enviarACRM(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Payload inválido');

  // 1) Sanea payload (quita "No bank selected", trims, etc.)
  const clean = sanitizePayload_(payload);

  // 2) Validación de obligatorios (igual que antes)
  const oblig = ['dealId','tipoOperacion','tipoHipoteca','importePropiedad','importeHipoteca','consultor'];
  const vacios = oblig.filter(k => !clean[k]);
  if (vacios.length) throw new Error('Campos obligatorios vacíos: ' + vacios.join(', '));

  // 3) Normaliza y valida porcentaje de gastos (0–100).
  //    Nota: lo usamos para consistencia del dato recibido, pero NO lo enviaremos ni guardaremos.
  if (clean.porcGastos != null && clean.porcGastos !== '') {
    var pg = parseFloat(String(clean.porcGastos).replace(',','.'));
    if (isFinite(pg)) {
      pg = Math.max(0, Math.min(pg, 100));
      clean.porcGastos = pg.toFixed(2); // guarda como "10.00" en el objeto temporal
    } else {
      delete clean.porcGastos;
    }
  }

  // 4) Quitar 'porcGastos' del objeto final (no registrar ni enviar al CRM)
  if ('porcGastos' in clean) delete clean.porcGastos;

  // 5) Registrar en la hoja
  const sh  = logSheet_();
  sh.appendRow([ new Date(), clean.dealId, clean.consultor, JSON.stringify(clean) ]);

  // 6) Enviar a Zapier (CRM)
  const resp = UrlFetchApp.fetch(ZAP_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(clean),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('Zapier respondió ' + code + ': ' + resp.getContentText());
  }
}

function extraerDatos(dealId) {
  const id = dealId == null ? '' : String(dealId).trim();
  if (!id) throw new Error('Deal ID obligatorio');

  const resp = UrlFetchApp.fetch(EXTRAER_HOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ dealId: id }),
    muteHttpExceptions: true,
    headers: { Accept: 'application/json' }
  });

  const status = resp.getResponseCode();
  const text = resp.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('Hook de extracción respondió ' + status + ': ' + text);
  }

  let parsed = null;
  try { parsed = text ? JSON.parse(text) : null; } catch (e) { parsed = null; }

  // 1) “Desenvolver” la estructura de Make
  let normalized = unwrapMake_(parsed) || {};

  // 2) Quitar las comillas raras de las claves (lo que te sale como "\"dealId\"")
  normalized = dequoteKeys_(normalized);

  return {
    dealId: id,
    data: normalized,   // ← ahora con claves normales: dealId, nombre1, etc.
    raw: text,
    httpStatus: status
  };
}

function extraerDatosDebug(dealId) {
  const id = dealId == null ? '' : String(dealId).trim();
  if (!id) throw new Error('Deal ID obligatorio');

  const resp = UrlFetchApp.fetch(EXTRAER_HOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ dealId: id }),
    muteHttpExceptions: true,
    headers: { Accept: 'application/json' }
  });

  const status = resp.getResponseCode();
  const text   = resp.getContentText();

  let parsed = null, parseError = null;
  try { parsed = text ? JSON.parse(text) : null; } catch (e) { parseError = String(e); }

  let unwrapped = null, unwrapError = null;
  try { unwrapped = unwrapMake_(parsed || text); } catch (e) { unwrapError = String(e); }

  // ← clave: limpia las comillas de las claves también en el debug
  unwrapped = dequoteKeys_(unwrapped);

  return {
    dealId: id,
    httpStatus: status,
    size: text ? text.length : 0,
    rawPreview: text ? text.slice(0, 1200) : '',
    parseError,
    parsedType: parsed ? Object.prototype.toString.call(parsed) : null,
    parsedKeys: (parsed && typeof parsed === 'object') ? Object.keys(parsed) : null,
    unwrapError,
    unwrappedType: unwrapped ? Object.prototype.toString.call(unwrapped) : null,
    unwrappedKeys: (unwrapped && typeof unwrapped === 'object') ? Object.keys(unwrapped) : null,
    data: unwrapped,
    raw: text
  };
}



/**
 * Intenta “desenvolver” la respuesta típica de Make hasta llegar al objeto útil.
 * Soporta:
 *  - { body: { ... } }
 *  - { Body: { ... } }
 *  - { data: { body: { ... } } }
 *  - [ { body: { ... } } ]  (arrays/bundles)
 *  - objetos directos con los campos finales
 */
function unwrapMake_(x) {
  if (!x) return null;

  // Si x es JSON en string, intenta parsearlo
  if (typeof x === 'string') {
    try { x = JSON.parse(x); } catch(e) { /* keep as string */ }
  }

  // Si viene como array (bundle 1, bundle 2...)
  if (Array.isArray(x)) {
    for (var i = 0; i < x.length; i++) {
      var got = unwrapMake_(x[i]);
      if (got && typeof got === 'object' && Object.keys(got).length) return got;
    }
    return null;
  }

  // Accesos directos típicos (y si body/Body son string, parsea)
  if (x && typeof x === 'object') {
    if (x.body != null) {
      if (typeof x.body === 'string') {
        try { return unwrapMake_(JSON.parse(x.body)); } catch(e) { /* ignore */ }
      }
      if (typeof x.body === 'object') return unwrapMake_(x.body);
    }
    if (x.Body != null) {
      if (typeof x.Body === 'string') {
        try { return unwrapMake_(JSON.parse(x.Body)); } catch(e) { /* ignore */ }
      }
      if (typeof x.Body === 'object') return unwrapMake_(x.Body);
    }
    if (x.data != null) {
      return unwrapMake_(x.data); // recursivo: data puede contener body string/objeto
    }
  }

  // Si llega aquí y es objeto plano, úsalo tal cual
  return (x && typeof x === 'object') ? x : null;
}

/** Quita comillas literales de las claves: {"\"dealId\"": "123"} -> {"dealId": "123"}  */
function dequoteKeys_(x) {
  if (Array.isArray(x)) {
    return x.map(dequoteKeys_);
  }
  if (x && typeof x === 'object') {
    var out = {};
    Object.keys(x).forEach(function (k) {
      var nk = String(k).replace(/^"+|"+$/g, ''); // quita comillas al principio/fin
      out[nk] = dequoteKeys_(x[k]);
    });
    return out;
  }
  return x;
}



/* ───────────── UTILIDAD: incluir HTML parciales ───────────── */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function lookupConsultorNombre_(numero) {
  const key = numero == null ? '' : String(numero).trim();
  if (!key) return '';

  if (Object.prototype.hasOwnProperty.call(CONSULTOR_NOMBRES_ESTATICOS, key)) {
    const nombreEstatico = CONSULTOR_NOMBRES_ESTATICOS[key];
    if (nombreEstatico == null) return '';
    return String(nombreEstatico).trim();
  }

  if (!/\d/.test(key)) return '';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(CONSULTOR_SHEET_NAME);
    if (!sh) return '';
    const lastRow = sh.getLastRow();
    if (!lastRow) return '';
    if (sh.getLastColumn() < 2) return '';

    const values = sh.getRange(1, 1, lastRow, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const fila = values[i];
      if (!fila || fila.length < 2) continue;
      const numeroHoja = fila[0];
      const nombreHoja = fila[1];
      if (numeroHoja == null || nombreHoja == null) continue;
      if (String(numeroHoja).trim() === key) {
        return String(nombreHoja).trim();
      }
    }
  } catch (err) {
    Logger.log('lookupConsultorNombre_ error: %s', err && err.message);
  }

  return '';
}

function getConsultorNombre(numero) {
  return lookupConsultorNombre_(numero);
}

/* =======================================================================
   EURÍBOR 12M – Servicio para el front (lee EuriborDiario.es)
   ======================================================================= */

/** Devuelve { valor, fechaISO, fuente, oficial }  */
function getEuriborActual(){
  // 1) Intentar EuriborDiario.es (página específica del día)
  try{
    const d = fetchFromEuriborDiario_();
    if (d && isFinite(d.valor)) return d; // indicativo (diario o media mensual provisional)
  }catch(e){ Logger.log('EuriborDiario error: %s', e && e.message); }

  // 2) Fallback: BdE (media mensual, oficial)
  try{
    const o = fetchFromBde_();
    if (o && isFinite(o.valor)) return o;
  }catch(e){ Logger.log('BdE error: %s', e && e.message); }

  // 3) Sin datos
  return { valor: null, fechaISO: null, fuente: 'sin datos', oficial: false };
}

/* ======= EuriborDiario.es =======
   Intenta primero /euribor-hoy. Si no, la home. Devuelve {valor, fechaISO, fuente, oficial:false}
*/
function fetchFromEuriborDiario_(){
  const urls = [
    'https://www.euribordiario.es/euribor-hoy',
    'https://www.euribordiario.es/'
  ];
  const uaHeaders = { 'User-Agent':'Mozilla/5.0', 'Accept-Language':'es-ES,es;q=0.9' };

  for (var i=0;i<urls.length;i++){
    const res = UrlFetchApp.fetch(urls[i], { muteHttpExceptions:true, followRedirects:true, headers: uaHeaders });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) continue;
    const html = res.getContentText();

    // 1) Buscar una mención explícita a "Euríbor 12 meses" seguido de un %
    //    o "Euríbor hoy" con un %.
    var m = html.match(/Eur[ií]bor[^\n]{0,80}?12[^%]*?([0-9]{1,2}[\.,][0-9]{3})\s*%/i)
          || html.match(/Eur[ií]bor\s+hoy[^%]*?([0-9]{1,2}[\.,][0-9]{3})\s*%/i)
          || html.match(/media\s+del\s+Eur[ií]bor[^\d]*([0-9]{1,2}[\.,]\d{3})\s*%/i);
    if (!m) continue;

    const valor = parsePct_(m[1]);
    if (!isValidEuribor_(valor)) continue; // evita 1884% y similares

    // 2) Fecha: "Publicado el día 12 de agosto de 2025" | "12/08/2025" | "Hoy, 12 de agosto de 2025"
    var f = html.match(/(\d{1,2}\/\d{1,2}\/\d{4})/)
          || html.match(/(?:Publicado\s+el\s+d[ií]a|Hoy,)\s*([0-9]{1,2}\s+de\s+[A-Za-zñÑ]+\s+de\s+[0-9]{4})/i);
    const fechaISO = f ? (f[1].includes('/') ? toISO_(parseFechaES_(f[1])) : toISO_((parseFechaESLargo_(f[1])))) : new Date().toISOString();

    return { valor: valor, fechaISO: fechaISO, fuente: 'EuriborDiario.es', oficial: false };
  }
  throw new Error('No se pudo extraer dato de EuriborDiario.es');
}

/* ======= BdE (fallback oficial media mensual) ======= */
function fetchFromBde_(){
  const url = 'https://app.bde.es/bie_rest/resources/series/' + encodeURIComponent(SERIE_BDE_EURIBOR12M) + '/datos?last=1';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions:true });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('HTTP '+code+' en BdE');
  const json = JSON.parse(res.getContentText());
  if (!Array.isArray(json) || !json.length) throw new Error('BdE sin datos');
  const o = json[0];
  const valor = parsePct_(o.valor);
  const fechaISO = new Date(o.fecha).toISOString();
  if (!isValidEuribor_(valor)) throw new Error('BdE valor fuera de rango');
  return { valor: valor, fechaISO: fechaISO, fuente: 'Banco de España (API)', oficial: true };
}

/* ======= Helpers ======= */
function parsePct_(s){
  if (!s) return null;
  var t = String(s).trim().replace(/[%\s]+/g,'');
  var hasDot = t.indexOf('.')>-1, hasComma = t.indexOf(',')>-1;
  var dec = '.';
  if (hasDot && hasComma) dec = (t.lastIndexOf('.')>t.lastIndexOf(',')) ? '.' : ',';
  else if (hasComma && !hasDot) dec = ',';
  var norm = (dec===',') ? t.replace(/\./g,'').replace(',', '.') : t.replace(/,/g,'');
  var v = parseFloat(norm);
  return isFinite(v) ? v : null;
}

function isValidEuribor_(v){ return typeof v==='number' && v>-5 && v<15; }

function parseFechaES_(txt){ // dd/mm/yyyy
  var [d,m,y] = txt.split('/');
  return new Date(parseInt(y,10), parseInt(m,10)-1, parseInt(d,10));
}
function parseFechaESLargo_(txt){ // 12 de agosto de 2025
  var meses = {enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,setiembre:8,octubre:9,noviembre:10,diciembre:11};
  var m = (txt||'').toLowerCase().match(/(\d{1,2})\s+de\s+([a-zñ]+)\s+de\s+(\d{4})/);
  if (!m) return new Date();
  return new Date(parseInt(m[3],10), meses[m[2]]||0, parseInt(m[1],10));
}

function toISO_(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'"); }
