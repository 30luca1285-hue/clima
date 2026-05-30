// ============================================================
// CLIMA — Monitoraggio Netatmo su Google Sheets
// ============================================================
// SETUP RAPIDO:
//  1. Vai su https://dev.netatmo.com/apps/ → crea app
//  2. Imposta Redirect URI = URL del tuo web app (vedi SETUP nel menu)
//  3. Dal menu 🌡️ Clima → "1. Configura credenziali"
//  4. Dal menu 🌡️ Clima → "2. Autorizza con Netatmo"
//  5. Dal menu 🌡️ Clima → "3. Avvia aggiornamento automatico"
// ============================================================

const CFG = {
  DATA_SHEET:              'Dati',
  DASHBOARD_SHEET:         'Dashboard',
  DATA_SHEET_STUDIO:       'Dati_Studio',
  DASHBOARD_SHEET_STUDIO:  'Dashboard_Studio',
  INTERVAL_MIN:            10,           // minuti tra un fetch e l'altro
};

const API = {
  TOKEN:    'https://api.netatmo.com/oauth2/token',
  AUTH:     'https://api.netatmo.com/oauth2/authorize',
  STATIONS: 'https://api.netatmo.com/api/getstationsdata',
  MEASURE:  'https://api.netatmo.com/api/getmeasure',
};

// Restituisce config di una "stazione" — in realtà 2 sensori sullo stesso device Netatmo:
//   ESTERNO (default) = NAModule1 (modulo esterno T/H/Pioggia/Pressione) — quello legacy, sheet "Dati"/Foglio1.
//   STUDIO            = NAMain    (base interna stessa stazione, T/H interni stanza + Pressione + CO2) — nuovo tab.
// Entrambi condividono device_id, refresh token, credenziali.
function _stationCfg(stationKey) {
  const sk = (typeof stationKey === 'string' ? stationKey : 'ESTERNO').toUpperCase();
  const props = PropertiesService.getScriptProperties();
  const deviceId = props.getProperty('DEVICE_ID');

  if (sk === 'STUDIO') {
    return {
      key:           'STUDIO',
      isMain:        true,                    // legge device.dashboard_data (NAMain)
      deviceId:      deviceId,
      moduleId:      deviceId,                // per getmeasure su NAMain: module_id = device_id
      rainModuleId:  null,                    // NAMain non ha pluviometro
      deviceName:    props.getProperty('DEVICE_NAME') || 'Studio (interno)',
      dataSheetName: 'Dati_Studio',
      dashSheetName: 'Dashboard_Studio',
      rainCacheKey:  null,
      jsonCacheKey:  '_DATA_CACHE_STUDIO',
      hasFoglio1:    false,
    };
  }
  // ESTERNO (default, legacy) = NAModule1, sheet "Dati"/"Dashboard"/"Foglio1"
  return {
    key:           'ESTERNO',
    isMain:        false,
    deviceId:      deviceId,
    moduleId:      props.getProperty('MODULE_ID'),
    rainModuleId:  props.getProperty('RAIN_MODULE_ID'),
    deviceName:    props.getProperty('MODULE_NAME') || 'Esterno',
    dataSheetName: CFG.DATA_SHEET,
    dashSheetName: CFG.DASHBOARD_SHEET,
    rainCacheKey:  'RAIN_DAILY_CACHE',
    jsonCacheKey:  '_DATA_CACHE',
    hasFoglio1:    true,
  };
}

// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const studio = ui.createMenu('Studio (interno NAMain)')
    .addItem('Aggiorna ora (Studio)', 'fetchAndSaveDataStudio')
    .addItem('Aggiorna dashboard Studio', 'updateDashboardStudio')
    .addItem('Import 30gg storici Studio', 'importHistoricalDataStudio');

  ui.createMenu('🌡️ Clima')
    .addItem('0. Autorizza script (prima volta)', 'autorizzaScript')
    .addItem('1. Configura credenziali Netatmo', 'setupCredentials')
    .addItem('2. Autorizza con Netatmo', 'startAuth')
    .addItem('3. Avvia aggiornamento automatico', 'setupTrigger')
    .addSeparator()
    .addItem('Rileva stazioni Netatmo', 'rilevaStazioni')
    .addItem('Aggiorna dati ora (Esterno)', 'fetchAndSaveData')
    .addItem('Aggiorna solo dashboard (Esterno)', 'updateDashboard')
    .addItem('Archivia mese precedente in Foglio1', 'archiviaMesePrecedente')
    .addItem('Importa dati storici (API)', 'importHistoricalData')
    .addItem('Recupera pioggia mancante', 'recoverRainData')
    .addSeparator()
    .addSubMenu(studio)
    .addSeparator()
    .addItem('Stato sistema', 'showStatus')
    .addItem('Ferma aggiornamento automatico', 'removeTrigger')
    .addToUi();
}

// Wrapper menu Studio (interno NAMain)
function updateDashboardStudio() {
  updateDashboard('STUDIO');
}

// ─────────────────────────────────────────────────────────────
// STEP 1 — CREDENZIALI
// ─────────────────────────────────────────────────────────────

function setupCredentials() {
  const ui = SpreadsheetApp.getUi();

  let r = ui.prompt('Setup Netatmo (1/2)', 'Inserisci il CLIENT ID\n(da dev.netatmo.com → la tua app → Client ID):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const clientId = r.getResponseText().trim();
  if (!clientId) { ui.alert('Client ID vuoto, ripeti.'); return; }

  r = ui.prompt('Setup Netatmo (2/2)', 'Inserisci il CLIENT SECRET:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const clientSecret = r.getResponseText().trim();
  if (!clientSecret) { ui.alert('Client Secret vuoto, ripeti.'); return; }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('CLIENT_ID', clientId);
  props.setProperty('CLIENT_SECRET', clientSecret);

  ui.alert(
    '✅ Credenziali salvate!\n\n' +
    'Prossimo step: menu 🌡️ Clima → "2. Autorizza con Netatmo"\n\n' +
    'PRIMA però imposta il Redirect URI nella tua app Netatmo:\n' +
    getWebAppUrl()
  );
}

// ─────────────────────────────────────────────────────────────
// STEP 2 — OAUTH FLOW
// ─────────────────────────────────────────────────────────────

function startAuth() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('CLIENT_ID');

  if (!clientId) {
    ui.alert('Prima configura le credenziali (Step 1).');
    return;
  }

  const redirectUri = getWebAppUrl();
  const authUrl =
    API.AUTH +
    '?client_id='     + encodeURIComponent(clientId) +
    '&redirect_uri='  + encodeURIComponent(redirectUri) +
    '&scope=read_station' +
    '&response_type=code';

  ui.alert(
    'Autorizzazione Netatmo',
    'Visita questo URL nel browser, accedi con il tuo account Netatmo e concedi l\'accesso.\n\n' +
    'Il sistema si autorizzerà automaticamente.\n\n' +
    '👉 ' + authUrl,
    ui.ButtonSet.OK
  );
}

// doGet viene chiamato da Netatmo dopo l'autorizzazione e dal frontend per i dati
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.code)                   return handleOAuthCallback(params.code);
  if (params.action === 'getData')   return serveJsonData(params.station);
  if (params.action === 'listStations') return listStationsJson();
  if (params.action === 'clearCache') {
    _cacheInvalidate('STUDIO');
    _cacheInvalidate('AZIENDA');
    return ContentService.createTextOutput('cache cleared').setMimeType(ContentService.MimeType.TEXT);
  }

  // Frontend: index.html con tab Azienda/Studio
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Clima Olio Galluzzi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Espone le "stazioni" (sensori) configurate per il frontend.
// Stesso device Netatmo, due tab: Esterno (NAModule1) + Studio (NAMain).
function listStationsJson() {
  const out = [];
  ['ESTERNO', 'STUDIO'].forEach(sk => {
    const cfg = _stationCfg(sk);
    if (cfg.deviceId) out.push({ key: cfg.key, nome: cfg.deviceName });
  });
  return ContentService
    .createTextOutput(JSON.stringify({ stazioni: out }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Aggiorna la cache pioggia giornaliera nelle ScriptProperties (max 1 volta ogni 30 min).
// Chiamata da fetchAndSaveData dove auth e UrlFetch funzionano di sicuro.
// cacheKey opzionale: 'RAIN_DAILY_CACHE' (azienda, default) o 'RAIN_DAILY_CACHE_STUDIO'.
function updateRainDailyCacheIfNeeded(deviceId, rainModuleId, token, cacheKey) {
  const PROP_KEY = cacheKey || 'RAIN_DAILY_CACHE';
  const props    = PropertiesService.getScriptProperties();
  const existing = props.getProperty(PROP_KEY);
  if (existing) {
    try {
      const parsed = JSON.parse(existing);
      if (Date.now() - parsed.updated < 30 * 60 * 1000) return; // fresca di meno di 30 min
    } catch(_) {}
  }

  try {
    const now    = new Date();
    const cutoff = new Date(now.getTime() - 32 * 24 * 3600 * 1000);
    const params =
      'device_id='  + encodeURIComponent(deviceId) +
      '&module_id=' + encodeURIComponent(rainModuleId) +
      '&scale=1day' +
      '&type=sum_rain' +
      '&date_begin=' + Math.floor(cutoff.getTime() / 1000) +
      '&date_end='   + Math.floor(now.getTime() / 1000) +
      '&optimize=false&real_time=true';

    const resp = UrlFetchApp.fetch(API.MEASURE + '?' + params, {
      headers: { Authorization: 'Bearer ' + (token || getValidToken()) },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('updateRainDailyCache HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText());
      return;
    }

    const rawBody = JSON.parse(resp.getContentText()).body;
    const result  = {};

    if (Array.isArray(rawBody)) {
      rawBody.forEach(chunk => {
        const begTime  = chunk.beg_time;
        const stepTime = chunk.step_time || 86400;
        (chunk.value || []).forEach((val, i) => {
          const mm  = (val && val[0] != null) ? val[0] : 0;
          const d   = new Date((begTime + i * stepTime) * 1000);
          const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          result[key] = Math.round(mm * 10) / 10;
        });
      });
    } else if (rawBody && typeof rawBody === 'object') {
      Object.keys(rawBody).forEach(tsStr => {
        const d   = new Date(parseInt(tsStr) * 1000);
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const val = rawBody[tsStr];
        const mm  = Array.isArray(val) ? (val[0] != null ? val[0] : 0) : (typeof val === 'number' ? val : 0);
        result[key] = Math.round(mm * 10) / 10;
      });
    }

    props.setProperty(PROP_KEY, JSON.stringify({ updated: Date.now(), data: result }));
    Logger.log(PROP_KEY + ' aggiornata: ' + Object.keys(result).length + ' giorni');

  } catch(e) {
    Logger.log('updateRainDailyCacheIfNeeded error: ' + e);
  }
}

// Legge i totali giornalieri dalla cache ScriptProperties (precalcolata da fetchAndSaveData).
// stationKey: 'AZIENDA' (default) o 'STUDIO'.
function getRainDailyMap(stationKey) {
  const sk    = (stationKey || 'AZIENDA').toUpperCase();
  const propKey = sk === 'STUDIO' ? 'RAIN_DAILY_CACHE_STUDIO' : 'RAIN_DAILY_CACHE';
  const props    = PropertiesService.getScriptProperties();
  const existing = props.getProperty(propKey);
  if (!existing) return {};
  try {
    return JSON.parse(existing).data || {};
  } catch(_) {
    return {};
  }
}

// ─────────────────────────────────────────────────────────────
// CACHE JSON (CacheService, TTL 9 min, chunking per > 100 KB)
// ─────────────────────────────────────────────────────────────

const CACHE_KEY   = 'CLIMA_JSON';
const CACHE_CHUNK = 90000; // byte per chunk (limite CacheService: 100 KB)
const CACHE_TTL   = 540;   // secondi (9 minuti)

function _cacheKeyFor(stationKey) {
  const sk = (stationKey || 'AZIENDA').toUpperCase();
  return sk === 'STUDIO' ? CACHE_KEY + '_STUDIO' : CACHE_KEY;
}

function _cacheWrite(json, stationKey) {
  try {
    const base   = _cacheKeyFor(stationKey);
    const cache  = CacheService.getScriptCache();
    const n      = Math.ceil(json.length / CACHE_CHUNK);
    const obj    = {};
    obj[base + '_N'] = String(n);
    for (let i = 0; i < n; i++) {
      obj[base + '_' + i] = json.slice(i * CACHE_CHUNK, (i + 1) * CACHE_CHUNK);
    }
    cache.putAll(obj, CACHE_TTL);
  } catch(e) {
    Logger.log('_cacheWrite error: ' + e);
  }
}

function _cacheRead(stationKey) {
  try {
    const base  = _cacheKeyFor(stationKey);
    const cache = CacheService.getScriptCache();
    const nStr  = cache.get(base + '_N');
    if (!nStr) return null;
    const n     = parseInt(nStr);
    const parts = [];
    for (let i = 0; i < n; i++) {
      const part = cache.get(base + '_' + i);
      if (part === null) return null;
      parts.push(part);
    }
    return parts.join('');
  } catch(e) {
    Logger.log('_cacheRead error: ' + e);
    return null;
  }
}

function _cacheInvalidate(stationKey) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(_cacheKeyFor(stationKey) + '_N');
  } catch(e) {}
}

// ─────────────────────────────────────────────────────────────

function serveJsonData(stationKey) {
  const cfg = _stationCfg(stationKey);

  // Prova a servire dalla cache specifica della stazione
  const cached = _cacheRead(cfg.key);
  if (cached) {
    return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(cfg.dataSheetName);
  // Foglio1 storico mensile è SOLO per stazione STUDIO (i suoi storici sono lì da 2020)
  const mensile   = cfg.hasFoglio1 ? leggiDatiStorici() : {};
  const now       = new Date();

  let attuale    = null;
  let giornaliero = [];
  let raw         = [];

  if (dataSheet && dataSheet.getLastRow() > 1) {
    const cutoff30d  = new Date(now.getTime() - 30 * 24 * 3600 * 1000);
    const lastRow    = dataSheet.getLastRow();
    const LOOK_BACK  = 5000;
    const startRow   = Math.max(2, lastRow - LOOK_BACK + 1);
    const numRows    = lastRow - startRow + 1;
    const nCols      = cfg.isMain ? 4 : 5;
    raw = dataSheet.getRange(startRow, 1, numRows, nCols).getValues()
      .filter(r => r[0] instanceof Date && r[0] >= cutoff30d && r[1] !== '' && !isNaN(parseFloat(r[1])))
      .map(r => cfg.isMain
        ? { ts: r[0].getTime(), t: parseFloat(r[1]), h: parseFloat(r[2]), p: 0,
            co2: r[3] !== '' && r[3] !== null ? parseFloat(r[3]) : null }
        : { ts: r[0].getTime(), t: parseFloat(r[1]), h: parseFloat(r[2]), p: parseFloat(r[3]) || 0,
            press: r[4] !== '' && r[4] !== null ? parseFloat(r[4]) : null }
      );

    if (raw.length) {
      attuale = raw[raw.length - 1];

      // Aggrega per giorno
      const byDay = {};
      raw.forEach(r => {
        const d   = new Date(r.ts);
        const key = d.getFullYear() + '-' +
                    String(d.getMonth() + 1).padStart(2, '0') + '-' +
                    String(d.getDate()).padStart(2, '0');
        if (!byDay[key]) byDay[key] = [];
        byDay[key].push(r);
      });

      // Totali giornalieri esatti da cache ScriptProperties (aggiornata da fetchAndSaveData)
      const rainDailyMap = getRainDailyMap(cfg.key);

      giornaliero = Object.keys(byDay).sort().map(data => {
        const recs  = byDay[data];
        const temps = recs.map(r => r.t);
        const hums  = recs.map(r => r.h);
        const dps   = recs.map(r => {
          const a = 17.27, b = 237.3;
          const alpha = (a * r.t / (b + r.t)) + Math.log(r.h / 100);
          return b * alpha / (a - alpha);
        });
        // Pioggia: usa il totale giornaliero esatto da getmeasure (scale=1day).
        // Fallback ai delta di sum_rain_1 se getmeasure non è disponibile.
        let pioTot;
        if (rainDailyMap[data] != null) {
          pioTot = rainDailyMap[data];
        } else {
          pioTot = 0;
          for (let i = 1; i < recs.length; i++) {
            const delta = (recs[i].p || 0) - (recs[i-1].p || 0);
            if (delta > 0) pioTot += delta;
          }
          pioTot = Math.round(pioTot * 10) / 10;
        }
        const avg = arr => arr.reduce((a, b) => a + b) / arr.length;
        return { data, tMin: Math.min(...temps), tMedia: avg(temps), tMax: Math.max(...temps),
                       hMin: Math.min(...hums),  hMedia: avg(hums),  hMax: Math.max(...hums),
                       rdMin: Math.min(...dps),  rdMedia: avg(dps),  rdMax: Math.max(...dps),
                       pioggia: pioTot };
      });

      // Integra tutti i mesi presenti in raw non già presenti in Foglio 1
      const avg = arr => arr.reduce((a, b) => a + b) / arr.length;
      const mesiInRaw = new Set(raw.map(r => { const d = new Date(r.ts); return d.getFullYear() + '-' + d.getMonth(); }));
      mesiInRaw.forEach(key => {
        const [yrN, mN] = key.split('-').map(Number);
        const yr = String(yrN);
        if (!mensile[yr]) mensile[yr] = {};
        if (!mensile[yr][mN]) {
          const mr    = raw.filter(r => { const d = new Date(r.ts); return d.getFullYear() === yrN && d.getMonth() === mN; });
          const temps = mr.map(r => r.t);
          const hums  = mr.map(r => r.h);
          if (temps.length) {
            let pioggiaMese = null;
            const giornalieroMese = giornaliero.filter(g => {
              const d = new Date(g.data);
              return d.getFullYear() === yrN && d.getMonth() === mN;
            });
            if (giornalieroMese.length) {
              pioggiaMese = Math.round(giornalieroMese.reduce((s, g) => s + (g.pioggia || 0), 0) * 10) / 10;
            }
            const isCurrent = yrN === now.getFullYear() && mN === now.getMonth();
            mensile[yr][mN] = { tMin: Math.min(...temps), tMedia: avg(temps), tMax: Math.max(...temps),
                                hMin: Math.min(...hums),  hMedia: avg(hums),  hMax: Math.max(...hums),
                                pioggia: pioggiaMese, fonte: isCurrent ? 'live' : 'calcolato' };
          }
        }
      });
    }
  }

  const json = JSON.stringify({
    aggiornato: now.getTime(),
    stazione:   cfg.key,
    nomeStazione: cfg.deviceName,
    attuale, mensile, giornaliero, raw
  });
  _cacheWrite(json, cfg.key);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function handleOAuthCallback(code) {
  try {
    const props        = PropertiesService.getScriptProperties();
    const clientId     = props.getProperty('CLIENT_ID');
    const clientSecret = props.getProperty('CLIENT_SECRET');
    const redirectUri  = getWebAppUrl();

    const payload =
      'grant_type=authorization_code' +
      '&client_id='     + encodeURIComponent(clientId) +
      '&client_secret=' + encodeURIComponent(clientSecret) +
      '&code='          + encodeURIComponent(code) +
      '&redirect_uri='  + encodeURIComponent(redirectUri);

    Logger.log('TOKEN REQUEST payload: ' + payload);
    Logger.log('redirect_uri usato: ' + redirectUri);

    const resp = UrlFetchApp.fetch(API.TOKEN, {
      method: 'POST',
      contentType: 'application/x-www-form-urlencoded',
      payload: payload,
      muteHttpExceptions: true,
    });

    const responseText = resp.getContentText();
    Logger.log('TOKEN RESPONSE (' + resp.getResponseCode() + '): ' + responseText);

    if (resp.getResponseCode() !== 200) {
      return HtmlService.createHtmlOutput(
        '<h2>❌ Errore token</h2>' +
        '<p><b>HTTP ' + resp.getResponseCode() + '</b></p>' +
        '<pre>' + responseText + '</pre>' +
        '<hr><p><b>redirect_uri inviato:</b><br><code>' + redirectUri + '</code></p>' +
        '<p><b>client_id:</b> ' + clientId + '</p>' +
        '<p><b>code (primi 20 car):</b> ' + code.substring(0, 20) + '...</p>'
      );
    }

    const data = JSON.parse(responseText);
    saveTokens(data);
    detectStations(data.access_token);

    return HtmlService.createHtmlOutput(
      '<h2>✅ Autorizzazione completata!</h2>' +
      '<p>Puoi chiudere questa finestra.</p>' +
      '<p>Torna al foglio → menu 🌡️ Clima → <strong>"3. Avvia aggiornamento automatico"</strong></p>'
    );
  } catch (err) {
    Logger.log('ERRORE handleOAuthCallback: ' + err.toString());
    return HtmlService.createHtmlOutput('<h2>❌ Errore</h2><p>' + err.toString() + '</p>');
  }
}

// ─────────────────────────────────────────────────────────────
// GESTIONE TOKEN
// ─────────────────────────────────────────────────────────────

function saveTokens(data) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ACCESS_TOKEN',  data.access_token);
  props.setProperty('TOKEN_EXPIRY',  String(Date.now() + data.expires_in * 1000));
  if (data.refresh_token) props.setProperty('REFRESH_TOKEN', data.refresh_token);
}

function getValidToken() {
  const props  = PropertiesService.getScriptProperties();
  const expiry = parseInt(props.getProperty('TOKEN_EXPIRY') || '0');

  if (Date.now() > expiry - 5 * 60 * 1000) {
    refreshAccessToken();
  }
  return props.getProperty('ACCESS_TOKEN');
}

function refreshAccessToken() {
  const props        = PropertiesService.getScriptProperties();
  const clientId     = props.getProperty('CLIENT_ID');
  const clientSecret = props.getProperty('CLIENT_SECRET');
  const refreshToken = props.getProperty('REFRESH_TOKEN');

  if (!refreshToken) throw new Error('Nessun refresh token. Esegui il setup.');

  const resp = UrlFetchApp.fetch(API.TOKEN, {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload:
      'grant_type=refresh_token' +
      '&client_id='     + encodeURIComponent(clientId) +
      '&client_secret=' + encodeURIComponent(clientSecret) +
      '&refresh_token=' + encodeURIComponent(refreshToken),
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) throw new Error('Refresh token fallito: ' + resp.getContentText());
  saveTokens(JSON.parse(resp.getContentText()));
}

// ─────────────────────────────────────────────────────────────
// RILEVAMENTO STAZIONE
// ─────────────────────────────────────────────────────────────

function detectStations(accessToken) {
  const resp = UrlFetchApp.fetch(API.STATIONS, {
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) throw new Error('Errore getstationsdata: ' + resp.getContentText());

  const body    = JSON.parse(resp.getContentText()).body;
  const devices = body.devices || [];
  if (!devices.length) throw new Error('Nessuna stazione trovata.');

  // Una sola stazione: il NAMain è il device stesso (Studio interno), NAModule1 = esterno, NAModule3 = pluviometro.
  const device = devices[0];
  const extMod = device.modules.find(m => m.type === 'NAModule1');
  const rain3  = device.modules.find(m => m.type === 'NAModule3');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('DEVICE_ID',   device._id);
  props.setProperty('DEVICE_NAME', device.station_name || 'Stazione Studio');
  if (extMod) {
    props.setProperty('MODULE_ID',   extMod._id);
    props.setProperty('MODULE_NAME', extMod.module_name || 'Esterno');
  }
  if (rain3) {
    props.setProperty('RAIN_MODULE_ID', rain3._id);
  }
  Logger.log('Stazione: ' + device.station_name + ' (' + device._id + ')' +
             ' | NAMain → tab Studio' +
             ' | NAModule1 esterno: ' + (extMod ? extMod.module_name : 'no') +
             ' | NAModule3 pluviometro: ' + (rain3 ? (rain3.module_name || rain3._id) : 'no'));
}

// ─────────────────────────────────────────────────────────────
// FETCH DATI
// ─────────────────────────────────────────────────────────────

function fetchAndSaveData(stationKey) {
  try {
    const cfg = _stationCfg(stationKey);
    const props = PropertiesService.getScriptProperties();

    if (!cfg.deviceId || !cfg.moduleId) {
      Logger.log('[' + cfg.key + '] Setup non completato: mancano DEVICE_ID o MODULE_ID.');
      return;
    }

    const token = getValidToken();
    const resp  = UrlFetchApp.fetch(API.STATIONS + '?device_id=' + encodeURIComponent(cfg.deviceId), {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('[' + cfg.key + '] Errore API stazioni: ' + resp.getContentText());
      return;
    }

    const body    = JSON.parse(resp.getContentText()).body;
    const device  = (body.devices || [])[0];
    if (!device) { Logger.log('[' + cfg.key + '] Dispositivo non trovato nella risposta.'); return; }

    // Source dei sensori: per STUDIO usa la base NAMain (device.dashboard_data),
    // per ESTERNO usa il modulo NAModule1 (module.dashboard_data).
    let sourceDash;
    if (cfg.isMain) {
      sourceDash = device.dashboard_data;
      if (!sourceDash) { Logger.log('[' + cfg.key + '] NAMain dashboard_data assente'); return; }
    } else {
      const module = device.modules.find(m => m._id === cfg.moduleId);
      if (!module) { Logger.log('[' + cfg.key + '] Modulo esterno non trovato.'); return; }
      sourceDash = module.dashboard_data;
      if (!sourceDash) { Logger.log('[' + cfg.key + '] NAModule1 dashboard_data assente'); return; }
    }
    const timestamp = new Date(sourceDash.time_utc * 1000);
    const temp      = sourceDash.Temperature;
    const hum       = sourceDash.Humidity;
    const co2       = cfg.isMain && sourceDash.CO2 != null ? sourceDash.CO2 : null;

    // Pioggia: solo per ESTERNO (NAMain non ha pluviometro). 0 anche se NAModule3 non presente/offline.
    let rain = 0;
    if (!cfg.isMain) {
      const rainMod = device.modules.find(m => m.type === 'NAModule3');
      if (rainMod && !cfg.rainModuleId) {
        props.setProperty('RAIN_MODULE_ID', rainMod._id);
        cfg.rainModuleId = rainMod._id;
      }
      rain = (rainMod && rainMod.dashboard_data)
             ? (rainMod.dashboard_data.sum_rain_1 != null ? rainMod.dashboard_data.sum_rain_1
                : rainMod.dashboard_data.Rain != null     ? rainMod.dashboard_data.Rain : 0)
             : 0;
      if (rainMod) {
        const rd = rainMod.dashboard_data;
        Logger.log('[' + cfg.key + '] NAModule3: ' + (rainMod.module_name || rainMod._id) +
                   ' | sum_rain_1=' + (rd ? rd.sum_rain_1 : 'N/A') +
                   ' | reachable=' + rainMod.reachable);
      }
    }

    // Pressione: sempre da NAMain (anche per ESTERNO usa device.dashboard_data — è la stessa stazione)
    const press = (device.dashboard_data && device.dashboard_data.Pressure != null)
                  ? device.dashboard_data.Pressure : null;

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateDataSheet(ss, cfg.dataSheetName);

    // Salta se il dato è già presente (stesso minuto)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const lastTs = sheet.getRange(lastRow, 1).getValue();
      if (lastTs instanceof Date && Math.abs(lastTs - timestamp) < 60000) {
        Logger.log('[' + cfg.key + '] Dato già presente, skip.');
        return;
      }
    }

    // Schema righe diverse: Studio 4 col (T/H/CO2), Esterno 5 col (T/H/Pioggia/Pressione)
    if (cfg.isMain) {
      sheet.appendRow([timestamp, temp, hum, co2]);
      Logger.log('[' + cfg.key + '] Salvato: ' + timestamp.toLocaleString('it-IT') + ' | T=' + temp + '°C H=' + hum + '% CO2=' + co2 + 'ppm');
    } else {
      sheet.appendRow([timestamp, temp, hum, rain, press]);
      Logger.log('[' + cfg.key + '] Salvato: ' + timestamp.toLocaleString('it-IT') + ' | T=' + temp + '°C H=' + hum + '% pioggia=' + rain + 'mm press=' + press + 'hPa');
    }

    // Invalida cache JSON specifica di questa stazione
    _cacheInvalidate(cfg.key);

    // Aggiorna cache pioggia giornaliera (max 1 volta ogni 30 min) — solo Esterno ha NAModule3
    if (!cfg.isMain) {
      const rainMod2 = device.modules.find(m => m.type === 'NAModule3');
      const rainModuleId = cfg.rainModuleId || (rainMod2 ? rainMod2._id : null);
      if (rainModuleId) updateRainDailyCacheIfNeeded(cfg.deviceId, rainModuleId, token, cfg.rainCacheKey);
    }

    if (cfg.hasFoglio1) {
      // Solo STUDIO ha Foglio1 → archivia mese precedente come prima
      archiviaMesePrecedente();
      updateDashboard();
    } else {
      updateDashboard(cfg.key);
    }

  } catch (err) {
    Logger.log('Errore fetchAndSaveData: ' + err.toString());
  }
}

// Wrapper per il trigger della "stazione" Studio (in realtà NAMain dello stesso device)
function fetchAndSaveDataStudio() {
  fetchAndSaveData('STUDIO');
}

// ─────────────────────────────────────────────────────────────
// IMPORT DATI STORICI
// ─────────────────────────────────────────────────────────────

function importHistoricalData() {
  const ui = SpreadsheetApp.getUi();
  const props    = PropertiesService.getScriptProperties();
  const deviceId = props.getProperty('DEVICE_ID');
  const moduleId = props.getProperty('MODULE_ID');

  if (!deviceId || !moduleId) {
    ui.alert('Setup non completato. Prima autorizza il sistema (Step 1 e 2).');
    return;
  }

  const r = ui.prompt(
    'Import dati storici',
    'Data di inizio import (formato YYYY-MM-DD).\n\nPer default: 2020-01-01',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const inputDate = r.getResponseText().trim() || '2020-01-01';
  const startDate = new Date(inputDate + 'T00:00:00');
  if (isNaN(startDate)) { ui.alert('Data non valida.'); return; }

  ui.alert(
    'Import avviato',
    'L\'import partirà da ' + startDate.toLocaleDateString('it-IT') + '.\n' +
    'Potrebbero volerci alcuni minuti. Controlla i log per lo stato.\n\n' +
    'NOTA: L\'API Netatmo conserva i dati fino a ~2 anni (piano gratuito).',
    ui.ButtonSet.OK
  );

  try {
    _doImport(startDate, 'AZIENDA');
    ui.alert('✅ Import completato! Controlla il foglio Dati.');
  } catch (err) {
    ui.alert('❌ Errore durante l\'import:\n' + err.toString());
    Logger.log('Errore importHistoricalData: ' + err.toString());
  }
}

// Import 30 giorni "Studio" (NAMain) — Luca 29/05/2026
function importHistoricalDataStudio() {
  const ui  = SpreadsheetApp.getUi();
  const cfg = _stationCfg('STUDIO');
  if (!cfg.deviceId) {
    ui.alert('Stazione non configurata. Prima esegui setup (step 1-2).');
    return;
  }
  const startDate = new Date(Date.now() - 30 * 24 * 3600 * 1000);
  startDate.setHours(0,0,0,0);
  ui.alert('Import Studio (interno)', 'Importerò gli ultimi 30 giorni (scale 30min) di ' + cfg.deviceName + ' (NAMain).\nControlla i log per lo stato.', ui.ButtonSet.OK);
  try {
    _doImport(startDate, 'STUDIO');
    ui.alert('✅ Import Studio completato! Foglio "' + cfg.dataSheetName + '" popolato.');
  } catch(err) {
    ui.alert('❌ Errore import Studio:\n' + err.toString());
    Logger.log('Errore importHistoricalDataStudio: ' + err);
  }
}

function _doImport(startDate, stationKey) {
  const cfg      = _stationCfg(stationKey);
  const deviceId = cfg.deviceId;
  // Per NAMain (Studio interno): getmeasure usa module_id = device_id
  const moduleId = cfg.isMain ? cfg.deviceId : cfg.moduleId;
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = getOrCreateDataSheet(ss, cfg.dataSheetName);

  // Prendi il timestamp più recente già presente nel foglio
  let importFrom = startDate;
  const lastRow  = sheet.getLastRow();
  if (lastRow > 1) {
    const lastTs = sheet.getRange(lastRow, 1).getValue();
    if (lastTs instanceof Date && lastTs > importFrom) {
      importFrom = new Date(lastTs.getTime() + 60000);
    }
  }

  const endDate   = new Date();
  const STEP_MS   = 20 * 24 * 3600 * 1000; // 20 giorni per chunk (30min scale, max ~960 valori)
  let   totalRows = 0;
  let   current   = new Date(importFrom);

  while (current < endDate) {
    const chunkEnd = new Date(Math.min(current.getTime() + STEP_MS, endDate.getTime()));
    const token    = getValidToken();

    const types = cfg.isMain ? 'Temperature,Humidity,CO2' : 'Temperature,Humidity';
    const params =
      'device_id='  + encodeURIComponent(deviceId) +
      '&module_id=' + encodeURIComponent(moduleId) +
      '&scale=30min' +
      '&type=' + types +
      '&date_begin=' + Math.floor(current.getTime() / 1000) +
      '&date_end='   + Math.floor(chunkEnd.getTime() / 1000) +
      '&optimize=false' +
      '&real_time=false';

    const resp = UrlFetchApp.fetch(API.MEASURE + '?' + params, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() === 200) {
      const body   = JSON.parse(resp.getContentText()).body;
      const newRows = [];

      // Gestisce entrambi i formati: array di chunk (default) o oggetto {timestamp: [val...]}.
      // Per Studio (isMain) i value sono [T, H, CO2], per Esterno [T, H].
      const pushRow = (ts, vals) => {
        if (cfg.isMain) {
          if (vals[0] != null && vals[1] != null) {
            newRows.push([new Date(ts), vals[0], vals[1], vals[2] != null ? vals[2] : null]);
          }
        } else {
          if (vals[0] != null && vals[1] != null) {
            newRows.push([new Date(ts), vals[0], vals[1]]);
          }
        }
      };

      if (Array.isArray(body)) {
        body.forEach(chunk => {
          const begTime  = chunk.beg_time;
          const stepTime = chunk.step_time || 1800;
          (chunk.value || []).forEach((tuple, i) => {
            if (tuple) pushRow((begTime + i * stepTime) * 1000, tuple);
          });
        });
      } else if (body && typeof body === 'object') {
        Object.keys(body).forEach(tsStr => {
          const ts  = parseInt(tsStr);
          const val = body[tsStr];
          if (Array.isArray(val)) pushRow(ts * 1000, val);
          else if (typeof val === 'object') {
            pushRow(ts * 1000, [val.Temperature, val.Humidity, val.CO2]);
          }
        });
      }

      if (newRows.length > 0) {
        const nCols = cfg.isMain ? 4 : 3;
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, nCols).setValues(newRows);
        totalRows += newRows.length;
      }

      Logger.log('Import chunk: ' + current.toLocaleDateString('it-IT') + ' → ' + chunkEnd.toLocaleDateString('it-IT') + ' | righe: ' + newRows.length);
    } else {
      Logger.log('Errore chunk ' + current.toLocaleDateString('it-IT') + ': ' + resp.getContentText());
    }

    current = new Date(chunkEnd.getTime() + 60000);
    Utilities.sleep(300);
  }

  // Ordina per data crescente
  if (sheet.getLastRow() > 2) {
    const nCols = cfg.isMain ? 4 : 3;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, nCols).sort(1);
  }

  Logger.log('[' + cfg.key + '] Import totale: ' + totalRows + ' righe.');
  _cacheInvalidate(cfg.key);
  updateDashboard(cfg.key);
}

// ─────────────────────────────────────────────────────────────
// RECUPERO PIOGGIA MANCANTE
// ─────────────────────────────────────────────────────────────

function recoverRainData() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const deviceId = props.getProperty('DEVICE_ID');
  if (!deviceId) { ui.alert('Setup non completato. Prima autorizza il sistema.'); return; }

  const r = ui.prompt(
    'Recupera pioggia mancante',
    'Inserisci la data (o range) da recuperare.\n\nEsempi:\n  2026-03-17\n  2026-03-17/2026-03-18\n\nDefault: oggi',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const input = r.getResponseText().trim();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const parts  = (input || today).split('/');
  const startDate = new Date((parts[0] || today) + 'T00:00:00');
  const endDate   = new Date((parts[1] || parts[0] || today) + 'T23:59:59');

  if (isNaN(startDate) || isNaN(endDate)) { ui.alert('Data non valida.'); return; }

  try {
    const updated = _doRecoverRain(deviceId, startDate, endDate);
    ui.alert('✅ Recupero completato!\n\n' + updated + ' righe aggiornate nel foglio Dati.');
    updateDashboard();
  } catch (err) {
    ui.alert('❌ Errore:\n' + err.toString());
    Logger.log('Errore recoverRainData: ' + err.toString());
  }
}

function _doRecoverRain(deviceId, startDate, endDate) {
  const token = getValidToken();

  // 1. Trova il modulo pioggia (NAModule3)
  const stResp = UrlFetchApp.fetch(API.STATIONS + '?device_id=' + encodeURIComponent(deviceId), {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (stResp.getResponseCode() !== 200) throw new Error('getstationsdata: ' + stResp.getContentText());

  const device  = (JSON.parse(stResp.getContentText()).body.devices || [])[0];
  if (!device) throw new Error('Dispositivo non trovato.');
  const rainMod = device.modules.find(m => m.type === 'NAModule3');
  if (!rainMod) throw new Error('Pluviometro (NAModule3) non trovato nella stazione.');
  const rainModuleId = rainMod._id;
  Logger.log('Pluviometro: ' + (rainMod.module_name || rainModuleId));

  // 2. Scarica pioggia via getmeasure (scale=30min → sum_rain per slot)
  const params =
    'device_id='  + encodeURIComponent(deviceId) +
    '&module_id=' + encodeURIComponent(rainModuleId) +
    '&scale=30min' +
    '&type=sum_rain' +
    '&date_begin=' + Math.floor(startDate.getTime() / 1000) +
    '&date_end='   + Math.floor(endDate.getTime() / 1000) +
    '&optimize=false' +
    '&real_time=false';

  const mResp = UrlFetchApp.fetch(API.MEASURE + '?' + params, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (mResp.getResponseCode() !== 200) throw new Error('getmeasure: ' + mResp.getContentText());

  // Costruisce mappa  slotStart_ms → {slotStart, slotEnd, mm}  (per ogni slot da 30min)
  const rainMap = {};
  const rawBody = JSON.parse(mResp.getContentText()).body;
  Logger.log('getmeasure body tipo: ' + (Array.isArray(rawBody) ? 'array' : typeof rawBody));

  if (Array.isArray(rawBody)) {
    // Formato optimize=false con beg_time/step_time/value
    rawBody.forEach(chunk => {
      const begTime  = chunk.beg_time;
      const stepTime = chunk.step_time || 1800;
      (chunk.value || []).forEach((val, i) => {
        const mm = (val && val[0] != null) ? val[0] : 0;
        if (mm > 0) {
          const slotStart = (begTime + i * stepTime) * 1000;
          const slotEnd   = slotStart + stepTime * 1000;
          rainMap[slotStart] = { slotStart, slotEnd, mm };
        }
      });
    });
  } else if (rawBody && typeof rawBody === 'object') {
    // Formato oggetto {timestamp_str: [val]}
    Object.keys(rawBody).forEach(tsStr => {
      const slotStart = parseInt(tsStr) * 1000;
      const slotEnd   = slotStart + 1800 * 1000;
      const val = rawBody[tsStr];
      const mm  = Array.isArray(val) ? (val[0] != null ? val[0] : 0) : (typeof val === 'number' ? val : 0);
      if (mm > 0) rainMap[slotStart] = { slotStart, slotEnd, mm };
    });
  }
  Logger.log('Slot con pioggia trovati: ' + Object.keys(rainMap).length);

  // 3. Leggi il foglio Dati e aggiorna colonna D (pioggia) per ogni riga nel range
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Foglio Dati vuoto o non trovato.');

  const numRows  = sheet.getLastRow() - 1;
  const tsValues = sheet.getRange(2, 1, numRows, 1).getValues();   // colonna A
  const startMs  = startDate.getTime();
  const endMs    = endDate.getTime();
  let   updated  = 0;

  // Per ogni riga nel range, assegna la pioggia dello slot 30min che la contiene
  const rainUpdates = tsValues.map(([ts]) => {
    if (!(ts instanceof Date)) return null;
    const tsMs = ts.getTime();
    if (tsMs < startMs || tsMs > endMs) return null;

    let mm = 0;
    for (const key of Object.keys(rainMap)) {
      const { slotStart, slotEnd, mm: slotMm } = rainMap[key];
      if (tsMs >= slotStart && tsMs < slotEnd) {
        // Distribuisce equamente la pioggia del bucket sulle righe che cadono dentro
        mm = slotMm;
        break;
      }
    }
    return mm;
  });

  // Per ogni slot piovoso, conta quante righe ci cadono dentro per distribuire correttamente
  for (const key of Object.keys(rainMap)) {
    const { slotStart, slotEnd, mm } = rainMap[key];
    const rowsInSlot = [];
    tsValues.forEach(([ts], i) => {
      if (ts instanceof Date) {
        const tsMs = ts.getTime();
        if (tsMs >= slotStart && tsMs < slotEnd) rowsInSlot.push(i);
      }
    });
    if (rowsInSlot.length === 0) continue;
    // Prima riga dello slot prende tutta la pioggia, le altre 0
    // (come fa il fetch live con sum_rain_1)
    rowsInSlot.forEach((i, idx) => { rainUpdates[i] = idx === 0 ? mm : 0; });
  }

  // Scrivi le celle aggiornate
  rainUpdates.forEach((mm, i) => {
    if (mm === null) return;
    sheet.getRange(i + 2, 4).setValue(mm);  // colonna D = pioggia
    updated++;
  });

  Logger.log('Righe aggiornate: ' + updated);
  return updated;
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────

function updateDashboard(stationKey) {
  const cfg       = _stationCfg(stationKey);
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dash      = ss.getSheetByName(cfg.dashSheetName) || ss.insertSheet(cfg.dashSheetName);
  const dataSheet = ss.getSheetByName(cfg.dataSheetName);
  const now       = new Date();
  const MESI      = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
                     'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];

  // ── Leggi dati live da foglio Dati ──
  let liveRecords = [];
  if (dataSheet && dataSheet.getLastRow() > 1) {
    liveRecords = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues()
      .filter(r => r[0] instanceof Date && r[1] !== '' && !isNaN(parseFloat(r[1])))
      .map(r => ({ ts: r[0], t: parseFloat(r[1]), h: parseFloat(r[2]) }));
  }
  const last = liveRecords.length ? liveRecords[liveRecords.length - 1] : null;

  // ── Storici mensili da Foglio 1 SOLO per stazione STUDIO (la AZIENDA non ha pregressi) ──
  const storici = cfg.hasFoglio1 ? leggiDatiStorici() : {};

  // ── Costruisci mappa unificata: { anno: { mese: {tMin,tMedia,tMax,hMin,hMedia,hMax,pioggia,fonte} } } ──
  // Prima carica tutto da Foglio 1
  const dati = {};
  Object.keys(storici).forEach(yr => {
    dati[yr] = {};
    Object.keys(storici[yr]).forEach(m => {
      dati[yr][m] = { ...storici[yr][m], fonte: 'storico' };
    });
  });

  // Poi, per l'anno corrente, calcola i mesi live da Dati e sovrascrive SOLO se non già in Foglio 1
  const annoCorrente = now.getFullYear();
  const meseCorrente = now.getMonth();
  if (!dati[annoCorrente]) dati[annoCorrente] = {};

  for (let m = 0; m <= meseCorrente; m++) {
    if (dati[annoCorrente][m] && dati[annoCorrente][m].fonte === 'storico') continue; // Foglio 1 ha priorità
    const recs = liveRecords.filter(r => r.ts.getFullYear() === annoCorrente && r.ts.getMonth() === m);
    if (!recs.length) continue;
    const temps = recs.map(r => r.t);
    const hums  = recs.map(r => r.h);
    const avg   = arr => arr.reduce((a, b) => a + b) / arr.length;
    dati[annoCorrente][m] = {
      tMin:    Math.min(...temps),
      tMedia:  avg(temps),
      tMax:    Math.max(...temps),
      hMin:    Math.min(...hums),
      hMedia:  avg(hums),
      hMax:    Math.max(...hums),
      pioggia: null,
      fonte:   m === meseCorrente ? 'live' : 'calcolato',
    };
  }

  // ── Calcola record storici da Foglio 1 (su tutti gli anni) ──
  let recMinT = null, recMaxT = null, recMinH = null, recMaxH = null;
  let recMinTInfo = '', recMaxTInfo = '', recMinHInfo = '', recMaxHInfo = '';
  Object.keys(dati).forEach(yr => {
    Object.keys(dati[yr]).forEach(m => {
      const s = dati[yr][m];
      const label = MESI[m] + ' ' + yr;
      if (s.tMin !== null && (recMinT === null || s.tMin < recMinT)) { recMinT = s.tMin; recMinTInfo = label; }
      if (s.tMax !== null && (recMaxT === null || s.tMax > recMaxT)) { recMaxT = s.tMax; recMaxTInfo = label; }
      if (s.hMin !== null && (recMinH === null || s.hMin < recMinH)) { recMinH = s.hMin; recMinHInfo = label; }
      if (s.hMax !== null && (recMaxH === null || s.hMax > recMaxH)) { recMaxH = s.hMax; recMaxHInfo = label; }
    });
  });

  // ── Scrivi ──
  dash.clearContents();
  dash.clearFormats();
  let R = 1;

  const f = v => v !== null && v !== undefined ? (typeof v === 'number' ? v.toFixed(1) : String(v)) : '—';

  function merge(r, c1, c2, val, bg, fg, bold, size, align) {
    const range = dash.getRange(r, c1, 1, c2 - c1 + 1).merge().setValue(val);
    if (bg)    range.setBackground(bg);
    if (fg)    range.setFontColor(fg);
    if (bold)  range.setFontWeight('bold');
    if (size)  range.setFontSize(size);
    if (align) range.setHorizontalAlignment(align);
    return range;
  }

  function cell(r, c, val, bg, fg, bold, size, align) {
    const range = dash.getRange(r, c).setValue(val);
    if (bg)    range.setBackground(bg);
    if (fg)    range.setFontColor(fg);
    if (bold)  range.setFontWeight('bold');
    if (size)  range.setFontSize(size);
    if (align) range.setHorizontalAlignment(align);
    return range;
  }

  // ── TITOLO ──
  const titolo = cfg.key === 'STUDIO'
    ? '🌡️  CLIMA — STUDIO (interno NAMain)'
    : '🌡️  CLIMA — STAZIONE ESTERNA (NAModule1)';
  merge(R,1,8, titolo,'#1a73e8','#ffffff',true,16,'center');
  dash.setRowHeight(R, 42); R++;

  // ── ULTIMO VALORE LIVE ──
  const liveStr = last
    ? 'Ultimo aggiornamento: ' + last.ts.toLocaleString('it-IT') +
      '   |   T: ' + last.t.toFixed(1) + ' °C   |   U: ' + last.h.toFixed(0) + ' %'
    : 'Nessun dato live — avvia "Aggiorna dati ora"';
  merge(R,1,8, liveStr, '#e8f0fe','#1a237e',true,11,'center'); R++;
  R++; // spazio

  // ── RECORD STORICI ──
  const labelRecord = cfg.hasFoglio1
    ? '🏆  RECORD STORICI (2020 → oggi)'
    : '🏆  RECORD (dati raw disponibili in ' + cfg.dataSheetName + ')';
  merge(R,1,8, labelRecord,'#37474f','#ffffff',true,11,'left'); R++;
  const hdrs = ['','Valore','Periodo','','','','',''];
  hdrs.forEach((h,i) => cell(R,i+1,h,'#cfd8dc',null,true,9,'center'));
  dash.getRange(R,1).setHorizontalAlignment('left'); R++;

  function recRow(label, val, unit, info, fgVal) {
    cell(R,1,label,null,null,true,10,'left');
    cell(R,2, val !== null ? val.toFixed(1)+unit : '—', null, fgVal, true, 13, 'center');
    merge(R,3,8, info, null,'#555555',false,10,'left');
    R++;
  }
  recRow('Minima temperatura', recMinT, '°C', recMinTInfo, '#0d47a1');
  recRow('Massima temperatura', recMaxT, '°C', recMaxTInfo, '#b71c1c');
  recRow('Minima umidità',  recMinH, '%',  recMinHInfo, '#0d47a1');
  recRow('Massima umidità', recMaxH, '%',  recMaxHInfo, '#b71c1c');
  R++;

  // ── TABELLA UNIFICATA PER ANNO ──
  merge(R,1,8,'📊  DATI MENSILI PER ANNO','#1a73e8','#ffffff',true,11,'left'); R++;

  const anni = Object.keys(dati).map(Number).sort((a,b) => b - a);

  anni.forEach(yr => {
    // Intestazione anno
    merge(R,1,8,'Anno ' + yr, '#455a64','#ffffff',true,12,'left'); R++;

    // Header colonne
    ['Mese','T.min °C','T.media °C','T.max °C','U.min %','U.media %','U.max %','Pioggia mm'].forEach((h,i) => {
      cell(R,i+1,h,'#cfd8dc',null,true,9,'center');
    });
    dash.getRange(R,1).setHorizontalAlignment('left');
    R++;

    MESI.forEach((nome, m) => {
      const s   = dati[yr] && dati[yr][m] ? dati[yr][m] : null;
      const bg  = m % 2 === 0 ? '#f8f9fa' : '#ffffff';
      const isLive = s && s.fonte === 'live';
      const label  = nome + (isLive ? ' ⟳' : '');

      cell(R,1,label,bg,isLive?'#1a73e8':null,isLive,10,'left');
      if (s) {
        cell(R,2, f(s.tMin),   bg, s.tMin!==null&&s.tMin<0?'#1565c0':null, true,  10,'center');
        cell(R,3, f(s.tMedia), bg, null,                                     false, 10,'center');
        cell(R,4, f(s.tMax),   bg, s.tMax!==null&&s.tMax>35?'#b71c1c':null, true,  10,'center');
        cell(R,5, f(s.hMin),   bg, null, false, 10,'center');
        cell(R,6, f(s.hMedia), bg, null, false, 10,'center');
        cell(R,7, f(s.hMax),   bg, null, false, 10,'center');
        cell(R,8, s.pioggia!==null ? s.pioggia : '—', bg, null, false, 10,'center');
      } else {
        [2,3,4,5,6,7,8].forEach(c => cell(R,c,'—',bg,'#cccccc',false,10,'center'));
      }
      R++;
    });
    R++; // spazio tra anni
  });

  // ── Colonne ──
  dash.setColumnWidth(1, 140);
  [2,3,4,5,6,7,8].forEach(c => dash.setColumnWidth(c, 88));

  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────────────────────
// TRIGGER
// ─────────────────────────────────────────────────────────────

function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers
    .filter(t => ['fetchAndSaveData','fetchAndSaveDataStudio'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Esterno (NAModule1) → trigger sempre
  ScriptApp.newTrigger('fetchAndSaveData').timeBased().everyMinutes(CFG.INTERVAL_MIN).create();
  // Studio (NAMain) → trigger sempre (è lo stesso device)
  ScriptApp.newTrigger('fetchAndSaveDataStudio').timeBased().everyMinutes(CFG.INTERVAL_MIN).create();

  SpreadsheetApp.getUi().alert(
    '✅ Aggiornamento attivo ogni ' + CFG.INTERVAL_MIN + ' min su:\n' +
    '• Esterno (NAModule1) → foglio Dati\n' +
    '• Studio (NAMain)     → foglio Dati_Studio'
  );
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['fetchAndSaveData','fetchAndSaveDataStudio'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('Trigger rimossi (esterno + studio).');
}

// Esegue una rilevazione del device Netatmo collegato (riusa detectStations)
function rilevaStazioni() {
  const ui = SpreadsheetApp.getUi();
  try {
    detectStations(getValidToken());
    const props = PropertiesService.getScriptProperties();
    ui.alert('✅ Device rilevato:\n\n' +
      'Nome: ' + props.getProperty('DEVICE_NAME') + '\n' +
      'Modulo esterno (NAModule1): ' + (props.getProperty('MODULE_NAME') || 'non trovato') + '\n' +
      'NAMain (Studio interno) → automatico\n' +
      'Pluviometro (NAModule3): ' + (props.getProperty('RAIN_MODULE_ID') ? '✓' : 'non trovato'));
  } catch(e) {
    ui.alert('❌ Errore: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────
// STATO SISTEMA
// ─────────────────────────────────────────────────────────────

function showStatus() {
  const props    = PropertiesService.getScriptProperties();
  const expiry   = props.getProperty('TOKEN_EXPIRY');
  const trigExt    = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'fetchAndSaveData');
  const trigStudio = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'fetchAndSaveDataStudio');

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const cfgExt     = _stationCfg('ESTERNO');
  const cfgStudio  = _stationCfg('STUDIO');
  const dataExt    = ss.getSheetByName(cfgExt.dataSheetName);
  const dataStudio = ss.getSheetByName(cfgStudio.dataSheetName);
  const rowsExt    = dataExt    ? Math.max(0, dataExt.getLastRow() - 1)    : 0;
  const rowsStudio = dataStudio ? Math.max(0, dataStudio.getLastRow() - 1) : 0;

  SpreadsheetApp.getUi().alert('Stato Sistema', [
    'CLIENT_ID:       ' + (props.getProperty('CLIENT_ID')     ? '✅' : '❌ mancante'),
    'CLIENT_SECRET:   ' + (props.getProperty('CLIENT_SECRET') ? '✅' : '❌ mancante'),
    'Refresh token:   ' + (props.getProperty('REFRESH_TOKEN') ? '✅' : '❌ mancante (autorizza)'),
    'Token scade:     ' + (expiry ? new Date(parseInt(expiry)).toLocaleString('it-IT') : '—'),
    '─── ESTERNO (NAModule1) ───',
    'DEVICE:          ' + (cfgExt.deviceId ? props.getProperty('DEVICE_NAME') + ' (' + cfgExt.deviceId + ')' : '—'),
    'MODULO esterno:  ' + (props.getProperty('MODULE_NAME') || '—'),
    'Record foglio:   ' + rowsExt,
    'Trigger attivo:  ' + (trigExt.length ? '✅ ogni ' + CFG.INTERVAL_MIN + ' min' : '❌'),
    '─── STUDIO (NAMain interno) ───',
    'Record foglio:   ' + rowsStudio,
    'Trigger attivo:  ' + (trigStudio.length ? '✅ ogni ' + CFG.INTERVAL_MIN + ' min' : '❌'),
    '─────────────────────',
    'Web App URL:',
    getWebAppUrl(),
  ].join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────────────────────

function getOrCreateDataSheet(ss, sheetName) {
  const name = sheetName || CFG.DATA_SHEET;
  const isStudio = name === 'Dati_Studio';
  let sheet = ss.getSheetByName(name);
  // Schema dedicato per Studio (NAMain interno): no pioggia/pressione, sì CO2
  const HEADERS = isStudio
    ? ['Data/Ora', 'Temperatura (°C)', 'Umidità (%)', 'CO2 (ppm)']
    : ['Data/Ora', 'Temperatura (°C)', 'Umidità (%)', 'Pioggia (mm)', 'Pressione (hPa)'];
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    if (!isStudio) sheet.setColumnWidth(5, 130);
    sheet.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange('B2:B').setNumberFormat('0.0');
    sheet.getRange('C2:C').setNumberFormat('0');
    sheet.getRange('D2:D').setNumberFormat(isStudio ? '0' : '0.0');  // CO2 intero, pioggia con decimale
    if (!isStudio) sheet.getRange('E2:E').setNumberFormat('0.0');
  } else {
    HEADERS.forEach((h, i) => {
      const c = sheet.getRange(1, i + 1);
      if (c.getValue() !== h) c.setValue(h).setFontWeight('bold').setBackground('#e8eaf6');
    });
  }
  return sheet;
}

function getWebAppUrl() {
  // URL fisso del deployment @versioned — NON cambia con i push
  return 'https://script.google.com/macros/s/AKfycbzl8r8hzoZMs0n0gslLyIdYRN0Q3oF7CnjSqsIuFlycqSquisPJmsyoUQ8QFBzPQFYF/exec';
}

// Chiama questa funzione UNA VOLTA dall'editor (▶ Esegui) per autorizzare tutto
function autorizzaScript() {
  const props = PropertiesService.getScriptProperties();

  // Tocca ogni servizio per forzare i permessi
  SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers();

  // Salva l'URL del web app
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) props.setProperty('WEBAPP_URL', url);
  } catch (_) {}

  const url = getWebAppUrl();
  props.setProperty('WEBAPP_URL', url);

  Logger.log('✅ Script autorizzato!');
  Logger.log('Redirect URI per Netatmo: ' + url);
  Logger.log('Ora torna sul foglio → menu 🌡️ Clima → step 1, 2, 3');
}

// ─────────────────────────────────────────────────────────────
// ARCHIVIAZIONE AUTOMATICA MESE PRECEDENTE IN FOGLIO 1
// ─────────────────────────────────────────────────────────────

function archiviaMesePrecedente() {
  const ui   = SpreadsheetApp.getUi();
  const now  = new Date();
  const d    = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const anno = d.getFullYear();
  const mese = d.getMonth(); // 0-based

  const MESI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
                'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
  const meseName = MESI[mese] + ' ' + anno;

  try {

  // Già presente in Foglio1 con dati reali? Niente da fare.
  const storici = leggiDatiStorici();
  if (storici[anno] && storici[anno][mese] && storici[anno][mese].tMedia !== null) {
    ui.alert('ℹ️ ' + meseName + ' è già presente in Foglio1 con dati.');
    return;
  }

  // Leggi tutti i record del mese dal foglio Dati
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CFG.DATA_SHEET);
  if (!dataSheet || dataSheet.getLastRow() < 2) {
    ui.alert('❌ Foglio Dati non trovato o vuoto.');
    return;
  }

  const rows = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 5).getValues()
    .filter(r => r[0] instanceof Date && r[0].getFullYear() === anno && r[0].getMonth() === mese
                 && r[1] !== '' && !isNaN(parseFloat(r[1])));

  if (!rows.length) {
    ui.alert('❌ Nessun dato trovato in Dati per ' + meseName + '.');
    return;
  }

  const temps = rows.map(r => parseFloat(r[1]));
  const hums  = rows.map(r => parseFloat(r[2]));
  const dps   = rows.map(r => {
    const t = parseFloat(r[1]), h = parseFloat(r[2]);
    const a = 17.27, b = 237.3;
    const alpha = (a * t / (b + t)) + Math.log(h / 100);
    return b * alpha / (a - alpha);
  });
  const avg = arr => arr.reduce((a, b) => a + b) / arr.length;

  // Pioggia dalla cache giornaliera
  const rainDailyMap = getRainDailyMap();
  const giorni = Object.keys(rainDailyMap).filter(k => {
    const dd = new Date(k); return dd.getFullYear() === anno && dd.getMonth() === mese;
  });
  const pioggia = giorni.length
    ? Math.round(giorni.reduce((s, k) => s + (rainDailyMap[k] || 0), 0) * 10) / 10
    : null;

  const newRow = [
    MESI[mese],
    +avg(temps).toFixed(2),
    +Math.min(...temps).toFixed(1),
    +Math.max(...temps).toFixed(1),
    +avg(hums).toFixed(1),
    +Math.min(...hums).toFixed(1),
    +Math.max(...hums).toFixed(1),
    pioggia,
    +avg(dps).toFixed(2),
    +Math.min(...dps).toFixed(1),
    +Math.max(...dps).toFixed(1),
  ];

  // Scrivi in Foglio1: cerca la riga esistente del mese, altrimenti aggiungi
  const sheet     = ss.getSheetByName('Foglio1') || ss.getSheetByName('Foglio 1') || ss.getSheets()[0];
  const sheetData = sheet.getDataRange().getValues();

  let annoRow  = -1;
  let meseRow  = -1; // riga (0-based) già esistente con il nome del mese
  let nextAnnoRow = sheetData.length;
  for (let i = 0; i < sheetData.length; i++) {
    const v = String(sheetData[i][0]).trim();
    if (v === String(anno))      { annoRow = i; }
    else if (annoRow >= 0 && v === MESI[mese]) { meseRow = i; }
    else if (annoRow >= 0 && /^\d{4}$/.test(v)) { nextAnnoRow = i; break; }
  }

  if (meseRow >= 0) {
    // Riga del mese già presente (ma vuota): aggiorna i valori
    sheet.getRange(meseRow + 1, 1, 1, newRow.length).setValues([newRow]);
  } else if (annoRow < 0) {
    // Anno non esiste ancora: aggiungi in fondo
    sheet.appendRow([String(anno)]);
    sheet.appendRow(newRow);
  } else if (nextAnnoRow >= sheetData.length) {
    // Anno esiste ed è l'ultimo: appendi in fondo
    sheet.appendRow(newRow);
  } else {
    // Inserisci prima del prossimo anno (1-based)
    sheet.insertRowBefore(nextAnnoRow + 1);
    sheet.getRange(nextAnnoRow + 1, 1, 1, newRow.length).setValues([newRow]);
  }

  Logger.log('archiviaMesePrecedente: salvato ' + meseName + ' in Foglio1.');
  updateDashboard();
  ui.alert('✅ ' + meseName + ' salvato in Foglio1 e Dashboard aggiornato (' + rows.length + ' record elaborati).');

  } catch(e) {
    ui.alert('❌ Errore: ' + e.message);
    Logger.log('archiviaMesePrecedente errore: ' + e);
  }
}

// ─────────────────────────────────────────────────────────────
// LETTURA DATI STORICI DA FOGLIO 1
// ─────────────────────────────────────────────────────────────

function leggiDatiStorici() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  // Prova per nome, poi fallback al primo foglio
  const sheet = ss.getSheetByName('Foglio1') || ss.getSheetByName('Foglio 1') || ss.getSheets()[0];
  const raw   = sheet.getDataRange().getValues();

  const MESI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
                'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];

  const parseN = v => {
    if (v === '' || v === null || v === undefined) return null;
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).replace(',', '.'));
    return isNaN(n) ? null : n;
  };

  const result = {}; // { 2020: { 0: {tMedia,tMin,...}, 1: {...}, ... }, ... }
  let anno = null;

  raw.forEach(row => {
    const prima = String(row[0]).trim();
    if (/^\d{4}$/.test(prima)) {
      anno = parseInt(prima);
      result[anno] = {};
      return;
    }
    if (anno && MESI.includes(prima)) {
      result[anno][MESI.indexOf(prima)] = {
        tMedia:  parseN(row[1]),
        tMin:    parseN(row[2]),
        tMax:    parseN(row[3]),
        hMedia:  parseN(row[4]),
        hMin:    parseN(row[5]),
        hMax:    parseN(row[6]),
        pioggia: parseN(row[7]),
        rdMedia: parseN(row[8]),
        rdMin:   parseN(row[9]),
        rdMax:   parseN(row[10]),
      };
    }
  });

  return result;
}

function startOf(unit, date, offsetDays) {
  const d = new Date(date);
  if (offsetDays) d.setDate(d.getDate() + offsetDays);
  d.setHours(0, 0, 0, 0);
  return d;
}

function calcStats(recs) {
  if (!recs || !recs.length) return { minT: null, maxT: null, avgT: null, minH: null, maxH: null, avgH: null };
  const temps = recs.map(r => r.t).filter(v => !isNaN(v));
  const hums  = recs.map(r => r.h).filter(v => !isNaN(v));
  const avg   = arr => arr.length ? arr.reduce((a, b) => a + b) / arr.length : null;
  return {
    minT: temps.length ? Math.min(...temps) : null,
    maxT: temps.length ? Math.max(...temps) : null,
    avgT: avg(temps),
    minH: hums.length  ? Math.min(...hums)  : null,
    maxH: hums.length  ? Math.max(...hums)  : null,
    avgH: avg(hums),
  };
}
