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
  DATA_SHEET:      'Dati',
  DASHBOARD_SHEET: 'Dashboard',
  INTERVAL_MIN:    30,           // minuti tra un fetch e l'altro
};

const API = {
  TOKEN:    'https://api.netatmo.com/oauth2/token',
  AUTH:     'https://api.netatmo.com/oauth2/authorize',
  STATIONS: 'https://api.netatmo.com/api/getstationsdata',
  MEASURE:  'https://api.netatmo.com/api/getmeasure',
};

// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌡️ Clima')
    .addItem('0. Autorizza script (prima volta)', 'autorizzaScript')
    .addItem('1. Configura credenziali Netatmo', 'setupCredentials')
    .addItem('2. Autorizza con Netatmo', 'startAuth')
    .addItem('3. Avvia aggiornamento automatico', 'setupTrigger')
    .addSeparator()
    .addItem('Aggiorna dati ora', 'fetchAndSaveData')
    .addItem('Aggiorna solo dashboard', 'updateDashboard')
    .addItem('Importa dati storici (API)', 'importHistoricalData')
    .addSeparator()
    .addItem('Stato sistema', 'showStatus')
    .addItem('Ferma aggiornamento automatico', 'removeTrigger')
    .addToUi();
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

// doGet viene chiamato da Netatmo dopo l'autorizzazione
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.code)            return handleOAuthCallback(params.code);
  if (params.action === 'getData') return serveJsonData();

  const props = PropertiesService.getScriptProperties();
  const ready = !!props.getProperty('REFRESH_TOKEN');
  return HtmlService.createHtmlOutput(
    '<h2>🌡️ Clima — Netatmo Monitor</h2>' +
    '<p>Stato: ' + (ready ? '✅ Autorizzato' : '⏳ In attesa di autorizzazione') + '</p>' +
    '<p><b>API dati:</b> <code>' + getWebAppUrl() + '?action=getData</code></p>'
  );
}

function serveJsonData() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CFG.DATA_SHEET);
  const mensile   = leggiDatiStorici();   // da Foglio 1
  const now       = new Date();

  let attuale    = null;
  let giornaliero = [];
  let raw         = [];

  if (dataSheet && dataSheet.getLastRow() > 1) {
    raw = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 3).getValues()
      .filter(r => r[0] instanceof Date && r[1] !== '' && !isNaN(parseFloat(r[1])))
      .map(r => ({ ts: r[0].getTime(), t: parseFloat(r[1]), h: parseFloat(r[2]) }));

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

      giornaliero = Object.keys(byDay).sort().map(data => {
        const recs  = byDay[data];
        const temps = recs.map(r => r.t);
        const hums  = recs.map(r => r.h);
        const dps   = recs.map(r => {
          const a = 17.27, b = 237.3;
          const alpha = (a * r.t / (b + r.t)) + Math.log(r.h / 100);
          return b * alpha / (a - alpha);
        });
        const avg = arr => arr.reduce((a, b) => a + b) / arr.length;
        return { data, tMin: Math.min(...temps), tMedia: avg(temps), tMax: Math.max(...temps),
                       hMin: Math.min(...hums),  hMedia: avg(hums),  hMax: Math.max(...hums),
                       rdMin: Math.min(...dps),  rdMedia: avg(dps),  rdMax: Math.max(...dps) };
      });

      // Integra mese corrente in mensile se non già presente in Foglio 1
      const yr = String(now.getFullYear());
      const m  = now.getMonth();
      if (!mensile[yr]) mensile[yr] = {};
      if (!mensile[yr][m]) {
        const mr    = raw.filter(r => { const d = new Date(r.ts); return d.getFullYear() === now.getFullYear() && d.getMonth() === m; });
        const temps = mr.map(r => r.t);
        const hums  = mr.map(r => r.h);
        if (temps.length) {
          const avg = arr => arr.reduce((a, b) => a + b) / arr.length;
          mensile[yr][m] = { tMin: Math.min(...temps), tMedia: avg(temps), tMax: Math.max(...temps),
                             hMin: Math.min(...hums),  hMedia: avg(hums),  hMax: Math.max(...hums),
                             pioggia: null, fonte: 'live' };
        }
      }
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ aggiornato: now.getTime(), attuale, mensile, giornaliero, raw }))
    .setMimeType(ContentService.MimeType.JSON);
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

  // Se c'è più di una stazione, usa la prima
  const device = devices[0];
  const extMod = device.modules.find(m => m.type === 'NAModule1'); // NAModule1 = outdoor

  const props = PropertiesService.getScriptProperties();
  props.setProperty('DEVICE_ID',   device._id);
  props.setProperty('DEVICE_NAME', device.station_name || 'Stazione');

  if (extMod) {
    props.setProperty('MODULE_ID',   extMod._id);
    props.setProperty('MODULE_NAME', extMod.module_name || 'Esterno');
  }

  Logger.log('Stazione: ' + device.station_name + ' | Modulo esterno: ' + (extMod ? extMod.module_name : 'non trovato'));
}

// ─────────────────────────────────────────────────────────────
// FETCH DATI
// ─────────────────────────────────────────────────────────────

function fetchAndSaveData() {
  try {
    const props    = PropertiesService.getScriptProperties();
    const deviceId = props.getProperty('DEVICE_ID');
    const moduleId = props.getProperty('MODULE_ID');

    if (!deviceId || !moduleId) {
      Logger.log('Setup non completato: mancano DEVICE_ID o MODULE_ID.');
      return;
    }

    const token = getValidToken();
    const resp  = UrlFetchApp.fetch(API.STATIONS + '?device_id=' + encodeURIComponent(deviceId), {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('Errore API stazioni: ' + resp.getContentText());
      return;
    }

    const body    = JSON.parse(resp.getContentText()).body;
    const device  = (body.devices || [])[0];
    if (!device) { Logger.log('Dispositivo non trovato nella risposta.'); return; }

    const module  = device.modules.find(m => m._id === moduleId);
    if (!module)  { Logger.log('Modulo esterno non trovato.'); return; }

    const dash      = module.dashboard_data;
    const timestamp = new Date(dash.time_utc * 1000);
    const temp      = dash.Temperature;
    const hum       = dash.Humidity;

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateDataSheet(ss);

    // Salta se il dato è già presente (stesso minuto)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const lastTs = sheet.getRange(lastRow, 1).getValue();
      if (lastTs instanceof Date && Math.abs(lastTs - timestamp) < 60000) {
        Logger.log('Dato già presente, skip.');
        return;
      }
    }

    sheet.appendRow([timestamp, temp, hum]);
    Logger.log('Salvato: ' + timestamp.toLocaleString('it-IT') + ' — ' + temp + '°C — ' + hum + '%');

    updateDashboard();

  } catch (err) {
    Logger.log('Errore fetchAndSaveData: ' + err.toString());
  }
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
    _doImport(startDate);
    ui.alert('✅ Import completato! Controlla il foglio Dati.');
  } catch (err) {
    ui.alert('❌ Errore durante l\'import:\n' + err.toString());
    Logger.log('Errore importHistoricalData: ' + err.toString());
  }
}

function _doImport(startDate) {
  const props    = PropertiesService.getScriptProperties();
  const deviceId = props.getProperty('DEVICE_ID');
  const moduleId = props.getProperty('MODULE_ID');
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = getOrCreateDataSheet(ss);

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

    const params =
      'device_id='  + encodeURIComponent(deviceId) +
      '&module_id=' + encodeURIComponent(moduleId) +
      '&scale=30min' +
      '&type=Temperature,Humidity' +
      '&date_begin=' + Math.floor(current.getTime() / 1000) +
      '&date_end='   + Math.floor(chunkEnd.getTime() / 1000) +
      '&optimize=false' +
      '&real_time=false';

    const resp = UrlFetchApp.fetch(API.MEASURE + '?' + params, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() === 200) {
      const body   = JSON.parse(resp.getContentText()).body || [];
      const newRows = [];

      body.forEach(chunk => {
        const begTime  = chunk.beg_time;
        const stepTime = chunk.step_time || 1800;
        (chunk.value || []).forEach((pair, i) => {
          if (pair[0] !== null && pair[1] !== null) {
            newRows.push([
              new Date((begTime + i * stepTime) * 1000),
              pair[0], // Temperature
              pair[1], // Humidity
            ]);
          }
        });
      });

      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);
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
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).sort(1);
  }

  Logger.log('Import totale: ' + totalRows + ' righe.');
  updateDashboard();
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────

function updateDashboard() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dash      = ss.getSheetByName(CFG.DASHBOARD_SHEET) || ss.insertSheet(CFG.DASHBOARD_SHEET);
  const dataSheet = ss.getSheetByName(CFG.DATA_SHEET);
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

  // ── Leggi dati storici mensili da Foglio 1 ──
  const storici = leggiDatiStorici();

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
  merge(R,1,8,'🌡️  CLIMA — STAZIONE ESTERNA','#1a73e8','#ffffff',true,16,'center');
  dash.setRowHeight(R, 42); R++;

  // ── ULTIMO VALORE LIVE ──
  const liveStr = last
    ? 'Ultimo aggiornamento: ' + last.ts.toLocaleString('it-IT') +
      '   |   T: ' + last.t.toFixed(1) + ' °C   |   U: ' + last.h.toFixed(0) + ' %'
    : 'Nessun dato live — avvia "Aggiorna dati ora"';
  merge(R,1,8, liveStr, '#e8f0fe','#1a237e',true,11,'center'); R++;
  R++; // spazio

  // ── RECORD STORICI ──
  merge(R,1,8,'🏆  RECORD STORICI (2020 → oggi)','#37474f','#ffffff',true,11,'left'); R++;
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
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'fetchAndSaveData')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('fetchAndSaveData').timeBased().everyMinutes(CFG.INTERVAL_MIN).create();

  SpreadsheetApp.getUi().alert('✅ Aggiornamento automatico attivo ogni ' + CFG.INTERVAL_MIN + ' minuti.');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'fetchAndSaveData')
    .forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('Trigger rimosso. I dati non verranno più aggiornati automaticamente.');
}

// ─────────────────────────────────────────────────────────────
// STATO SISTEMA
// ─────────────────────────────────────────────────────────────

function showStatus() {
  const props    = PropertiesService.getScriptProperties();
  const expiry   = props.getProperty('TOKEN_EXPIRY');
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'fetchAndSaveData');

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CFG.DATA_SHEET);
  const rows      = dataSheet ? Math.max(0, dataSheet.getLastRow() - 1) : 0;

  SpreadsheetApp.getUi().alert('Stato Sistema', [
    'CLIENT_ID:       ' + (props.getProperty('CLIENT_ID')     ? '✅' : '❌ mancante'),
    'CLIENT_SECRET:   ' + (props.getProperty('CLIENT_SECRET') ? '✅' : '❌ mancante'),
    'Refresh token:   ' + (props.getProperty('REFRESH_TOKEN') ? '✅' : '❌ mancante (autorizza)'),
    'Token scade:     ' + (expiry ? new Date(parseInt(expiry)).toLocaleString('it-IT') : '—'),
    'DEVICE_ID:       ' + (props.getProperty('DEVICE_ID')   || '—'),
    'MODULE esterno:  ' + (props.getProperty('MODULE_NAME') || '—'),
    '─────────────────────',
    'Record nel foglio: ' + rows,
    'Trigger attivo:  ' + (triggers.length ? '✅ ogni ' + CFG.INTERVAL_MIN + ' min' : '❌'),
    '─────────────────────',
    'Web App URL:',
    getWebAppUrl(),
  ].join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────────────────────

function getOrCreateDataSheet(ss) {
  let sheet = ss.getSheetByName(CFG.DATA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CFG.DATA_SHEET);
    sheet.getRange(1, 1, 1, 3)
      .setValues([['Data/Ora', 'Temperatura (°C)', 'Umidità (%)']])
      .setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 120);
    sheet.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange('B2:B').setNumberFormat('0.0');
    sheet.getRange('C2:C').setNumberFormat('0');
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
