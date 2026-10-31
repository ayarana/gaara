const CONFIG = {
  CACHE_SECONDS: 300,
  CACHE_PREFIX: 'gaara_'
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'data';
    if (action === 'ping') return json_({ ok: true, ts: new Date().toISOString() });
    if (action !== 'data') throw new Error('Ação desconhecida');
    const fresh = e && e.parameter && e.parameter.fresh === '1';
    return ContentService.createTextOutput(getData(fresh)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function authorize() {
  ss_().getName();
  CacheService.getScriptCache().get('x');
  Logger.log('Autorizado! Planilha: ' + ss_().getName());
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_() {
  const ss = ss_();
  const name = PropertiesService.getScriptProperties().getProperty('SHEET_NAME');
  const sh = name ? ss.getSheetByName(name) : ss.getSheets()[0];
  if (!sh) throw new Error('Aba não encontrada: ' + name);
  return sh;
}

function getData(fresh) {
  const P = CONFIG.CACHE_PREFIX;
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(P + 'n') || 0);
  const builtAt = Number(cache.get(P + 't') || 0);

  const skipCache = fresh && Date.now() - builtAt > 20000;
  if (n && !skipCache) {
    const keys = []; for (let i = 0; i < n; i++) keys.push(P + i);
    const parts = cache.getAll(keys);
    if (keys.every(k => parts[k] !== undefined)) return keys.map(k => parts[k]).join('');
  }
  const json = JSON.stringify(buildData_());
  try {
    const map = {}, SIZE = 90000; let i = 0;
    for (let p = 0; p < json.length; p += SIZE) map[P + (i++)] = json.slice(p, p + SIZE);
    map[P + 'n'] = String(i);
    map[P + 't'] = String(Date.now());
    cache.putAll(map, CONFIG.CACHE_SECONDS);
  } catch (e) {  }
  return json;
}

function buildData_() {
  return { ok: true, updated: new Date().toISOString(), rows: sheet_().getDataRange().getDisplayValues() };
}

function clearCache() {
  CacheService.getScriptCache().remove(CONFIG.CACHE_PREFIX + 'n');
}
