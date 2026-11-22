// Minimal Google Apps Script mock to run Finance HQ end-to-end on real data.
const fs = require('fs');
const raw = JSON.parse(fs.readFileSync('/tmp/sheet.json', 'utf8'));
const isoRe = /^\d{4}-\d{2}-\d{2}T/;
const conv = v => (typeof v === 'string' && isoRe.test(v)) ? new Date(v) : v;
const data = {};
Object.keys(raw).forEach(k => data[k] = raw[k].map(r => r.map(conv)));
const calls = { mail: [], cal: [], toast: [] };
function chain() { const p = new Proxy(function () {}, { get: (t, k) => k === 'then' ? undefined : chain, apply: () => p }); return p; }
function makeSheet(name) {
  if (!data[name]) data[name] = [];
  const g = data[name];
  const sheet = {
    getName: () => name,
    getLastRow: () => { let n = g.length; while (n > 0 && g[n - 1].every(v => v === '' || v == null)) n--; return n; },
    getMaxRows: () => Math.max(g.length, 1000), getMaxColumns: () => 26,
    getRange: (a, b, c, d) => {
      if (typeof a === 'string') return chainRange(1, 1, 1, 1);
      return chainRange(a, b, c || 1, d || 1);
    },
    getBandings: () => [], getCharts: () => [], getConditionalFormatRules: () => [], getIndex: () => 1, isSheetHidden: () => false,
    insertRowsBefore: () => sheet, appendRow: r => { g.push(r); return sheet; }
  };
  function chainRange(r, c, nr, nc) {
    const rng = new Proxy({}, { get: (t, k) => {
      if (k === 'getValues') return () => { const o = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (g[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } o.push(row); } return o; };
      if (k === 'getValue') return () => { const v = (g[r - 1] || [])[c - 1]; return v === undefined ? '' : v; };
      if (k === 'setValues') return vals => { vals.forEach((row, i) => { while (g.length < r + i) g.push([]); row.forEach((v, j) => g[r - 1 + i][c - 1 + j] = v); }); return rng; };
      if (k === 'setValue') return v => { for (let i = 0; i < nr; i++) { while (g.length < r + i) g.push([]); for (let j = 0; j < nc; j++) g[r - 1 + i][c - 1 + j] = v; } return rng; };
      if (k === 'clearContent') return () => { for (let i = 0; i < nr; i++) if (g[r - 1 + i]) for (let j = 0; j < nc; j++) g[r - 1 + i][c - 1 + j] = ''; return rng; };
      if (k === 'getRow') return () => r; if (k === 'getColumn') return () => c; if (k === 'getNumRows') return () => nr; if (k === 'getNumColumns') return () => nc; if (k === 'getLastColumn') return () => c + nc - 1;
      if (k === 'getSheet') return () => sheet;
      if (k === 'applyRowBanding') return () => chain();
      return () => rng;
    } });
    return rng;
  }
  sheet.clear = () => { g.length = 0; return sheet; };
  return new Proxy(sheet, { get: (t, k) => k in t ? t[k] : () => chain() });
}
const sheets = {};
const ss = new Proxy({
  getSheetByName: n => data[n] ? (sheets[n] = sheets[n] || makeSheet(n)) : null,
  insertSheet: n => { data[n] = []; return (sheets[n] = makeSheet(n)); },
  getId: () => 'ID', getUrl: () => 'https://sheet', getSheets: () => Object.keys(data).map(n => sheets[n] = sheets[n] || makeSheet(n)), getNumSheets: () => Object.keys(data).length,
  toast: (m) => calls.toast.push(m)
}, { get: (t, k) => k in t ? t[k] : () => chain() });
const props = process.env.FRESH ? { SIMPLEFIN_SETUP_TOKEN: Buffer.from('https://claim.test/c').toString('base64') } : { SIMPLEFIN_ACCESS_URL: 'https://u:p@x', FIN_SHEET_ID: 'ID' };
if (process.env.FRESH) Object.keys(data).forEach(k => delete data[k]);
const fmtDate = (d, tz, f) => {
  d = new Date(d); const p = n => String(n).padStart(2, '0');
  const M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], D = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return f.replace(/yyyy|YYYY|MMMM|MMM|MM|dd|EEE|ww|h|mm|a|d/g, t => ({ yyyy: d.getFullYear(), YYYY: d.getFullYear(), MMMM: M[d.getMonth()], MMM: M[d.getMonth()], MM: p(d.getMonth() + 1), dd: p(d.getDate()), EEE: D[d.getDay()], ww: '47', h: d.getHours() % 12 || 12, mm: p(d.getMinutes()), a: d.getHours() < 12 ? 'AM' : 'PM', d: d.getDate() }[t]));
};
global.SpreadsheetApp = new Proxy({ openById: () => ss, getActiveSpreadsheet: () => null, create: () => ss, getUi: () => chain(),
  BandingTheme: { LIGHT_GREY: 1 }, newConditionalFormatRule: () => chain(), newDataValidation: () => chain() }, { get: (t, k) => k in t ? t[k] : chain });
global.PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => props[k] = v, deleteProperty: k => delete props[k], deleteAllProperties: () => Object.keys(props).forEach(k => delete props[k]) }) };
global.Session = { getScriptTimeZone: () => 'America/New_York', getEffectiveUser: () => ({ getEmail: () => 'me@x.com' }) };
global.Utilities = { formatDate: fmtDate, sleep: () => {}, base64Encode: x => Buffer.from(String(x)).toString('base64'), base64Decode: x => Buffer.from(x, 'base64'), newBlob: b => ({ getDataAsString: () => Buffer.from(b).toString() }) };
global.Logger = { log: m => console.log('LOG', String(m).slice(0, 160)) };
const trig = process.env.FRESH ? [] : ['financeDailySync', 'phase2Daily', 'archiveMonthly', 'phase3Daily', 'weeklyDigest', 'onCategoryEdit'];
global.ScriptApp = { getProjectTriggers: () => trig.map(h => ({ getHandlerFunction: () => h, h: h, getTriggerSourceId: () => (global.TRIG_SRC || 'ID') })),
  newTrigger: h => { const t = new Proxy({}, { get: (o, k) => k === 'create' ? () => { trig.push(h); return {}; } : () => t }); return t; },
  deleteTrigger: t => { const i = trig.indexOf(t.h); if (i >= 0) trig.splice(i, 1); }, WeekDay: { SUNDAY: 0 }, getScriptId: () => 'SCRIPTID', getOAuthToken: () => 'TOKEN' };
calls.trig = trig;
global.LockService = { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) };
global.MailApp = { sendEmail: o => calls.mail.push(o.subject) };
global.CalendarApp = { getDefaultCalendar: () => ({ createEvent: (t) => { calls.cal.push(t); return { removeAllReminders: () => {}, addPopupReminder: () => {} }; } }) };
const emptyIt = () => ({ hasNext: () => false, next: () => null });
const folder = { getFilesByName: emptyIt, createFile: b => { calls.pdf = (calls.pdf || []).concat([b.getName()]); return {}; } };
global.DriveApp = { getFoldersByName: emptyIt, createFolder: () => folder };
const blobOf = (html) => { const b = { name: '', html: html, getAs: () => b, setName: n => { b.name = n; return b; }, getName: () => b.name }; return b; };
global.HtmlService = { createHtmlOutput: h => { const o = { html: h, getBlob: () => blobOf(h), setTitle: () => o, addMetaTag: () => o, setWidth: () => o, setHeight: () => o }; calls.lastHtml = h; return o; } };
global.UrlFetchApp = { fetch: (url, opt) => { if (global.SF_HANDLER) return global.SF_HANDLER(url, opt); throw new Error('no network in test'); } };
global.Charts = { ChartHiddenDimensionStrategy: { SHOW_BOTH: 3 } };
module.exports = { data, calls, props };
