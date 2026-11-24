// End-to-end simulation: a friend on the OLD version gets a real update to the NEW version.
process.env.FRESH = '1';
const M = require('/tmp/mock.js'); const fs = require('fs'), vm = require('vm');
const sf = JSON.parse(fs.readFileSync('/tmp/sf.json', 'utf8'));
const OLD = fs.readFileSync('/mnt/user-data/uploads/fhq/FinanceHQ.gs', 'utf8');      // what friends run today (2026.11.21.5)
const NEW = fs.readFileSync('/home/claude/pub/FinanceHQ.gs', 'utf8');                 // the release being published
const ver = c => c.match(/^const FHQ_VERSION = '([\d.]+)'/m)[1];
const drive = {}; const VFILE = '1hboJqNvt9l5ER8Y7DBI1XoWO7VrK0Zez', CFILE = '19qFP1Z10JbC3QBqZYMmsH-CInRCtvVeg';
let driveDown = false;
global.SF_HANDLER = (url) => {
  const ok = o => ({ getResponseCode: () => 200, getContentText: () => o });
  if (/uc\?export=download/.test(url)) { if (driveDown) return { getResponseCode: () => 503, getContentText: () => '' }; const id = url.split('id=')[1]; return drive[id] ? ok(drive[id]) : { getResponseCode: () => 404, getContentText: () => '' }; }
  if (/claim/.test(url)) return ok('https://user:pass@bridge.test/simplefin');
  const st = +url.match(/start-date=(\d+)/)[1], en = (url.match(/end-date=(\d+)/) || [])[1];
  return ok(JSON.stringify({ errors: [], accounts: sf.accounts.map(a => Object.assign({}, a, { transactions: a.transactions.filter(t => t.posted >= st && (!en || t.posted < +en)) })) }));
};
const mails = []; global.MailApp = { sendEmail: (a, b, c) => mails.push(typeof a === 'object' ? a : { to: a, subject: b, body: c }) };
const G = ['SpreadsheetApp','PropertiesService','Session','Utilities','Logger','ScriptApp','LockService','MailApp','CalendarApp','DriveApp','HtmlService','UrlFetchApp','Charts','console','Date','JSON','Math'];
const load = code => { const c = {}; G.forEach(k => c[k] = global[k]); vm.createContext(c); vm.runInContext(code, c); return c; };
let pass = 0, fail = 0;
const check = (name, cond, extra) => { cond ? pass++ : fail++; console.log((cond ? '  ✔ ' : '  ✘ ') + name + (extra ? '  — ' + extra : '')); };
const P = () => M.props, logText = () => (M.data.Log || []).map(r => r[1]).join('\n');
const errLines = () => (M.data.Log || []).map(r => String(r[1])).filter(t => /error|fail|exception|typeerror|not defined/i.test(t) && !/^(Help request|Published)/.test(t));

console.log(`\nOLD (friends today): ${ver(OLD)}   NEW (release): ${ver(NEW)}`);
console.log('\n1 · A friend sets up Finance HQ on the old version');
let A = load(OLD);
A.setupFinanceHQ();
A.saveSetup(JSON.stringify(Object.assign(JSON.parse(A.getSetupData()), { save: 15 })));
const txBefore = M.data.Transactions.length, setBefore = JSON.stringify(A.settings_());
M.data.Transactions[5][5] = 'Groceries'; M.data.Transactions[5][6] = 'manual'; const myFix = JSON.stringify(M.data.Transactions[5]);   // a category the friend fixed by hand
check('connected and synced', A.isConnected_() && txBefore > 50, (txBefore - 1) + ' transactions');
check('menu says "Check for updates"', A.updateLabel_() === "Check for updates (you have " + ver(OLD) + ")");
mails.length = 0;

console.log('\n2 · Before anything is published, the daily check stays quiet');
drive[VFILE] = JSON.stringify({ version: ver(OLD), notes: 'old', code: CFILE }); drive[CFILE] = OLD;
A.phase2Daily();
check('no update email', mails.length === 0);
check('no update flag', !P().UPDATE_AVAILABLE);

console.log('\n3 · Drive is unreachable for a day');
driveDown = true; A.checkForUpdate_(); driveDown = false;
check('fails silently — no email, no error in Log', mails.length === 0 && !/update file/i.test(logText()));

console.log('\n4 · Tomisin publishes ' + ver(NEW));
drive[VFILE] = JSON.stringify({ version: ver(NEW), notes: 'New: Finance HQ → Get help… sends Tomisin a message from inside the sheet.', code: CFILE, published: new Date().toISOString() });
drive[CFILE] = NEW;

console.log('\n5 · Next morning the friend\'s copy runs its daily job');
A.phase2Daily();
const um = mails.filter(m => /update/i.test(m.subject));
check('exactly one update email', um.length === 1, um[0] && um[0].subject);
check('email has the release notes and install steps', um[0] && /Get help/.test(um[0].body) && /Install update/.test(um[0].body));
check('menu now shows ⬆ Install update', A.updateLabel_() === '⬆ Install update (' + ver(NEW) + ')', A.updateLabel_());
check('Health shows the update is waiting', (M.data.Health || []).some(r => /update/i.test(r.join(' ')) && r.join(' ').includes(ver(NEW))));

console.log('\n6 · Day after — still not installed');
A.phase2Daily();
check('no second email for the same version', mails.filter(m => /update/i.test(m.subject)).length === 1);

console.log('\n7 · Friend opens Finance HQ → ⬆ Install update');
const d = JSON.parse(A.getUpdateCode());
check('window gets the full new code', d.code === NEW && d.version === ver(NEW), (d.code.length / 1024).toFixed(0) + ' KB');
check('window links to the friend\'s own script', /script\.google\.com\/home\/projects\/SCRIPTID\/edit/.test(d.editor));
drive[CFILE] = NEW.slice(0, 40000);
let partial = ''; try { A.getUpdateCode(); } catch (e) { partial = e.message; }
check('a half-uploaded release is refused', /incomplete/.test(partial), partial);
drive[CFILE] = NEW;

console.log('\n8 · Friend pastes the new code, saves, and opens the sheet');
const B = load(NEW);
B.onOpen(); check('window sees the new version once saved', B.installedVersion() === ver(NEW)); B.finishUpdate();
check('first run on the new version is recorded', P().RUNNING_VERSION === ver(NEW) && logText().includes('Running Finance HQ ' + ver(NEW)));
check('update flag cleared, menu back to "Check for updates"', !P().UPDATE_AVAILABLE && B.updateLabel_() === 'Check for updates (you have ' + ver(NEW) + ')');
check('bank connection kept (no re-connect needed)', B.isConnected_());
check('all transactions kept', M.data.Transactions.length >= txBefore, (M.data.Transactions.length - 1) + ' rows');
{ const fx = JSON.parse(myFix), now = M.data.Transactions.find(r => r[7] === fx[7]); check('hand-fixed category kept', now && now[5] === 'Groceries', JSON.stringify(fx).slice(0,120) + ' => ' + JSON.stringify(now).slice(0,160)); }
check('settings kept (savings goal 15%)', JSON.stringify(B.settings_()) === setBefore && B.pctSetting_(B.settings_()['Savings goal (% of pay)'] || B.settings_()['Savings goal']) !== 0);
check('triggers intact', ['financeDailySync', 'phase2Daily'].every(h => M.calls.trig.includes(h)), M.calls.trig.join(', '));
check('no errors in the Log', errLines().length === 0, errLines().slice(-2).join(' | '));
B.phase2Daily();
check('no more update emails once installed', mails.filter(m => /update/i.test(m.subject)).length === 1);

console.log('\n9 · The new feature works on the updated copy: Finance HQ → Get help…');
const pv = JSON.parse(B.getSupportPreview()).details;
check('details preview has version and sync status', pv.includes(ver(NEW)) && /Bank connected: yes/.test(pv));
check('no dollar amounts in the details', !/\$\s?\d/.test(pv));
const r = B.sendSupport('My Chase card stopped updating on Tuesday', true); const sm = mails[mails.length - 1];
check('email goes to Tomisin, reply-to the friend', sm.to === 'financehq.app@gmail.com' && sm.replyTo === 'me@x.com', sm.subject);
check('email has message + details, no amounts', /Chase card stopped/.test(sm.body) && /TECHNICAL DETAILS/.test(sm.body) && !/\$\s?\d/.test(sm.body));
const r2 = B.sendSupport('Second issue here', false);
check('box unticked → message only', !/TECHNICAL DETAILS/.test(mails[mails.length - 1].body));
let tooShort = ''; try { B.sendSupport('hi', true); } catch (e) { tooShort = e.message; }
check('empty message is refused', /little about/.test(tooShort));
for (let i = 0; i < 3; i++) B.sendSupport('Message number ' + i, false);
let capped = ''; try { B.sendSupport('Sixth message today', false); } catch (e) { capped = e.message; }
check('capped at 5 a day', /5 messages today/.test(capped));

console.log('\n10 · Two releases in one week');
{ const bump = v => NEW.replace("const FHQ_VERSION = '" + ver(NEW) + "'", "const FHQ_VERSION = '" + v + "'");
  drive[VFILE] = JSON.stringify({ version: '2099.01.01.1', notes: 'small fix', code: CFILE }); drive[CFILE] = bump('2099.01.01.1');
  P().UPDATE_TOLD_AT = String(Date.now() - 2 * 864e5);   // the new version emailed 2 days ago
  const n0 = mails.filter(m => /update/i.test(m.subject)).length; B.checkForUpdate_();
  console.log('DBG', mails.filter(m => /update/i.test(m.subject)).length, n0, B.updateLabel_(), PropertiesService.getScriptProperties().getProperty('UPDATE_TOLD_AT'), PropertiesService.getScriptProperties().getProperty('UPDATE_TOLD')); check('no second email within a week, menu still offers it', mails.filter(m => /update/i.test(m.subject)).length === n0 && /2099\.01\.01\.1/.test(B.updateLabel_()));
  P().UPDATE_TOLD_AT = String(Date.now() - 7 * 864e5); B.checkForUpdate_();
  check('a week later the newest one is emailed', mails.filter(m => /update/i.test(m.subject)).length === n0 + 1 && /2099\.01\.01\.1/.test(mails[mails.length - 1].subject)); }

console.log('\n11 · A friend with update emails turned off');
B.setSetting_ ? 0 : 0;
const S = M.data.Settings, row = S.findIndex(r => /Update emails/.test(r[0]));
if (row >= 0) S[row][1] = 'No';
drive[VFILE] = JSON.stringify({ version: '2026.12.9.9', notes: 'x', code: CFILE }); drive[CFILE] = NEW.replace("const FHQ_VERSION = '" + ver(NEW) + "'", "const FHQ_VERSION = '2026.12.9.9'");
const before = mails.length; B.checkForUpdate_();
check('no email, but the menu still offers the update', mails.length === before && /Install update \(2026\.12\.9\.9\)/.test(B.updateLabel_()), row >= 0 ? 'setting found' : 'setting row missing');

console.log(`\n${pass} passed, ${fail} failed`);
console.log('\n--- support email body sample ---\n' + mails.filter(m=>/help/i.test(m.subject))[0].body);
