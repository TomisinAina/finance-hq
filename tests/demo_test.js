process.env.FRESH = '1';
const M = require('/tmp/mock.js'); const fs = require('fs'), vm = require('vm');
const mails = []; global.MailApp = { sendEmail: (a, b) => mails.push(typeof a === 'object' ? a.subject : b) };
global.SF_HANDLER = () => { throw new Error('network should not be used in demo'); };
const code = fs.readFileSync('/home/claude/pub/FinanceHQ.gs', 'utf8');
const ctx = vm.createContext(global); vm.runInContext(code, ctx);
const L = console.log; console.log = () => {};
let err = ''; try { vm.runInContext('startDemo(); financeDailySync(); phase2Daily(); phase3Daily(); weeklyDigest(true);', ctx); } catch (e) { err = e.stack; }
console.log = L;
if (err) console.log('ERR', err.split('\n').slice(0, 3).join(' / '));
const show = n => (M.data[n] || []).map(r => r.filter(v => v !== '' && v != null).map(v => v instanceof Date ? v.toISOString().slice(0, 10) : typeof v === 'number' ? Math.round(v * 100) / 100 : String(v).slice(0, 60)).join(' | ')).filter(Boolean);
const d = show('Dashboard'); console.log(d.slice(0, 14).join('\n'));
console.log('\nTX rows', M.data.Transactions.length - 1, '| accounts', M.data.Accounts.length - 1, '| mails', mails.length, mails.slice(0, 3), '| cal', M.calls.cal.length);
console.log('HEALTH', show('Health').filter(l => /Look|Fix/.test(l)).join(' || '));
console.log('LOG errors', show('Log').filter(l => /error|fail/i.test(l)).slice(-3));
let c = ''; try { vm.runInContext('connectBank("x")', ctx); } catch (e) { c = e.message; } console.log('connect on demo:', c);
const app = JSON.parse(vm.runInContext('getAppData()', ctx)); console.log('app keys ok', !!app.tx, app.tx.length);
const before = M.data.Transactions.length; vm.runInContext('financeDailySync()', ctx); console.log('resync dupes?', M.data.Transactions.length - before);
