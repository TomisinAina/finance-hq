const SCEN = process.argv[2], NOW = new Date(process.argv[3] + 'T15:00:00-05:00').getTime();
const RD = Date; class FD extends RD { constructor(...a) { a.length ? super(...a) : super(NOW); } static now() { return NOW; } }
global.Date = FD;
process.env.FRESH = '1';
const M = require('/tmp/mock.js'); const fs = require('fs'), vm = require('vm');
const sf = JSON.parse(fs.readFileSync('/tmp/pay/' + SCEN + '.json', 'utf8'));
global.SF_HANDLER = (url) => {
  if (/claim/.test(url)) return { getResponseCode: () => 200, getContentText: () => 'https://u:p@bridge.test/simplefin' };
  const st = +url.match(/start-date=(\d+)/)[1], en = (url.match(/end-date=(\d+)/) || [])[1];
  const out = { errors: [], accounts: sf.accounts.map(a => Object.assign({}, a, { transactions: a.transactions.filter(t => t.posted >= st && (!en || t.posted < +en)) })) };
  return { getResponseCode: () => 200, getContentText: () => JSON.stringify(out) };
};
global.MailApp = { sendEmail: () => {} };
const code = fs.readFileSync(process.env.CODE || '/home/claude/pub/FinanceHQ.gs', 'utf8');
const ctx = vm.createContext(global); vm.runInContext(code, ctx);
const L = console.log; console.log = () => {};
try { vm.runInContext('setupFinanceHQ(); saveSetup(JSON.stringify(Object.assign(JSON.parse(getSetupData()), { save: 20 }))); phase2Daily();', ctx); } catch (e) { L('ERR', e.message); }
console.log = L;
const dash = (M.data.Dashboard || []).map(r => r.filter(v => v !== '' && v != null).map(String).join(' | '));
const line = dash.find(l => /pay ~?\$|no pay/i.test(l) && /spent/.test(l)) || '';
const tile = dash.findIndex(l => /LEFT TO SPEND/.test(l));
const val = tile >= 0 ? dash[tile + 1] : '';
const rules = (M.data.Rules || []).filter(r => /auto: looks like a paycheck/.test(r[2])).map(r => r[0]);
const health = (M.data.Health || []).filter(r => /Look|Fix/.test(r[0])).map(r => r[1]).join('; ');
const save = dash.find(l => /Move to savings/.test(l)) || '';
try{ fs.writeFileSync('/tmp/pay/app_'+SCEN+'.json', vm.runInContext('getAppData()', ctx)); }catch(e){ L('APPERR', e.message); }
console.log(JSON.stringify({ scen: SCEN, today: process.argv[3], left: val.split(' | ')[0], math: line.split(' | ').find(x => /spent/.test(x)), save: save.split(' | ').slice(0, 2).join(' '), autoRules: rules, health }));
