/**
 * FINANCE HQ — a free, private personal-finance tracker in Google Sheets.
 * Banks + cards sync daily through SimpleFIN Bridge (read-only, ~$15/yr). Everything else runs free on Google.
 * Dashboard · statement-period budgets · subscriptions · rewards check · loan planner · cash-flow forecast
 * alerts (email + phone reminders) · weekly digest · monthly PDF archive · phone web app.
 *
 * SETUP — see README. Short version:
 *   1. New Google Sheet -> Extensions -> Apps Script. Delete the sample code, paste this whole file, save.
 *   2. Project Settings (gear) -> Script properties -> add  SIMPLEFIN_SETUP_TOKEN = <your SimpleFIN setup token>.
 *   3. Run  setupFinanceHQ  and allow the permissions. Takes 1–3 minutes.
 *   4. Open the sheet and use the menu  Finance HQ → Quick setup…  — one short form for cards, budget, debts and trips.
 *   5. Optional phone app: Deploy -> New deployment -> Web app -> Execute as Me, access Only myself.
 *
 * Read-only: this can never move money. Your data stays in your own Google account.
 */

function setupFinanceHQ() {
  financeSetup();     // connect SimpleFIN, pull history, daily sync ~6 AM
  phase2Setup();      // Dashboard, Subscriptions, CardCheck, Debt, Settings; refresh ~7 AM
  archiveSetup();     // monthly PDF summary into Drive folder "Finance HQ Archive"
  autoRewards_();     // match your cards to known rewards rates
  phase3Setup();      // alerts ~8 AM, weekly digest Sundays 9 AM (sends a test digest now)
  PERIOD_ = null; rebuild_(true);   // final pass with the suggested limits in place
  Logger.log('Finance HQ is set up. Open the sheet, then use the menu  Finance HQ → Quick setup…  (one short form). ' + finSS_().getUrl());
}

// ===== PART 1 — sync engine (SimpleFIN -> Google Sheet) =====

// Which spreadsheet: MY_SHEET_ID (optional, in a personal file) > saved property > the sheet this script is
// attached to (Extensions > Apps Script) > a new "Finance HQ" sheet in your Drive.
let FIN_SS_CACHE_ = null;
function finSS_() {
  if (FIN_SS_CACHE_) return FIN_SS_CACHE_;
  const P = PropertiesService.getScriptProperties();
  const mine = typeof MY_SHEET_ID !== 'undefined' && MY_SHEET_ID, stored = P.getProperty('FIN_SHEET_ID');
  let active = null; try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  // A copied sheet brings its script (and saved settings) along. If this copy isn't the sheet those settings belong to,
  // start clean — so a copy can never reuse someone else's bank connection.
  if (!mine && active && stored && stored !== active.getId()) { P.deleteAllProperties(); Logger.log('New copy detected — starting fresh.'); }
  const id = mine || (active ? '' : P.getProperty('FIN_SHEET_ID'));
  let ss = id ? SpreadsheetApp.openById(id) : active;
  if (!ss) ss = SpreadsheetApp.create('Finance HQ');
  if (P.getProperty('FIN_SHEET_ID') !== ss.getId()) P.setProperty('FIN_SHEET_ID', ss.getId());
  return (FIN_SS_CACHE_ = ss);
}
const SF_ACCESS = 'SIMPLEFIN_ACCESS_URL';
const DAY = 24 * 3600 * 1000;

const TABS = {
  Accounts:     ['Account ID', 'Institution', 'Account', 'Type', 'Balance', 'Available', 'Balance date', 'Last synced', 'Status'],
  Transactions: ['Date', 'Account', 'Institution', 'Description', 'Amount', 'Category', 'Matched rule', 'Txn ID', 'Account ID'],
  Balances:     ['Date', 'Account ID', 'Account', 'Type', 'Balance'],
  NetWorth:     ['Date', 'Assets', 'Liabilities', 'Net worth', 'Cash', 'Investments', 'Credit cards', 'Loans'],
  Manual:       ['Name', 'Type', 'Balance', 'Rate', 'Monthly payment', 'Notes'],
  Rules:        ['If description contains', 'Category', 'Notes'],
  Log:          ['When', 'Message']
};

// First match wins, so specific rules go above general ones (e.g. UBER EATS before UBER).
const DEFAULT_RULES = [
  ['PAYMENT THANK YOU', 'Transfer', 'card payment'], ['AUTOPAY', 'Transfer', ''], ['AUTOMATIC PAYMENT', 'Transfer', ''],
  ['ONLINE PAYMENT', 'Transfer', ''], ['CARD PAYMENT', 'Transfer', ''], ['MOBILE PAYMENT', 'Transfer', ''],
  ['CHASE CREDIT CRD', 'Transfer', ''], ['AMEX EPAYMENT', 'Transfer', ''], ['AMERICAN EXPRESS ACH', 'Transfer', ''],
  ['STRIPE TRANSFER', 'Income', 'freelance payouts'], ['TRANSFER', 'Transfer', ''], ['ROBINHOOD', 'Investing', ''], ['VANGUARD', 'Investing', ''], ['FIDELITY', 'Investing', ''],
  ['DEPT EDUCATION', 'Student loans', ''], ['STUDENT LN', 'Student loans', ''], ['AIDVANTAGE', 'Student loans', ''],
  ['MOHELA', 'Student loans', ''], ['NELNET', 'Student loans', ''], ['EDFINANCIAL', 'Student loans', ''],
  ['PAYROLL', 'Income', ''], ['DIR DEP', 'Income', ''], ['DIRECT DEP', 'Income', ''], ['INTEREST PAID', 'Income', ''], ['DIVIDEND', 'Income', ''],
  ['INTEREST CHARGE', 'Fees', ''], ['LATE FEE', 'Fees', ''], ['ANNUAL FEE', 'Fees', ''], ['FOREIGN TRANSACTION', 'Fees', ''],
  ['UBER EATS', 'Dining', ''], ['DOORDASH', 'Dining', ''], ['GRUBHUB', 'Dining', ''], ['STARBUCKS', 'Dining', ''], ['CHIPOTLE', 'Dining', ''],
  ['SWEETGREEN', 'Dining', ''], ['CAVA', 'Dining', ''], ['MCDONALD', 'Dining', ''], ['DUNKIN', 'Dining', ''], ['CHICK-FIL-A', 'Dining', ''],
  ['TST*', 'Dining', 'Toast restaurant POS'], ['SQ *', 'Dining', 'Square — often food; fix individually'], ['RESTAURANT', 'Dining', ''], ['BAR & GRILL', 'Dining', ''],
  ['WHOLE FOODS', 'Groceries', ''], ['TRADER JOE', 'Groceries', ''], ['SAFEWAY', 'Groceries', ''], ['GIANT', 'Groceries', ''],
  ['HARRIS TEETER', 'Groceries', ''], ['WEGMANS', 'Groceries', ''], ['ALDI', 'Groceries', ''], ['LIDL', 'Groceries', ''], ['COSTCO', 'Groceries', ''], ['H MART', 'Groceries', ''],
  ['SHELL', 'Gas', ''], ['EXXON', 'Gas', ''], ['SUNOCO', 'Gas', ''], ['WAWA', 'Gas', ''], ['SHEETZ', 'Gas', ''], ['CHEVRON', 'Gas', ''], ['BP#', 'Gas', ''], ['7-ELEVEN', 'Gas', ''],
  ['UBER', 'Transport', ''], ['LYFT', 'Transport', ''], ['WMATA', 'Transport', ''], ['PARKING', 'Transport', ''], ['EZPASS', 'Transport', ''], ['E-ZPASS', 'Transport', ''],
  ['NETFLIX', 'Subscriptions', ''], ['SPOTIFY', 'Subscriptions', ''], ['HULU', 'Subscriptions', ''], ['HBO', 'Subscriptions', ''], ['DISNEY', 'Subscriptions', ''],
  ['APPLE.COM/BILL', 'Subscriptions', ''], ['YOUTUBE', 'Subscriptions', ''], ['ANTHROPIC', 'Subscriptions', ''], ['CLAUDE', 'Subscriptions', ''], ['OPENAI', 'Subscriptions', ''],
  ['MICROSOFT', 'Subscriptions', ''], ['PLAYSTATION', 'Subscriptions', ''], ['XBOX', 'Subscriptions', ''], ['PRIME VIDEO', 'Subscriptions', ''], ['SIMPLEFIN', 'Subscriptions', ''],
  ['DELTA AIR', 'Travel', ''], ['UNITED AIR', 'Travel', ''], ['AMERICAN AIR', 'Travel', ''], ['SOUTHWEST', 'Travel', ''], ['JETBLUE', 'Travel', ''],
  ['HYATT', 'Travel', ''], ['HILTON', 'Travel', ''], ['MARRIOTT', 'Travel', ''], ['AIRBNB', 'Travel', ''], ['EXPEDIA', 'Travel', ''], ['CHASE TRAVEL', 'Travel', ''],
  ['VERIZON', 'Bills & utilities', ''], ['T-MOBILE', 'Bills & utilities', ''], ['AT&T', 'Bills & utilities', ''], ['COMCAST', 'Bills & utilities', ''], ['XFINITY', 'Bills & utilities', ''],
  ['PEPCO', 'Bills & utilities', ''], ['DOMINION', 'Bills & utilities', ''], ['BGE', 'Bills & utilities', ''], ['WASHINGTON GAS', 'Bills & utilities', ''], ['GEICO', 'Bills & utilities', 'insurance'], ['PROGRESSIVE', 'Bills & utilities', 'insurance'],
  ['CVS', 'Health', ''], ['WALGREENS', 'Health', ''], ['PHARMACY', 'Health', ''],
  ['AMAZON', 'Shopping', ''], ['AMZN', 'Shopping', ''], ['TARGET', 'Shopping', ''], ['WALMART', 'Shopping', ''], ['BEST BUY', 'Shopping', ''], ['NIKE', 'Shopping', ''], ['UNIQLO', 'Shopping', '']
];

// Loans and anything SimpleFIN can't see go in the Manual tab: Name | Loan | Balance | Rate (4.5%) | Monthly payment | Notes
const DEFAULT_MANUAL = [];

// ---------------- setup ----------------
function financeSetup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(SF_ACCESS)) {
    const tok = (props.getProperty('SIMPLEFIN_SETUP_TOKEN') || '').trim();
    if (!tok) throw new Error('Add SIMPLEFIN_SETUP_TOKEN under Project Settings -> Script properties first.');
    const claimUrl = Utilities.newBlob(Utilities.base64Decode(tok)).getDataAsString().trim();
    const res = UrlFetchApp.fetch(claimUrl, { method: 'post', payload: '', muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      throw new Error('SimpleFIN claim failed (' + res.getResponseCode() + '). Setup tokens work only once - ' +
                      'generate a new one in SimpleFIN and replace SIMPLEFIN_SETUP_TOKEN.');
    }
    props.setProperty(SF_ACCESS, res.getContentText().trim());
    props.deleteProperty('SIMPLEFIN_SETUP_TOKEN');
    Logger.log('SimpleFIN connected.');
  }
  ensureTabs_();
  financeBackfill();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'financeDailySync')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('financeDailySync').timeBased().everyDays(1).atHour(6).create();
  Logger.log('Daily sync scheduled (~6 AM).');
}

// ---------------- sync ----------------
function financeBackfill() {
  // Pulls history in 30-day windows going back up to 12 months, stopping once a bank has nothing older
  // (most connections give ~90 days). SimpleFIN limits requests per day, so don't run this often.
  const now = Date.now();
  let total = 0, last = null, months = 0;
  for (let i = 0; i < 12; i++) {
    const end = now - i * 30 * DAY, start = end - 30 * DAY;
    const data = sfFetch_(start, end);
    if (i === 0) last = data;
    const added = writeTransactions_(data);
    total += added;
    months = i + 1;
    if (i >= 2 && added === 0) break;   // no older history available
  }
  Logger.log('History reached back about ' + months * 30 + ' days.');
  writeAccounts_(last);
  snapshot_();
  sortTransactions_();
  log_('Backfill done: ' + total + ' transactions, ' + (last.accounts || []).length + ' accounts.');
}

function financeDailySync() {
  const data = sfFetch_(Date.now() - 10 * DAY, null);
  const added = writeTransactions_(data);
  writeAccounts_(data);
  snapshot_();
  sortTransactions_();
  log_('Daily sync: ' + added + ' new transactions, ' + (data.accounts || []).length + ' accounts.');
}

function sfFetch_(start, end) {
  const access = PropertiesService.getScriptProperties().getProperty(SF_ACCESS);
  if (!access) throw new Error('Not connected - run financeSetup first.');
  if (access === DEMO_ACCESS) return demoBank_(start, end);
  const m = access.match(/^(https?:\/\/)([^:]+):([^@]+)@(.+)$/);
  let url = m[1] + m[4] + '/accounts?start-date=' + Math.floor(start / 1000);
  if (end) url += '&end-date=' + Math.floor(end / 1000);
  const auth = Utilities.base64Encode(decodeURIComponent(m[2]) + ':' + decodeURIComponent(m[3]));
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Basic ' + auth }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('SimpleFIN error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  const data = JSON.parse(res.getContentText());
  (data.errors || []).forEach(e => log_('SimpleFIN: ' + e));
  return data;
}

// ---------------- writers ----------------
function writeTransactions_(data) {
  const sh = sheet_('Transactions');
  const ids = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 8, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const rules = loadRules_();
  const rows = [];
  (data.accounts || []).forEach(a => {
    (a.transactions || []).forEach(t => {
      if (t.pending || ids.has(String(t.id))) return;
      const desc = [t.description, t.payee].filter(Boolean).join(' | ') || t.memo || '';
      const [cat, rule] = categorize_(desc, rules, t.amount);
      rows.push([new Date(((t.transacted_at || t.posted) * 1000)), a.name, orgName_(a), desc, Number(t.amount),
                 cat, rule, String(t.id), a.id]);
      ids.add(String(t.id));
    });
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

function writeAccounts_(data) {
  const sh = sheet_('Accounts');
  const existing = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, TABS.Accounts.length).getValues() : [];
  const byId = {}; existing.forEach((r, i) => byId[r[0]] = i);
  const now = new Date();
  const seen = new Set();
  (data.accounts || []).forEach(a => {
    seen.add(a.id);
    const row = [a.id, orgName_(a), a.name, '', Number(a.balance), a['available-balance'] != null ? Number(a['available-balance']) : '',
                 new Date(Number(a['balance-date']) * 1000), now, 'OK'];
    if (a.id in byId) {
      const old = existing[byId[a.id]];
      row[3] = old[3] || guessType_(a);           // keep a Type you edited
      existing[byId[a.id]] = row;
    } else {
      row[3] = guessType_(a);
      existing.push(row);
    }
  });
  existing.forEach(r => { if (!seen.has(r[0])) r[8] = 'Not returned by SimpleFIN - check connection'; });
  if (existing.length) sh.getRange(2, 1, existing.length, TABS.Accounts.length).setValues(existing);
}

function snapshot_() {
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const acc = sheet_('Accounts'), bal = sheet_('Balances'), nw = sheet_('NetWorth'), man = sheet_('Manual');
  const accounts = acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, TABS.Accounts.length).getValues() : [];
  const manual = man.getLastRow() > 1 ? man.getRange(2, 1, man.getLastRow() - 1, 3).getValues().filter(r => r[0] !== '') : [];

  // one balance snapshot per account per day
  const lastBalDate = bal.getLastRow() > 1 ? fmt_(bal.getRange(bal.getLastRow(), 1).getValue(), tz) : '';
  if (lastBalDate !== today && accounts.length) {
    bal.getRange(bal.getLastRow() + 1, 1, accounts.length, 5)
       .setValues(accounts.map(r => [new Date(), r[0], r[2], r[3], r[4]]));
  }

  let cash = 0, inv = 0, cards = 0, loans = 0, otherAssets = 0, otherDebt = 0;
  accounts.forEach(r => {
    const t = String(r[3]), b = Number(r[4]) || 0;
    if (t === 'Credit card') cards += Math.abs(b); // card balance = what you owe (SimpleFIN usually shows it negative)
    else if (t === 'Loan') loans += Math.abs(b);
    else if (t === 'Investment') inv += b;
    else if (t === 'Checking' || t === 'Savings') cash += b;
    else if (b >= 0) otherAssets += b; else otherDebt += -b;
  });
  const linkedNames = new Set(accounts.filter(r => r[3] === 'Loan').map(r => String(r[2]).trim().toUpperCase()));
  manual.forEach(r => {
    const b = Number(r[2]) || 0;
    if (linkedNames.has(String(r[0]).trim().toUpperCase())) return;     // already counted from the linked account
    if (isDebt_(r[1], r[0], '')) loans += Math.abs(b); else otherAssets += b;
  });
  const assets = cash + inv + otherAssets, liab = cards + loans + otherDebt;
  const row = [new Date(), round_(assets), round_(liab), round_(assets - liab), round_(cash), round_(inv), round_(cards), round_(loans)];
  const lastNw = nw.getLastRow() > 1 ? fmt_(nw.getRange(nw.getLastRow(), 1).getValue(), tz) : '';
  if (lastNw === today) nw.getRange(nw.getLastRow(), 1, 1, row.length).setValues([row]);
  else nw.appendRow(row);
}

function sortTransactions_() {
  const sh = sheet_('Transactions');
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, TABS.Transactions.length).sort({ column: 1, ascending: false });
}

// ---------------- categories ----------------
function recategorize() {
  // Re-applies Rules to rows still marked Uncategorized (your manual edits are never overwritten).
  const sh = sheet_('Transactions');
  if (sh.getLastRow() < 2) return;
  const rules = loadRules_();
  const rng = sh.getRange(2, 4, sh.getLastRow() - 1, 4); // D:G
  const vals = rng.getValues();
  let n = 0;
  vals.forEach(r => {
    if (r[2] === 'Uncategorized' || r[2] === '') {
      const [cat, rule] = categorize_(r[0], rules, r[1]);
      if (cat !== 'Uncategorized') { r[2] = cat; r[3] = rule; n++; }
    }
  });
  rng.setValues(vals);
  Logger.log('Recategorized ' + n + ' transactions.');
}

function loadRules_() {
  const sh = sheet_('Rules');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(r => r[0] !== '' && r[1] !== '').map(r => [String(r[0]).toUpperCase(), String(r[1])]);
}

// A rule can also require an amount: "TRANSFER TO CHECKING @150" = that text AND exactly $150 (in or out).
function categorize_(desc, rules, amt) {
  const d = String(desc).toUpperCase(), a = Math.abs(Number(amt) || 0);
  for (const [k, c] of rules) {
    const m = k.match(/^(.*?)\s*@\s*\$?([\d,]+(?:\.\d+)?)\s*$/);
    if (m) { if (amt !== undefined && Math.abs(a - Number(m[2].replace(/,/g, ''))) < 0.005 && (!m[1] || d.indexOf(m[1].trim()) !== -1)) return [c, 'rule: ' + k]; continue; }
    if (d.indexOf(k) !== -1) return [c, 'rule: ' + k];
  }
  return ['Uncategorized', ''];
}

// ---------------- helpers ----------------
function ensureTabs_() {
  const ss = finSS_();
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  const rules = ss.getSheetByName('Rules');
  if (rules.getLastRow() < 2) rules.getRange(2, 1, DEFAULT_RULES.length, 3).setValues(DEFAULT_RULES);
  const man = ss.getSheetByName('Manual');
  if (man.getLastRow() < 2 && DEFAULT_MANUAL.length) man.getRange(2, 1, DEFAULT_MANUAL.length, 6).setValues(DEFAULT_MANUAL);
  const money = '$#,##0.00;[Red]-$#,##0.00';
  ss.getSheetByName('Accounts').getRange('E:F').setNumberFormat(money);
  ss.getSheetByName('Transactions').getRange('E:E').setNumberFormat(money);
  ss.getSheetByName('Transactions').getRange('A:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName('Balances').getRange('E:E').setNumberFormat(money);
  ss.getSheetByName('NetWorth').getRange('B:H').setNumberFormat(money);
  ss.getSheetByName('Manual').getRange('C:C').setNumberFormat(money);
  ss.getSheetByName('Manual').getRange('D:D').setNumberFormat('0.00%');
  const first = ss.getSheets()[0];
  if (first.getName() !== 'Accounts' && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

function sheet_(name) {
  const ss = finSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) { ensureTabs_(); sh = ss.getSheetByName(name) || ss.insertSheet(name); }
  return sh;
}

function orgName_(a) { return (a.org && (a.org.name || a.org.domain)) || ''; }

function guessType_(a) {
  const s = (a.name + ' ' + orgName_(a)).toLowerCase().replace(/credit union/g, '');
  if (/loan|mortgage|aidvantage|mohela|nelnet|edfinancial|student/.test(s)) return 'Loan';
  if (/credit|card|sapphire|gold|platinum|freedom|amex|american express|visa|mastercard/.test(s)) return 'Credit card';
  if (/401|ira|roth|brokerage|individual|invest|robinhood|vanguard|fidelity|schwab|retirement/.test(s)) return 'Investment';
  if (/saving|high yield|hysa|money market/.test(s)) return 'Savings';
  if (/checking|debit|spend/.test(s)) return 'Checking';
  return Number(a.balance) < 0 ? 'Credit card' : 'Checking';
}

function log_(msg) {
  Logger.log(msg);
  try { sheet_('Log').appendRow([new Date(), msg]); } catch (e) {}
}
function fmt_(d, tz) { return d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d).slice(0, 10); }
function round_(x) { return Math.round(x * 100) / 100; }


// ===== PART 2 — analysis: Dashboard, budget periods, subscriptions, CardCheck, debt plan =====

const P2_TABS = {
  Dashboard: [],
  Trips: ['Trip', 'Start', 'End', 'Notes'],
  Subscriptions: ['Merchant', 'Typical charge', 'Cadence', 'Last charged', 'Next expected', 'Per year', 'Price change?', 'Charges seen'],
  CardCheck: ['Date', 'Merchant', 'Category', 'Amount', 'Card used', 'Better card', 'Points missed'],
  Rewards: ['Card', 'Category', 'Points per $'],
  Debt: [],
  Settings: ['Setting', 'Value', 'Notes']
};

const DEFAULT_SETTINGS = [
  ['Dining target per month', '', 'Your dining budget per period — suggested from your spending, edit me'],
  ['Debt payments per month (total)', '', 'Only if you did not enter payments per debt on the Manual tab'],
  ['Safety net target (months)', 3, 'Months of spending to keep in cash'],
  ['Cash held for others', 0, 'Money in your accounts that is not yours (subtracted from your cash)'],
  ['Travel fund per month', '', 'Optional'],
  ['Expected yearly purchases', '', 'Known annual buys, so they are not flagged as unusual'],
  ['Spending month (Auto / Statement / Calendar)', 'Auto', 'Auto = statement periods when most spending is on cards']
];

// Trips: add rows in the Trips tab (Trip | Start | End) so trip spending is kept apart from everyday spending.
const DEFAULT_TRIPS = [];

// Card rewards live in the Rewards tab: Card | Category | Points per $ (card names as shown on the Dashboard).
// Add an "Everything else" row per card for its base rate. Leave the tab empty to skip the CardCheck.
function loadRewards_() {
  const sh = sheet_('Rewards');
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, 3).setValues([['Card', 'Category', 'Points per $']]).setFontWeight('bold');
  const o = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const cat = String(r[1]).trim(), card = String(r[0]).trim();
    if (card && cat && Number(r[2]) > 0) (o[cat] = o[cat] || {})[card] = Number(r[2]);
  });
  return o;
}

// Rules added on top of Phase 1's list (first match wins)
const P2_RULES = [
  ['MOBILE PAYMENT', 'Transfer', 'card payment'], ['PAYMENT THANK YOU', 'Transfer', 'card payment'], ['AUTOPAY PAYMENT', 'Transfer', 'card payment'],
  ['EPAYMENT', 'Transfer', 'card payment'], ['CRD PMT', 'Transfer', 'card payment'], ['CARD PMT', 'Transfer', 'card payment'], ['PAYMENT RECEIVED', 'Transfer', 'card payment'],
  ['MORTGAGE', 'Housing', ''], ['RENT PAYMENT', 'Housing', ''], ['RENTCAFE', 'Housing', ''], ['APPFOLIO', 'Housing', ''], ['BILT RENT', 'Housing', ''], ['ZILLOW RENT', 'Housing', ''],
  ['TOYOTA FIN', 'Loan payments', ''], ['HONDA FIN', 'Loan payments', ''], ['GM FINANCIAL', 'Loan payments', ''], ['FORD CREDIT', 'Loan payments', ''],
  ['ALLY AUTO', 'Loan payments', ''], ['CAPITAL ONE AUTO', 'Loan payments', ''], ['CHASE AUTO', 'Loan payments', ''], ['CARMAX AUTO', 'Loan payments', ''], ['TESLA FIN', 'Loan payments', ''],
  ['NAVIENT', 'Student loans', ''], ['SALLIE MAE', 'Student loans', ''], ['TITHE', 'Giving', ''], ['PUSHPAY', 'Giving', ''], ['GOFUNDME', 'Giving', ''], ['DONATION', 'Giving', ''], ['STATE FARM', 'Insurance', ''], ['ALLSTATE', 'Insurance', ''], ['LEMONADE', 'Insurance', ''],
  ['ONLINE PAYMENT', 'Transfer', 'card payment'], ['THANK YOU-MOBILE', 'Transfer', 'card payment'], ['DUNKIN', 'Dining', ''],
  ['TRF TO', 'Transfer', ''], ['TRF FR', 'Transfer', ''], ['TFR FR', 'Transfer', ''], ['STRIPE TRANSFER', 'Income', 'freelance payouts'], ['TRANSFER', 'Transfer', ''],
  ['ZELLE', 'P2P transfers', ''], ['VENMO', 'P2P transfers', ''], ['CASH APP', 'P2P transfers', ''], ['APPLE CASH', 'P2P transfers', ''], ['AMEX SEND', 'P2P transfers', ''],
  ['BRANCH DEPOSIT', 'Transfer', ''], ['REAL TIME DEPOSIT', 'Transfer', ''],
  ['AFFIRM', 'Affirm', ''], ['KLARNA', 'Affirm', ''], ['AFTERPAY', 'Affirm', ''],
  ['FANDUEL', 'Sports betting', ''], ['BETFAIR', 'Sports betting', ''], ['DRAFTKINGS', 'Sports betting', ''],
  ['ATM REBATE', 'Fees', 'rebate'], ['ATM FEE', 'Fees', ''], ['ATM WITHDRAWAL', 'Cash (ATM)', ''], ['SHR WITHDRAWAL', 'Cash (ATM)', ''],
  ['MEMBERSHIP FEE', 'Fees', 'card annual fee'], ['DINING CREDIT', 'Dining', 'card credit'],
  ['PLANET FITNESS', 'Gym', ''], ['LA FITNESS', 'Gym', ''], ['EQUINOX', 'Gym', ''], ['ONELIFE', 'Gym', ''],
  ['THECUT', 'Haircuts', ''], ['BOOKSY', 'Haircuts', ''],
  ['PLAYSTATION', 'Games', ''], ['XBOX', 'Games', ''], ['STEAM', 'Games', ''], ['NINTENDO', 'Games', ''], ['PARAMOUNT', 'Subscriptions', ''],
  ['ENTERPRISE RENT', 'Travel', 'rental car'], ['NATIONAL CAR', 'Travel', ''], ['RENTALTOLL', 'Travel', ''], ['ERACTOLL', 'Travel', ''],
  ['AIRPORT', 'Travel', ''], ['FLIXBUS', 'Travel', ''], ['WAYMO', 'Transport', ''],
  ['STUBHUB', 'Entertainment', ''], ['AMC ', 'Entertainment', ''], ['FEVO', 'Entertainment', ''], ['ARMK', 'Entertainment', 'arena'],
  ['CAR WASH', 'Car & errands', ''], ['WHIP CLEAN', 'Car & errands', ''], ['DMV', 'Car & errands', ''], ['CLEANERS', 'Car & errands', ''],
  ['COLPARK', 'Car & errands', 'parking'], ['AUTOZONE', 'Car & errands', ''], ['PAYMENTUS', 'Bills & utilities', ''],
  ['SPA', 'Personal care', ''], ['MOBIL', 'Gas', ''], ['ROYAL FARMS', 'Dining', ''], ['PUBLIX', 'Groceries', ''], ['SUPERMAX', 'Groceries', ''],
  ['DOLLAR TREE', 'Shopping', ''], ['MARSHALLS', 'Shopping', ''], ['HOLLISTER', 'Shopping', ''], ['TIKTOK SHOP', 'Shopping', ''], ['DEPOP', 'Shopping', ''], ['EDIKTED', 'Shopping', ''],
  ['TAVERN', 'Dining', ''], [' BAR', 'Dining', ''], ['GRILL', 'Dining', ''], ['CAFE', 'Dining', ''], ['RAMEN', 'Dining', ''], ['TACO', 'Dining', ''],
  ['WINGSTOP', 'Dining', ''], ['SHAKE SHACK', 'Dining', ''], ['POTBELLY', 'Dining', ''], ['BAKERY', 'Dining', ''], ['POLLO', 'Dining', ''], ['PIZZA', 'Dining', '']
];

const NOT_SPENDING = ['Transfer', 'Investing', 'Investing (brokerage)', 'Income', 'Reimbursement', 'P2P transfers'];

// ---------------- setup & daily ----------------
function phase2Setup() {
  ensureTabs_();                       // Phase 1 tabs
  const ss = finSS_();
  Object.keys(P2_TABS).forEach(n => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (P2_TABS[n].length && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, P2_TABS[n].length).setValues([P2_TABS[n]]).setFontWeight('bold'); sh.setFrozenRows(1);
    }
  });
  const st = ss.getSheetByName('Settings');
  if (st.getLastRow() < 2) st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  const tr = ss.getSheetByName('Trips');
  if (tr.getLastRow() < 2 && DEFAULT_TRIPS.length) tr.getRange(2, 1, DEFAULT_TRIPS.length, 4).setValues(DEFAULT_TRIPS);
  addRules_();
  fixAccountTypes_();
  recategorizeAll_();
  ss.setActiveSheet(ss.getSheetByName('Dashboard')); ss.moveActiveSheet(1);
  phase2Daily();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'phase2Daily').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('phase2Daily').timeBased().everyDays(1).atHour(7).create();
  Logger.log('Phase 2 ready. Dashboard, Subscriptions, CardCheck and Debt built; refreshes daily ~7 AM.');
}

function phase2Daily() { try { afterUpdate_(); checkForUpdate_(); } catch (e) {} try { ensureCategoryLock_(); ensureYearlyTrigger_(); } catch (e) {} try { checkinIfDue_(); } catch (e) {} try { feedbackDue_(); } catch (e) {} PropertiesService.getScriptProperties().deleteProperty('DASH_VIEW'); rebuild_(true); try { archiveIfDue_(); } catch (e) { log_('Monthly archive failed: ' + e); } try { paydayNudge_(); } catch (e) { log_('Payday nudge skipped: ' + e); } }
function rebuild_(withStyle) {
  if (addRules_() | autoIncomeRules_()) recategorizeAll_();   // new rules -> re-apply to history once
  ensureSettings_();
  recategorize();                      // Phase 1: fills any new Uncategorized rows
  try { ensureYearlyTab_(); } catch (e) {}
  const tx = loadTx_();
  const step = (name, f) => { try { return f(); } catch (e) { log_(name + ' failed: ' + (e && e.message ? e.message : e)); } };
  const debt = step('Debt tab', () => buildDebt_()) || '';
  step('Subscriptions', () => buildSubscriptions_(tx));
  const missed = step('CardCheck', () => buildCardCheck_(tx)) || 0;
  step('Dashboard', () => buildDashboard_(tx, missed, debt));
  step('Budget', () => buildBudget_(tx));
  step('Insights', () => buildInsights_(tx));
  try { const S = settings_(); writeForecast_(forecast_(tx, S, cardCycles_(tx, S, new Date()), 45)); } catch (e) { log_('Forecast skipped: ' + e); }
  try { if (typeof buildHealth_ === 'function') buildHealth_(); } catch (e) { log_('Health check skipped: ' + e); }
  if (withStyle) try { styleAll_(); } catch (e) { log_('Formatting skipped: ' + e); }
}

// ---------------- data ----------------
function loadTx_() {
  const sh = sheet_('Transactions');
  if (sh.getLastRow() < 2) return [];
  const trips = loadTrips_();
  const acc = {}; const ash = sheet_('Accounts');
  if (ash.getLastRow() > 1) ash.getRange(2, 1, ash.getLastRow() - 1, 4).getValues().forEach(r => acc[r[0]] = r[3]);
  const out = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().filter(r => r[0] instanceof Date).map(r => {
    const t = { date: r[0], account: String(r[1]), inst: String(r[2]), desc: String(r[3]), amt: Number(r[4]) || 0,
                cat: String(r[5]), id: r[7], type: acc[r[8]] || '' };
    t.trip = '';
    if (t.type === 'Credit card' || t.type === 'Checking') {
      for (const tp of trips) if (t.date >= tp.start && t.date < tp.endNext) { t.trip = tp.name; break; }
    }
    if (NOT_SPENDING.indexOf(t.cat) !== -1 || t.cat === 'Student loans' || t.cat === 'Affirm') t.trip = '';
    t.month = Utilities.formatDate(t.date, Session.getScriptTimeZone(), 'yyyy-MM');
    return t;
  });
  const P = periodCfg_(out);
  out.forEach(t => { t.calMonth = t.month; t.month = txPeriod_(t, P); });
  return out;
}

// ---------- Spending month: statement periods vs calendar months ----------
// "Nov 2026" in statement mode = the period that closes in November (e.g. Oct 21 – Nov 20).
// Card purchases follow their own card's closing day; checking spending and pay follow the latest closing day.
let PERIOD_ = null;
function periodCfg_(tx) {
  if (PERIOD_) return PERIOD_;
  const S = settings_();
  const mode = String(S['Spending month (Auto / Statement / Calendar)'] || 'Auto').trim().toLowerCase();
  const cuts = {};
  Object.keys(S).forEach(k => {
    const m = String(k).match(/^(.+) closes \(day of month\)$/), d = Number(S[k]);
    if (m && cardTracked_(m[1], S) && d >= 1 && d <= 31) cuts[m[1]] = d;
  });
  const days = Object.keys(cuts).map(k => cuts[k]);
  const since = new Date(Date.now() - 90 * 864e5);
  const sp = (tx || []).filter(t => t.date >= since && t.amt < 0 && NOT_SPENDING.indexOf(t.cat) === -1 &&
    t.type !== 'Investment' && t.cat !== 'Student loans');
  const tot = sp.reduce((s, t) => s - t.amt, 0), card = sp.filter(t => t.type === 'Credit card').reduce((s, t) => s - t.amt, 0);
  const share = tot ? card / tot : 0;
  const on = days.length > 0 && (mode.indexOf('statement') === 0 || (mode.indexOf('cal') !== 0 && share >= 0.6));
  PERIOD_ = { on: on, cuts: cuts, main: days.length ? Math.max.apply(null, days) : 0, share: share };
  return PERIOD_;
}
function periodKeyFor_(d, cut) {
  const tz = Session.getScriptTimeZone();
  if (!cut) return Utilities.formatDate(d, tz, 'yyyy-MM');
  const y = d.getFullYear(), m = d.getMonth(), c = Math.min(cut, new Date(y, m + 1, 0).getDate());
  return Utilities.formatDate(d.getDate() > c ? new Date(y, m + 1, 1) : new Date(y, m, 1), tz, 'yyyy-MM');
}
function txPeriod_(t, P) {
  if (!P.on) return periodKeyFor_(t.date, 0);
  const cut = t.type === 'Credit card' ? (P.cuts[cardName_(t.account, t.inst)] || P.main) : P.main;
  return periodKeyFor_(t.date, cut);
}
function currentPeriod_(when) { const P = PERIOD_ || periodCfg_(loadTx_()); return periodKeyFor_(when || new Date(), P.on ? P.main : 0); }
function periodBounds_(key) {
  const P = PERIOD_ || { on: false };
  const [y, m] = key.split('-').map(Number);
  if (!P.on) return [new Date(y, m - 1, 1), new Date(y, m, 0)];
  const clamp = (yy, mm) => new Date(yy, mm, Math.min(P.main, new Date(yy, mm + 1, 0).getDate()));
  return [new Date(clamp(y, m - 2).getTime() + 864e5), clamp(y, m - 1)];
}
function periodRange_(key) {
  const tz = Session.getScriptTimeZone(), b = periodBounds_(key);
  return Utilities.formatDate(b[0], tz, 'MMM d') + ' – ' + Utilities.formatDate(b[1], tz, 'MMM d');
}

function loadTrips_() {
  const sh = sheet_('Trips');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().filter(r => r[0] && r[1] && r[2]).map(r => {
    const s = parseDay_(r[1]), e = parseDay_(r[2]);
    return s && e ? { name: r[0], start: s, endNext: new Date(e.getFullYear(), e.getMonth(), e.getDate() + 1) } : null;
  }).filter(Boolean);
}

// Earlier periods to average over — only ones your history fully covers (a half-synced first month would drag averages down).
function priorPeriods_(key, n, tx) {
  const first = tx.length ? tx.reduce((m, t) => t.date < m ? t.date : m, tx[0].date) : null;
  const start = first ? new Date(first.getFullYear(), first.getMonth(), first.getDate()) : null;
  const [y, m] = key.split('-').map(Number), tz = Session.getScriptTimeZone(), out = [], partial = [];
  for (let i = 1; i <= n; i++) {
    const k = Utilities.formatDate(new Date(y, m - 1 - i, 1), tz, 'yyyy-MM');
    const b = periodBounds_(k);
    if (!start || b[0] >= start) out.push(k); else if (b[1] >= start) partial.push(k);
  }
  return out.length ? out : partial;
}
function settings_() {
  const sh = sheet_('Settings'); const o = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => o[r[0]] = r[1]);
  return o;
}

function isSpend_(t) {
  if (NOT_SPENDING.indexOf(t.cat) !== -1 || t.type === 'Investment') return false;
  // money coming INTO checking/savings that isn't categorized is a deposit, not a refund
  if (t.amt > 0 && t.type !== 'Credit card' && (!t.cat || t.cat === 'Uncategorized')) return false;
  return true;
}

// Statement cycles per card. Settings: "<Card> closes (day of month)" and "<Card> due (day of month)".
function cardCycles_(tx, S, now) {
  const tz = Session.getScriptTimeZone();
  const acc = sheet_('Accounts');
  const list = acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, 5).getValues().filter(x => x[3] === 'Credit card') : [];
  const clampDay = (y, m, d) => new Date(y, m, Math.min(d, new Date(y, m + 1, 0).getDate()));
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return list.filter(a => cardTracked_(cardName_(a[2], a[1]), S)).map(a => {
    const name = cardName_(a[2], a[1]);
    const cl = Number(S[name + ' closes (day of month)']), du = Number(S[name + ' due (day of month)']);
    const configured = cl >= 1 && cl <= 31 && du >= 1 && du <= 31;
    let lastClose, nextClose, prevClose, due;
    if (configured) {
      lastClose = clampDay(today.getFullYear(), today.getMonth(), cl);
      if (lastClose > today) lastClose = clampDay(today.getFullYear(), today.getMonth() - 1, cl);
      nextClose = clampDay(lastClose.getFullYear(), lastClose.getMonth() + 1, cl);
      prevClose = clampDay(lastClose.getFullYear(), lastClose.getMonth() - 1, cl);
      due = clampDay(lastClose.getFullYear(), lastClose.getMonth(), du);
      if (due <= lastClose) due = clampDay(lastClose.getFullYear(), lastClose.getMonth() + 1, du);
    } else {           // fallback: calendar months
      lastClose = new Date(today.getFullYear(), today.getMonth(), 0);
      nextClose = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      prevClose = new Date(today.getFullYear(), today.getMonth() - 1, 0);
      due = null;
    }
    const mine = tx.filter(t => t.account === a[2]);
    const isPay = t => t.amt > 0 && (t.cat === 'Transfer' || /PAYMENT|THANK YOU/i.test(t.desc));
    const spent = -mine.filter(t => t.date > lastClose && !isPay(t)).reduce((s, t) => s + t.amt, 0);
    const paidAmt = mine.filter(t => t.date > lastClose && isPay(t)).reduce((s, t) => s + t.amt, 0);
    // Statement balance from the bank's real balance: owed now − new charges since close + payments since close
    const owed = -Number(a[4]) || 0;
    const statement = configured ? Math.max(0, owed - spent + paidAmt)
      : Math.max(0, -mine.filter(t => t.date > prevClose && t.date <= lastClose && !isPay(t)).reduce((s, t) => s + t.amt, 0));
    const paid = statement <= 1 || paidAmt >= statement - 1;
    const daysToDue = due ? Math.round((due - today) / 86400000) : 99;
    let status = 'set dates in Settings';
    if (due) status = paid ? 'Paid ✓' : (daysToDue < 0 ? 'Past due?' : daysToDue === 0 ? 'Due today' : 'Due in ' + daysToDue + ' days');
    const f = d => Utilities.formatDate(d, tz, 'MMM d');
    return { name: name, configured: configured, spent: Math.round(spent), statement: Math.round(statement), paid: paid, daysToDue: daysToDue,
             cycle: f(new Date(lastClose.getTime() + 86400000)) + ' – ' + f(nextClose), closes: f(nextClose),
             due: due ? f(due) : '—', status: status, account: a[2],
             dueISO: due ? Utilities.formatDate(due, tz, 'yyyy-MM-dd') : '', closesISO: Utilities.formatDate(nextClose, tz, 'yyyy-MM-dd'),
             start: Utilities.formatDate(new Date(lastClose.getTime() + 86400000), tz, 'yyyy-MM-dd') };
  });
}

function cardName_(account, inst) {
  // Optional nicknames: MY_CARD_NAMES = [[/gold/i, 'Amex Gold'], ...] in a personal file (matched on account + bank).
  const both = String(account) + ' ' + String(inst);
  if (typeof MY_CARD_NAMES !== 'undefined') for (const p of MY_CARD_NAMES) if (p[0].test(both)) return p[1];
  let n = String(account).replace(/\s*\(\d+\)\s*$/, '').replace(/[®™]/g, '').trim();
  const last4 = (String(account).match(/(\d{4})\)?\s*$/) || [])[1];
  if (/^(credit card|card|credit)$/i.test(n)) n = String(inst).replace(/\s*(bank|n\.?a\.?)$/i, '').trim() + ' card';
  return last4 && !/\d{4}/.test(n) ? n + ' ' + last4 : n;
}
function cardTracked_(name, S) { return !/^n/i.test(String((S || settings_())[name + ' track (Yes/No)'] || 'Yes')); }

function ensureSettings_() {
  const sh = sheet_('Settings');
  const have = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : []);
  const add = [];
  DEFAULT_SETTINGS.concat(GOAL_SETTINGS).forEach(s => {
    if (s[0] === 'Debt payments per month (total)' && have.has('Student loan payment per month')) return;
    if (!have.has(s[0])) add.push(s);
  });
  const acc = sheet_('Accounts');
  if (acc.getLastRow() > 1) acc.getRange(2, 1, acc.getLastRow() - 1, 4).getValues().filter(x => x[3] === 'Credit card').forEach(x => {
    const n = cardName_(x[2], x[1]);
    [[n + ' closes (day of month)', '', 'statement closing day, e.g. 15 — on your statement'],
     [n + ' due (day of month)', '', 'payment due day, e.g. 10'],
     [n + ' track (Yes/No)', 'Yes', 'No = hide this card (e.g. an add-on you do not pay)']].forEach(s => { if (!have.has(s[0]) && !add.find(a => a[0] === s[0])) add.push(s); });
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
}

// ---------------- subscriptions ----------------
function buildSubscriptions_(tx) {
  const sh = sheet_('Subscriptions');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent();
  if (sh.getRange(1, 9).getValue() === '') sh.getRange(1, 9).setValue('Status');
  const marks = getSubStatus_();
  const groups = {};
  const notSub = ['Student loans', 'Loan payments', 'Housing', 'Affirm', 'Cash (ATM)', 'Fees', 'Uncategorized'];
  tx.filter(t => t.amt < 0 && isSpend_(t) && (t.type === 'Credit card' || t.type === 'Checking') && notSub.indexOf(t.cat) === -1)
    .forEach(t => { const k = merchantKey_(t.desc); (groups[k] = groups[k] || []).push(t); });
  const now = new Date();
  const out = [];
  Object.keys(groups).forEach(k => {
    const g = groups[k].sort((a, b) => a.date - b.date);
    // keep the most common amount band (subscriptions repeat at ~the same price)
    const amts = g.map(t => -t.amt);
    const ref = median_(amts);
    const same = g.filter(t => Math.abs(-t.amt - ref) <= Math.max(1, ref * 0.15));
    if (same.length < 2) return;
    const gaps = []; for (let i = 1; i < same.length; i++) gaps.push((same[i].date - same[i - 1].date) / 86400000);
    const gap = median_(gaps);
    const subLike = ['Subscriptions', 'Dating apps', 'Personal subscriptions', 'Gym', 'Games'].indexOf(same[same.length - 1].cat) !== -1;
    const everyday = ['Dining', 'Groceries', 'Gas', 'Cash (ATM)', 'Shopping', 'Transport', 'Fees'].indexOf(same[same.length - 1].cat) !== -1;
    let cadence = '';
    if (gap >= 25 && gap <= 35 && !everyday) cadence = 'Monthly';
    else if (gap >= 6 && gap <= 8 && subLike) cadence = 'Weekly';
    else if (gap >= 80 && gap <= 100 && !everyday) cadence = 'Quarterly';
    else return;
    const last = same[same.length - 1], first = same[0];
    const stopped = (now - last.date) / 86400000 > gap * 1.8 + 3;
    if (stopped && !marks[k]) return;                                   // stopped and never marked — drop it
    const per = cadence === 'Monthly' ? 12 : cadence === 'Weekly' ? 52 : 4;
    const change = Math.abs(-last.amt - -first.amt) > 0.5 ? (-last.amt > -first.amt ? 'UP from $' : 'down from $') + (-first.amt).toFixed(2) : '';
    out.push([k, -last.amt, cadence, last.date, stopped ? 'stopped ✓' : new Date(last.date.getTime() + gap * 86400000), Math.round(-last.amt * per), change, same.length, (marks[k] || {}).s || '']);
  });
  out.sort((a, b) => (a[8] === 'cancelled') - (b[8] === 'cancelled') || b[5] - a[5]);
  if (out.length) {
    sh.getRange(2, 1, out.length, 9).setValues(out);
    sh.getRange(2, 2, out.length, 1).setNumberFormat('$#,##0.00'); sh.getRange(2, 6, out.length, 1).setNumberFormat('$#,##0');
    sh.getRange(2, 4, out.length, 2).setNumberFormat('mmm d');
  }
  const r = out.length + 3;
  sh.getRange(r, 1, 2, 2).setValues([['Total per year', out.filter(x => x[8] !== 'cancelled').reduce((s, x) => s + x[5], 0)],
                                     ['Known yearly purchases', settings_()['Expected yearly purchases'] || '']]).setFontWeight('bold');
  sh.getRange(r, 2).setNumberFormat('$#,##0');
}

function merchantKey_(desc) {
  return String(desc).toUpperCase().split('|')[0]
    .replace(/APLPAY |TST\* ?|SQ \* ?|POS DEBIT[- ]+VISA CHECK CARD \d+ - |POS DEBIT- +\d+ \d+ /g, '')
    .replace(/[0-9#*.]+/g, ' ').replace(/\s+/g, ' ').trim().split(' ').slice(0, 2).join(' ');
}
function median_(a) { if (!a.length) return 0; const s = a.slice().sort((x, y) => x - y); const m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }

// ---------------- card check ----------------
function buildCardCheck_(tx) {
  const sh = sheet_('CardCheck');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  const R = loadRewards_(), base = R['Everything else'] || {}, S = settings_();
  const since = new Date(Date.now() - 90 * 86400000);
  const rows = [];
  const allCards = [...new Set([].concat.apply([], Object.keys(R).map(c => Object.keys(R[c]))))].filter(c => cardTracked_(c, S));
  if (allCards.length < 2) return 0;
  tx.filter(t => t.date >= since && t.amt < 0 && isSpend_(t) && t.type === 'Credit card').forEach(t => {
    const card = cardName_(t.account, t.inst);
    if (!cardTracked_(card, S)) return;
    const rateOf = c => (R[t.cat] && R[t.cat][c]) || base[c] || 1;
    const best = allCards.reduce((b, c) => rateOf(c) > rateOf(b) ? c : b, allCards[0]);
    const own = rateOf(card);
    if (!best || card === best || rateOf(best) <= own) return;
    const missed = Math.round(-t.amt * (rateOf(best) - own));
    if (missed > 0) rows.push([t.date, merchantKey_(t.desc), t.cat, -t.amt, card, best, missed]);
  });
  rows.sort((a, b) => b[0] - a[0]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 7).setValues(rows);
    sh.getRange(2, 4, rows.length, 1).setNumberFormat('$#,##0.00'); sh.getRange(2, 1, rows.length, 1).setNumberFormat('mmm d');
  }
  return rows.reduce((s, r) => s + r[6], 0);
}

// ---------------- debt planner (any kind of debt) ----------------
const DEBT_TYPES = ['Credit card debt', 'Personal loan', 'Medical debt', 'Car loan', 'Student loan (private)', 'Student loan (federal)', 'Mortgage', 'Other debt'];
const ASSET_TYPE = 'Asset (not debt)';
const DEBT_TIPS = {
  'Credit card debt': 'Top priority. Card interest (often 20%+) costs more than almost any investment earns. Pay more than the minimum, stop adding to the balance, and look at a 0% balance-transfer card if your credit is good.',
  'Personal loan': 'Usually 8–20%. Worth paying down early once any card debt is gone — check there is no prepayment penalty.',
  'Medical debt': 'Before paying interest on it, ask the provider for an itemized bill, financial assistance, or a 0% payment plan.',
  'Car loan': 'Above ~7%, extra payments are worth it. Under ~5%, paying on schedule and saving the difference is reasonable. Make sure extra money goes to principal.',
  'Student loan (private)': 'No forgiveness programs. If your rate is high and your credit has improved, refinancing can lower it. Pay these before low-rate federal loans.',
  'Student loan (federal)': 'If you are going for PSLF or income-driven forgiveness, extra payments may be wasted — check studentaid.gov first. Otherwise, loans under ~4–5% are often fine on schedule while you build savings and invest.',
  'Mortgage': 'Usually your cheapest debt. One extra payment a year can take years off a 30-year loan. Drop PMI once you reach 20% equity, and consider refinancing if rates fall ~1% below yours.',
  'Other debt': 'Pay highest interest first. Ask the lender whether extra payments go to principal.'
};

function debtType_(type, name, notes) {
  const t = String(type || '').trim();
  if (DEBT_TYPES.indexOf(t) !== -1 || t === ASSET_TYPE) return t;
  const s = (t + ' ' + name + ' ' + (notes || '')).toLowerCase();
  const servicer = /sallie|navient|nelnet|mohela|aidvantage|edfinancial|sofi|earnest|great lakes|fedloan/.test(s);
  if (!servicer && !/loan|debt|mortgage|card|owe|heloc|financ|lease/.test(s)) return ASSET_TYPE;
  if (/student|federal|direct|sallie|navient|nelnet|mohela|aidvantage|edfinancial|sofi|earnest|great lakes|fedloan/.test(s))
    return /private|sallie|sofi|earnest/.test(s) ? 'Student loan (private)' : 'Student loan (federal)';
  if (/mortgage|home loan|heloc|house|condo/.test(s)) return 'Mortgage';
  if (/\bauto\b|auto loan|\bcar\b|vehicle|lease|honda|toyota|ford|nissan|tesla|bmw|mercedes|hyundai|kia|chevy|jeep|subaru|lexus|audi|mazda|volkswagen/.test(s)) return 'Car loan';
  if (/card/.test(s)) return 'Credit card debt';
  if (/medical|hospital/.test(s)) return 'Medical debt';
  if (/personal/.test(s)) return 'Personal loan';
  return 'Other debt';
}
function isDebt_(type, name, notes) { return debtType_(type, name, notes) !== ASSET_TYPE; }

function linkedLoans_() {
  // Loan accounts linked through SimpleFIN: name -> live balance owed
  const a = sheet_('Accounts'), o = {};
  if (a.getLastRow() > 1) a.getRange(2, 1, a.getLastRow() - 1, 5).getValues().filter(r => r[3] === 'Loan').forEach(r => o[String(r[2]).trim().toUpperCase()] = { name: String(r[2]), inst: String(r[1]), bal: Math.abs(Number(r[4]) || 0) });
  return o;
}
function loadDebts_() {
  const man = sheet_('Manual'), linked = linkedLoans_(), used = {};
  const rng = man.getLastRow() > 1 ? man.getRange(2, 1, man.getLastRow() - 1, 6) : null, v = rng ? rng.getValues() : [];
  let changed = false;
  const debts = [];
  v.forEach(r => {
    const L = linked[String(r[0]).trim().toUpperCase()];
    if (L) { used[String(r[0]).trim().toUpperCase()] = 1; r[2] = L.bal; }            // live balance wins
    if (!r[0] || !(Number(r[2]) > 0)) return;
    const ty = debtType_(r[1], r[0], r[5]);
    if (ty !== String(r[1]).trim()) { r[1] = ty; changed = true; }      // tidy old "Loan" rows into a real type
    if (ty === ASSET_TYPE) return;
    let rate = Number(r[3]) || 0; if (rate > 1) rate = rate / 100;      // 4.5 -> 4.5%
    debts.push({ name: String(r[0]), type: ty, bal: Number(r[2]), rate: rate, pay: Number(r[4]) || 0 });
  });
  if (changed && rng) rng.setValues(v);
  Object.keys(linked).forEach(k => {
    if (used[k] || !(linked[k].bal > 0)) return;
    debts.push({ name: linked[k].name, type: debtType_('', linked[k].name, linked[k].inst), bal: linked[k].bal, rate: 0, pay: 0, noRate: true });
  });
  // No per-debt payments? Split a total from Settings by balance, else estimate a minimum.
  const S = settings_();
  const total = Number(S['Debt payments per month (total)']) || Number(S['Student loan payment per month']) || 0;
  const missing = debts.filter(d => !d.pay), missBal = missing.reduce((s, d) => s + d.bal, 0);
  missing.forEach(d => {
    d.estimated = !total;
    d.pay = total ? Math.round(total * d.bal / missBal * 100) / 100
      : d.type === 'Credit card debt' ? Math.max(25, d.bal * 0.03) : Math.max(25, d.bal * d.rate / 12 + d.bal / 120);
  });
  return debts;
}

function simulateDebts_(debts, extra, method) {
  const b = debts.map(d => ({ name: d.name, bal: d.bal, rate: d.rate, pay: d.pay, done: 0, interest: 0 }));
  let m = 0, interest = 0;
  while (b.some(x => x.bal > 0.01) && m < 600) {
    m++;
    let pool = extra;
    b.forEach(x => { if (x.bal > 0.01) { const i = x.bal * x.rate / 12; x.bal += i; x.interest += i; interest += i; } });
    b.forEach(x => { if (x.bal > 0.01) x.bal -= Math.min(x.pay, x.bal); else pool += x.pay; });   // finished debts free up their payment
    b.filter(x => x.bal > 0.01).sort(method === 'snowball' ? (p, q) => p.bal - q.bal : (p, q) => q.rate - p.rate)
      .forEach(x => { if (pool > 0) { const a = Math.min(pool, x.bal); x.bal -= a; pool -= a; } });
    b.forEach(x => { if (x.bal <= 0.01 && !x.done) x.done = m; });
  }
  return { months: b.some(x => x.bal > 0.01) ? null : m, interest: interest, debts: b };
}
// kept for older callers
function simulate_(loans, payment) {
  const per = loans.length ? payment / loans.length : 0;
  const r = simulateDebts_(loans.map(l => ({ name: l.name, bal: l.bal, rate: l.rate, pay: per })), 0, 'avalanche');
  return { months: r.months || 600, interest: r.interest, order: r.debts.slice().sort((a, c) => a.done - c.done).map(d => ({ name: d.name, month: d.done, rate: d.rate })) };
}

function buildDebt_() {
  const sh = sheet_('Debt'), tz = Session.getScriptTimeZone();
  sh.clear(); sh.getRange('A:H').breakApart(); sh.setConditionalFormatRules([]);
  uiPage_(sh);
  sh.setHiddenGridlines(true);
  for (let c = 1; c <= 8; c++) sh.setColumnWidth(c, c === 2 ? 190 : 118);
  sh.getRange('A:H').setFontFamily('Arial').setFontColor(UI.ink).setVerticalAlignment('middle');
  const debts = loadDebts_();
  const when = n => n === null ? 'Never at this payment' : Utilities.formatDate(new Date(new Date().getFullYear(), new Date().getMonth() + n, 1), tz, 'MMM yyyy');
  sh.getRange('A1:F1').merge().setValue('Debt plan').setFontSize(18).setFontWeight('bold');
  sh.getRange('G1:H1').merge().setValue('Updated ' + Utilities.formatDate(new Date(), tz, 'MMM d')).setFontColor(UI.mute).setHorizontalAlignment('right');
  if (!debts.length) {
    sh.getRange('A3:H3').merge().setValue('No debts yet. Add them on the Manual tab: Name | Type | Balance | Rate | Monthly payment.').setFontColor(UI.mute);
    return 'no debts on the Manual tab';
  }
  const base = simulateDebts_(debts, 0, 'avalanche');
  const owed = debts.reduce((s, d) => s + d.bal, 0), pays = debts.reduce((s, d) => s + d.pay, 0);
  let r = 3;
  r = uiTiles_(sh, r, [
    ['TOTAL OWED', owed, '$#,##0', UI.bad, debts.length + ' debt' + (debts.length > 1 ? 's' : '')],
    ['MONTHLY PAYMENTS', pays, '$#,##0', UI.ink, debts.some(d => d.estimated) ? 'some estimated — add yours on Manual' : 'from your Manual tab'],
    ['DEBT-FREE', when(base.months), '@', base.months === null ? UI.bad : UI.good, base.months === null ? 'payments too low' : Math.round(base.months / 12 * 10) / 10 + ' years'],
    ['INTEREST STILL TO PAY', base.interest, '$#,##0', UI.warn, 'at your current payments']
  ]);

  r = uiBand_(sh, r, 'Your debts — extra money goes to #1 first (highest rate)');
  r = uiHeader_(sh, r, ['#', 'Debt', 'Type', 'Balance', 'Rate', 'Payment', 'Paid off', 'Interest left']);
  const order = debts.map((d, i) => Object.assign({}, d, { sim: base.debts[i] })).sort((a, c) => c.rate - a.rate);
  const t0 = r;
  order.forEach((d, i) => {
    sh.getRange(r, 7).setNumberFormat('@');
    sh.getRange(r, 1, 1, 8).setValues([[i + 1, d.name, d.type, d.bal, d.rate, d.pay, when(d.sim.done || null), d.sim.interest]]);
    if (d.estimated) sh.getRange(r, 6).setNote('Estimated — put your real payment in the Manual tab.');
    if (d.noRate) sh.getRange(r, 5).setNote('Interest rate unknown — add this debt on the Manual tab (same name) with its rate.');
    r++;
  });
  sh.getRange(t0, 4, order.length, 1).setNumberFormat('$#,##0'); sh.getRange(t0, 5, order.length, 1).setNumberFormat('0.00%');
  sh.getRange(t0, 6, order.length, 1).setNumberFormat('$#,##0'); sh.getRange(t0, 8, order.length, 1).setNumberFormat('$#,##0');
  sh.getRange(t0, 1, order.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  uiStripes_(sh, t0, order.length, 8);
  r++;

  r = uiBand_(sh, r, 'What paying a little extra each month does');
  r = uiHeader_(sh, r, ['Extra / month', '', 'Debt-free', 'Total interest', 'Interest saved', 'Months sooner', '', '']);
  const e0 = r;
  [0, 50, 100, 250, 500].forEach(x => {
    const s = simulateDebts_(debts, x, 'avalanche');
    sh.getRange(r, 3).setNumberFormat('@');
    sh.getRange(r, 1, 1, 6).setValues([[x ? '+$' + x : 'Current plan', '', when(s.months), s.interest,
      x ? base.interest - s.interest : '', x && base.months && s.months ? base.months - s.months : '']]);
    r++;
  });
  sh.getRange(e0, 4, 5, 2).setNumberFormat('$#,##0'); sh.getRange(e0, 5, 5, 2).setFontColor(UI.good).setFontWeight('bold');
  uiStripes_(sh, e0, 5, 8);
  r++;

  if (debts.length > 1) {
    const sn = simulateDebts_(debts, 50, 'snowball'), av = simulateDebts_(debts, 50, 'avalanche');
    r = uiBand_(sh, r, 'Two ways to attack it (with +$50/month)');
    r = uiHeader_(sh, r, ['Method', '', 'How it works', '', '', 'Debt-free', 'Total interest', '']);
    const m0 = r;
    sh.getRange(m0, 6, 2, 1).setNumberFormat('@');
    sh.getRange(r, 1, 2, 7).setValues([
      ['Avalanche', '', 'Highest interest rate first — saves the most money', '', '', when(av.months), av.interest],
      ['Snowball', '', 'Smallest balance first — quick wins that keep you motivated', '', '', when(sn.months), sn.interest]]);
    sh.getRange(m0, 7, 2, 1).setNumberFormat('$#,##0');
    [m0, m0 + 1].forEach(x => sh.getRange(x, 3, 1, 3).merge().setWrap(true));
    uiStripes_(sh, m0, 2, 8); r += 3;
  }

  r = uiBand_(sh, r, 'Tips for the debts you have');
  [...new Set(debts.map(d => d.type))].sort((a, c) => DEBT_TYPES.indexOf(a) - DEBT_TYPES.indexOf(c)).forEach(ty => {
    sh.getRange(r, 1, 1, 2).merge().setValue(ty).setFontWeight('bold').setVerticalAlignment('top');
    sh.getRange(r, 3, 1, 6).merge().setValue(DEBT_TIPS[ty] || DEBT_TIPS['Other debt']).setWrap(true).setVerticalAlignment('top');
    sh.setRowHeight(r, 48); r++;
  });
  sh.getRange(r + 1, 1, 1, 8).merge().setValue('General information, not financial advice. Rates and balances come from your Manual tab.').setFontColor(UI.mute).setFontSize(9);
  const first = order[0];
  return '$' + Math.round(owed).toLocaleString() + ' — debt-free ' + when(base.months) + '; pay extra on ' + first.name + ' (' + (first.rate * 100).toFixed(1) + '%) first';
}

// ===== PART 6 — look & feel (every tab except the Dashboard) =====
const UI = { ink: '#2A2A2A', mute: '#8A857A', band: '#2F6F4E', tile: '#FFFFFF', stripe: '#FAF9F6', good: '#2F6F4E', bad: '#B8423A', warn: '#B86A2B', accent: '#3C6C8A', edit: '#FFF8E1', page: '#F7F6F2', line: '#E6E4DD', head: '#EFEDE6' };
const TAB_GUIDE = [
  // [tab, colour, what it is, do you edit it?]
  ['Dashboard', '#2F6F4E', 'Your money at a glance: what is left, bills due, spending by category, month by month.', 'No'],
  ['Budget', '#2F6F4E', 'Your limit for each category, suggestions that keep your savings goal on track, and how each period went. Change a limit right here.', 'Limits only'],
  ['Insights', '#2F6F4E', 'Plain-English notes: are you on pace, what is over, what to move to savings, what changed since last period.', 'No'],
  ['Start here', '#2F6F4E', 'This guide.', 'No'],
  ['Health', '#2F6F4E', 'Checks that everything is working and flags anything that looks wrong — with how to fix it.', 'No'],
  ['Transactions', '#3C6C8A', 'Every purchase, payment and deposit. Change a Category with the dropdown — it stays fixed after that.', 'Category only'],
  ['Accounts', '#3C6C8A', 'Every linked account and its latest balance.', 'No'],
  ['Subscriptions', '#3C6C8A', 'Charges that repeat (streaming, apps, gym). Mark each one keep / could cancel / cancelled — it feeds the Savings tracker.', 'Status only'],
  ['Forecast', '#3C6C8A', 'Where your checking balance is headed over the next 45 days.', 'No'],
  ['Debt', '#3C6C8A', 'Your debt plan: payoff dates, interest, and what paying extra would do.', 'No'],
  ['CardCheck', '#3C6C8A', 'Purchases where another card would have earned more points.', 'No'],
  ['NetWorth', '#3C6C8A', 'A daily record of what you own minus what you owe.', 'No'],
  ['Settings', '#C89B1E', 'Your savings goals, budget limits, card dates and alert choices. Yellow cells are yours to change.', 'Yes'],
  ['Rules', '#C89B1E', 'How purchases get a category: "if the description contains X, use category Y".', 'Yes'],
  ['Trips', '#C89B1E', 'Trip dates, so travel spending is kept apart from everyday spending.', 'Yes'],
  ['Manual', '#C89B1E', 'Debts and anything your bank link cannot see (student, car, mortgage loans…).', 'Yes'],
  ['Yearly costs', '#C89B1E', 'Bills that come once a year (annual card fee, car insurance paid yearly, Prime…). 1/12 is set aside every period so they never surprise you.', 'Yes'],
  ['Rewards', '#C89B1E', 'Points per $ for each card, so CardCheck knows your best card.', 'Optional'],
  ['Balances', '#9CA3AF', 'Behind-the-scenes balance history.', 'No (hidden)'],
  ['Log', '#9CA3AF', 'Behind-the-scenes sync log.', 'No (hidden)']
];

function uiBand_(sh, r, text) {
  sh.getRange(r, 1, 1, 8).merge().setValue(text.toUpperCase()).setBackground(UI.head).setFontColor('#5C574D')
    .setFontWeight('bold').setFontSize(9).setBorder(true, false, true, false, false, false, UI.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(r, 30); return r + 1;
}
function uiPage_(sh) {   // cream page like KJ's dashboard
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 80), 8).setBackground(UI.page);
}
function uiHeader_(sh, r, labels) {
  sh.getRange(r, 1, 1, labels.length).setValues([labels]).setFontColor(UI.mute).setFontSize(9).setFontWeight('bold');
  return r + 1;
}
function uiStripes_(sh, r, n, w) {
  for (let i = 0; i < n; i++) sh.getRange(r + i, 1, 1, w).setBackground(i % 2 ? UI.stripe : '#FFFFFF');
  if (n > 0) sh.getRange(r, 1, n, w).setBorder(false, true, true, true, false, false, UI.line, SpreadsheetApp.BorderStyle.SOLID);
}
function uiTiles_(sh, r, tiles) {
  tiles.forEach((t, i) => {
    const c = 1 + i * 2;
    sh.getRange(r, c, 1, 2).merge().setValue(t[0]).setFontSize(9).setFontColor(UI.mute).setFontWeight('bold');
    const v = sh.getRange(r + 1, c, 1, 2).merge().setNumberFormat(t[2]).setFontSize(18).setFontWeight('bold').setFontColor(t[3]);
    v.setValue(t[1]);
    sh.getRange(r + 2, c, 1, 2).merge().setValue(t[4]).setFontSize(9).setFontColor(UI.mute);
    sh.getRange(r, c, 3, 2).setBackground(UI.tile).setHorizontalAlignment('left')
      .setBorder(true, true, true, true, false, false, UI.line, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.setRowHeight(r + 1, 36);
  return r + 4;
}

// Plain table look: dark header, zebra rows, no gridlines, frozen header, friendly widths.
function styleTable_(sh, widths, opts) {
  opts = opts || {};
  const n = widths.length, maxR = Math.max(sh.getMaxRows(), 2);
  sh.setHiddenGridlines(true);
  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, maxR, n).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1, maxR, n).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(UI.head).setFirstRowColor('#FFFFFF').setSecondRowColor(UI.stripe);
  sh.getRange(1, 1, 1, n).setFontColor('#5C574D').setFontWeight('bold').setWrap(true).setFontSize(9)
    .setBorder(false, false, true, false, false, false, UI.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(1, 34); sh.setFrozenRows(1);
  widths.forEach((w, i) => { if (w === 0) sh.hideColumns(i + 1); else { sh.showColumns(i + 1); sh.setColumnWidth(i + 1, w); } });
  if (opts.headers) sh.getRange(1, 1, 1, opts.headers.length).setValues([opts.headers]);
  if (opts.notes) opts.notes.forEach((t, i) => { if (t) sh.getRange(1, i + 1).setNote(t); });
  (opts.money || []).forEach(c => sh.getRange(2, c, maxR - 1, 1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00'));
  (opts.dates || []).forEach(c => sh.getRange(2, c, maxR - 1, 1).setNumberFormat('mmm d, yyyy'));
  (opts.edit || []).forEach(c => sh.getRange(2, c, maxR - 1, 1).setBackground(UI.edit));
  if (opts.greenPositive) {
    const rg = sh.getRange(2, opts.greenPositive, maxR - 1, 1);
    sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(UI.good).setRanges([rg]).build()]);
  }
}
function dropdown_(rng, list) {
  rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build());
}

function categoryList_() {
  const set = new Set(['Dining', 'Groceries', 'Gas', 'Shopping', 'Travel', 'Transport', 'Entertainment', 'Bills & utilities', 'Health',
    'Subscriptions', 'Gym', 'Haircuts', 'Personal care', 'Car & errands', 'Games', 'Cash (ATM)', 'Fees', 'Student loans', 'Affirm',
    'Housing', 'Loan payments', 'Insurance', 'Giving', 'Kids', 'Pets', 'Gifts', 'Education',
    'Income', 'Reimbursement', 'Transfer', 'P2P transfers', 'Investing', 'Investing (brokerage)', 'Uncategorized']);
  const rs = sheet_('Rules');
  if (rs.getLastRow() > 1) rs.getRange(2, 2, rs.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) set.add(String(r[0])); });
  return [...set].sort();
}

function styleAll_() {
  const ss = finSS_(), S = settings_();
  const cats = categoryList_();

  // Start here guide
  const g = ss.getSheetByName('Start here') || ss.insertSheet('Start here');
  g.clear(); g.getBandings().forEach(b => b.remove()); g.setHiddenGridlines(true);
  g.getRange('A1').setValue('Welcome to Finance HQ').setFontSize(18).setFontWeight('bold');
  g.getRange('A2').setValue('Everything updates by itself each morning. Green tabs = overview, blue = your money, orange = things you can set, grey = behind the scenes.')
    .setFontColor(UI.mute).setWrap(false);
  g.getRange(4, 1, 1, 3).setValues([['Tab', 'What it shows', 'Do you edit it?']]);
  g.getRange(5, 1, TAB_GUIDE.length, 3).setValues(TAB_GUIDE.map(t => [t[0], t[2], t[3]]));
  TAB_GUIDE.forEach((t, i) => g.getRange(5 + i, 1).setFontColor(t[1]).setFontWeight('bold'));
  g.getRange(1, 1, Math.max(g.getMaxRows(), 40), 3).setBackground(UI.page);
  g.getRange(4, 1, 1, 3).setBackground(UI.head).setFontColor('#5C574D').setFontWeight('bold');
  uiStripes_(g, 5, TAB_GUIDE.length, 3);
  g.getRange(5, 2, TAB_GUIDE.length, 1).setWrap(true);
  g.setColumnWidth(1, 140); g.setColumnWidth(2, 560); g.setColumnWidth(3, 130);
  const tipRow = 6 + TAB_GUIDE.length;
  g.getRange(tipRow, 1, 4, 1).setValues([['Quick tips'], ['• Wrong category? Pick the right one in Transactions — it will not be changed again.'],
    ['• Same store keeps landing wrong? Add it to Rules (write "mine" in Notes so it wins).'], ['• Menu  Finance HQ → Quick setup…  edits cards, budget, debts and trips in one form. Finance HQ → Refresh now syncs right away.']]);
  g.getRange(tipRow, 1).setFontWeight('bold');
  const tr = tipRow + 5, terms = [['Terms & privacy'],
    ['• Your money data stays in this sheet in your own Google account. Finance HQ can only read your accounts — it can never move money.'],
    ['• Weekly check-ins: once a week this copy sends the person who shared it a few usage counts — your Google email, whether your bank is syncing, how many days you used the phone app and sheet, how many categories you fixed, whether transfers were ticked, and whether setup is complete.'],
    ['• Never sent: amounts, balances, account or card names, store names or transactions. See exactly what is sent: Finance HQ → What check-ins share. Turn it off: Settings → "Share usage check-ins…" → No.'],
    ['• Updates: once a day this copy checks whether the person who shared it has published a newer Finance HQ, and emails you if so. Nothing installs unless you paste it in (Finance HQ → ⬆ Install update). Emails off: Settings → "Update emails" → No.'],
    ['• Get help (Finance HQ → Get help…) emails the person who shared this from your own Gmail: your message, plus — only if you leave the box ticked — the technical details shown before you send (version, sync status, health checks, recent errors, with dollar amounts removed).'],
    ['• Provided as-is for personal use, not financial advice. Bank syncing is by SimpleFIN under your own SimpleFIN account.']];
  g.getRange(tr, 1, terms.length, 1).setValues(terms);
  g.getRange(tr, 1).setFontWeight('bold');
  g.getRange(tr + 1, 1, terms.length - 1, 1).setFontColor(UI.mute);
  g.getRange('A:C').setFontFamily('Arial').setVerticalAlignment('middle');

  // Tab colours, order, hidden tabs
  TAB_GUIDE.forEach((t, i) => {
    const sh = ss.getSheetByName(t[0]); if (!sh) return;
    sh.setTabColor(t[1]);
    if (sh.getIndex() !== i + 1) { ss.setActiveSheet(sh); ss.moveActiveSheet(Math.min(i + 1, ss.getNumSheets())); }
    if (/hidden/.test(t[3])) { if (!sh.isSheetHidden()) sh.hideSheet(); }
  });

  // Each tab
  const T = n => ss.getSheetByName(n);
  if (T('Transactions')) {
    const sh = T('Transactions');
    styleTable_(sh, [100, 200, 140, 330, 110, 160, 150, 0, 0], {
      headers: ['Date', 'Account', 'Bank', 'Description', 'Amount', 'Category', 'How it was categorized'],
      notes: ['', '', '', 'As your bank shows it', 'Money out is red, money in is green', 'Pick from the list to fix it. Your choice is locked in.', '"rule: X" = matched a Rule · "manual" = you set it'],
      money: [5], dates: [1], edit: [6], greenPositive: 5 });
    sh.getRange(2, 7, sh.getMaxRows() - 1, 1).setFontColor(UI.mute).setFontSize(9);
    dropdown_(sh.getRange(2, 6, sh.getMaxRows() - 1, 1), cats);
  }
  if (T('Accounts')) {
    const sh = T('Accounts');
    styleTable_(sh, [0, 190, 270, 110, 120, 120, 120, 120, 80], {
      headers: ['Account ID', 'Bank', 'Account', 'Type', 'Balance', 'Available', 'Balance as of', 'Last synced', 'Status'],
      notes: ['', '', '', 'Checking, Savings, Credit card, Investment or Loan', 'Card balances show as negative = what you owe'], money: [5, 6], dates: [7, 8] });
    const st = sh.getRange(2, 9, sh.getMaxRows() - 1, 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK').setFontColor(UI.good).setBold(true).setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextDoesNotContain('OK').whenCellNotEmpty().setFontColor(UI.bad).setBold(true).setRanges([st]).build()]);
  }
  if (T('Subscriptions')) {
    const sh = T('Subscriptions');
    styleTable_(sh, [220, 110, 95, 105, 105, 105, 140, 90, 130], {
      headers: ['Subscription', 'Usual charge', 'How often', 'Last charged', 'Next charge', 'Cost per year', 'Price change?', 'Times seen', 'Status'],
      notes: ['Found automatically from repeating charges', '', '', '', '', '', '', '', 'keep · could cancel · cancelled — feeds the Savings tracker on the Dashboard'], money: [2, 6], edit: [9] });
    const pc = sh.getRange(2, 7, sh.getMaxRows() - 1, 1), stc = sh.getRange(2, 9, sh.getMaxRows() - 1, 1);
    dropdown_(stc, SUB_STATUSES);
    sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextContains('UP').setFontColor(UI.bad).setBold(true).setRanges([pc]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('could cancel').setBackground(KJ.yellowBg).setFontColor('#8A6D16').setRanges([stc]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('cancelled').setBackground(KJ.greenBg).setFontColor(KJ.green).setRanges([stc]).build()]);
  }
  if (T('CardCheck')) styleTable_(T('CardCheck'), [100, 220, 130, 110, 160, 160, 120], {
    headers: ['Date', 'Where', 'Category', 'Amount', 'Card you used', 'Better card', 'Points missed'],
    notes: ['Last 90 days. Uses your Rewards tab.'], money: [4] });
  if (T('NetWorth')) styleTable_(T('NetWorth'), [120, 130, 130, 130, 130, 130, 130, 130], {
    headers: ['Date', 'What you own', 'What you owe', 'Net worth', 'Cash', 'Investments', 'Credit cards', 'Loans'],
    notes: ['One row per day'], money: [2, 3, 4, 5, 6, 7, 8], dates: [1], greenPositive: 4 });
  if (T('Rules')) {
    const sh = T('Rules');
    styleTable_(sh, [280, 200, 260], {
      headers: ['If the description contains…', '…use this category', 'Notes (write "mine" to make your rule win)'],
      notes: ['Not case-sensitive. First match from the top wins. Add @amount to also match the amount, e.g. "TRANSFER TO CHECKING @150".'], edit: [1, 2, 3] });
    dropdown_(sh.getRange(2, 2, sh.getMaxRows() - 1, 1), cats);
  }
  if (T('Trips')) {
    const sh = T('Trips');
    styleTable_(sh, [200, 130, 130, 300], { headers: ['Trip', 'First day', 'Last day', 'Notes'], dates: [2, 3], edit: [1, 2, 3, 4] });
  }
  if (T('Manual')) {
    const sh = T('Manual');
    styleTable_(sh, [220, 190, 130, 90, 140, 300], {
      headers: ['Name', 'Type', 'Balance', 'Interest rate', 'Monthly payment', 'Notes'],
      notes: ['e.g. "Honda loan"', 'Pick the kind of debt — the Debt tab gives tips for each', 'What you owe today', 'Yearly rate, e.g. 4.5%', 'Your required payment'],
      money: [3, 5], edit: [1, 2, 3, 4, 5, 6] });
    sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setNumberFormat('0.00%');
    dropdown_(sh.getRange(2, 2, sh.getMaxRows() - 1, 1), DEBT_TYPES.concat([ASSET_TYPE]));
  }
  if (T('Rewards')) {
    const sh = T('Rewards');
    styleTable_(sh, [220, 200, 120], { headers: ['Card', 'Category', 'Points per $'], notes: ['Card name as shown on the Dashboard', 'Add "Everything else" for the base rate'], edit: [1, 2, 3] });
    const cards = sheet_('Accounts').getLastRow() > 1 ? [...new Set(sheet_('Accounts').getRange(2, 1, sheet_('Accounts').getLastRow() - 1, 4).getValues()
      .filter(a => a[3] === 'Credit card').map(a => cardName_(a[2], a[1])))] : [];
    if (cards.length) dropdown_(sh.getRange(2, 1, sh.getMaxRows() - 1, 1), cards);
    dropdown_(sh.getRange(2, 2, sh.getMaxRows() - 1, 1), cats.concat(['Everything else']));
  }
  styleYearly_();
  styleSettings_();
  ensureCategoryLock_();
  if (typeof ensureMenu_ === 'function') ensureMenu_();
}

// Settings: grouped into sections, yellow = yours to change, Yes/No dropdowns.
function styleSettings_() {
  const sh = sheet_('Settings');
  if (sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().filter(r => r[0] !== '' && !/^— /.test(String(r[0])));
  if (!rows.length) return;
  const group = k => /closes \(day|due \(day|track \(Yes/.test(k) ? 'Cards' : /^Alert|alert|digest|insights|Bill reminder|check-ins|Payday|Update emails/i.test(k) ? 'Alerts and emails'
    : /% of pay\)/.test(k) ? 'Savings goals' : /monthly limit|Dining target/.test(k) ? 'Spending limits (per period)' : 'Budget basics';
  const order = ['Savings goals', 'Budget basics', 'Spending limits (per period)', 'Cards', 'Alerts and emails'];
  const out = [];
  order.forEach(gname => {
    const g = rows.filter(r => group(String(r[0])) === gname);
    if (gname === 'Spending limits (per period)') g.sort((a, b) => /Dining target/.test(a[0]) ? -1 : /Dining target/.test(b[0]) ? 1 : String(a[0]).localeCompare(String(b[0])));
    if (gname === 'Cards') g.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
    if (g.length) { out.push(['— ' + gname.toUpperCase() + ' —', '', '']); g.forEach(r => out.push(r)); }
  });
  sh.getBandings().forEach(b => b.remove());
  sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 3).clearContent().clearFormat().clearDataValidations();
  sh.getRange(2, 1, out.length, 3).setValues(out);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Your value', 'What it means']]).setBackground(UI.head).setFontColor('#5C574D').setFontWeight('bold');
  sh.setRowHeight(1, 34); sh.setFrozenRows(1);
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 150); sh.setColumnWidth(3, 460);
  sh.getRange(1, 1, out.length + 1, 3).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  out.forEach((r, i) => {
    const row = i + 2, k = String(r[0]);
    if (/^— /.test(k)) { sh.getRange(row, 1, 1, 3).setBackground(UI.page).setFontWeight('bold').setFontColor(UI.band).setFontSize(9); sh.setRowHeight(row, 30); return; }
    sh.getRange(row, 2).setBackground(UI.edit).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(row, 3).setFontColor(UI.mute).setWrap(true);
    if (/\(Yes\/No\)/.test(k)) dropdown_(sh.getRange(row, 2), ['Yes', 'No']);
    if (/Auto \/ Statement \/ Calendar/.test(k)) dropdown_(sh.getRange(row, 2), ['Auto', 'Statement', 'Calendar']);
    if (/monthly limit|target per month|alert over|under \(\$\)|payment per month|held for others|fund per month|payments per month/.test(k)) sh.getRange(row, 2).setNumberFormat('$#,##0');
    if (/% of pay\)/.test(k)) sh.getRange(row, 2).setNumberFormat('0"%"');
  });
}

// Changing a Category in Transactions locks it ("manual") so rules never overwrite it.
function ensureCategoryLock_() {
  // The edit trigger must belong to THIS spreadsheet; one left over from another copy never fires here.
  const id = finSS_().getId();
  let ok = false;
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onCategoryEdit').forEach(t => {
    let src = ''; try { src = t.getTriggerSourceId(); } catch (e) {}
    if (src === id && !ok) ok = true; else ScriptApp.deleteTrigger(t);
  });
  if (ok) return;
  try { ScriptApp.newTrigger('onCategoryEdit').forSpreadsheet(finSS_()).onEdit().create(); } catch (e) { log_('Category lock trigger not created: ' + e); }
}
function onCategoryEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet(), name = sh.getName();
  bumpUsage_('edit');
  if (name === 'Transactions' && e.range.getColumn() <= 6 && e.range.getLastColumn() >= 6 && e.range.getRow() >= 2)
    { sh.getRange(e.range.getRow(), 7, e.range.getNumRows(), 1).setValue('manual'); bumpUsage_('cat'); }   // your category choice is locked
  if (name === 'Dashboard' && e.range.getRow() === 2 && e.range.getColumn() === 7) {               // the month picker
    // Sheets may turn "Oct 2026" into a date, so accept a date or the text.
    const raw = e.range.getValue(), tz = Session.getScriptTimeZone();
    const v = raw instanceof Date ? Utilities.formatDate(raw, tz, 'MMM yyyy') : String(raw).replace(' (current)', '').trim();
    const mm = v.match(/^(\w{3})\w* (\d{4})$/);
    const mi = mm ? ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(mm[1].toLowerCase()) : -1;
    const key = mi >= 0 ? mm[2] + '-' + String(mi + 1).padStart(2, '0') : '';
    const props = PropertiesService.getScriptProperties();
    if (key && key !== currentPeriod_(new Date())) props.setProperty('DASH_VIEW', key); else props.deleteProperty('DASH_VIEW');
    toast_(finSS_(), 'Loading ' + (v || 'this period') + '…', 'Finance HQ', 10);
    PERIOD_ = null; const tx = loadTx_();
    buildDashboard_(tx, buildCardCheck_(tx), buildDebt_());
    return;
  }
  if (name === 'Dashboard' && e.range.getColumn() === 1 && e.range.getNumRows() === 1) {          // a "Transfers to make" checkbox
    const key = String(sh.getRange(e.range.getRow(), 9).getValue()).split('|');
    if (key.length === 2) {
      const on = e.range.getValue() === true;
      setTick(key[0], key[1], on);
      sh.getRange(e.range.getRow(), 2).setFontLine(on ? 'line-through' : 'none').setFontColor(on ? KJ.mute : KJ.ink);
      sh.getRange(e.range.getRow(), 7).setValue(on ? 'done ✓' : 'not yet').setFontColor(on ? KJ.green : KJ.orange);
    }
    return;
  }
  if (name === 'Subscriptions' && e.range.getColumn() <= 9 && e.range.getLastColumn() >= 9 && e.range.getRow() >= 2) {
    const r0 = e.range.getRow(), v = sh.getRange(r0, 1, e.range.getNumRows(), 9).getValues();
    v.forEach(x => { if (x[0] && !/total|known/i.test(x[0])) setSubStatus(String(x[0]), SUB_STATUSES.indexOf(String(x[8]).toLowerCase()) !== -1 ? String(x[8]).toLowerCase() : ''); });
    liveRebuild_(name); return;
  }
  if (name === 'Budget' && (e.range.getColumn() === 2 || e.range.getColumn() === 4) && e.range.getNumColumns() === 1) {   // a limit or "Use it?" box
    let changed = false;
    for (let i = 0; i < e.range.getNumRows(); i++) {
      const row = e.range.getRow() + i, key = String(sh.getRange(row, 9).getValue());
      if (key.indexOf('limit|') !== 0) continue;
      const cat = key.slice(6);
      if (e.range.getColumn() === 4) { if (sh.getRange(row, 4).getValue() === true && Number(sh.getRange(row, 3).getValue()) > 0) { setLimit_(cat, sh.getRange(row, 3).getValue()); changed = true; } }
      else { const v = sh.getRange(row, 2).getValue(); if (v === '' || Number(v) >= 0) { setLimit_(cat, v === '' ? '' : Number(v)); changed = true; } }
    }
    if (changed) { bumpUsage_('limit'); liveRebuild_('Settings'); }
    return;
  }
  if (LIVE_TABS.indexOf(name) !== -1 && e.range.getRow() >= 2) liveRebuild_(name);  // refresh the Dashboard right away
}

function formatSheets() { styleAll_(); }


// Paychecks that don't say "payroll": the same payer depositing a similar amount every 1–5 weeks becomes an Income rule.
function autoIncomeRules_() {
  const tx = loadTx_(), since = new Date(Date.now() - 100 * 864e5), groups = {};
  tx.filter(t => t.date >= since && t.amt >= 100 && (t.type === 'Checking' || t.type === 'Savings') && (t.cat === 'Uncategorized' || !t.cat))
    .forEach(t => { const k = incomeKey_(t.desc); if (k) (groups[k] = groups[k] || []).push(t); });
  const sh = sheet_('Rules');
  const have = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toUpperCase()) : []);
  const add = [];
  Object.keys(groups).forEach(k => {
    const g = groups[k].sort((a, b) => a.date - b.date);
    const days = [...new Set(g.map(t => Utilities.formatDate(t.date, Session.getScriptTimeZone(), 'yyyy-MM-dd')))];
    if (days.length < 2 || have.has(k)) return;
    const gaps = []; for (let i = 1; i < days.length; i++) gaps.push((parseDay_(days[i]) - parseDay_(days[i - 1])) / 864e5);
    const gap = median_(gaps), amts = g.map(t => t.amt), med = median_(amts);
    const steady = amts.filter(a => Math.abs(a - med) <= med * 0.35).length >= Math.ceil(amts.length * 0.6);
    const gig = /UPWORK|FIVERR|STRIPE|ETSY|SHOPIFY|TOPTAL|SQUARE INC|GUMROAD|PATREON|TWITCH|YOUTUBE|TIKTOK|SUBSTACK|WHOP|THUMBTACK|TASKRABBIT/.test(k) && g.length >= 3;
    if (((gap >= 6 && gap <= 35 && steady) || gig) && !/TRANSFER|TRF|ZELLE|VENMO|CASH APP|REFUND|RETURN|REVERSAL|INTEREST/i.test(k)) add.push([k, 'Income', 'auto: looks like a paycheck — change if not']);
  });
  if (!add.length) return false;
  sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  log_('Paycheck detected automatically: ' + add.map(a => a[0]).join(', '));
  return true;
}

function payeeKey_(desc) {
  const up = String(desc).toUpperCase().split('|')[0];
  const stop = /^(ACH|TRANSACTION|DEBIT|CREDIT|PAYMENT|PMT|WEB|BILL|PAY|PPD|CCD|ID|TO|ONLINE|ELECTRONIC|POS|PURCHASE|RECURRING|AUTOPAY|-)$/;
  const words = up.replace(/[^A-Z&' ]+/g, ' ').split(/\s+/).filter(w => w && w.length > 1 && !stop.test(w));
  for (let n = Math.min(2, words.length); n >= 1; n--) { const k = words.slice(0, n).join(' '); if (k.length >= 4 && up.indexOf(k) !== -1) return k; }
  return '';
}
function incomeKey_(desc) {
  // payer name without bank filler words ("Deposit - ACME CORP PAYROLL 1234" -> "ACME CORP PAYROLL"); must appear as-is in the description
  const up = String(desc).toUpperCase().split('|')[0];
  const stop = /^(DEPOSIT|DEP|ACH|CREDIT|DIRECT|PPD|CCD|ID|FROM|ELECTRONIC|ONLINE|CO|ENTRY|DESCR|EXTERNAL|INCOMING|TRANSFER|-)$/;
  const words = up.replace(/[^A-Z&' ]+/g, ' ').split(/\s+/).filter(w => w && !stop.test(w) && w.length > 1);
  for (let n = Math.min(3, words.length); n >= 1; n--) {
    const k = words.slice(0, n).join(' ');
    if (k.length >= 5 && up.indexOf(k) !== -1) return k;
  }
  return '';
}

// ---------------- one-time fixes ----------------
function addRules_() {
  // Order (first match wins): your rows marked "personal" or "mine" in Notes > built-in rules > everything else.
  // Returns true if anything changed, so history gets re-categorized once.
  const sh = sheet_('Rules');
  const n = Math.max(0, sh.getLastRow() - 1);
  const rows = n ? sh.getRange(2, 1, n, 3).getValues().filter(r => r[0] !== '') : [];
  const keys = new Set(P2_RULES.map(r => r[0].toUpperCase()));
  const mineFirst = r => /^(personal|mine)$/i.test(String(r[2]).trim());
  const bad = r => String(r[0]).trim().toUpperCase() === 'MOBIL' && /not "mobile"/.test(String(r[2]));
  const first = rows.filter(r => mineFirst(r) && !bad(r));
  const rest = rows.filter(r => !mineFirst(r) && !keys.has(String(r[0]).toUpperCase()) && !bad(r));
  const all = first.concat(P2_RULES.map(r => [r[0], r[1], r[2] || ''])).concat(rest);
  const same = all.length === rows.length && all.every((r, k) => String(r[0]) === String(rows[k][0]) && String(r[1]) === String(rows[k][1]));
  if (same) return false;
  if (n) sh.getRange(2, 1, n, 3).clearContent();
  sh.getRange(2, 1, all.length, 3).setValues(all);
  return true;
}

function fixAccountTypes_() {
  const sh = sheet_('Accounts');
  if (sh.getLastRow() < 2) return;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  v.forEach(r => {
    const s = (r[1] + ' ' + r[2]).toLowerCase();
    if (/credit union/.test(s) && r[3] === 'Credit card' && !/card|visa/.test(String(r[2]).toLowerCase())) {
      r[3] = /checking/.test(s) ? 'Checking' : 'Savings';
    }
  });
  sh.getRange(2, 1, v.length, 4).setValues(v);
  snapshot_();   // Phase 1: refresh today's net worth with the corrected types
}

function recategorizeAll_() {
  // Re-applies all rules to every transaction. Activity inside investment accounts -> Investing (brokerage).
  const sh = sheet_('Transactions');
  const acc = sheet_('Accounts'), inv = new Set(acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, 4).getValues().filter(a => a[3] === 'Investment').map(a => a[0]) : []);
  if (sh.getLastRow() < 2) return;
  const rules = loadRules_();
  const rng = sh.getRange(2, 1, sh.getLastRow() - 1, 9);
  const v = rng.getValues();
  v.forEach(r => {
    if (String(r[6]).toLowerCase() === 'manual') return;     // put "manual" in Matched rule to lock your own category
    if (inv.has(r[8])) { r[5] = 'Investing (brokerage)'; r[6] = 'account'; return; }
    const [cat, rule] = categorize_(r[3], rules, r[4]); r[5] = cat; r[6] = rule;
  });
  rng.setValues(v);
}


// ===== PART 3 — monthly PDF archive (Drive folder "Finance HQ Archive"), made the day after each period closes =====

const ARCHIVE_FOLDER = 'Finance HQ Archive';

function archiveSetup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'archiveMonthly').forEach(t => ScriptApp.deleteTrigger(t));
  const tx = loadTx_(), P = PERIOD_ || { on: false };
  // statement mode: run the morning after the period closes; calendar mode: the 1st
  const runDay = P.on && P.main < 28 ? P.main + 1 : 1;
  ScriptApp.newTrigger('archiveMonthly').timeBased().onMonthDay(runDay).atHour(8).create();
  ensureYearlyTrigger_();             // January: last year's "Year in Review" PDF
  const thisMonth = currentPeriod_();
  const months = [...new Set(tx.map(t => t.month))].filter(m => m < thisMonth).sort();
  months.forEach(m => archiveMonth_(m, tx));
  Logger.log('Archived: ' + months.join(', ') + '  ->  Drive folder "' + ARCHIVE_FOLDER + '"');
}

// Kept for old schedules: the daily refresh now handles the monthly archive itself.
function archiveMonthly() { phase2Daily(); }

// Runs with every daily refresh: when a period has closed and isn't archived yet, save its PDF and send the summary email.
// Doesn't depend on a fixed day of the month, so it can't drift if your statement dates change.
function archiveIfDue_() {
  const P = PropertiesService.getScriptProperties();
  PERIOD_ = null;
  const tx = loadTx_(), cur = currentPeriod_();
  const [y, m] = cur.split('-').map(Number), last = Utilities.formatDate(new Date(y, m - 2, 15), Session.getScriptTimeZone(), 'yyyy-MM');
  let done = null; try { done = JSON.parse(P.getProperty('ARCHIVED') || 'null'); } catch (e) {}
  if (!Array.isArray(done)) {                       // first run: earlier periods were archived at setup
    done = [...new Set(tx.map(t => t.month))].filter(k => k < last);
  }
  if (done.indexOf(last) !== -1 || !tx.some(t => t.month === last)) { P.setProperty('ARCHIVED', JSON.stringify(done.slice(-36))); return; }
  const pdf = archiveMonth_(last, tx);
  if (typeof monthlyInsightsEmail_ === 'function') { try { monthlyInsightsEmail_(last, tx, pdf); } catch (e) { log_('Monthly email failed: ' + e); } }
  done.push(last);
  P.setProperty('ARCHIVED', JSON.stringify(done.slice(-36)));
  log_('Monthly archive: ' + periodRange_(last) + ' saved' + (pdf ? ' and emailed' : '') + '.');
}

function archiveMonth_(m, tx) {
  const tz = Session.getScriptTimeZone(), S = settings_();
  const [y, mo] = m.split('-').map(Number);
  const title = Utilities.formatDate(new Date(y, mo - 1, 15), tz, 'MMMM yyyy') + ((PERIOD_ || {}).on ? ' (statement period ' + periodRange_(m) + ')' : '');
  const prevKeys = [1, 2, 3].map(i => Utilities.formatDate(new Date(y, mo - 1 - i, 15), tz, 'yyyy-MM'));
  const prev = prevKeys[0];
  const $ = n => (n < 0 ? '−$' : '$') + Math.round(Math.abs(n)).toLocaleString('en-US');
  const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

  const inMonth = k => tx.filter(t => t.month === k);
  const spendOf = k => inMonth(k).filter(isSpend_);
  const sum = (a, f) => a.reduce((s, t) => s + f(t), 0);
  const cur = inMonth(m), sp = spendOf(m);
  const pay = sum(cur.filter(t => t.cat === 'Income'), t => t.amt);
  const home = -sum(sp.filter(t => !t.trip), t => t.amt), trips = -sum(sp.filter(t => t.trip), t => t.amt);
  const left = pay - home - trips, rate = pay ? left / pay : 0;
  const prevSpend = -sum(spendOf(prev), t => t.amt);
  const firstDay = tx.length ? tx.reduce((a, t) => t.date < a ? t.date : a, tx[0].date) : null;
  const partial = firstDay && Utilities.formatDate(firstDay, tz, 'yyyy-MM') === m && firstDay.getDate() > 1;

  // categories with limits (Settings "<Cat> monthly limit", Dining target, else 3-month average)
  const byCat = (arr) => { const o = {}; arr.forEach(t => o[t.cat] = (o[t.cat] || 0) - t.amt); return o; };
  const c = byCat(sp), cp = byCat(spendOf(prev));
  const limit = k => NO_LIMIT_CATS.indexOf(k) !== -1 ? 0 : limitFor_(k, m, tx, S);
  const cats = Object.keys(c).filter(k => c[k] > 0.5).sort((a, b) => c[b] - c[a]);
  const catRows = cats.map(k => {
    const L = limit(k), r = L ? c[k] / L : 0, col = !L ? '#6B7280' : r >= 1 ? '#DC2626' : r >= 0.3 ? '#D97706' : '#059669';
    const d = c[k] - (cp[k] || 0);
    return '<tr><td>' + esc(k) + '</td><td class="n">' + $(c[k]) + '</td><td class="n">' + (L ? $(L) : '—') +
      '</td><td class="n" style="color:' + col + '">' + (L ? Math.round(r * 100) + '%' : '') + '</td><td class="n" style="color:' +
      (d > 0 ? '#DC2626' : '#059669') + '">' + (cp[k] !== undefined || d ? (d > 0 ? '+' : '') + $(d) : '') + '</td></tr>';
  }).join('');

  const cardRows = [...new Set(sp.filter(t => t.type === 'Credit card').map(t => cardName_(t.account, t.inst)))].filter(n => cardTracked_(n, S)).map(n => {
    const a = sp.filter(t => t.type === 'Credit card' && cardName_(t.account, t.inst) === n);
    return a.length ? '<tr><td>' + n + '</td><td class="n">' + $(-sum(a, t => t.amt)) + '</td><td class="n">' + a.length + '</td></tr>' : '';
  }).join('');
  const big = sp.filter(t => t.amt < 0).sort((a, b) => a.amt - b.amt).slice(0, 10).map(t =>
    '<tr><td>' + Utilities.formatDate(t.date, tz, 'MMM d') + '</td><td>' + esc(merchantKey_(t.desc)) + '</td><td>' + esc(t.cat) +
    '</td><td class="n">' + $(-t.amt) + '</td></tr>').join('');
  const tripNames = [...new Set(sp.filter(t => t.trip).map(t => t.trip))];

  // net worth snapshot closest to month end
  const nw = sheet_('NetWorth'), end = new Date(periodBounds_(m)[1].getTime() + 864e5);
  let nwTxt = 'no snapshot yet';
  if (nw.getLastRow() > 1) {
    const rows = nw.getRange(2, 1, nw.getLastRow() - 1, 8).getValues().filter(r => r[0] instanceof Date && r[0] < end);
    if (rows.length) { const r = rows[rows.length - 1]; nwTxt = $(r[3]) + '  (cash ' + $(r[4]) + ' · investments ' + $(r[5]) + ' · cards ' + $(-r[6]) + ' · loans ' + $(-r[7]) + ')'; }
  }

  const tile = (l, v, s, col) => '<td class="tile"><div class="tl">' + l + '</div><div class="tv" style="color:' + (col || '#111827') + '">' + v + '</div><div class="ts">' + s + '</div></td>';
  const html = '<html><head><style>body{font-family:Arial,Helvetica,sans-serif;color:#111827;font-size:11px;margin:28px}' +
    'h1{font-size:22px;margin:0}h2{font-size:13px;margin:22px 0 6px;padding:5px 8px;background:#1F2937;color:#fff}' +
    'table{width:100%;border-collapse:collapse}td,th{padding:5px 6px;border-bottom:1px solid #E5E7EB;text-align:left}th{color:#6B7280;font-weight:normal;font-size:10px}' +
    '.n{text-align:right}.tile{background:#F3F4F6;text-align:center;border:4px solid #fff;padding:10px}.tl{font-size:9px;color:#6B7280;letter-spacing:1px}' +
    '.tv{font-size:20px;font-weight:bold;margin:4px 0}.ts{font-size:9px;color:#6B7280}.sub{color:#6B7280}</style></head><body>' +
    '<h1>' + title + ' — money summary</h1><div class="sub">Finance HQ · generated ' + Utilities.formatDate(new Date(), tz, 'MMM d, yyyy') +
    (partial ? ' · <b>partial month</b> (data starts ' + Utilities.formatDate(firstDay, tz, 'MMM d') + ')' : '') + '</div>' +
    '<table style="margin-top:14px"><tr>' +
    tile('LEFT OVER', $(left), 'pay minus all spending', left >= 0 ? '#059669' : '#DC2626') +
    tile('SAVINGS RATE', Math.round(rate * 100) + '%', 'goal 20%+', rate >= 0.2 ? '#059669' : rate >= 0 ? '#D97706' : '#DC2626') +
    tile('PAY', $(pay), 'take-home') +
    tile('SPENDING', $(home + trips), (prevSpend ? (home + trips > prevSpend ? '▲ ' : '▼ ') + $(Math.abs(home + trips - prevSpend)) + ' vs the period before' : '')) +
    '</tr></table>' +
    (trips ? '<p class="sub">Includes ' + $(trips) + ' of trip spending (' + esc(tripNames.join(', ')) + ').</p>' : '') +
    '<h2>SPENDING BY CATEGORY</h2><table><tr><th>Category</th><th class="n">Spent</th><th class="n">Limit</th><th class="n">Used</th><th class="n">vs last month</th></tr>' + catRows + '</table>' +
    '<h2>CREDIT CARDS</h2><table><tr><th>Card</th><th class="n">Spent this month</th><th class="n">Purchases</th></tr>' + cardRows + '</table>' +
    '<h2>BIGGEST PURCHASES</h2><table><tr><th>Date</th><th>Merchant</th><th>Category</th><th class="n">Amount</th></tr>' + big + '</table>' +
    (typeof insights_ === 'function' ? '<h2>INSIGHTS</h2><ul>' + insights_(m, tx, S).map(s => '<li style="margin:4px 0">' + esc(s) + '</li>').join('') + '</ul>' : '') +
    '<h2>NET WORTH AT MONTH END</h2><p>' + nwTxt + '</p></body></html>';

  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(m + ' Finance Summary.pdf');
  const folder = archiveFolder_();
  const old = folder.getFilesByName(blob.getName()); while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(blob);
  return blob;
}

function archiveFolder_() {
  const it = DriveApp.getFoldersByName(ARCHIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ARCHIVE_FOLDER);
}


// ===== PART 4 — alerts, cash-flow forecast, weekly digest, monthly insights =====

const P3_SETTINGS = [
  ['Alerts: email (Yes/No)', 'Yes', 'Emails each batch of alerts'],
  ['Alerts: phone reminders (Yes/No)', 'Yes', 'Adds a Google Calendar event with a pop-up, so your phone notifies you'],
  ['Alert email', '', 'Blank = the Google account that runs this script'],
  ['Large charge alert over ($)', 200, 'Any single purchase above this'],
  ['Checking low point alert under ($)', 500, 'Forecast warns if checking is projected to dip below this'],
  ['Bill reminder (days before due)', 3, 'Phone reminder this many days before each card is due'],
  ['Weekly digest (Yes/No)', 'Yes', 'Sunday 9 AM email'],
  ['Monthly insights email (Yes/No)', 'Yes', 'Sent when each period closes, with the PDF attached'],
  ['Update emails (Yes/No)', 'Yes', 'Email me when a new version of Finance HQ is ready (installing takes about a minute)'],
  ['Payday nudge (Yes/No)', 'Yes', 'When a paycheck lands: how much to move to savings and what\'s left to spend'],
  ['Share usage check-ins with the Finance HQ creator (Yes/No)', 'Yes', 'Weekly counts only (is it syncing, days used) — never amounts or account details. Menu → What check-ins share']
];
// Spending that's hard to cut: suggested limit = your average. Everything else: average minus 20%.
const ESSENTIAL_CATS = ['Groceries', 'Gas', 'Bills & utilities', 'Student loans', 'Loan payments', 'Housing', 'Gym', 'Subscriptions', 'Haircuts', 'Transport', 'Insurance', 'Health'];
const NO_LIMIT_CATS = ['Uncategorized', 'Fees', 'Affirm', 'Sports betting', 'Giving'];
const FIXED_CATS = ['Student loans', 'Loan payments', 'Housing', 'Insurance', 'Bills & utilities', 'Giving'];   // fixed bills: no limit suggested

// ---------------------------------------------------------------- setup
function phase3Setup() {
  ensureSettings_();
  ensureP3Settings_();
  const tx = loadTx_();
  suggestLimits_(tx);
  ScriptApp.getProjectTriggers().filter(t => ['phase3Daily', 'weeklyDigest'].indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('phase3Daily').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('weeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
  PropertiesService.getScriptProperties().deleteProperty('P3_SEEN');
  phase3Daily();            // first run only records what already exists (no alert flood)
  weeklyDigest(true);       // test digest so you can see it
  Logger.log('Phase 3 ready: daily alerts 8 AM, digest Sundays 9 AM. Test digest sent to ' + alertEmail_());
}

function ensureP3Settings_() {
  const sh = sheet_('Settings');
  const have = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : []);
  const add = P3_SETTINGS.filter(s => !have.has(s[0]));
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
}

function suggestLimits_(tx) {
  // Writes "<Category> monthly limit" rows only where you don't have one yet. Edit them freely.
  const S = settings_(), sh = sheet_('Settings');
  const cur = currentPeriod_(), [y, m] = cur.split('-').map(Number);
  const prev = priorPeriods_(cur, 3, tx);
  const avg = {};
  prev.forEach(k => { const st = periodStats_(k, tx); Object.keys(st.byCat).forEach(c => avg[c] = (avg[c] || 0) + st.byCat[c] / prev.length); });
  let fit = {}; try { fit = budgetPlan_(tx, S).sugg || {}; } catch (e) {}     // limits that leave room for your savings goal
  const add = [];
  Object.keys(avg).sort((a, b) => avg[b] - avg[a]).forEach(c => {
    if (avg[c] < 15 || NO_LIMIT_CATS.indexOf(c) !== -1 || FIXED_CATS.indexOf(c) !== -1 || S[c + ' monthly limit'] !== undefined) return;
    if (c === 'Dining') {
      if (!Number(S['Dining target per month'])) setSetting_('Dining target per month', fit[c] || Math.max(50, Math.round(avg[c] * 0.8 / 10) * 10), 'suggested from your spending (average $' + Math.round(avg[c]) + ') — edit me');
      return;
    }
    const ess = ESSENTIAL_CATS.indexOf(c) !== -1;
    const v = fit[c] || (ess ? Math.ceil(avg[c] / 10) * 10 : Math.max(10, Math.round(avg[c] * 0.8 / 10) * 10));
    add.push([c + ' monthly limit', v, 'suggested from your spending (average $' + Math.round(avg[c]) + ') so your savings goal fits — edit me']);
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
}

// ---------------------------------------------------------------- shared numbers
function periodStats_(key, tx) {
  const inP = tx.filter(t => t.month === key);
  const sp = inP.filter(isSpend_);
  const byCat = {};
  sp.forEach(t => byCat[t.cat] = (byCat[t.cat] || 0) - t.amt);
  Object.keys(byCat).forEach(c => { if (byCat[c] < 0.5) delete byCat[c]; });
  const pay = inP.filter(t => t.cat === 'Income').reduce((s, t) => s + t.amt, 0);
  const spend = -sp.reduce((s, t) => s + t.amt, 0);
  const trips = -sp.filter(t => t.trip).reduce((s, t) => s + t.amt, 0);
  return { key: key, pay: pay, spend: spend, trips: trips, left: pay - spend, byCat: byCat, tx: inP };
}
// Share of pay you set aside before spending (savings + giving + tax goals), as a fraction.
function setAsidePct_(S) {
  return ['Savings goal (% of pay)', 'Giving goal (% of pay)', 'Tax set-aside (% of pay)'].reduce((a, k) => a + pctSetting_(S[k]), 0) / 100;
}
// Regular bills (student loans, rent, insurance, phone…) you usually pay each period but haven't paid yet this period.
// Held back from "Left to spend" so that money isn't counted as free.
function billsComing_(tx, key, S) {
  S = S || settings_();
  const keys = priorPeriods_(key, 3, tx);
  if (!keys.length) return { amt: 0, list: [] };
  const giveSetAside = pctSetting_(S['Giving goal (% of pay)']) > 0;
  const avg = catAverages_(keys, tx), list = [];
  FIXED_CATS.filter(c => !(c === 'Giving' && giveSetAside)).forEach(c => {
    if (!(avg[c] >= 5)) return;
    const paid = tx.filter(t => t.month === key && t.cat === c && isSpend_(t)).reduce((s, t) => s - t.amt, 0);
    const left = Math.round(Math.max(0, avg[c] - paid) * 100) / 100;
    if (left >= 5) list.push({ cat: c, amt: left });
  });
  return { amt: Math.round(list.reduce((s, x) => s + x.amt, 0) * 100) / 100, list: list };
}
// Paychecks still expected before this period ends (any account), from your usual payday rhythm and amount.
function payStillComing_(tx, key) {
  const tz = Session.getScriptTimeZone(), now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const end = periodBounds_(key)[1], days = {};
  tx.filter(t => t.cat === 'Income' && t.amt > 50).forEach(t => { const k = Utilities.formatDate(t.date, tz, 'yyyy-MM-dd'); days[k] = (days[k] || 0) + t.amt; });
  const pk = Object.keys(days).sort();
  if (pk.length < 2) return { amt: 0, n: 0 };
  const gaps = []; for (let i = 1; i < pk.length; i++) { const g = Math.round((parseDay_(pk[i]) - parseDay_(pk[i - 1])) / 864e5); if (g >= 6) gaps.push(g); }
  const gap = Math.round(median_(gaps.slice(-4))) || 14, amt = median_(pk.slice(-4).map(k => days[k]));
  const last = parseDay_(pk[pk.length - 1]);
  if ((today - last) / 864e5 > gap * 2 + 5) return { amt: 0, n: 0, stale: true };          // paychecks stopped: don't promise money that isn't coming
  // Irregular income (freelance, gig work, commissions): no steady rhythm, so expect what a normal period brought in.
  const recent = gaps.slice(-6), steady = recent.length >= 2 && recent.filter(g => Math.abs(g - gap) <= 4).length >= Math.ceil(recent.length * 0.75)
    && pk.slice(-6).map(k => days[k]).filter(a => Math.abs(a - amt) <= amt * 0.35).length >= Math.ceil(Math.min(6, pk.length) * 0.6);
  if (!steady) {
    const prev = priorPeriods_(key, 3, tx);
    if (!prev.length) return { amt: 0, n: 0 };
    const inc = k => tx.filter(t => t.month === k && t.cat === 'Income' && t.amt > 0).reduce((s2, t) => s2 + t.amt, 0);
    const usual = prev.reduce((s2, k) => s2 + inc(k), 0) / prev.length, got = inc(key);
    return { amt: Math.round(Math.max(0, usual - got)), n: 0, irregular: true };
  }
  let d = last, tot = 0, n = 0;
  for (let i = 0; i < 12; i++) { d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + gap); if (d > end) break; if (d > today) { tot += amt; n++; } }
  return { amt: Math.round(tot), n: n };
}
function limitFor_(cat, key, tx, S) {
  const v = Number(S[cat + ' monthly limit']);
  if (v > 0) return v;
  if (cat === 'Dining' && Number(S['Dining target per month'])) return Number(S['Dining target per month']);
  const prev = priorPeriods_(key, 3, tx);
  if (!prev.length) return 0;
  const a = prev.reduce((s, k) => s + (periodStats_(k, tx).byCat[cat] || 0), 0) / prev.length;
  return a >= 15 ? Math.round(a / 10) * 10 : 0;
}
function checkingBalance_() {
  const a = sheet_('Accounts');
  return a.getLastRow() > 1 ? a.getRange(2, 1, a.getLastRow() - 1, 5).getValues().filter(r => r[3] === 'Checking').reduce((s, r) => s + (Number(r[4]) || 0), 0) : 0;
}
const fmt$_ = n => (n < 0 ? '−$' : '$') + Math.round(Math.abs(n)).toLocaleString('en-US');
const yes_ = v => /^y/i.test(String(v || 'Yes'));
function alertEmail_() { return String(settings_()['Alert email'] || '').trim() || Session.getEffectiveUser().getEmail(); }

// ---------------------------------------------------------------- cash-flow forecast
// Projects checking day by day: estimated paychecks in, card bills out on their due dates, everyday debit spending.
function forecast_(tx, S, cards, days) {
  days = days || 45;
  const tz = Session.getScriptTimeZone(), now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const horizon = new Date(today.getFullYear(), today.getMonth(), today.getDate() + days);
  const since = new Date(today.getTime() - 90 * 864e5);
  const dayOf = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const events = [];   // {date, what, amt, kind: 'in' | 'bill'}

  // paychecks into checking: typical gap + amount from history
  const payDays = {};
  tx.filter(t => t.cat === 'Income' && t.type === 'Checking' && t.amt > 50).forEach(t => {
    const k = Utilities.formatDate(t.date, tz, 'yyyy-MM-dd'); payDays[k] = (payDays[k] || 0) + t.amt;
  });
  const pk = Object.keys(payDays).sort();
  if (pk.length) {
    const amts = pk.slice(-4).map(k => payDays[k]);
    const gaps = []; for (let i = 1; i < pk.length; i++) { const g = Math.round((parseDay_(pk[i]) - parseDay_(pk[i - 1])) / 864e5); if (g >= 6) gaps.push(g); }
    const gap = Math.round(median_(gaps.slice(-4))) || 14, amt = median_(amts);
    let d = parseDay_(pk[pk.length - 1]);
    if ((today - d) / 864e5 > gap * 2 + 5) d = horizon;                                   // paychecks stopped: project none
    for (let n = 0; n < 40; n++) {
      d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + gap);
      if (d > horizon) break;
      if (d > today) events.push({ date: d, what: 'Paycheck (estimated)', amt: amt, kind: 'in' });
    }
  }

  // card bills on their due dates
  const isPay = t => t.amt > 0 && (t.cat === 'Transfer' || /PAYMENT|THANK YOU/i.test(t.desc));
  (cards || []).forEach(c => {
    const mine = tx.filter(t => t.account === c.account);
    const cardDaily = -mine.filter(t => t.date >= since && t.amt < 0 && isSpend_(t)).reduce((s, t) => s + t.amt, 0) / 90;
    if (c.configured && c.dueISO) {
      const due = parseDay_(c.dueISO);
      if (!c.paid && c.statement > 0 && due >= today) events.push({ date: due, what: c.name + ' bill', amt: -c.statement, kind: 'bill' });
      // next statement = spending so far + typical spending until it closes
      const close = parseDay_(c.closesISO), next = new Date(due.getFullYear(), due.getMonth() + 1, due.getDate());
      const est = c.spent + cardDaily * Math.max(0, (close - today) / 864e5);
      if (next <= horizon && est > 1) events.push({ date: next, what: c.name + ' bill (estimated)', amt: -est, kind: 'bill' });
    } else {
      // no dates set: assume you pay it about when you paid it last, roughly a month of spending
      const pays = mine.filter(isPay).sort((a, b) => a.date - b.date);
      if (!pays.length || cardDaily < 1) return;
      let d = dayOf(pays[pays.length - 1].date);
      for (let n = 0; n < 3; n++) {
        d = new Date(d.getFullYear(), d.getMonth() + 1, d.getDate());
        if (d > horizon) break;
        if (d > today) events.push({ date: d, what: c.name + ' bill (estimated — add card dates)', amt: -cardDaily * 30, kind: 'bill' });
      }
    }
  });

  // everyday debit-card spending from checking (averaged)
  const debitDaily = -tx.filter(t => t.type === 'Checking' && t.date >= since && t.amt < 0 && isSpend_(t)).reduce((s, t) => s + t.amt, 0) / 90;

  const start = checkingBalance_();
  const acc = sheet_('Accounts'), savings = acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, 5).getValues().filter(r => r[3] === 'Savings').reduce((s, r) => s + (Number(r[4]) || 0), 0) : 0;
  let bal = start, low = { bal: start, date: today, cause: '' };
  const daily = [], rows = [];
  for (let i = 1; i <= days; i++) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() + i);
    bal -= debitDaily;
    const todays = events.filter(e => e.date.getTime() === d.getTime());
    todays.forEach(e => { bal += e.amt; rows.push([d, e.what, e.amt, bal]); });
    daily.push({ date: d, bal: bal, inn: todays.filter(e => e.amt > 0).reduce((s, e) => s + e.amt, 0),
      bills: -todays.filter(e => e.amt < 0).reduce((s, e) => s + e.amt, 0) });
    if (bal < low.bal) low = { bal: bal, date: d, cause: todays.filter(e => e.amt < 0).map(e => e.what).join(' + ') || low.cause };
  }
  return { start: start, savings: savings, rows: rows, daily: daily, low: low, end: bal, debitDaily: debitDaily, days: days,
    billsTotal: -events.filter(e => e.amt < 0 && e.date <= horizon && e.date > today).reduce((s, e) => s + e.amt, 0),
    payTotal: events.filter(e => e.amt > 0 && e.date <= horizon && e.date > today).reduce((s, e) => s + e.amt, 0),
    unconfigured: (cards || []).filter(c => !c.configured).map(c => c.name) };
}
function parseDay_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const d = new Date(v); return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function writeForecast_(f) {
  const ss = finSS_(), tz = Session.getScriptTimeZone(), S = settings_();
  const sh = ss.getSheetByName('Forecast') || ss.insertSheet('Forecast');
  sh.clear(); sh.getRange('A:H').breakApart(); sh.getBandings().forEach(b => b.remove()); sh.setConditionalFormatRules([]);
  sh.setHiddenGridlines(true); sh.setFrozenRows(0);
  uiPage_(sh);
  [120, 220, 120, 120, 120, 120, 110, 110].forEach((w, c) => sh.setColumnWidth(c + 1, w));
  sh.getRange('A:H').setFontFamily('Arial').setFontColor(UI.ink).setVerticalAlignment('middle');
  const day = d => Utilities.formatDate(d, tz, 'EEE MMM d');
  const floor = Number(S['Checking low point alert under ($)']) || 500;
  const ok = f.low.bal >= floor, neg = f.low.bal < 0;

  sh.getRange('A1:F1').merge().setValue('Will I have enough in checking?').setFontSize(18).setFontWeight('bold');
  sh.getRange('G1:H1').merge().setValue('Updated ' + Utilities.formatDate(new Date(), tz, 'MMM d, h:mm a')).setFontColor(UI.mute).setHorizontalAlignment('right');
  // plain-English answer
  const move = Math.ceil((floor - f.low.bal) / 50) * 50;
  const msg = ok
    ? '✅ You are covered. Your checking account stays above ' + fmt$_(Math.floor(f.low.bal / 10) * 10) + ' for the next ' + f.days + ' days, even after your card bills.'
    : (neg ? '❌ ' : '⚠️ ') + 'Heads up: checking could drop to ' + fmt$_(f.low.bal) + ' on ' + day(f.low.date) +
      (f.low.cause ? ' (after your ' + f.low.cause.replace(/ \(estimated[^)]*\)/g, '') + ')' : '') +
      '. Move about ' + fmt$_(move) + ' from savings before then to stay above ' + fmt$_(floor) + (f.savings ? ' (you have ' + fmt$_(f.savings) + ' in savings).' : '.');
  sh.getRange('A2:H2').merge().setValue(msg).setFontSize(12).setFontWeight('bold').setWrap(true)
    .setFontColor(ok ? UI.good : neg ? UI.bad : UI.warn).setBackground(ok ? '#ECFDF5' : neg ? '#FEF2F2' : '#FFFBEB');
  sh.setRowHeight(2, 46);

  let r = uiTiles_(sh, 4, [
    ['IN CHECKING NOW', f.start, '$#,##0', UI.ink, 'from your bank'],
    ['LOWEST POINT', f.low.bal, '$#,##0', ok ? UI.good : UI.bad, f.low.date ? day(f.low.date) : 'today'],
    ['PAYCHECKS COMING', f.payTotal, '$#,##0', UI.good, 'next ' + f.days + ' days (into checking)'],
    ['BILLS COMING', f.billsTotal, '$#,##0', UI.warn, 'card bills, next ' + f.days + ' days']
  ]);

  // week by week
  r = uiBand_(sh, r, 'Week by week');
  r = uiHeader_(sh, r, ['Week of', '', 'Money in', 'Card bills', 'Everyday spending', 'Checking at week end', '', '']);
  const maxBal = Math.max(1, f.start, ...f.daily.map(x => x.bal));
  const w0 = r;
  for (let i = 0; i < f.daily.length; i += 7) {
    const wk = f.daily.slice(i, i + 7), endBal = wk[wk.length - 1].bal;
    const minBal = Math.min(...wk.map(x => x.bal));
    sh.getRange(r, 1, 1, 6).setValues([[Utilities.formatDate(wk[0].date, tz, 'MMM d') + ' – ' + Utilities.formatDate(wk[wk.length - 1].date, tz, 'MMM d'), '',
      wk.reduce((s, x) => s + x.inn, 0), wk.reduce((s, x) => s + x.bills, 0), f.debitDaily * wk.length, endBal]]);
    sh.getRange(r, 1, 1, 2).merge();
    sh.getRange(r, 3).setFontColor(UI.good); sh.getRange(r, 4).setFontColor(UI.ink);
    sh.getRange(r, 6).setFontWeight('bold').setFontColor(minBal < floor ? UI.bad : UI.ink);
    sh.getRange(r, 7, 1, 2).merge().setFormula('=SPARKLINE(' + Math.max(0, Math.round(endBal)) + ',{"charttype","bar";"max",' + Math.round(maxBal) +
      ';"color1","' + (minBal < floor ? UI.bad : UI.accent) + '"})');
    r++;
  }
  sh.getRange(w0, 3, r - w0, 4).setNumberFormat('$#,##0;[Red]-$#,##0');
  uiStripes_(sh, w0, r - w0, 8);
  r++;

  // event list
  r = uiBand_(sh, r, "What's coming in and out");
  r = uiHeader_(sh, r, ['Date', 'What', 'Amount', 'Checking after', '', '', '', '']);
  const e0 = r;
  if (f.rows.length) {
    f.rows.forEach(x => {
      sh.getRange(r, 1, 1, 4).setValues([[x[0], x[1], x[2], x[3]]]);
      sh.getRange(r, 3).setFontColor(x[2] >= 0 ? UI.good : UI.ink);
      if (x[3] < floor) sh.getRange(r, 4).setFontColor(UI.bad).setFontWeight('bold');
      r++;
    });
    sh.getRange(e0, 1, r - e0, 1).setNumberFormat('ddd mmm d');
    sh.getRange(e0, 3, r - e0, 2).setNumberFormat('$#,##0;[Red]-$#,##0');
    uiStripes_(sh, e0, r - e0, 8);
  } else { sh.getRange(r, 1, 1, 8).merge().setValue('Nothing scheduled in the next ' + f.days + ' days.').setFontColor(UI.mute); r++; }
  r++;

  r = uiBand_(sh, r, 'How this is worked out');
  const notes = [
    '• Paychecks: the usual amount and timing of your recent deposits into checking.',
    '• Card bills: your statement balance on each due date, then an estimate for the next statement.' + (f.unconfigured.length ? ' Add closing/due days for ' + f.unconfigured.join(', ') + ' (Finance HQ → Quick setup) for exact dates.' : ''),
    '• Everyday spending: about ' + fmt$_(f.debitDaily) + ' a day — your average debit-card spending over the last 90 days.',
    '• Money you move between your own accounts isn\'t included. Warning level: ' + fmt$_(floor) + ' (change it in Settings).'
  ];
  notes.forEach(n => { sh.getRange(r, 1, 1, 8).merge().setValue(n).setFontColor(UI.mute).setFontSize(10).setWrap(true); r++; });
}

function cardsP3_(tx, S) {
  return cardCycles_(tx, S, new Date());
}

// ---------------------------------------------------------------- daily alerts
function phase3Daily() {
  ensureP3Settings_();
  const S = settings_(), tx = loadTx_(), now = new Date(), tz = Session.getScriptTimeZone();
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('P3_SEEN'), firstRun = !raw;
  const seen = JSON.parse(raw || '{}');
  const out = [];
  const add = (key, text, urgent) => { if (!seen[key]) { seen[key] = Date.now(); if (!firstRun) out.push({ text: text, urgent: !!urgent }); } };
  const cards = cardsP3_(tx, S);

  // 1) large purchases (last 3 days)
  const big = Number(S['Large charge alert over ($)']) || 200, recent = new Date(now.getTime() - 3 * 864e5);
  tx.filter(t => t.date >= recent && t.amt < -big && isSpend_(t)).forEach(t =>
    add('big:' + t.id, 'Large charge: ' + fmt$_(-t.amt) + ' at ' + merchantKey_(t.desc) + ' (' + t.cat + ')'));

  // 2) categories near / over their limit this period
  const cur = currentPeriod_(now), st = periodStats_(cur, tx);
  Object.keys(st.byCat).forEach(c => {
    const L = limitFor_(c, cur, tx, S); if (!L || NO_LIMIT_CATS.indexOf(c) !== -1) return;
    const r = st.byCat[c] / L;
    if (r >= 1) add('over:' + cur + ':' + c, c + ' is over its limit: ' + fmt$_(st.byCat[c]) + ' of ' + fmt$_(L), true);
    else if (r >= 0.8) add('near:' + cur + ':' + c, c + ' is at ' + Math.round(r * 100) + '% of its limit (' + fmt$_(L - st.byCat[c]) + ' left)');
  });

  // 3) new subscriptions and price increases
  const sub = sheet_('Subscriptions');
  if (sub.getLastRow() > 1) sub.getRange(2, 1, sub.getLastRow() - 1, 7).getValues().forEach(r => {
    if (!r[0] || !r[2] || /total|known/i.test(r[0])) return;
    add('sub:' + r[0], 'New recurring charge: ' + r[0] + ' ' + fmt$_(r[1]) + ' ' + String(r[2]).toLowerCase() + ' (' + fmt$_(r[5]) + '/yr)');
    if (/UP/i.test(r[6])) add('subup:' + r[0] + ':' + r[1], 'Price increase: ' + r[0] + ' is now ' + fmt$_(r[1]) + ' (' + r[6] + ')', true);
  });

  // 4) bill reminders as phone reminders on (due date − N days)
  const lead = Number(S['Bill reminder (days before due)']) || 3;
  cards.forEach(c => {
    if (!c.configured || c.paid || c.statement <= 0 || !c.dueISO) return;
    const k = 'billcal:' + c.name + ':' + c.dueISO;
    if (!seen[k]) {
      seen[k] = Date.now();
      const due = new Date(c.dueISO + 'T09:00:00'), when = new Date(due.getTime() - lead * 864e5);
      if (yes_(S['Alerts: phone reminders (Yes/No)']) && when > now)
        phoneReminder_('Pay ' + c.name + ' ' + fmt$_(c.statement) + ' — due ' + c.due, when, 'Statement balance ' + fmt$_(c.statement) + ', due ' + c.due + '. Finance HQ.');
    }
    if (c.daysToDue <= lead) add('billdue:' + c.name + ':' + c.dueISO, c.name + ' ' + fmt$_(c.statement) + ' is due ' + c.due + ' (' + c.status + ')', true);
  });

  // 5) cash-flow forecast low point
  const f = forecast_(tx, S, cards, 45); writeForecast_(f);
  const floor = Number(S['Checking low point alert under ($)']) || 500;
  if (f.low.bal < floor) add('low:' + Utilities.formatDate(f.low.date, tz, 'yyyy-MM-dd'),
    'Checking is projected to dip to ' + fmt$_(f.low.bal) + ' on ' + Utilities.formatDate(f.low.date, tz, 'EEE MMM d') +
    ' — move ' + fmt$_(Math.ceil((floor - f.low.bal) / 50) * 50) + ' from savings before then.', true);

  // 6) health problems (reminds weekly while unresolved)
  try {
    const wk = Utilities.formatDate(now, tz, 'YYYY-ww');
    buildHealth_().filter(c => c[0] === 2).forEach(c => add('health:' + c[1] + ':' + wk, 'Finance HQ needs attention — ' + c[1] + ': ' + c[2], true));
  } catch (e) {}

  // keep 6 months of history
  Object.keys(seen).forEach(k => { if (Date.now() - seen[k] > 180 * 864e5) delete seen[k]; });
  props.setProperty('P3_SEEN', JSON.stringify(seen));
  if (out.length) deliverAlerts_(out, S);
}

function deliverAlerts_(alerts, S) {
  const props = PropertiesService.getScriptProperties();
  const log = JSON.parse(props.getProperty('P3_WEEK') || '[]').concat(alerts.map(a => [Date.now(), a.text]));
  props.setProperty('P3_WEEK', JSON.stringify(log.filter(x => Date.now() - x[0] < 8 * 864e5).slice(-40)));
  alerts.sort((a, b) => b.urgent - a.urgent);
  if (yes_(S['Alerts: email (Yes/No)']))
    mail_({ to: alertEmail_(), subject: 'Finance HQ: ' + alerts[0].text + (alerts.length > 1 ? ' (+' + (alerts.length - 1) + ' more)' : ''),
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px"><ul>' + alerts.map(a => '<li style="margin:6px 0">' + (a.urgent ? '⚠️ ' : '') + a.text + '</li>').join('') +
        '</ul><p style="color:#6B7280">Finance HQ · <a href="' + finSS_().getUrl() + '">open sheet</a></p></div>' });
  if (yes_(S['Alerts: phone reminders (Yes/No)']))
    phoneReminder_('Finance HQ: ' + alerts[0].text + (alerts.length > 1 ? ' (+' + (alerts.length - 1) + ')' : ''), new Date(Date.now() + 3 * 60000), alerts.map(a => '• ' + a.text).join('\n'));
}

function phoneReminder_(title, start, desc) {
  if (isDemo_()) return;
  const ev = CalendarApp.getDefaultCalendar().createEvent(title.slice(0, 120), start, new Date(start.getTime() + 15 * 60000), { description: desc || '' });
  ev.removeAllReminders(); ev.addPopupReminder(0);
}

// ---------------------------------------------------------------- weekly digest
function weeklyDigest(force) {
  const S = settings_();
  if (force !== true && !yes_(S['Weekly digest (Yes/No)'])) return;
  const tx = loadTx_(), tz = Session.getScriptTimeZone(), now = new Date();
  const cur = currentPeriod_(now), st = periodStats_(cur, tx), b = periodBounds_(cur);
  const daysLeft = Math.max(1, Math.round((b[1] - new Date(now.getFullYear(), now.getMonth(), now.getDate())) / 864e5) + 1);
  const coming = payStillComing_(tx, cur), cl = checklist_(tx, S, cur, st.pay + coming.amt), billsDue = billsComing_(tx, cur, S);
  const left = cl.pay > 0 ? cl.pay - cl.setAside - billsDue.amt - st.spend : null;   // same Left to spend as the Dashboard
  const leftTxt = left === null ? 'No pay yet' : fmt$_(left) + ' left to spend';
  const weekAgo = new Date(now.getTime() - 7 * 864e5);
  const weekSpend = -tx.filter(t => t.date >= weekAgo && isSpend_(t)).reduce((s, t) => s + t.amt, 0);
  const cards = cardsP3_(tx, S), f = forecast_(tx, S, cards, 45);
  const cats = Object.keys(st.byCat).sort((a, c) => st.byCat[c] - st.byCat[a]).slice(0, 6).map(c => {
    const L = limitFor_(c, cur, tx, S), r = L ? st.byCat[c] / L : 0;
    const col = !L ? '#6B7280' : r >= 1 ? '#DC2626' : r >= 0.3 ? '#D97706' : '#059669';
    return '<tr><td style="padding:4px 8px">' + c + '</td><td style="padding:4px 8px;text-align:right;color:' + col + '"><b>' + fmt$_(st.byCat[c]) + '</b></td><td style="padding:4px 8px;color:#6B7280">' + (L ? 'of ' + fmt$_(L) : '') + '</td></tr>';
  }).join('');
  const bills = cards.filter(c => c.configured && !c.paid && c.statement > 0).map(c => '<li>' + c.name + ' ' + fmt$_(c.statement) + ' — ' + c.status + ' (' + c.due + ')</li>').join('') || '<li>All statements paid ✓</li>';
  const week = JSON.parse(PropertiesService.getScriptProperties().getProperty('P3_WEEK') || '[]').filter(x => x[0] >= weekAgo.getTime()).map(x => '<li>' + x[1] + '</li>').join('');
  const unc = st.tx.filter(t => t.cat === 'Uncategorized' && t.amt < 0);
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#111827;max-width:560px">' +
    '<h2 style="margin:0">Your week in money</h2><p style="color:#6B7280;margin-top:4px">Period ' + periodRange_(cur) + '</p>' +
    '<p style="font-size:28px;margin:8px 0;color:' + (left === null || left >= 0 ? '#059669' : '#DC2626') + '"><b>' + (left === null ? fmt$_(st.spend) + '</b> spent so far</p>' : fmt$_(left) + '</b> left to spend</p>') +
    '<p>' + (daysLeft === 1 ? 'Last day of this period' : daysLeft + ' days to go') + (left > 0 ? ' · about <b>' + fmt$_(left / daysLeft) + '/day</b>' : '') + ' · spent ' + fmt$_(weekSpend) + ' this week</p>' +
    '<h3>Top categories this period</h3><table>' + cats + '</table>' +
    '<h3>Bills</h3><ul>' + bills + '</ul>' +
    '<h3>Checking forecast</h3><p>Lowest point in the next 45 days: <b style="color:' + (f.low.bal < (Number(S['Checking low point alert under ($)']) || 500) ? '#DC2626' : '#059669') + '">' + fmt$_(f.low.bal) + '</b> on ' + Utilities.formatDate(f.low.date, tz, 'EEE MMM d') + '</p>' +
    (week ? '<h3>Alerts this week</h3><ul>' + week + '</ul>' : '') +
    (unc.length ? '<p style="color:#6B7280">' + unc.length + ' uncategorized purchases (' + fmt$_(-unc.reduce((s, t) => s + t.amt, 0)) + ') — add Rules so they land in the right place.</p>' : '') +
    '<p style="color:#6B7280"><a href="' + finSS_().getUrl() + '">Open Finance HQ</a></p></div>';
  mail_({ to: alertEmail_(), subject: 'Finance HQ weekly: ' + leftTxt + ', ' + (daysLeft === 1 ? 'last day of the period' : daysLeft + ' days to go'), htmlBody: html });
}

// ---------------------------------------------------------------- monthly insights (used by Archive)
function insights_(key, tx, S) {
  // Same notes as the Insights tab's review, as plain sentences (used by the monthly PDF and email).
  return insightsReview_(key, tx, S).map(x => x[1] + (x[2] ? ' — ' + x[2] : '.'));
}

function monthlyInsightsEmail_(key, tx, pdfBlob) {
  const S = settings_();
  if (!yes_(S['Monthly insights email (Yes/No)'])) return;
  const title = Utilities.formatDate(new Date(key + '-15T12:00:00'), Session.getScriptTimeZone(), 'MMMM yyyy');
  mail_({ to: alertEmail_(), subject: 'Finance HQ: your ' + title + ' review',
    htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;max-width:560px"><h2>' + title + ' (' + periodRange_(key) + ')</h2><ul>' +
      insights_(key, tx, S).map(s => '<li style="margin:6px 0">' + s + '</li>').join('') + '</ul><p style="color:#6B7280">Full summary attached.</p></div>',
    attachments: pdfBlob ? [pdfBlob] : [] });
}


// ===== PART 5 — phone web app (Deploy → Web app → Execute as Me → Only myself) =====

function doGet() {
  bumpUsage_('app');
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('Finance HQ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppData() {
  const tz = Session.getScriptTimeZone(), now = new Date();
  const S = settings_();
  const tx = loadTx_();
  const cards = cardCycles_(tx, S, now).filter(c => c.configured || c.spent || c.statement);
  const limits = {}, curKey = currentPeriod_(now);
  [...new Set(tx.map(t => t.cat))].forEach(c => { if (NO_LIMIT_CATS.indexOf(c) === -1 && FIXED_CATS.indexOf(c) === -1) { const L = limitFor_(c, curKey, tx, S); if (L > 0) limits[c] = L; } });
  const ash = sheet_('Accounts');
  const accts = ash.getLastRow() > 1 ? ash.getRange(2, 1, ash.getLastRow() - 1, 5).getValues() : [];
  const checking = accts.filter(a => a[3] === 'Checking').reduce((s, a) => s + (Number(a[4]) || 0), 0);
  const cutoff = new Date(now.getFullYear(), now.getMonth() - 12, 1);
  const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const rows = tx.filter(t => t.date >= cutoff).map(t => ({
    d: fmt(t.date), m: t.month, n: prettyName_(t.desc), raw: t.desc, a: Math.round(t.amt * 100) / 100,
    c: t.cat || 'Uncategorized', card: t.type === 'Credit card' ? cardName_(t.account, t.inst) : '',
    acct: t.account, trip: t.trip, spend: isSpend_(t), inc: t.cat === 'Income', id: String(t.id || '')
  }));
  rows.sort((x, y) => x.d < y.d ? 1 : x.d > y.d ? -1 : 0);
  const P = PERIOD_ || { on: false }, ranges = {};
  rows.forEach(t => { if (!ranges[t.m]) ranges[t.m] = periodRange_(t.m); });
  const current = currentPeriod_(now); ranges[current] = periodRange_(current);
  return JSON.stringify({
    current: current, statement: P.on, ranges: ranges, periodEnd: fmt(periodBounds_(current)[1]),
    daysLeft: Math.max(1, Math.round((periodBounds_(current)[1] - new Date(now.getFullYear(), now.getMonth(), now.getDate())) / 864e5) + 1),
    today: fmt(now), tx: rows, cards: cards, checking: Math.round(checking), coming: payStillComing_(tx, current), keepPct: setAsidePct_(S),
    plan: (() => { const comingP = payStillComing_(tx, curKey), c = checklist_(tx, S, curKey, periodStats_(curKey, tx).pay + comingP.amt);
      const bl = billsComing_(tx, curKey, S);
      return { pay: c.pay, got: c.got, setAside: Math.round(c.setAside * 100) / 100, parts: c.parts, bills: bl.amt, billList: bl.list, goal: c.goalSave, pct: pctSetting_(S['Savings goal (% of pay)']), saved: c.moves.into, gross: c.moves.gross, out: c.moves.takenOut,
      items: c.items.map(i => ({ k: i.k, label: i.label, amt: Math.round(i.amt), note: i.note, left: Math.round(i.left), ticked: i.ticked, auto: i.auto })) }; })(),
    subs: (() => { const t = subsTracker_(); return { could: t.could, saved: t.saved, couldList: t.couldList.map(subName_), savedList: t.savedList.map(subName_) }; })(),
    limits: limits, cats: categoryList_(), demo: isDemo_(), fixedCats: FIXED_CATS.concat(NO_LIMIT_CATS)
  });
}

function prettyName_(desc) {
  const s = String(desc).split('|')[0]
    .replace(/APLPAY |TST\* ?|SQ \* ?|POS DEBIT[- ]+VISA CHECK CARD \d+ - |POS DEBIT- +\d+ \d+ /gi, '')
    .replace(/[0-9#*.]+/g, ' ').replace(/\s+/g, ' ').trim().split(' ').slice(0, 3).join(' ');
  return s.toLowerCase().replace(/\b\w/g, ch => ch.toUpperCase()) || String(desc).slice(0, 30);
}

const APP_HTML = `<!DOCTYPE html><html><head><base target="_top">
<style>
:root{--bg:#0f1115;--card:#181b22;--card2:#20242d;--line:#2a2f3a;--ink:#eef0f4;--mute:#8b93a3;--good:#34d399;--bad:#f87171;--warn:#fbbf24;--acc:#60a5fa}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}
.wrap{max-width:760px;margin:0 auto;padding:14px 14px 60px}
.top{display:flex;align-items:center;justify-content:space-between;margin:4px 0 14px}
.top h1{font-size:18px;font-weight:600;margin:0}
.sw{display:flex;align-items:center;gap:6px;background:var(--card);border:1px solid var(--line);border-radius:999px;padding:4px}
.sw button{background:none;border:0;color:var(--ink);font-size:18px;width:34px;height:30px;border-radius:999px;cursor:pointer}
.sw button:hover{background:var(--card2)}.sw button:disabled{opacity:.25}
.sw span{min-width:96px;text-align:center;font-weight:600}
.pills{display:flex;gap:6px;overflow-x:auto;margin:-4px 0 12px;padding-bottom:2px;scrollbar-width:none}
.pills::-webkit-scrollbar{display:none}
.pl{flex:none;border:0;border-radius:999px;padding:6px 12px;font-size:13px;background:var(--card);color:var(--mute);cursor:pointer}
.pl.on{background:var(--good);color:#0b1220;font-weight:600}
.two{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-bottom:12px}
.two .card{margin:0}
.chk{display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:12px;background:var(--card2);margin-top:8px;cursor:pointer}
.bx{width:20px;height:20px;border-radius:6px;border:2px solid #4b5563;flex:none;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700}
.chk.on .bx{background:var(--good);border-color:var(--good);color:#0b1220}
.chk.on .ct{text-decoration:line-through;color:var(--mute)}
.yel{background:#2a2410;border-color:#5a4514}.grn{background:#0f2a1f;border-color:#1d5c40}
.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px;margin-bottom:12px}
.lbl{font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:var(--mute);margin-bottom:8px}
.bill{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-top:1px solid var(--line)}
.bill:first-of-type{border-top:0}.big{font-size:34px;font-weight:700;letter-spacing:-.02em}
.sub{color:var(--mute);font-size:13px}.good{color:var(--good)}.bad{color:var(--bad)}.warn{color:var(--warn)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:12px}
.grid .card{margin:0;cursor:pointer;transition:border-color .15s}.grid .card:hover,.grid .card.on{border-color:var(--acc)}
.pill{display:inline-block;font-size:12px;padding:2px 9px;border-radius:999px;background:var(--card2)}
.row{display:flex;align-items:center;gap:10px;padding:7px 0;cursor:pointer;border-radius:8px}
.row:hover .nm{color:var(--acc)}.row.on .nm{color:var(--acc);font-weight:600}
.nm{width:130px;flex:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar{flex:1;height:10px;background:var(--card2);border-radius:99px;overflow:hidden;position:relative}
.bar i{display:block;height:100%;border-radius:99px;background:var(--acc)}
.amt{width:80px;text-align:right;font-variant-numeric:tabular-nums}
.meter{height:14px;background:var(--card2);border-radius:99px;position:relative;margin:10px 0 6px;overflow:visible}
.meter i{display:block;height:100%;border-radius:99px}.meter b{position:absolute;top:-4px;bottom:-4px;width:2px;background:var(--ink);opacity:.7}
.split{display:flex;justify-content:space-between;align-items:baseline}
.tx{display:flex;justify-content:space-between;gap:10px;padding:9px 0;border-top:1px solid var(--line)}
.tx .l{min-width:0}.tx .t{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tx .r{text-align:right;flex:none;font-variant-numeric:tabular-nums}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}
.chip{background:var(--acc);color:#0b1220;border:0;border-radius:999px;padding:4px 10px;font-size:13px;font-weight:600;cursor:pointer}
input{width:100%;background:var(--card2);border:1px solid var(--line);color:var(--ink);border-radius:10px;padding:9px 12px;font-size:15px;margin-bottom:6px}
.load{padding:80px 0;text-align:center;color:var(--mute)}
.tx.tap{cursor:pointer}.tx.tap:hover .t{color:var(--acc)}
.edit{background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:12px;margin:2px 0 10px}
.edit select{width:100%;background:var(--card);border:1px solid var(--line);color:var(--ink);border-radius:10px;padding:9px 10px;font-size:15px;margin-bottom:8px}
.edit label{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--mute);margin-bottom:10px}
.edit input[type=checkbox]{width:18px;height:18px;margin:0;flex:none}
.btn{border:0;border-radius:999px;padding:8px 16px;font-size:14px;font-weight:600;cursor:pointer;background:var(--card);color:var(--ink)}
.btn.go{background:var(--good);color:#0b1220}
.flash{background:#0f2a1f;border:1px solid #1d5c40;color:var(--good);border-radius:10px;padding:8px 10px;font-size:13px;margin-bottom:8px}
</style></head><body><div class="wrap" id="app"><div class="load">Loading your finances…</div></div>
<script>
var TOP=[],D,M,MONTHS,F={cat:'',card:null,q:''};
var $=function(n){return '$'+Math.round(Math.abs(n)).toLocaleString();};
var sgn=function(n){return (n<0?'−':'')+$(n);};
var esc=function(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});};
var mName=function(m){var p=m.split('-');return new Date(+p[0],+p[1]-1,1).toLocaleString('en-US',{month:'short',year:'numeric'});};
var LIST=[],FLASH='';
function setData(j,keepM){D=JSON.parse(j);var s={};D.tx.forEach(function(t){s[t.m]=1;});s[D.current]=1;MONTHS=Object.keys(s).sort();
  M=keepM&&MONTHS.indexOf(keepM)>-1?keepM:D.current;}
google.script.run.withSuccessHandler(function(j){setData(j);render();}).withFailureHandler(function(e){document.getElementById('app').innerHTML='<div class="load">Couldn\\'t load: '+esc(e.message)+'</div>';}).getAppData();

function go(d){var i=MONTHS.indexOf(M)+d;if(i>=0&&i<MONTHS.length){M=MONTHS[i];F.cat='';F.card=null;render();}}
function pick(m){M=m;F.cat='';F.card=null;render();}
function tick(i){var it=D.plan.items[i];it.ticked=!it.ticked;render();google.script.run.withFailureHandler(function(){it.ticked=!it.ticked;render();}).setTick(D.current,it.k,it.ticked);}
function setCat(c){F.cat=F.cat===c?'':c;F.card=null;render();}
function setCard(i){F.card=F.card===i?null:i;F.cat='';render();}
function clr(){F.cat='';F.card=null;F.q='';render();}
function ed(i){var id=i<0?'':LIST[i].id;F.ed=F.ed===id?'':id;FLASH='';render();}
function saveCat(i){var t=LIST[i],c=document.getElementById('ec').value,r=document.getElementById('er').checked,m=document.getElementById('em');
  m.textContent='Saving…';
  google.script.run.withSuccessHandler(function(j){setData(j,M);F.ed='';FLASH='Saved: '+t.n+' → '+c+(D.ruleMade?' · rule added, so '+D.ruleMade+' is always '+c:'');render();
    google.script.run.withFailureHandler(function(){}).refreshAfterPhoneEdit();})
  .withFailureHandler(function(e){m.textContent='Could not save: '+e.message;}).setTxCategory(t.id,c,r);}

function render(){
  var cur=M===D.current,mt=D.tx.filter(function(t){return t.m===M;});
  var inc=mt.filter(function(t){return t.inc;}).reduce(function(s,t){return s+t.a;},0);
  var byCat={};mt.filter(function(t){return t.spend;}).forEach(function(t){byCat[t.c]=(byCat[t.c]||0)-t.a;});
  var cats=Object.keys(byCat).map(function(k){return [k,byCat[k]];}).filter(function(x){return x[1]>0.5;}).sort(function(a,b){return b[1]-a[1];});
  var spent=Object.keys(byCat).reduce(function(s,k){return s+byCat[k];},0),coming=cur&&D.coming?D.coming.amt:0,kp=cur?(Number(D.keepPct)||0):0,payAll=cur?Math.max(inc,Number(D.plan&&D.plan.pay)||0):inc,keep=cur?(Number(D.plan&&D.plan.setAside)||payAll*kp):0,bills=cur?(Number(D.plan&&D.plan.bills)||0):0,left=payAll-keep-bills-spent,top=cats.filter(function(c){return (D.fixedCats||[]).indexOf(c[0])<0;}).slice(0,3),rest=Math.max(0,spent-top.reduce(function(s,x){return s+x[1];},0));
  var h='';
  h+='<div class="top"><h1>Finance HQ'+(D.demo?' <span style="font-size:13px;font-weight:600;color:#F59E0B">DEMO</span>':'')+'</h1><div class="sub" style="text-align:right">'+(D.ranges[M]||'')+'</div></div>';
  h+='<div class="pills">'+MONTHS.slice(-6).map(function(m){return '<button class="pl'+(m===M?' on':'')+'" onclick="pick(\\''+m+'\\')">'+mName(m).split(' ')[0]+(m===D.current?' · now':'')+'</button>';}).join('')+'</div>';

  // Bills due (always current)
  var open=D.cards.filter(function(c){return c.configured&&!c.paid&&c.statement>0;}),due=open.reduce(function(s,c){return s+c.statement;},0);
  h+='<div class="card"><div class="lbl">Bills due</div>';
  if(!open.length)h+='<div class="good" style="font-weight:600">All statements paid ✓</div>';
  open.sort(function(a,b){return a.daysToDue-b.daysToDue;}).forEach(function(c){
    var cl=c.daysToDue<=3?'bad':c.daysToDue<=7?'warn':'sub';
    h+='<div class="bill"><div><div style="font-weight:600">'+esc(c.name)+'</div><div class="sub">Due '+c.due+'</div></div><div style="text-align:right"><div style="font-weight:700;font-size:18px">'+$(c.statement)+'</div><div class="'+cl+'" style="font-size:13px">'+esc(c.status)+'</div></div></div>';});
  if(open.length){var ok=D.checking>=due;
    h+='<div class="sub" style="margin-top:8px;padding-top:8px;border-top:1px solid var(--line)">Checking '+$(D.checking)+' · '+(ok?'<span class="good">covers it ✓</span>':'<span class="bad">move '+$(due-D.checking)+' from savings before the due date</span>')+'</div>';}
  h+='</div>';

  // Left this month
  if(cur&&!(payAll>0)){h+='<div class="card"><div class="lbl">Spent so far '+(D.statement?'this period':'this month')+'</div><div class="big">'+$(spent)+'</div><div class="sub">No paycheck '+(inc>0?'':'found')+' this period yet. If a deposit is pay, set it to Income (tap it in Recent).</div><div style="margin-bottom:10px"></div>';}
  else{h+='<div class="card"><div class="lbl">'+(D.statement?(cur?'Left to spend this period':'Left over · '+(D.ranges[M]||mName(M))):(cur?'Left to spend this month':'Left over in '+mName(M)))+'</div>';
  h+='<div class="big '+(left>=0?'good':'bad')+'">'+sgn(left)+'</div><div class="sub">'+(cur?'Pay '+(coming?'~':'')+$(payAll)+(coming?' ('+$(inc)+' so far + '+(D.coming.irregular?'your usual for the rest':D.coming.n+' paycheck'+(D.coming.n>1?'s':'')+' coming')+')':'')+((D.plan&&D.plan.parts)||[]).map(function(x){return ' − '+x.label+' '+$(x.amt);}).join('')+(bills?' − bills still due '+$(bills):'')+' − spent '+$(spent):'Pay '+$(inc)+' − spending '+$(spent))+'</div>';
  if(cur){var dl=Number(D.daysLeft)||1;
    h+='<div style="margin:6px 0 10px;font-size:14px"><b>'+(dl<=1?'Last day':dl+' days left to spend')+'</b>'+(left>0&&dl>1?' <span class="sub">· about '+$(left/Math.max(dl,1))+'/day</span>':'')+'</div>';}
  else h+='<div style="margin-bottom:10px"></div>';}
  TOP=top;top.forEach(function(x,n){var lim=D.limits[x[0]]||0,r=lim?x[1]/lim:0,col=!lim?'sub':r>=1?'bad':r>=0.3?'warn':'good';
    h+='<div class="row'+(F.cat===x[0]?' on':'')+'" onclick="setCat(TOP['+n+'][0])" style="flex-wrap:wrap"><span class="nm">'+esc(x[0])+'</span><span class="bar"><i style="width:'+(lim?Math.max(2,Math.min(r,1)*100):Math.max(2,x[1]/spent*100))+'%;background:'+(lim?'var(--'+col+')':'#4b5563')+'"></i></span><span class="amt '+col+'">'+$(x[1])+'</span>'+
    '<span class="sub" style="width:100%;padding-left:140px;margin-top:-2px;font-size:12px">'+(!lim?'no limit set':(r>=1?'over by '+$(x[1]-lim):$(lim-x[1])+' left')+' of '+$(lim)+' limit')+'</span></div>';});
  if(rest>0.5)h+='<div class="row" style="cursor:default"><span class="nm sub">Everything else</span><span class="bar"><i style="width:'+(rest/spent*100)+'%;background:#4b5563"></i></span><span class="amt sub">'+$(rest)+'</span></div>';
  h+='</div>';

  // Saving plan (current period)
  if(cur&&D.plan){var P=D.plan;
    h+='<div class="two"><div class="card"><div class="lbl">Savings grew by</div><div style="font-size:26px;font-weight:700;color:var('+(P.saved>=0?'--acc':'--bad')+')">'+sgn(P.saved)+'</div><div class="sub">'+(P.goal?'goal '+$(P.goal)+(P.saved>=P.goal-0.5?' ✓':'')+(P.gross||P.out?' · in '+$(P.gross)+' · out '+$(P.out):''):P.pct?'goal '+P.pct+'% of pay — once pay lands':'set a goal in Settings')+'</div></div>'+
      '<div class="card"><div class="lbl">Pay this period</div><div style="font-size:26px;font-weight:700">'+$(P.got!=null?P.got:P.pay)+'</div><div class="sub">'+(P.got!=null&&P.pay>P.got+0.5?'so far · expecting ~'+$(P.pay)+' total':'landed this period')+'</div></div></div>';
    h+='<div class="card"><div class="split"><div class="lbl">Transfers to make</div><div class="sub">tap when done</div></div>';
    if(!P.items.length)h+='<div class="sub">Set a savings goal in Settings to get a checklist here.</div>';
    P.items.forEach(function(it,i){h+='<div class="chk'+(it.ticked?' on':'')+'" onclick="tick('+i+')"><span class="bx">'+(it.ticked?'✓':'')+'</span><div style="flex:1"><div class="ct" style="font-weight:600">'+esc(it.label)+' · '+$(it.amt)+'</div><div class="sub">'+esc(it.note)+'</div></div><div class="sub" style="text-align:right">'+(it.ticked?'<span class="good">done</span>':(it.k==='tax'?'set aside':$(it.left)+' to go'))+'</div></div>';});
    h+='</div>';
    var S2=D.subs;
    h+='<div class="two"><div class="card yel"><div class="lbl" style="color:var(--warn)">Could cancel</div><div style="font-size:24px;font-weight:700;color:var(--warn)">'+$(S2.could)+'<small style="font-size:13px">/yr</small></div><div class="sub" style="color:#d9b44a">'+(S2.couldList.length?esc(S2.couldList.join(', ')):'mark in Subscriptions tab')+'</div></div>'+
      '<div class="card grn"><div class="lbl" style="color:var(--good)">Already saved</div><div style="font-size:24px;font-weight:700;color:var(--good)">'+$(S2.saved)+'<small style="font-size:13px">/yr</small></div><div class="sub" style="color:#6ee7b7">'+(S2.savedList.length?esc(S2.savedList.join(', ')):'cancelled subs add up here')+'</div></div></div>';
  }

  // Card cycles
  if(D.cards.length){h+='<div class="grid">';D.cards.forEach(function(c,i){
    h+='<div class="card'+(F.card===i?' on':'')+'" onclick="setCard('+i+')"><div class="split"><div style="font-weight:600">'+esc(c.name)+'</div><span class="pill">'+(c.configured?'closes '+c.closes:'calendar month')+'</span></div>'+
    '<div style="font-size:26px;font-weight:700;margin-top:6px">'+$(c.spent)+'</div><div class="sub">this cycle · '+c.cycle+'</div>'+
    '<div class="sub" style="margin-top:6px">Last statement '+$(c.statement)+' · <span class="'+(c.paid?'good':'warn')+'">'+esc(c.status)+'</span></div></div>';});h+='</div>';}

  // All categories
  h+='<div class="card"><div class="lbl">Spending by category · tap to filter</div>';
  cats.forEach(function(x){h+=catRow(x[0],x[1],cats.length?cats[0][1]:1);});
  if(!cats.length)h+='<div class="sub">No spending yet this period.</div>';
  h+='</div>';

  // Transactions
  var list,chip='';
  if(F.card!==null){var c=D.cards[F.card];list=D.tx.filter(function(t){return t.acct===c.account&&t.d>=c.start;});chip=c.name+' · this cycle';}
  else{list=mt;if(F.cat){list=list.filter(function(t){return t.c===F.cat;});chip=F.cat;}}
  if(F.q){var q=F.q.toLowerCase();list=list.filter(function(t){return (t.n+' '+t.raw+' '+t.c).toLowerCase().indexOf(q)>-1;});}
  h+='<div class="card"><div class="split"><div class="lbl">Transactions · tap one to change its category</div><div class="sub">'+list.length+'</div></div>';
  if(FLASH)h+='<div class="flash">'+esc(FLASH)+'</div>';
  if(chip)h+='<div class="chips"><button class="chip" onclick="clr()">'+esc(chip)+'  ✕</button></div>';
  h+='<input id="q" placeholder="Search merchant or category" value="'+esc(F.q)+'" oninput="F.q=this.value;draw()">';
  h+='<div id="list">'+txHtml(list)+'</div></div>';
  document.getElementById('app').innerHTML=h;
}
function limitFor(k){if(D.limits[k])return D.limits[k];
  var i=MONTHS.indexOf(M),ms=MONTHS.slice(Math.max(0,i-3),i);if(!ms.length)ms=[M];
  var tot=-D.tx.filter(function(t){return t.spend&&t.c===k&&ms.indexOf(t.m)>-1;}).reduce(function(s,t){return s+t.a;},0);
  return Math.max(10,Math.round(tot/ms.length/10)*10);}
function draw(){var pos=document.getElementById('q').selectionStart;render();var q=document.getElementById('q');q.focus();q.setSelectionRange(pos,pos);}
function catRow(k,v,max){return '<div class="row'+(F.cat===k?' on':'')+'" onclick="setCat(\\''+esc(k).replace(/'/g,"\\\\'")+'\\')"><span class="nm">'+esc(k)+'</span><span class="bar"><i style="width:'+Math.max(2,v/max*100)+'%"></i></span><span class="amt">'+$(v)+'</span></div>';}
function txHtml(l){if(!l.length)return '<div class="sub" style="padding:10px 0">Nothing here.</div>';
  LIST=l.slice(0,250);
  return LIST.map(function(t,i){var d=new Date(t.d+'T12:00:00').toLocaleString('en-US',{month:'short',day:'numeric'});
    var box=F.ed&&F.ed===t.id?'<div class="edit"><select id="ec">'+(D.cats||[]).map(function(x){return '<option'+(x===t.c?' selected':'')+'>'+esc(x)+'</option>';}).join('')+'</select>'+
      '<label><input type="checkbox" id="er"> Always use this for '+esc(t.n)+'</label><button class="btn go" onclick="saveCat('+i+')">Save</button> <button class="btn" onclick="ed(-1)">Cancel</button><div id="em" class="sub" style="margin-top:8px"></div></div>':'';
    return '<div class="tx'+(t.id?' tap':'')+'"'+(t.id?' onclick="ed('+i+')"':'')+'><div class="l"><div class="t">'+esc(t.n)+'</div><div class="sub">'+d+' · '+esc(t.c)+(t.card?' · '+esc(t.card):'')+(t.trip?' · '+esc(t.trip):'')+'</div></div><div class="r '+(t.a>0?'good':'')+'">'+(t.a>0?'+':'')+'$'+Math.abs(t.a).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})+'</div></div>'+box;}).join('');}
</script></body></html>`;

// ===== PART 7 — Quick setup (one form instead of editing tabs) + Finance HQ menu =====
// Sheet menu "Finance HQ": Quick setup… · Refresh now · Start here. No need to open Apps Script after install.

// Points per $ for popular cards. Rates change — Quick setup shows the match so people can switch or skip.
const CARD_PRESETS = {
  'Amex Gold': { re: /gold/i, bank: /american express|amex/i, r: { 'Dining': 4, 'Groceries': 4, 'Travel': 1, 'Everything else': 1 } },
  'Amex Platinum': { re: /platinum/i, bank: /american express|amex/i, r: { 'Travel': 1, 'Everything else': 1 } },
  'Amex Blue Cash Preferred': { re: /blue cash pref/i, bank: /american express|amex/i, r: { 'Groceries': 6, 'Subscriptions': 6, 'Gas': 3, 'Transport': 3, 'Everything else': 1 } },
  'Amex Blue Cash Everyday': { re: /blue cash every/i, bank: /american express|amex/i, r: { 'Groceries': 3, 'Gas': 3, 'Shopping': 3, 'Everything else': 1 } },
  'Chase Sapphire Preferred': { re: /sapphire pref/i, bank: /chase/i, r: { 'Dining': 3, 'Gas': 3, 'Subscriptions': 3, 'Travel': 2, 'Everything else': 1 } },
  'Chase Sapphire Reserve': { re: /sapphire res/i, bank: /chase/i, r: { 'Dining': 3, 'Travel': 4, 'Everything else': 1 } },
  'Chase Freedom Unlimited': { re: /freedom unl/i, bank: /chase/i, r: { 'Dining': 3, 'Health': 3, 'Everything else': 1.5 } },
  'Chase Freedom Flex': { re: /freedom flex/i, bank: /chase/i, r: { 'Dining': 3, 'Health': 3, 'Everything else': 1 } },
  'Capital One Savor': { re: /savor/i, bank: /capital one/i, r: { 'Dining': 3, 'Entertainment': 3, 'Groceries': 3, 'Subscriptions': 3, 'Everything else': 1 } },
  'Capital One Venture': { re: /venture(?! ?x)/i, bank: /capital one/i, r: { 'Everything else': 2 } },
  'Capital One Venture X': { re: /venture ?x/i, bank: /capital one/i, r: { 'Everything else': 2 } },
  'Capital One Quicksilver': { re: /quicksilver/i, bank: /capital one/i, r: { 'Everything else': 1.5 } },
  'Citi Double Cash': { re: /double cash/i, bank: /citi/i, r: { 'Everything else': 2 } },
  'Citi Custom Cash': { re: /custom cash/i, bank: /citi/i, r: { 'Everything else': 1 } },
  'Wells Fargo Active Cash': { re: /active cash/i, bank: /wells/i, r: { 'Everything else': 2 } },
  'Wells Fargo Autograph': { re: /autograph/i, bank: /wells/i, r: { 'Dining': 3, 'Travel': 3, 'Gas': 3, 'Transport': 3, 'Subscriptions': 3, 'Bills & utilities': 3, 'Everything else': 1 } },
  'Bilt Mastercard': { re: /bilt/i, bank: /./, r: { 'Dining': 3, 'Travel': 2, 'Everything else': 1 } },
  'Discover it': { re: /discover/i, bank: /discover/i, r: { 'Everything else': 1 } },
  'Apple Card': { re: /apple card/i, bank: /apple|goldman/i, r: { 'Everything else': 2 } },
  'Navy Federal cashRewards': { re: /cashrewards/i, bank: /navy/i, r: { 'Everything else': 1.5 } }
};
const LENDER_RE = /DEPT EDUCATION|STUDENT LN|AIDVANTAGE|MOHELA|NELNET|EDFINANCIAL|NAVIENT|SALLIE MAE|SOFI LOAN|EARNEST|TOYOTA FIN|HONDA FIN|NISSAN|ALLY|GM FINANCIAL|FORD CREDIT|CAPITAL ONE AUTO|CHASE AUTO|WELLS FARGO AUTO|CARMAX|TESLA FIN|HYUNDAI MOTOR|KIA MOTORS|MORTGAGE|ROCKET|MR COOPER|PENNYMAC|LOANCARE|UPSTART|LENDINGCLUB|PROSPER|BEST EGG|MARCUS|LIGHTSTREAM|DISCOVER PERSONAL/;

function onOpen() { financeMenu_(); }
function onOpenInstalled() { financeMenu_(); }
function financeMenu_() {
  try {
    const m = SpreadsheetApp.getUi().createMenu('Finance HQ');
    if (!isConnected_()) {                       // brand-new copy: one clear first step
      m.addItem('Connect my bank…', 'showConnect').addItem('See a demo with sample data', 'startDemo').addItem('Get help…', 'showSupport').addItem('What check-ins share', 'showCheckinPreview').addToUi();
      return;
    }
    m.addItem('Quick setup…', 'showQuickSetup')
      .addItem('Refresh now', 'refreshNow')
      .addItem('Year in review (PDF)', 'yearInReviewNow')
      .addSeparator()
      .addItem('Health check', 'runHealthCheck')
      .addItem('Start here (guide)', 'openStartHere')
      .addItem('What check-ins share', 'showCheckinPreview')
      .addItem('Set up the phone app…', 'showPhoneSetup')
      .addItem('Send feedback', 'sendFeedback')
      .addItem('Get help…', 'showSupport')
      .addSeparator()
      .addItem(updateLabel_(), 'installUpdateNow')
      .addToUi();
  } catch (e) {}
}
function ensureMenu_() {
  // Scripts attached to the sheet get the menu automatically; a separate project needs an "on open" trigger.
  if (SpreadsheetApp.getActiveSpreadsheet()) return;
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onOpenInstalled')) return;
  try { ScriptApp.newTrigger('onOpenInstalled').forSpreadsheet(finSS_()).onOpen().create(); } catch (e) { log_('Menu trigger not created: ' + e); }
}
function openStartHere() { const sh = finSS_().getSheetByName('Start here'); if (sh) finSS_().setActiveSheet(sh); }

function refreshNow() {
  const ss = finSS_();
  toast_(ss, 'Syncing your accounts…', 'Finance HQ', 30);
  try { financeDailySync(); } catch (e) { toast_(ss, 'Sync problem: ' + e.message, 'Finance HQ', 10); }
  toast_(ss, 'Rebuilding…', 'Finance HQ', 30);
  phase2Daily();
  try { if (typeof phase3Daily === 'function') phase3Daily(); } catch (e) {}
  toast_(ss, 'All up to date ✓', 'Finance HQ', 5);
}

function presetFor_(account, inst) {
  const a = String(account), b = String(inst);
  for (const k of Object.keys(CARD_PRESETS)) { const p = CARD_PRESETS[k]; if (p.re.test(a) && p.bank.test(a + ' ' + b)) return k; }
  if (/sapphire/i.test(a)) return 'Chase Sapphire Preferred';
  return '';
}

function setSetting_(key, value, note) {
  const sh = sheet_('Settings');
  const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(value);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[key, value, note || '']]);
}

function showQuickSetup() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(QS_HTML).setWidth(760).setHeight(680), 'Finance HQ — quick setup');
}

function getSetupData() {
  ensureSettings_(); if (typeof ensureP3Settings_ === 'function') ensureP3Settings_();
  const S = settings_(), tz = Session.getScriptTimeZone();
  const acc = sheet_('Accounts');
  const accts = acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, 5).getValues() : [];
  const R = loadRewards_();
  const cards = [], seenCard = {};
  accts.filter(a => a[3] === 'Credit card').forEach(a => {
    const name = cardName_(a[2], a[1]); if (seenCard[name]) return; seenCard[name] = 1;
    const has = Object.keys(R).some(c => R[c][name]);
    cards.push({ name: name, account: String(a[2]), close: S[name + ' closes (day of month)'] || '', due: S[name + ' due (day of month)'] || '',
      track: cardTracked_(name, S), preset: has ? '(keep my Rewards tab)' : !cardTracked_(name, S) ? '(none / skip)' : presetFor_(a[2], a[1]) });
  });
  // debts already entered + ones spotted in your accounts and payments
  const man = sheet_('Manual');
  const debts = man.getLastRow() > 1 ? man.getRange(2, 1, man.getLastRow() - 1, 5).getValues().filter(r => r[0] !== '')
    .map(r => ({ name: String(r[0]), type: debtType_(r[1], r[0], ''), bal: r[2], rate: Number(r[3]) ? Math.round(Number(r[3]) * (Number(r[3]) < 1 ? 10000 : 100)) / 100 : '', pay: r[4], found: '' })) : [];
  const known = new Set(debts.map(d => d.name.toUpperCase()));
  accts.filter(a => a[3] === 'Loan').forEach(a => {
    if (!known.has(String(a[2]).toUpperCase())) debts.push({ name: String(a[2]), type: debtType_('', a[2], a[1]), bal: Math.abs(Number(a[4]) || 0), rate: '', pay: '', found: 'linked account' });
  });
  const since = new Date(Date.now() - 100 * 864e5), groups = {};
  const haveTypes = new Set(debts.map(d => d.type));
  loadTx_().filter(t => t.date >= since && t.amt < 0 && t.type !== 'Credit card' &&
      (t.cat === 'Student loans' || t.cat === 'Loan payments' || LENDER_RE.test(t.desc.toUpperCase())))
    .forEach(t => {
      const up = t.desc.toUpperCase(), m = up.match(LENDER_RE);
      const k = m ? m[0] : payeeKey_(t.desc);
      if (!k) return;
      (groups[k] = groups[k] || { amts: [], hint: t.cat + ' ' + up }).amts.push(-t.amt);
    });
  Object.keys(groups).forEach(k => {
    const ty = debtType_('', k, groups[k].hint);
    if (ty === ASSET_TYPE || haveTypes.has(ty) || [...known].some(n => n.indexOf(k) !== -1)) return;   // already listed
    debts.push({ name: k.replace(/\b\w+/g, w => w[0] + w.slice(1).toLowerCase()), type: ty, bal: '', rate: '', pay: Math.round(median_(groups[k].amts)), found: 'payments seen' });
  });
  const tr = sheet_('Trips');
  const trips = tr.getLastRow() > 1 ? tr.getRange(2, 1, tr.getLastRow() - 1, 3).getValues().filter(r => r[0] !== '')
    .map(r => ({ name: String(r[0]), start: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]).slice(0, 10), end: r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd') : String(r[2]).slice(0, 10) })) : [];
  const dining = explicitLimit_('Dining', S);
  const yearly = loadYearly_().map(y => ({ name: y.name, amt: y.amt, month: y.month >= 0 ? MON_[y.month] : '' }));
  return JSON.stringify({
    cards: cards, presets: Object.keys(CARD_PRESETS), debts: debts, types: DEBT_TYPES, trips: trips, yearly: yearly, months: MON_,
    dining: dining, email: String(S['Alert email'] || '') || Session.getEffectiveUser().getEmail(),
    phone: yes_(S['Alerts: phone reminders (Yes/No)']), digest: yes_(S['Weekly digest (Yes/No)']),
    cushion: Number(S['Cash held for others']) || 0,
    save: S['Savings goal (% of pay)'] === '' || S['Savings goal (% of pay)'] === undefined ? 20 : pctSetting_(S['Savings goal (% of pay)']),
    give: pctSetting_(S['Giving goal (% of pay)']), tax: pctSetting_(S['Tax set-aside (% of pay)'])
  });
}

function saveSetup(json) {
  const d = JSON.parse(json);
  // cards
  const rewardRows = [];
  d.cards.forEach(c => {
    setSetting_(c.name + ' closes (day of month)', Number(c.close) || '', 'statement closing day');
    setSetting_(c.name + ' due (day of month)', Number(c.due) || '', 'payment due day');
    setSetting_(c.name + ' track (Yes/No)', c.track ? 'Yes' : 'No', 'No = hide this card');
    if (c.preset && CARD_PRESETS[c.preset]) { const r = CARD_PRESETS[c.preset].r; Object.keys(r).forEach(cat => rewardRows.push([c.name, cat, r[cat]])); }
  });
  if (rewardRows.length) {
    const rw = sheet_('Rewards'), keepCards = new Set(rewardRows.map(r => r[0]));
    const old = rw.getLastRow() > 1 ? rw.getRange(2, 1, rw.getLastRow() - 1, 3).getValues().filter(r => r[0] && !keepCards.has(String(r[0]))) : [];
    const all = old.concat(rewardRows);
    if (rw.getLastRow() > 1) rw.getRange(2, 1, rw.getLastRow() - 1, 3).clearContent();
    rw.getRange(1, 1, 1, 3).setValues([['Card', 'Category', 'Points per $']]);
    rw.getRange(2, 1, all.length, 3).setValues(all);
  }
  // budget + alerts
  if (Number(d.dining) > 0) setLimit_('Dining', Number(d.dining));
  setSetting_('Cash held for others', Number(d.cushion) || 0);
  if (d.save !== undefined) setSetting_('Savings goal (% of pay)', Math.max(0, Math.min(90, Number(d.save) || 0)));
  if (d.give !== undefined) setSetting_('Giving goal (% of pay)', Math.max(0, Math.min(50, Number(d.give) || 0)));
  if (d.tax !== undefined) setSetting_('Tax set-aside (% of pay)', Math.max(0, Math.min(50, Number(d.tax) || 0)));
  if (d.email) setSetting_('Alert email', d.email);
  setSetting_('Alerts: phone reminders (Yes/No)', d.phone ? 'Yes' : 'No');
  setSetting_('Weekly digest (Yes/No)', d.digest ? 'Yes' : 'No');
  // debts (keeps any Manual rows that are assets)
  const man = sheet_('Manual');
  const keep = man.getLastRow() > 1 ? man.getRange(2, 1, man.getLastRow() - 1, 6).getValues().filter(r => r[0] !== '' && debtType_(r[1], r[0], r[5]) === ASSET_TYPE) : [];
  const notes = {}; if (man.getLastRow() > 1) man.getRange(2, 1, man.getLastRow() - 1, 6).getValues().forEach(r => notes[r[0]] = r[5]);
  const debtRows = d.debts.filter(x => x.name && Number(x.bal) > 0).map(x => {
    const rate = Number(x.rate) || 0;
    return [x.name, x.type, Number(x.bal), rate > 1 ? rate / 100 : rate, Number(x.pay) || '', notes[x.name] || ''];
  });
  const rows = debtRows.concat(keep);
  if (man.getLastRow() > 1) man.getRange(2, 1, man.getLastRow() - 1, 6).clearContent();
  if (rows.length) man.getRange(2, 1, rows.length, 6).setValues(rows);
  // trips
  const tr = sheet_('Trips'), tNotes = {};
  if (tr.getLastRow() > 1) tr.getRange(2, 1, tr.getLastRow() - 1, 4).getValues().forEach(r => tNotes[String(r[0])] = r[3]);   // keep your notes
  const tRows = d.trips.filter(t => t.name && t.start && t.end).map(t => [t.name, new Date(t.start + 'T12:00:00'), new Date(t.end + 'T12:00:00'), tNotes[t.name] || '']);
  if (tr.getLastRow() > 1) tr.getRange(2, 1, tr.getLastRow() - 1, 4).clearContent();
  if (tRows.length) tr.getRange(2, 1, tRows.length, 4).setValues(tRows);
  // yearly costs (keeps your notes)
  if (d.yearly) {
    const ys = ensureYearlyTab_(), yNotes = {};
    if (ys.getLastRow() > 1) ys.getRange(2, 1, ys.getLastRow() - 1, 4).getValues().forEach(r => yNotes[String(r[0]).trim()] = r[3]);
    const yRows = d.yearly.filter(y => String(y.name || '').trim() && Number(y.amt) > 0).map(y => [String(y.name).trim(), Number(y.amt), y.month || '', yNotes[String(y.name).trim()] || '']);
    if (ys.getLastRow() > 1) ys.getRange(2, 1, ys.getLastRow() - 1, 4).clearContent();
    if (yRows.length) { ys.getRange(2, 3, yRows.length, 1).setNumberFormat('@'); ys.getRange(2, 1, yRows.length, 4).setValues(yRows); }
  }
  PERIOD_ = null;
  phase2Daily();
  return 'Saved. Your Dashboard, Debt plan and budget periods are updated.';
}

const QS_HTML = `<!DOCTYPE html><html><head><base target="_top"><style>
body{font:14px/1.45 Arial,sans-serif;color:#111827;margin:0;padding:4px 6px 70px}
h3{font-size:13px;letter-spacing:.05em;text-transform:uppercase;color:#fff;background:#2F6F4E;padding:7px 10px;border-radius:6px;margin:18px 0 8px}
.hint{color:#6B7280;font-size:12px;margin:0 0 8px}
table{width:100%;border-collapse:collapse}td,th{padding:5px 4px;text-align:left;vertical-align:middle}th{font-size:11px;color:#6B7280;font-weight:normal}
tr:nth-child(even) td{background:#F7F8FA}
input,select{font:13px Arial;padding:5px 6px;border:1px solid #D1D5DB;border-radius:6px;background:#FFFBEB;width:100%;box-sizing:border-box}
input[type=checkbox]{width:auto}
.n{width:70px}.tag{font-size:11px;color:#059669}
.bar{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid #E5E7EB;padding:10px 12px;display:flex;justify-content:space-between;align-items:center}
button{font:14px Arial;padding:8px 16px;border-radius:8px;border:1px solid #D1D5DB;background:#fff;cursor:pointer}
.go{background:#059669;color:#fff;border-color:#059669;font-weight:bold}
.add{font-size:12px;padding:4px 10px;margin-top:6px}
#msg{color:#6B7280;font-size:13px}
</style></head><body><div id="app">Loading…</div>
<div class="bar"><span id="msg">Everything is optional — fill what you know, skip the rest.</span><span><button onclick="google.script.host.close()">Close</button> <button class="go" onclick="save()">Save</button></span></div>
<script>
var D;
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function opt(list,v){return list.map(function(x){return '<option'+(x===v?' selected':'')+'>'+esc(x)+'</option>';}).join('');}
google.script.run.withSuccessHandler(function(j){D=JSON.parse(j);draw();}).withFailureHandler(function(e){document.getElementById('app').textContent='Could not load: '+e.message;}).getSetupData();
function draw(){
  var h='';
  h+='<h3>1 · Your cards</h3><p class="hint">Closing and due days are on your statement (or in the card app). They let Finance HQ budget by statement and remind you before bills are due. Rewards are matched to your card automatically — change it if it is wrong.</p>';
  h+='<table><tr><th>Card</th><th>Statement closes (day)</th><th>Payment due (day)</th><th>Rewards</th><th>Show</th></tr>';
  D.cards.forEach(function(c,i){
    h+='<tr><td><b>'+esc(c.name)+'</b><br><span class="hint">'+esc(c.account)+'</span></td>'+
    '<td><input class="n" type="number" min="1" max="31" id="cl'+i+'" value="'+esc(c.close)+'" oninput="autoDue('+i+')"></td>'+
    '<td><input class="n" type="number" min="1" max="31" id="du'+i+'" value="'+esc(c.due)+'"></td>'+
    '<td><select id="pr'+i+'">'+opt(['(none / skip)','(keep my Rewards tab)'].concat(D.presets),c.preset||'(none / skip)')+'</select></td>'+
    '<td><input type="checkbox" id="tk'+i+'" '+(c.track?'checked':'')+'></td></tr>';});
  if(!D.cards.length)h+='<tr><td colspan="5" class="hint">No credit cards linked yet.</td></tr>';
  h+='</table>';
  h+='<h3>2 · Savings goal</h3><p class="hint">Pay yourself first: this share of each paycheck goes to savings or investing. 20% is a common target. You will get a checklist on the Dashboard.</p><table>'+
   '<tr><td>Save this % of pay</td><td><input id="gsave" type="number" min="0" max="90" value="'+esc(D.save)+'"></td></tr>'+
   '<tr><td>Giving (tithe / charity) % <span class="hint">optional</span></td><td><input id="ggive" type="number" min="0" max="50" value="'+esc(D.give)+'"></td></tr>'+
   '<tr><td>Tax set-aside % <span class="hint">only for side income with no tax taken out</span></td><td><input id="gtax" type="number" min="0" max="50" value="'+esc(D.tax)+'"></td></tr></table>';
  h+='<h3>3 · Budget and alerts</h3><table>'+
   '<tr><td>Dining budget per month</td><td><input id="dining" type="number" value="'+esc(D.dining||'')+'"></td></tr>'+
   '<tr><td>Money in your accounts that is not yours <span class="hint">(e.g. holding it for family)</span></td><td><input id="cushion" type="number" value="'+esc(D.cushion||0)+'"></td></tr>'+
   '<tr><td>Email for alerts and the weekly digest</td><td><input id="email" value="'+esc(D.email)+'"></td></tr>'+
   '<tr><td>Phone reminders (via Google Calendar)</td><td><input type="checkbox" id="phone" '+(D.phone?'checked':'')+'></td></tr>'+
   '<tr><td>Sunday weekly digest email</td><td><input type="checkbox" id="digest" '+(D.digest?'checked':'')+'></td></tr></table>'+
   '<p class="hint">Other category limits are suggested from your own spending — adjust them later in Settings if you like.</p>'+
   '<p class="hint">Weekly check-ins send the person who shared Finance HQ a few usage counts (never amounts or account details). Menu → What check-ins share · turn off in Settings.</p>';
  h+='<h3>4 · Debts</h3><p class="hint">Student loans, car loan, mortgage, personal loans… Rows marked found were spotted in your accounts — just add the balance and rate. Leave balance empty to skip a row.</p>';
  h+='<table><tr><th>Name</th><th>Type</th><th>Balance $</th><th>Rate %</th><th>Payment $/mo</th></tr><tbody id="debts">';
  D.debts.forEach(function(x,i){h+=debtRow(x,i);});
  h+='</tbody></table><button class="add" onclick="addDebt()">+ Add a debt</button>';
  h+='<h3>5 · Trips (optional)</h3><p class="hint">Spending during these dates is kept apart from your everyday budget.</p>';
  h+='<table><tr><th>Trip</th><th>First day</th><th>Last day</th></tr><tbody id="trips">';
  D.trips.forEach(function(t,i){h+=tripRow(t,i);});
  h+='</tbody></table><button class="add" onclick="addTrip()">+ Add a trip</button>';
  h+='<h3>6 · Yearly costs (optional)</h3><p class="hint">Bills that come once a year — annual card fee, car insurance paid yearly, Amazon Prime, registration. Finance HQ sets aside 1/12 each period so they never catch you off guard.</p>';
  h+='<table><tr><th>What</th><th>Yearly amount $</th><th>Month it is due</th></tr><tbody id="yearly">';
  D.yearly.forEach(function(y,i){h+=yearRow(y,i);});
  h+='</tbody></table><button class="add" onclick="addYear()">+ Add a yearly cost</button>';
  document.getElementById('app').innerHTML=h;
}
function debtRow(x,i){return '<tr><td><input id="dn'+i+'" value="'+esc(x.name)+'">'+(x.found?'<span class="tag">found: '+esc(x.found)+'</span>':'')+'</td><td><select id="dt'+i+'">'+opt(D.types,x.type)+'</select></td>'+
 '<td><input id="db'+i+'" type="number" value="'+esc(x.bal)+'"></td><td><input id="dr'+i+'" type="number" step="0.01" value="'+esc(x.rate)+'"></td><td><input id="dp'+i+'" type="number" value="'+esc(x.pay)+'"></td></tr>';}
function tripRow(t,i){return '<tr><td><input id="tn'+i+'" value="'+esc(t.name)+'"></td><td><input id="ts'+i+'" type="date" value="'+esc(t.start)+'"></td><td><input id="te'+i+'" type="date" value="'+esc(t.end)+'"></td></tr>';}
function addDebt(){var i=D.debts.length;D.debts.push({name:'',type:'Other debt',bal:'',rate:'',pay:''});document.getElementById('debts').insertAdjacentHTML('beforeend',debtRow(D.debts[i],i));}
function yearRow(y,i){return '<tr><td><input id="yn'+i+'" value="'+esc(y.name)+'"></td><td><input id="ya'+i+'" type="number" value="'+esc(y.amt)+'"></td><td><select id="ym'+i+'">'+opt([''].concat(D.months),y.month||'')+'</select></td></tr>';}
function addYear(){var i=D.yearly.length;D.yearly.push({name:'',amt:'',month:''});document.getElementById('yearly').insertAdjacentHTML('beforeend',yearRow(D.yearly[i],i));}
function addTrip(){var i=D.trips.length;D.trips.push({name:'',start:'',end:''});document.getElementById('trips').insertAdjacentHTML('beforeend',tripRow(D.trips[i],i));}
function autoDue(i){var c=+document.getElementById('cl'+i).value,d=document.getElementById('du'+i);if(c>=1&&c<=31&&!d.dataset.touched){var v=(c+25-1)%30+1;d.value=v;d.placeholder='estimate';}d.oninput=function(){d.dataset.touched=1;};}
function v(id){var e=document.getElementById(id);return e?(e.type==='checkbox'?e.checked:e.value):'';}
function save(){
  var bad=false;D.cards.forEach(function(c,i){['cl','du'].forEach(function(p){var x=v(p+i);if(x!==''&&(+x<1||+x>31))bad=true;});});
  if(+v('gsave')<0||+v('gsave')>90||+v('ggive')<0||+v('ggive')>50||+v('gtax')<0||+v('gtax')>50){document.getElementById('msg').innerHTML='<span style="color:#DC2626">Goal percentages look off — use numbers like 20 (for 20%).</span>';return;}
  if(bad){document.getElementById('msg').innerHTML='<span style="color:#DC2626">Closing and due days must be between 1 and 31.</span>';return;}
  var out={cards:D.cards.map(function(c,i){return{name:c.name,close:v('cl'+i),due:v('du'+i),track:v('tk'+i),preset:v('pr'+i)};}),
    dining:v('dining'),cushion:v('cushion'),email:v('email'),phone:v('phone'),digest:v('digest'),save:v('gsave'),give:v('ggive'),tax:v('gtax'),
    debts:D.debts.map(function(x,i){return{name:v('dn'+i),type:v('dt'+i),bal:v('db'+i),rate:v('dr'+i),pay:v('dp'+i)};}),
    trips:D.trips.map(function(t,i){return{name:v('tn'+i),start:v('ts'+i),end:v('te'+i)};}),
    yearly:D.yearly.map(function(y,i){return{name:v('yn'+i),amt:v('ya'+i),month:v('ym'+i)};})};
  document.getElementById('msg').textContent='Saving and rebuilding — about a minute…';
  document.querySelector('.go').disabled=true;
  google.script.run.withSuccessHandler(function(m){document.getElementById('msg').innerHTML='<span style="color:#059669">'+m+'</span>';setTimeout(function(){google.script.host.close();},1800);})
    .withFailureHandler(function(e){document.getElementById('msg').innerHTML='<span style="color:#DC2626">'+esc(e.message)+'</span>';document.querySelector('.go').disabled=false;}).saveSetup(JSON.stringify(out));
}
</script></body></html>`;

function autoRewards_() {
  const rw = sheet_('Rewards'); loadRewards_();
  if (rw.getLastRow() > 1) return;
  const acc = sheet_('Accounts'), rows = [], done = {};
  if (acc.getLastRow() > 1) acc.getRange(2, 1, acc.getLastRow() - 1, 4).getValues().filter(a => a[3] === 'Credit card').forEach(a => {
    const name = cardName_(a[2], a[1]), p = presetFor_(a[2], a[1]);
    if (!p || done[name]) return; done[name] = 1;
    Object.keys(CARD_PRESETS[p].r).forEach(cat => rows.push([name, cat, CARD_PRESETS[p].r[cat]]));
  });
  if (rows.length) rw.getRange(2, 1, rows.length, 3).setValues(rows);
}

// ===== PART 8 — Live updates + Health check =====

// Rebuild after any edit you make (Settings, Manual, Trips, Rewards, Rules, a Transactions category…).
// Starts immediately; several quick edits are folded into one rebuild. Formatting is skipped to keep it fast.
const LIVE_TABS = ['Settings', 'Manual', 'Trips', 'Rewards', 'Rules', 'Transactions', 'Accounts', 'Yearly costs'];
function liveRebuild_(sheetName) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('DIRTY_AT', String(Date.now()));
  if (sheetName === 'Rules') props.setProperty('DIRTY_RULES', '1');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;                  // a rebuild is already running — it will pick this edit up
  const ss = finSS_();
  try {
    let pass = 0;
    while (pass++ < 4) {
      const started = Date.now();
      Utilities.sleep(1500);                      // let a burst of typing finish
      toast_(ss, 'Updating your dashboard…', 'Finance HQ', 20);
      if (props.getProperty('DIRTY_RULES')) { props.deleteProperty('DIRTY_RULES'); recategorizeAll_(); }
      PERIOD_ = null;
      rebuild_(false);
      if (Number(props.getProperty('DIRTY_AT')) <= started) break;   // nothing new while we worked
    }
    toast_(ss, 'Dashboard updated ✓', 'Finance HQ', 3);
  } finally { lock.releaseLock(); }
}

// ---------------- health check ----------------
function buildHealth_() {
  const ss = finSS_(), tz = Session.getScriptTimeZone(), S = settings_(), now = new Date();
  const checks = [];   // [level 0=ok 1=warn 2=problem, check, found, fix]
  const add = (lvl, check, found, fix) => checks.push([lvl, check, found, fix || '']);
  const d$ = n => '$' + Math.round(Math.abs(n)).toLocaleString('en-US');

  // 1) connection + last sync
  const connected = !!PropertiesService.getScriptProperties().getProperty(SF_ACCESS);
  add(connected ? 0 : 2, 'Bank connection (SimpleFIN)', connected ? 'Connected' : 'Not connected',
    'Use Finance HQ → Connect my bank… and paste a SimpleFIN setup token (or add SIMPLEFIN_SETUP_TOKEN in Script properties and run setupFinanceHQ).');
  const logSh = ss.getSheetByName('Log');
  const logs = logSh && logSh.getLastRow() > 1 ? logSh.getRange(Math.max(2, logSh.getLastRow() - 300), 1, Math.min(300, logSh.getLastRow() - 1), 2).getValues() : [];
  const syncs = logs.filter(r => r[0] instanceof Date && /Daily sync|Backfill/i.test(r[1]));
  const accSyncs = (() => { const a = sheet_('Accounts'); return a.getLastRow() > 1 ? a.getRange(2, 8, a.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(d => d instanceof Date) : []; })();
  const lastSync = [syncs.length ? syncs[syncs.length - 1][0] : null].concat(accSyncs).filter(Boolean).reduce((m, d) => !m || d > m ? d : m, null);
  const hrs = lastSync ? (now - lastSync) / 36e5 : 999;
  add(hrs <= 30 ? 0 : hrs <= 54 ? 1 : 2, 'Last successful sync', lastSync ? Utilities.formatDate(lastSync, tz, 'EEE MMM d, h:mm a') : 'No sync recorded',
    'Use Finance HQ → Refresh now. If it keeps failing, check your connections on the SimpleFIN website.');
  const errs = logs.filter(r => r[0] instanceof Date && now - r[0] < 7 * 864e5 && /error|fail|problem|skipped/i.test(r[1]) && !/^(Automatic update skipped|Paste this|Tester tracker|Feedback answers|Published |Help request)/.test(r[1]));   // updates have their own row
  add(errs.length ? 1 : 0, 'Errors in the last 7 days', errs.length ? errs.length + ' — latest: ' + String(errs[errs.length - 1][1]).slice(0, 120) : 'None',
    'Open the Log tab (right-click any tab → Show hidden sheets) for details.');

  // 2) accounts
  const acc = sheet_('Accounts');
  const accts = acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, 9).getValues() : [];
  const broken = accts.filter(a => a[8] && a[8] !== 'OK');
  add(broken.length ? 2 : 0, 'All accounts reporting', broken.length ? broken.map(a => a[2]).join(', ') : accts.length + ' accounts OK',
    'Reconnect these on the SimpleFIN website (banks sometimes need you to sign in again).');
  const stale = accts.filter(a => a[6] instanceof Date && now - a[6] > 4 * 864e5);
  add(stale.length ? 1 : 0, 'Balances up to date', stale.length ? stale.map(a => a[2] + ' (' + Utilities.formatDate(a[6], tz, 'MMM d') + ')').join(', ') : 'All current',
    'These balances are more than 4 days old — check them in SimpleFIN.');

  // 3) automations
  const want = ['financeDailySync', 'phase2Daily', 'archiveMonthly', 'archiveYearly', 'phase3Daily', 'weeklyDigest', 'onCategoryEdit'];
  const trig = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const missing = want.filter(w => trig.indexOf(w) === -1), dupes = want.filter(w => trig.filter(t => t === w).length > 1);
  add(missing.length ? 2 : dupes.length ? 1 : 0, 'Automatic schedules', missing.length ? 'Missing: ' + missing.join(', ') : dupes.length ? 'Doubled: ' + dupes.join(', ') : 'All ' + want.length + ' running',
    'Run setupFinanceHQ again — it is safe and replaces the schedules without touching your data.');
  const editT = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onCategoryEdit');
  const editOk = editT.some(t => { try { return t.getTriggerSourceId() === ss.getId(); } catch (e) { return true; } });
  add(editOk ? 0 : 2, 'Instant updates when you edit', editOk ? 'On for this sheet' : 'Not connected to this sheet — edits will not refresh the Dashboard',
    'Use Finance HQ → Refresh now once (it reconnects automatically), or run phase2Daily.');

  // 4) cards
  const cards = [...new Set(accts.filter(a => a[3] === 'Credit card').map(a => cardName_(a[2], a[1])))].filter(n => cardTracked_(n, S));
  const noDates = cards.filter(n => !(Number(S[n + ' closes (day of month)']) >= 1 && Number(S[n + ' due (day of month)']) >= 1));
  const badDates = cards.filter(n => [S[n + ' closes (day of month)'], S[n + ' due (day of month)']].some(v => v !== '' && v !== undefined && !(Number(v) >= 1 && Number(v) <= 31)));
  add(badDates.length ? 2 : noDates.length ? 1 : 0, 'Card statement dates', badDates.length ? 'Not a valid day (1–31): ' + badDates.join(', ') : noDates.length ? 'Missing for ' + noDates.join(', ') + ' — budgets use calendar months for these' : 'Set for all ' + cards.length + ' cards',
    'Finance HQ → Quick setup… and enter the closing and due day from each statement.');

  // 5) transactions that look wrong
  const tx = loadTx_(), since90 = new Date(now - 90 * 864e5);
  const payAsSpend = tx.filter(t => t.type === 'Credit card' && t.amt > 100 && /PAYMENT|THANK YOU|AUTOPAY/i.test(t.desc) && isSpend_(t));
  add(payAsSpend.length ? 2 : 0, 'Card payments counted as spending', payAsSpend.length ? payAsSpend.length + ' payments (' + d$(payAsSpend.reduce((s, t) => s + t.amt, 0)) + ') sit in "' + [...new Set(payAsSpend.map(t => t.cat))].join('", "') + '"' : 'None',
    'Add a Rule (Notes: mine) for the payment text → Transfer, or set the Category to Transfer in Transactions.');
  const key = t => [Utilities.formatDate(t.date, tz, 'yyyy-MM-dd'), t.account, t.amt, t.desc].join('|');
  const cnt = {}; tx.filter(t => t.date >= since90).forEach(t => cnt[key(t)] = (cnt[key(t)] || 0) + 1);
  const dup = Object.keys(cnt).filter(k => cnt[k] > 1 && Math.abs(Number(k.split('|')[2])) >= 50);
  add(dup.length ? 1 : 0, 'Possible duplicate transactions', dup.length ? dup.length + ' (e.g. ' + dup[0].split('|').slice(0, 3).join(' · ') + ')' : 'None found',
    'Sometimes real (two identical coffees). If not, delete the extra row in Transactions.');
  const cur = currentPeriod_(now), unc = tx.filter(t => t.month === cur && t.cat === 'Uncategorized' && t.amt < 0);
  const uncAmt = -unc.reduce((s, t) => s + t.amt, 0);
  add(uncAmt > 150 ? 1 : 0, 'Uncategorized spending this period', unc.length ? unc.length + ' purchases, ' + d$(uncAmt) : 'None',
    'Pick a Category in Transactions (it locks), or add Rules for stores that repeat.');
  const bigIn = tx.filter(t => t.date >= since90 && t.amt >= 300 && t.type !== 'Credit card' && t.cat === 'Uncategorized');
  add(bigIn.length ? 1 : 0, 'Large deposits with no category', bigIn.length ? bigIn.length + ' (' + d$(bigIn.reduce((s, t) => s + t.amt, 0)) + ') — paychecks are only counted if categorized Income' : 'None',
    'Set these to Income, Transfer or Reimbursement in Transactions.');
  const lastPay = tx.filter(t => t.cat === 'Income' && t.amt > 0).reduce((m, t) => t.date > m ? t.date : m, new Date(0));
  add(now - lastPay > 35 * 864e5 ? 1 : 0, 'Paycheck detected recently', lastPay.getTime() ? 'Last: ' + Utilities.formatDate(lastPay, tz, 'MMM d') : 'No income found',
    'Add a Rule for your paycheck text (e.g. your employer name) → Income, Notes: mine.');

  // 6) your own inputs
  const man = sheet_('Manual');
  const debtRows = man.getLastRow() > 1 ? man.getRange(2, 1, man.getLastRow() - 1, 5).getValues().filter(r => r[0] && isDebt_(r[1], r[0], '') && Number(r[2]) > 0) : [];
  const linkedNoRate = Object.keys(linkedLoans_()).map(k => linkedLoans_()[k].name).filter(n => !debtRows.some(r => String(r[0]).trim().toUpperCase() === n.trim().toUpperCase()));
  const noRate = debtRows.filter(r => !Number(r[3])).map(r => r[0]).concat(linkedNoRate);
  const hasTotal = Number(S['Debt payments per month (total)']) || Number(S['Student loan payment per month']);
  const noPay = debtRows.filter(r => !Number(r[4]) && !hasTotal).map(r => r[0]);
  add(noRate.length ? 1 : noPay.length ? 1 : 0, 'Debts complete', noRate.length ? 'No interest rate: ' + noRate.join(', ') : noPay.length ? 'No monthly payment (estimated): ' + noPay.join(', ') : debtRows.length ? debtRows.length + ' debts complete' : 'No debts entered',
    'Finance HQ → Quick setup… → Debts, or the Manual tab.');
  const trips = sheet_('Trips');
  const tripBad = trips.getLastRow() > 1 ? trips.getRange(2, 1, trips.getLastRow() - 1, 3).getValues().filter(r => r[0] && (!(new Date(r[1]) > 0) || !(new Date(r[2]) > 0) || new Date(r[2]) < new Date(r[1]))).map(r => r[0]) : [];
  add(tripBad.length ? 2 : 0, 'Trip dates', tripBad.length ? 'Check dates for: ' + tripBad.join(', ') : 'OK', 'The last day must be on or after the first day.');
  const badNums = Object.keys(S).filter(k => /monthly limit|target per month/.test(k) && S[k] !== '' && !(Number(S[k]) >= 0));
  add(badNums.length ? 2 : 0, 'Spending limits are numbers', badNums.length ? badNums.join(', ') : 'OK', 'Type plain numbers in Settings, e.g. 250.');
  const R = loadRewards_(), rCards = [...new Set([].concat.apply([], Object.keys(R).map(c => Object.keys(R[c]))))];
  const unknown = rCards.filter(c => cards.indexOf(c) === -1 && cardTracked_(c, S));
  add(unknown.length ? 1 : 0, 'Rewards match your cards', unknown.length ? 'Not one of your cards: ' + unknown.join(', ') : rCards.length ? 'OK' : 'No rewards set (CardCheck is off)',
    'Card names in the Rewards tab must match the names on the Dashboard. Quick setup fills these for you.');
  const trk = subsTracker_();
  add(trk.still.length ? 2 : 0, 'Cancelled subscriptions', trk.still.length ? 'Still charging: ' + trk.still.map(subName_).join(', ') : trk.savedList.length ? trk.savedList.length + ' cancelled, none charging' : 'None marked',
    'Check the cancellation went through with the company, or set Status back to "keep" in Subscriptions.');
  add(Number(S['Savings goal (% of pay)']) > 0 ? 0 : 1, 'Savings goal', Number(S['Savings goal (% of pay)']) > 0 ? S['Savings goal (% of pay)'] + '% of pay' : 'Not set',
    'Set "Savings goal (% of pay)" in Settings — 20% is a common target.');
  { const av = PropertiesService.getScriptProperties().getProperty('UPDATE_AVAILABLE');
    add(av ? 1 : 0, 'Updates', av ? 'Version ' + JSON.parse(av).version + ' is ready (you have ' + FHQ_VERSION + ')' : 'Up to date (' + FHQ_VERSION + ')',
      'Finance HQ → ⬆ Install update — about a minute.'); }
  const yrH = yearlyPerPeriod_(), noMonth = yrH.list.filter(y => y.month < 0).map(y => y.name);
  add(noMonth.length ? 1 : 0, 'Yearly costs', noMonth.length ? 'No month picked for: ' + noMonth.join(', ') : yrH.total ? yrH.list.length + ' listed — $' + Math.round(yrH.per).toLocaleString() + ' set aside each period' : 'None listed (optional)',
    'Pick the month each one is due in the Yearly costs tab.');
  try {
    const BP = budgetPlan_(tx, S), freeH = BP.pay - BP.keep - BP.fixed - BP.limTot;
    add(BP.pay > 0 && freeH < -25 ? 1 : 0, 'Limits fit your pay', !(BP.pay > 0) ? 'No pay yet to compare with' : freeH >= -25 ? 'Yes — $' + Math.round(Math.max(0, freeH)).toLocaleString() + ' unplanned after saving' : 'Limits are $' + Math.round(-freeH).toLocaleString() + ' more than you have after saving',
      'Open the Budget tab and tick "Use it?" on the suggestions, or lower a few limits.');
  } catch (e) {}
  const P = PERIOD_ || {};
  add(0, 'Budget period', P.on ? 'Statement periods (closes ~' + P.main + 'th) — ' + Math.round((P.share || 0) * 100) + '% of spending is on cards' : 'Calendar months', 'Change in Settings → Spending month.');

  // ---- write the tab
  const sh = ss.getSheetByName('Health') || ss.insertSheet('Health');
  sh.clear(); sh.getBandings().forEach(b => b.remove()); sh.getRange('A:H').breakApart(); sh.setHiddenGridlines(true);
  uiPage_(sh);
  [70, 230, 150, 150, 120, 120, 120, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A:H').setFontFamily('Arial').setFontColor(UI.ink).setVerticalAlignment('middle');
  sh.getRange('A1:F1').merge().setValue('Health check').setFontSize(18).setFontWeight('bold');
  sh.getRange('G1:H1').merge().setValue('Checked ' + Utilities.formatDate(now, tz, 'MMM d, h:mm a')).setFontColor(UI.mute).setHorizontalAlignment('right');
  const nP = checks.filter(c => c[0] === 2).length, nW = checks.filter(c => c[0] === 1).length;
  let r = uiTiles_(sh, 3, [
    ['PROBLEMS', nP, '0', nP ? UI.bad : UI.good, nP ? 'fix these first' : 'none'],
    ['WORTH A LOOK', nW, '0', nW ? UI.warn : UI.good, nW ? 'not urgent' : 'none'],
    ['ALL GOOD', checks.length - nP - nW, '0', UI.good, 'of ' + checks.length + ' checks'],
    ['OVERALL', nP ? 'Needs attention' : nW ? 'Mostly fine' : 'Healthy ✓', '@', nP ? UI.bad : nW ? UI.warn : UI.good, 'updates with every refresh']
  ]);
  r = uiBand_(sh, r, 'Checks — problems first');
  r = uiHeader_(sh, r, ['Status', 'Check', 'What we found', '', '', 'How to fix', '', '']);
  checks.sort((a, b) => b[0] - a[0]).forEach((c, i) => {
    const icon = ['✅ OK', '⚠️ Look', '❌ Fix'][c[0]];
    sh.getRange(r, 1).setValue(icon).setFontWeight('bold').setFontColor([UI.good, UI.warn, UI.bad][c[0]]);
    sh.getRange(r, 2).setValue(c[1]).setFontWeight('bold').setWrap(true);
    sh.getRange(r, 3, 1, 3).merge().setValue(c[2]).setWrap(true);
    sh.getRange(r, 6, 1, 3).merge().setValue(c[0] ? c[3] : '').setWrap(true).setFontColor(UI.mute);
    sh.getRange(r, 1, 1, 8).setBackground(c[0] === 2 ? '#FEF2F2' : c[0] === 1 ? '#FFFBEB' : (i % 2 ? UI.stripe : '#FFFFFF'));
    sh.setRowHeight(r, 40); r++;
  });
  return checks;
}
function runHealthCheck() { buildHealth_(); finSS_().setActiveSheet(finSS_().getSheetByName('Health')); }

// ===== PART 9 — Savings goals, "Transfers to make" checklist, savings tracker =====

const KJ = { bg: '#F7F6F2', card: '#FFFFFF', line: '#E6E4DD', head: '#EFEDE6', ink: '#2A2A2A', soft: '#5C574D', mute: '#8A857A', faint: '#AAA39A',
  green: '#2F6F4E', greenBg: '#E1EFE6', red: '#B8423A', redBg: '#FBE7E3', orange: '#B86A2B', orangeBg: '#FBE9D8',
  blue: '#3C6C8A', blueBg: '#E3EEF6', yellow: '#C89B1E', yellowBg: '#F5EDCF', stripe: '#FAF9F6' };
const KJ_PIE = ['#2F6F4E', '#C89B1E', '#3C6C8A', '#B86A2B', '#B8423A', '#7A5C99', '#4F8A8B', '#D08C60', '#A8A398'];

const GOAL_SETTINGS = [
  ['Savings goal (% of pay)', 20, 'Pay yourself first — this % of each period\'s pay goes to savings or investing'],
  ['Giving goal (% of pay)', 0, 'Optional — tithe, church or charity'],
  ['Tax set-aside (% of pay)', 0, 'Only for income with no tax taken out (side gigs, 1099). Leave 0 for a normal paycheck']
];

// How much your savings + investments grew this period from money going in and out (not market gains):
//  counts deposits into savings (pay, transfers, cash) and money taken back out;
//  moving money between your own savings/investment accounts doesn't count either way;
//  money sent from checking to a brokerage counts as saved.
function savingsMoves_(tx, key) {
  const isStash = t => t.type === 'Savings' || t.type === 'Investment';
  const flowInv = t => /DEPOSIT|WITHDRAW|TRANSFER|CONTRIBUTION|ACH/i.test(t.desc) && !/DIVIDEND|INTEREST|REINVEST|BUY |SELL |SHARES/i.test(t.desc);
  const near = (a, b) => Math.abs(a.date - b.date) <= 3 * 864e5;
  const same = (a, b) => Math.abs(Math.abs(a.amt) - Math.abs(b.amt)) <= Math.max(0.01, Math.abs(a.amt) * 0.01) && Math.sign(a.amt) !== Math.sign(b.amt) && near(a, b);
  const win = tx.filter(t => Math.abs(t.date - periodBounds_(key)[0]) < 45 * 864e5 || t.month === key);
  const stashRows = win.filter(t => isStash(t) && (t.type === 'Savings' ? !/INTEREST|DIVIDEND/i.test(t.desc) : flowInv(t)));
  const used = new Set(), internal = new Set();
  stashRows.forEach(a => {                                      // pairs between two of your own savings/investment accounts
    if (internal.has(a)) return;
    const b = stashRows.find(x => x !== a && !internal.has(x) && x.account !== a.account && same(a, x));
    if (b) { internal.add(a); internal.add(b); }
  });
  let into = 0, out = 0, pay = 0;
  stashRows.filter(t => t.month === key && !internal.has(t)).forEach(t => {
    if (t.type === 'Savings' && t.amt < 0 && /^Investing/.test(t.cat)) return;   // sent to an investment account you haven't linked: still saved
    if (t.amt > 0) { into += t.amt; if (t.cat === 'Income') pay += t.amt; } else out += -t.amt;
  });
  const brokerDeposits = win.filter(t => t.type === 'Investment' && t.amt > 0 && flowInv(t));
  win.filter(t => t.month === key && t.type === 'Checking' && /^Investing/.test(t.cat) && t.amt < 0).forEach(t => {
    if (!brokerDeposits.some(d => same(t, d))) into += -t.amt;   // checking -> brokerage you haven't linked
  });
  return { into: into - out, gross: into, paidIn: pay, takenOut: out };
}

// Checklist for a period: Tax -> Giving -> Savings (each on what is left after the one before).
// "20", "20%" or 0.2 all mean 20%.
function pctSetting_(v) {
  const txt = String(v == null ? '' : v), n = Number(txt.replace('%', '').trim()) || 0;
  return Math.max(0, n > 0 && n < 1 && txt.indexOf('%') === -1 ? n * 100 : n);
}
// expected = pay you expect for the whole period (so far + paychecks still coming); the plan uses whichever is bigger.
function checklist_(tx, S, key, expected) {
  const st = periodStats_(key, tx), got = Math.max(0, st.pay), pay = Math.max(got, Number(expected) || 0);
  const pct = k => pctSetting_(S[k]) / 100;
  const t = pct('Tax set-aside (% of pay)'), g = pct('Giving goal (% of pay)'), s = pct('Savings goal (% of pay)');
  const taxAmt = pay * t, giveAmt = (pay - taxAmt) * g, saveAmt = (pay - taxAmt - giveAmt) * s;
  const moves = savingsMoves_(tx, key);
  const given = -tx.filter(x => x.month === key && x.cat === 'Giving' && isSpend_(x)).reduce((a, x) => a + x.amt, 0);
  const ticks = getTicks_()[key] || {};
  const items = [];
  if (t) items.push({ k: 'tax', label: 'Set aside for taxes', amt: taxAmt, note: Math.round(t * 100) + '% of pay', progress: 0 });
  if (g) items.push({ k: 'give', label: 'Giving', amt: giveAmt, note: Math.round(g * 100) + '% of after-tax pay', progress: given });
  const yr = typeof yearlyPerPeriod_ === 'function' ? yearlyPerPeriod_() : { per: 0, names: [] };
  if (s && pay > 0.5) items.push({ k: 'save', label: 'Move to savings', amt: saveAmt, note: Math.round(s * 100) + '% of ' + (t || g ? 'what\'s left' : pay > got + 0.5 ? 'expected pay $' + Math.round(pay).toLocaleString() : 'pay'), progress: moves.into });
  if (yr.per > 0) items.push({ k: 'yearly', label: 'Set aside for yearly costs', amt: yr.per,
    note: yr.names.slice(0, 3).join(', ') + (yr.names.length > 3 ? '…' : '') + ' ÷ 12 — keep it in savings', progress: Math.max(0, moves.into - saveAmt) });
  items.forEach(i => {
    i.auto = i.amt > 0.5 && i.progress >= i.amt - 0.5;
    i.ticked = ticks[i.k] === true || (ticks[i.k] !== false && i.auto);
    i.left = Math.max(0, i.amt - i.progress);
  });
  return { pay: pay, got: got, expected: pay > got + 0.5, items: items, moves: moves, goalSave: saveAmt, yearly: yr.per, setAside: taxAmt + giveAmt + saveAmt + yr.per,
    parts: items.map(i => ({ label: i.k === 'save' ? Math.round(s * 100) + '% to savings' : i.k === 'yearly' ? 'yearly costs' : i.k === 'give' ? 'giving' : 'taxes', amt: Math.round(i.amt * 100) / 100 })) };
}
function getTicks_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('TICKS') || '{}'); } catch (e) { return {}; } }
function setTick(key, item, value) {          // called from the phone app and the sheet checkbox
  bumpUsage_('tick');
  const all = getTicks_(); (all[key] = all[key] || {})[item] = !!value;
  Object.keys(all).sort().slice(0, -12).forEach(k => delete all[k]);     // keep a year
  PropertiesService.getScriptProperties().setProperty('TICKS', JSON.stringify(all));
  return true;
}

// Subscription status (keep / could cancel / cancelled), remembered by merchant.
const SUB_STATUSES = ['keep', 'could cancel', 'cancelled'];
function getSubStatus_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SUB_STATUS') || '{}'); } catch (e) { return {}; } }
function setSubStatus(merchant, status) {
  const all = getSubStatus_();
  if (!status) delete all[merchant]; else all[merchant] = { s: status, at: Date.now() };
  PropertiesService.getScriptProperties().setProperty('SUB_STATUS', JSON.stringify(all));
  return true;
}
function subsTracker_() {
  const sh = sheet_('Subscriptions'), out = { could: 0, saved: 0, couldList: [], savedList: [], still: [] };
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(r => {
    if (!r[0] || !r[2] || /total|known/i.test(r[0])) return;
    const st = String(r[8] || '').toLowerCase(), yr = Number(r[5]) || 0;
    if (st === 'could cancel') { out.could += yr; out.couldList.push(String(r[0])); }
    if (st === 'cancelled') {
      out.saved += yr; out.savedList.push(String(r[0]));
      const mark = (getSubStatus_()[r[0]] || {}).at;
      if (mark && r[3] instanceof Date && r[3].getTime() > mark + 2 * 864e5) out.still.push(String(r[0]));
    }
  });
  return out;
}
const titleCase_ = s => String(s).toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase());

// ---------------- Dashboard (light, KJ style) ----------------
function buildDashboard_(tx, missed, debt) {
  const sh = sheet_('Dashboard');
  sh.clear(); sh.setConditionalFormatRules([]); sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getRange('A:K').breakApart(); sh.getRange('A:K').clearDataValidations();
  sh.setHiddenGridlines(true); sh.setFrozenRows(0);
  const S = settings_(), tz = Session.getScriptTimeZone(), now = new Date();
  const W = 8, money = '$#,##0;[Red]-$#,##0';
  for (let c = 1; c <= W; c++) sh.setColumnWidth(c, 122);
  sh.setColumnWidth(9, 20); sh.setColumnWidth(10, 120); sh.setColumnWidth(11, 80);
  const maxR = Math.max(sh.getMaxRows(), 120);
  sh.getRange(1, 1, maxR, W + 1).setBackground(KJ.bg).setFontFamily('Arial').setFontColor(KJ.ink).setVerticalAlignment('middle').setFontSize(10);

  // ---- numbers
  const cur = currentPeriod_(now), PM = PERIOD_ || { on: false };
  const [cy, cm] = cur.split('-').map(Number);
  const months = []; for (let i = 5; i >= 0; i--) months.push(Utilities.formatDate(new Date(cy, cm - 1 - i, 1), tz, 'yyyy-MM'));
  // Month picker: any of the last 12 periods that has transactions. The choice is kept in DASH_VIEW.
  const pickable = []; for (let i = 0; i < 12; i++) { const m = Utilities.formatDate(new Date(cy, cm - 1 - i, 1), tz, 'yyyy-MM'); if (i === 0 || tx.some(t => t.month === m)) pickable.push(m); }
  const picked = PropertiesService.getScriptProperties().getProperty('DASH_VIEW') || '';
  const thisMonth = pickable.indexOf(picked) > 0 ? picked : cur, past = thisMonth !== cur;
  let full = priorPeriods_(thisMonth, 3, tx).reverse(); if (!full.length) full = [months[4]];
  const st = periodStats_(thisMonth, tx);
  const home = tx.filter(t => isSpend_(t) && !t.trip), byCat = {};
  home.forEach(t => { byCat[t.cat] = byCat[t.cat] || {}; byCat[t.cat][t.month] = (byCat[t.cat][t.month] || 0) - t.amt; });
  const spendM = m => Object.values(byCat).reduce((s, c) => s + (c[m] || 0), 0);
  const incM = m => tx.filter(t => t.month === m && t.cat === 'Income').reduce((s, t) => s + t.amt, 0);
  const tripM = m => -tx.filter(t => t.month === m && t.trip && isSpend_(t)).reduce((s, t) => s + t.amt, 0);
  const avg = f => full.reduce((s, m) => s + f(m), 0) / full.length;
  const baseAvg = avg(spendM);
  const nw = sheet_('NetWorth'), nwRows = nw.getLastRow() > 1 ? nw.getRange(2, 1, nw.getLastRow() - 1, 8).getValues().filter(r => r[0] instanceof Date) : [];
  const pb = periodBounds_(thisMonth), pStart = pb[0];
  const nwUpTo = past ? nwRows.filter(r => r[0] <= new Date(pb[1].getTime() + 864e5 - 1)) : nwRows;
  const nwRow = nwUpTo.length ? nwUpTo[nwUpTo.length - 1] : null;
  const nwStart = nwRows.filter(r => r[0] < pStart).pop() || nwRows[0];
  const cash = nwRow ? Number(nwRow[4]) - (Number(S['Cash held for others']) || 0) : 0;
  const cushion = baseAvg ? cash / baseAvg : 0, goal = Number(S['Safety net target (months)']) || 3;
  const daysLeft = Math.max(1, Math.round((pb[1] - new Date(now.getFullYear(), now.getMonth(), now.getDate())) / 864e5) + 1);
  const coming = past ? { amt: 0, n: 0 } : payStillComing_(tx, thisMonth);
  const cl = checklist_(tx, S, thisMonth, past ? 0 : st.pay + coming.amt);
  const bills = past ? { amt: 0, list: [] } : billsComing_(tx, thisMonth, S);
  const keepPct = setAsidePct_(S), payAll = cl.pay, spendable = payAll - cl.setAside - bills.amt - st.spend;   // what's yours to spend after saving and bills
  const rng = (r, c, h, w) => sh.getRange(r, c, h || 1, w || 1);
  const stripeRows = [];
  const card = (r0, r1) => {
    rng(r0, 1, r1 - r0 + 1, W).setBackground(KJ.card).setBorder(true, true, true, true, false, false, KJ.line, SpreadsheetApp.BorderStyle.SOLID);
    stripeRows.filter(x => x >= r0 && x <= r1).forEach(x => rng(x, 1, 1, W).setBackground(KJ.stripe));
  };
  const head = (r, text, caption) => {
    rng(r, 1, 1, 5).merge().setValue(text.toUpperCase()).setFontWeight('bold').setFontColor(KJ.mute).setFontSize(9);
    if (caption) rng(r, 6, 1, 3).merge().setValue(caption.toUpperCase()).setFontColor(KJ.faint).setFontSize(8).setHorizontalAlignment('right');
    sh.setRowHeight(r, 30);
  };
  let r = 1;

  // ---- title
  rng(r, 1, 1, 5).merge().setValue(isDemo_() ? 'Finance HQ · demo (sample data)' : 'Finance HQ').setFontSize(22).setFontWeight('bold');
  rng(r, 6, 1, 3).merge().setValue('Updated ' + Utilities.formatDate(now, tz, 'EEE MMM d, h:mm a')).setFontColor(KJ.mute).setHorizontalAlignment('right');
  sh.setRowHeight(r, 44); r++;
  rng(r, 1, 1, 5).merge().setValue((PM.on ? 'Statement period ' + periodRange_(thisMonth) : Utilities.formatDate(pb[0], tz, 'MMMM yyyy')) + '  ·  ' +
    (past ? 'finished period' : daysLeft === 1 ? 'last day' : daysLeft + ' days to go') + (Number(S['Savings goal (% of pay)']) ? '  ·  Goal: save ' + Number(S['Savings goal (% of pay)']) + '% of pay' : ''))
    .setFontColor(KJ.mute);
  {
    const label = m => Utilities.formatDate(new Date(+m.slice(0, 4), +m.slice(5, 7) - 1, 15), tz, 'MMM yyyy') + (m === cur ? ' (current)' : '');
    rng(r, 6).setValue('VIEWING ▸').setFontSize(8).setFontWeight('bold').setFontColor(KJ.mute).setHorizontalAlignment('right');
    rng(r, 7, 1, 2).merge().setNumberFormat('@').setValue(label(thisMonth)).setFontWeight('bold').setFontColor(past ? KJ.orange : KJ.green).setBackground(KJ.card)
      .setHorizontalAlignment('center').setBorder(true, true, true, true, false, false, past ? KJ.orange : KJ.line, SpreadsheetApp.BorderStyle.SOLID)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(pickable.map(label), true).setAllowInvalid(true).build())
      .setNote('Pick a month to see it here. Goes back to the current period with each morning refresh.');
    sh.setRowHeight(r, 28); r++;
    if (past) { rng(r, 1, 1, W).merge().setValue('Looking back at ' + label(thisMonth) + ' — pick "' + label(cur) + '" in the box above to go back to now.')
      .setFontColor(KJ.orange).setFontSize(9).setFontWeight('bold'); }
    r++;
  }

  // ---- tiles
  const tiles = [
    past ? ['LEFT OVER', st.left, money, st.left >= 0 ? KJ.green : KJ.red, 'pay $' + Math.round(st.pay).toLocaleString() + ' − spent $' + (Math.round(st.spend) || 0).toLocaleString()]
      : !(payAll > 0) ? ['SPENT SO FAR', st.spend, money, KJ.ink, (tx.some(t => t.cat === 'Income' && t.amt > 50) ? 'no pay expected this period' : 'no paycheck found yet') + ' — if a deposit is pay, set it to Income in Transactions']
      : ['LEFT TO SPEND', spendable, money, spendable >= 0 ? KJ.green : KJ.red,
        'pay ' + (coming.amt ? '~' : '') + '$' + Math.round(payAll).toLocaleString() + cl.parts.map(x => ' − ' + x.label + ' $' + Math.round(x.amt).toLocaleString()).join('') +
        (bills.amt ? ' − bills still due $' + Math.round(bills.amt).toLocaleString() : '') + ' − spent $' + (Math.round(st.spend) || 0).toLocaleString()],
    ['SAVINGS GREW BY', cl.moves.into, money, cl.moves.into >= 0 ? KJ.blue : KJ.red, 'in $' + Math.round(cl.moves.gross).toLocaleString() + ' · out $' + Math.round(cl.moves.takenOut).toLocaleString() + (cl.goalSave ? ' · goal $' + Math.round(cl.goalSave).toLocaleString() + (cl.moves.into >= cl.goalSave - 0.5 ? ' ✓' : '') : '')],
    ['NET WORTH', nwRow ? Number(nwRow[3]) : '—', money, KJ.ink, nwRow && nwStart && nwStart !== nwRow ? ((nwRow[3] - nwStart[3]) >= 0 ? '▲ $' : '▼ $') + Math.round(Math.abs(nwRow[3] - nwStart[3])).toLocaleString() + (past ? ' that period' : ' this period') : past && !nwRow ? 'no balance history that far back' : 'all accounts minus debts'],
    ['SAFETY NET', nwRow ? cushion : '—', '0.0" mo"', cushion >= goal ? KJ.green : KJ.orange, 'goal ' + goal + ' months ($' + Math.round(goal * baseAvg).toLocaleString() + ')']
  ];
  tiles.forEach((t, i) => {
    const c = 1 + i * 2;
    rng(r, c, 1, 2).merge().setValue(t[0]).setFontSize(8).setFontColor(KJ.mute).setFontWeight('bold');
    rng(r + 1, c, 1, 2).merge().setNumberFormat(t[2]).setValue(t[1]).setFontSize(20).setFontWeight('bold').setFontColor(t[3]);
    rng(r + 2, c, 1, 2).merge().setValue(t[4]).setFontSize(9).setFontColor(KJ.mute);
    rng(r, c, 3, 2).setBackground(KJ.card).setHorizontalAlignment('left').setBorder(true, true, true, true, false, false, KJ.line, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.setRowHeight(r + 1, 40); r += 4;

  // ---- bills by card
  const cards = past ? [] : cardCycles_(tx, S, now);              // bills only make sense for now
  if (cards.length) {
    const r0 = r; head(r, 'Bills by card', PM.on ? 'by statement cycle' : ''); r++;
    rng(r, 1, 1, W).setValues([['Card', '', 'This cycle', 'Spent so far', 'Statement', 'Due', 'Status', '']]).setFontColor(KJ.faint).setFontSize(8).setFontWeight('bold');
    rng(r, 4, 1, 2).setHorizontalAlignment('right'); r++;
    let due = 0;
    cards.forEach((c, i) => {
      if (!c.paid && c.configured) due += c.statement;
      rng(r, 1, 1, 2).merge().setValue(c.name).setFontWeight('bold');
      rng(r, 3, 1, 6).setValues([[c.cycle, c.spent, c.statement, c.due, c.status, '']]);
      rng(r, 4, 1, 2).setNumberFormat(money);
      rng(r, 7, 1, 2).merge().setFontWeight('bold').setFontColor(!c.configured ? KJ.mute : c.paid ? KJ.green : c.daysToDue <= 3 ? KJ.red : c.daysToDue <= 7 ? KJ.orange : KJ.ink);
      if (i % 2) stripeRows.push(r);
      r++;
    });
    const chk = checkingBalance_(), ok = chk >= due;
    rng(r, 1, 1, W).merge().setValue(due ? 'Bills to pay $' + Math.round(due).toLocaleString() + '  ·  checking $' + Math.round(chk).toLocaleString() +
      (ok ? '  ·  covered ✓' : '  ·  move $' + Math.round(due - chk).toLocaleString() + ' from savings before the due date') :
      (cards.some(c => c.configured) ? 'All card statements are paid ✓' : 'Add each card\'s closing and due day (Finance HQ → Quick setup) to track bills'))
      .setFontColor(due ? (ok ? KJ.green : KJ.red) : KJ.green).setFontWeight('bold');
    card(r0, r); r += 2;
  }

  // ---- transfers to make (checkboxes)
  {
    const r0 = r; head(r, 'Transfers to make', 'tick the box when it\'s done'); r++;
    if (!cl.items.length) {
      rng(r, 1, 1, W).merge().setValue('Set a savings goal in Settings ("Savings goal (% of pay)") and your transfers show up here.').setFontColor(KJ.mute); r++;
    } else {
      if (!cl.pay) { rng(r, 1, 1, W).merge().setValue('No pay yet this period — amounts fill in when your paycheck lands.').setFontColor(KJ.mute); r++; }
      cl.items.forEach(i => {
        const done = i.ticked;
        rng(r, 1).insertCheckboxes().setValue(done).setHorizontalAlignment('center');
        rng(r, 2, 1, 3).merge().setValue(i.label + ' — $' + Math.round(i.amt).toLocaleString()).setFontWeight('bold')
          .setFontColor(done ? KJ.mute : KJ.ink).setFontLine(done ? 'line-through' : 'none');
        rng(r, 5, 1, 2).merge().setValue(i.note).setFontColor(KJ.mute).setFontSize(9);
        rng(r, 7, 1, 2).merge().setValue(done ? (i.auto ? 'done ✓ (seen in your accounts)' : 'done ✓') : i.k === 'tax' ? 'tick when set aside' :
          i.progress > 0 ? '$' + Math.round(i.left).toLocaleString() + ' to go' : 'not yet').setFontColor(done ? KJ.green : KJ.orange).setFontSize(9).setHorizontalAlignment('right');
        rng(r, 9).setValue(thisMonth + '|' + i.k).setFontColor(KJ.bg);                 // hidden key for the tick
        sh.setRowHeight(r, 30); r++;
      });
      rng(r, 1, 1, W).merge().setValue('Savings progress counts money into your savings and investment accounts minus money taken back out (not market gains). Tick a box yourself any time.').setFontColor(KJ.faint).setFontSize(8).setWrap(true); r++;
    }
    card(r0, r - 1); r++;
  }

  // ---- savings tracker
  {
    const tr = subsTracker_(), r0 = r;
    card(r0, r0 + 4 + (tr.still.length ? 1 : 0));                     // white card first, tinted tiles on top
    head(r, 'Savings tracker', 'mark subscriptions in the Subscriptions tab'); r++;
    const half = (c, label, val, sub, fg, bg) => {
      rng(r, c, 1, 4).merge().setValue(label).setFontSize(8).setFontWeight('bold').setFontColor(fg);
      rng(r + 1, c, 1, 4).merge().setNumberFormat('$#,##0" / yr"').setValue(val).setFontSize(20).setFontWeight('bold').setFontColor(fg).setHorizontalAlignment('left');
      rng(r + 2, c, 1, 4).merge().setValue(sub).setFontSize(9).setFontColor(fg).setWrap(true);
      rng(r, c, 3, 4).setBackground(bg).setBorder(true, true, true, true, false, false, KJ.card, SpreadsheetApp.BorderStyle.SOLID_THICK);
    };
    half(1, 'COULD CANCEL (POTENTIAL SAVINGS)', tr.could, tr.couldList.length ? tr.couldList.map(subName_).join(', ') : 'Mark subscriptions "could cancel" to see this', KJ.yellow, KJ.yellowBg);
    half(5, 'CANCELLED (ALREADY SAVED)', tr.saved, tr.savedList.length ? tr.savedList.map(subName_).join(', ') : 'Mark ones you\'ve cancelled — it adds up here', KJ.green, KJ.greenBg);
    sh.setRowHeight(r + 1, 40); r += 3;
    rng(r, 1, 1, W).merge().setValue(tr.could ? 'Cancel everything marked "could cancel" and you\'d save $' + Math.round(tr.could).toLocaleString() + ' a year ($' + Math.round(tr.could / 12).toLocaleString() + ' a month).' :
      'Tip: in the Subscriptions tab, set Status to "could cancel" or "cancelled".').setFontColor(KJ.mute).setFontSize(9);
    if (tr.still.length) { r++; rng(r, 1, 1, W).merge().setValue('⚠ Still charging after you marked it cancelled: ' + tr.still.map(titleCase_).join(', ')).setFontColor(KJ.red).setFontWeight('bold'); }
    r += 2;
  }

  // ---- spending by category (ranked colour bars — no chart object, so it always draws)
  {
    const cats = Object.keys(st.byCat).map(c => [c, st.byCat[c], avg(m => (byCat[c] || {})[m] || 0)]).filter(x => x[1] > 0.5).sort((a, b) => b[1] - a[1]);
    const top = cats.slice(0, 9);
    if (cats.length > 9) top.push(['Everything else', cats.slice(9).reduce((s, x) => s + x[1], 0), cats.slice(9).reduce((s, x) => s + x[2], 0)]);
    const r0 = r; head(r, 'Spending by category', past ? 'that period' : PM.on ? 'this period' : 'this month'); r++;
    if (top.length) {
      const big = Math.max(...top.map(x => x[1])), tot = st.spend || 1;
      rng(r, 1, 1, W).setValues([['Category', '', 'Share of spending', '', '', past ? 'That period' : 'This period', 'Usual', '% of total']])
        .setFontColor(KJ.faint).setFontSize(8).setFontWeight('bold');
      rng(r, 6, 1, 3).setHorizontalAlignment('right'); r++;
      top.forEach((x, i) => {
        const col = x[0] === 'Everything else' ? '#A8A29E' : KJ_PIE[i % KJ_PIE.length], hot = x[2] > 0 && x[1] > x[2] * 1.25;
        const lbl = rng(r, 1, 1, 2).merge().setFontWeight('bold').setValue('●  ' + x[0]);
        try { lbl.setRichTextValue(SpreadsheetApp.newRichTextValue().setText('●  ' + x[0])
          .setTextStyle(0, 1, SpreadsheetApp.newTextStyle().setForegroundColor(col).setFontSize(13).build()).build()); }
        catch (e) {}
        rng(r, 3, 1, 3).merge().setFormula('=SPARKLINE(' + Math.round(x[1]) + ',{"charttype","bar";"max",' + Math.round(big) + ';"color1","' + col + '"})');
        rng(r, 6, 1, 3).setValues([[x[1], x[2], x[1] / tot]]).setHorizontalAlignment('right');
        rng(r, 6, 1, 2).setNumberFormat('$#,##0'); rng(r, 8).setNumberFormat('0%').setFontColor(KJ.mute);
        rng(r, 7).setFontColor(hot ? KJ.orange : KJ.mute);
        if (hot) rng(r, 7).setNote('Higher than usual (more than 25% above your recent average).');
        if (i % 2) stripeRows.push(r);
        sh.setRowHeight(r, 26); r++;
      });
      rng(r, 1, 1, 5).merge().setValue('Total').setFontWeight('bold');
      rng(r, 6, 1, 3).setValues([[st.spend, baseAvg, '']]).setFontWeight('bold').setHorizontalAlignment('right');
      rng(r, 6, 1, 2).setNumberFormat('$#,##0'); r++;
      rng(r, 1, 1, W).merge().setValue('Orange "Usual" = spending more than normal in that category.').setFontColor(KJ.faint).setFontSize(8); r++;
    } else { rng(r, 1, 1, W).merge().setValue('No spending yet this period.').setFontColor(KJ.mute); r++; }
    card(r0, r - 1); r++;
  }

  // ---- this period vs limits
  {
    const lim = Object.keys(Object.assign({}, st.byCat, ...Object.keys(S).filter(k => / monthly limit$/.test(k)).map(k => ({ [k.replace(/ monthly limit$/, '')]: 1 }))))
      .filter(c => NO_LIMIT_CATS.indexOf(c) === -1 && FIXED_CATS.indexOf(c) === -1)
      .map(c => [c, st.byCat[c] || 0, Number(S[c + ' monthly limit']) || (c === 'Dining' ? Number(S['Dining target per month']) || 0 : 0)]).filter(x => x[2] > 0)
      .sort((a, b) => b[1] / b[2] - a[1] / a[2]).slice(0, 12);
    if (lim.length) {
      const r0 = r; head(r, past ? 'That period vs your limits' : 'This period vs your limits', 'green = plenty left · amber = getting there · red = over'); r++;
      lim.forEach((x, i) => {
        const p = x[1] / x[2], col = p >= 1 ? KJ.red : p >= 0.3 ? KJ.yellow : KJ.green;
        rng(r, 1, 1, 2).merge().setValue(x[0]).setFontWeight('bold');
        rng(r, 3, 1, 2).setValues([[x[1], x[2]]]).setNumberFormat('$#,##0'); rng(r, 4).setFontColor(KJ.mute).setNumberFormat('"of "$#,##0');
        rng(r, 3).setFontColor(col).setFontWeight('bold');
        rng(r, 5, 1, 3).merge().setFormula('=SPARKLINE(' + Math.round(x[1]) + ',{"charttype","bar";"max",' + Math.round(x[2]) + ';"color1","' + col + '";"empty","zero"})');
        rng(r, 8).setValue(p >= 1 ? 'over by $' + Math.round(x[1] - x[2]).toLocaleString() : '$' + Math.round(x[2] - x[1]).toLocaleString() + ' left')
          .setFontColor(col).setFontSize(9).setHorizontalAlignment('right');
        if (i % 2) stripeRows.push(r);
        r++;
      });
      card(r0, r - 1); r++;
    }
  }

  // ---- period by period
  {
    const firstTx = tx.length ? tx.reduce((a, t) => t.date < a ? t.date : a, tx[0].date) : now;
    const firstDay = new Date(firstTx.getFullYear(), firstTx.getMonth(), firstTx.getDate());
    const shown = months.filter(m => tx.some(t => t.month === m) && (m === cur || periodBounds_(m)[0] >= firstDay));
    const r0 = r; head(r, PM.on ? 'Period by period' : 'Month by month', PM.on ? 'statement periods' : ''); r++;
    rng(r, 1, 1, W).setValues([[PM.on ? 'Period' : 'Month', 'Pay', 'Spent', 'Trips', 'Left over', 'Kept %', 'Net', '']]).setFontColor(KJ.faint).setFontSize(8).setFontWeight('bold');
    rng(r, 2, 1, 5).setHorizontalAlignment('right'); rng(r, 7, 1, 2).merge().setHorizontalAlignment('left'); r++;
    const maxAbs = Math.max(1, ...shown.map(m => Math.abs(incM(m) - spendM(m) - tripM(m))));
    shown.forEach((m, i) => {
      const lo = incM(m) - spendM(m) - tripM(m);
      rng(r, 1).setNumberFormat('@');
      rng(r, 1, 1, 6).setValues([[Utilities.formatDate(new Date(+m.slice(0, 4), +m.slice(5, 7) - 1, 15), tz, 'MMM yyyy') + (m === cur ? ' (so far)' : '') + (past && m === thisMonth ? '  ◂' : ''),
        incM(m), spendM(m), tripM(m), lo, incM(m) ? lo / incM(m) : '']]);
      rng(r, 2, 1, 5).setHorizontalAlignment('right'); rng(r, 2, 1, 4).setNumberFormat(money); rng(r, 6).setNumberFormat('0%');
      rng(r, 5).setFontColor(lo >= 0 ? KJ.green : KJ.red).setFontWeight('bold');
      rng(r, 7, 1, 2).merge().setFormula('=SPARKLINE(' + Math.round(Math.abs(lo)) + ',{"charttype","bar";"max",' + Math.round(maxAbs) + ';"color1","' + (lo >= 0 ? KJ.green : KJ.red) + '"})');
      if (i % 2) stripeRows.push(r);
      r++;
    });
    card(r0, r - 1); r++;
  }

  // ---- quick facts
  {
    const r0 = r; head(r, 'Quick facts'); r++;
    const tripTot = {}; tx.filter(t => t.trip && isSpend_(t)).forEach(t => tripTot[t.trip] = (tripTot[t.trip] || 0) - t.amt);
    [['Debt', debt],
     ['Points missed (90 days)', missed ? missed.toLocaleString() + ' points from using the wrong card — see CardCheck' : 'None — every purchase was on the right card ✓'],
     ['Trips', Object.keys(tripTot).length ? Object.keys(tripTot).map(k => k + ' $' + Math.round(tripTot[k]).toLocaleString()).join('  ·  ') : 'none yet'],
     ['Cash held for others', (Number(S['Cash held for others']) || 0) ? '$' + Number(S['Cash held for others']).toLocaleString() + ' left out of your safety net' : 'none']
    ].forEach(f => {
      rng(r, 1, 1, 2).merge().setValue(f[0]).setFontColor(KJ.mute).setFontWeight('bold');
      rng(r, 3, 1, 6).merge().setValue(f[1]).setWrap(true); r++;
    });
    card(r0, r - 1); r++;
  }
  rng(r, 1, 1, W).merge().setValue('Change goals and limits in Settings  ·  Finance HQ menu → Quick setup or Refresh now').setFontColor(KJ.faint).setFontSize(9).setHorizontalAlignment('center');
  sh.hideColumns(9, 3);
}


// Pop-up note at the bottom of the sheet. Silently skipped when run from the script editor or a schedule.
function toast_(ss, msg, title, secs) { try { ss.toast(msg, title, secs); } catch (e) {} }

// ===== PART 10 — Yearly costs · Budget tab · Insights tab · Year in review PDF · phone edits =====

const MON_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// ---------------- yearly costs (sinking funds) ----------------
// Yearly costs tab: Name | Yearly amount | Month it's due | Notes. 1/12 of the total is set aside every period.
function monthIndex_(v) {
  if (v instanceof Date) return v.getMonth();
  const s = String(v == null ? '' : v).trim();
  const i = MON_.map(m => m.toLowerCase()).indexOf(s.slice(0, 3).toLowerCase());
  if (i >= 0) return i;
  const n = Number(s); return n >= 1 && n <= 12 ? n - 1 : -1;
}
function loadYearly_() {
  const sh = finSS_().getSheetByName('Yearly costs');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => String(r[0]).trim() !== '' && Number(r[1]) > 0)
    .map(r => ({ name: String(r[0]).trim(), amt: Number(r[1]), month: monthIndex_(r[2]), notes: String(r[3] || '') }));
}
function yearlyPerPeriod_() {
  const list = loadYearly_(), total = list.reduce((s, y) => s + y.amt, 0);
  return { per: Math.round(total / 12 * 100) / 100, total: total, names: list.sort((a, b) => b.amt - a.amt).map(y => y.name), list: list };
}
// next due date of each yearly cost (the 1st of its month), soonest first
function yearlyUpcoming_(from) {
  from = from || new Date();
  const today = new Date(from.getFullYear(), from.getMonth(), 1);
  return loadYearly_().filter(y => y.month >= 0).map(y => {
    let d = new Date(today.getFullYear(), y.month, 1); if (d < today) d = new Date(today.getFullYear() + 1, y.month, 1);
    return Object.assign({ next: d, monthsAway: (d.getFullYear() - today.getFullYear()) * 12 + d.getMonth() - today.getMonth() }, y);
  }).sort((a, b) => a.next - b.next);
}
function ensureYearlyTab_() {
  const ss = finSS_();
  let sh = ss.getSheetByName('Yearly costs');
  if (!sh) {
    sh = ss.insertSheet('Yearly costs');
    sh.getRange(1, 1, 1, 4).setValues([['Yearly cost', 'Yearly amount', 'Month it\'s due', 'Notes']]);
  }
  return sh;
}
function styleYearly_() {
  const sh = ensureYearlyTab_();
  styleTable_(sh, [240, 140, 140, 360], {
    headers: ['Yearly cost', 'Yearly amount', 'Month it\'s due', 'Notes'],
    notes: ['Things you pay once a year: car insurance (if paid yearly), annual card fees, Amazon Prime, registration, renewals…',
      'The full yearly amount — Finance HQ sets aside 1/12 of it every period', 'Pick the month it usually charges'],
    money: [2], edit: [1, 2, 3, 4] });
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('@');            // keep "Feb" as text, never a date
  dropdown_(sh.getRange(2, 3, sh.getMaxRows() - 1, 1), MON_);
}

// ---------------- limits: read, write, and suggestions that keep your savings goal ----------------
function explicitLimit_(cat, S) {
  const v = Number(S[cat + ' monthly limit']);
  if (v > 0) return v;
  if (cat === 'Dining' && Number(S['Dining target per month']) > 0) return Number(S['Dining target per month']);
  return 0;
}
// writes to the setting the Dashboard actually reads (Dining may live in "Dining target per month")
function setLimit_(cat, value) {
  const S = settings_(), v = value === '' || value === null ? '' : Math.max(0, Math.round(Number(value) || 0));
  if (cat === 'Dining' && S['Dining monthly limit'] === undefined) setSetting_('Dining target per month', v, 'Your dining budget per period');
  else setSetting_(cat + ' monthly limit', v, 'your limit per period');
}
// average per period of everyday (non-trip) spending in each category
function catAverages_(keys, tx) {
  const avg = {};
  if (!keys.length) return avg;
  const set = new Set(keys);
  tx.filter(t => set.has(t.month) && isSpend_(t) && !t.trip).forEach(t => avg[t.cat] = (avg[t.cat] || 0) - t.amt / keys.length);
  Object.keys(avg).forEach(c => { if (avg[c] < 0.5) delete avg[c]; });
  return avg;
}
function budgetPlan_(tx, S) {
  const cur = currentPeriod_(), st = periodStats_(cur, tx);
  const coming = payStillComing_(tx, cur);
  const cl = checklist_(tx, S, cur, st.pay + coming.amt);
  const keys = priorPeriods_(cur, 3, tx);
  const avg = catAverages_(keys, tx);
  let pay = cl.pay;
  if (!(pay > 0) && keys.length) pay = keys.reduce((s, k) => s + periodStats_(k, tx).pay, 0) / keys.length;
  const keep = cl.setAside;
  const fixedCats = Object.keys(avg).filter(c => FIXED_CATS.indexOf(c) !== -1 || NO_LIMIT_CATS.indexOf(c) !== -1);
  const fixed = fixedCats.reduce((s, c) => s + avg[c], 0);
  const flex = [...new Set(Object.keys(avg).filter(c => avg[c] >= 15)
    .concat(Object.keys(S).filter(k => / monthly limit$/.test(k) && Number(S[k]) > 0).map(k => k.replace(/ monthly limit$/, '')))
    .concat(Number(S['Dining target per month']) > 0 ? ['Dining'] : []))]
    .filter(c => FIXED_CATS.indexOf(c) === -1 && NO_LIMIT_CATS.indexOf(c) === -1);
  const room = pay - keep - fixed;                                   // what's left for everyday categories
  const ess = flex.filter(c => ESSENTIAL_CATS.indexOf(c) !== -1), disc = flex.filter(c => ESSENTIAL_CATS.indexOf(c) === -1);
  const r10 = x => Math.max(10, Math.round(x / 10) * 10), up10 = x => Math.max(10, Math.ceil(x / 10) * 10);
  const sugg = {};
  ess.forEach(c => sugg[c] = avg[c] ? up10(avg[c]) : 0);
  const essTot = ess.reduce((s, c) => s + sugg[c], 0), discTot = disc.reduce((s, c) => s + (avg[c] || 0), 0);
  let factor = 1;
  if (pay > 0 && discTot > 0 && essTot + discTot > room) factor = Math.max(0.5, (room - essTot) / discTot);
  disc.forEach(c => sugg[c] = avg[c] ? (factor < 1 ? r10(avg[c] * factor) : up10(avg[c])) : 0);
  const suggTot = flex.reduce((s, c) => s + (sugg[c] || 0), 0);
  const limits = {}; flex.forEach(c => limits[c] = explicitLimit_(c, S));
  const limTot = flex.reduce((s, c) => s + (limits[c] || sugg[c] || 0), 0);
  return { cur: cur, st: st, cl: cl, coming: coming, keys: keys, avg: avg, pay: pay, keep: keep, fixed: fixed, fixedCats: fixedCats,
    room: room, flex: flex, sugg: sugg, suggTot: suggTot, limits: limits, limTot: limTot, factor: factor,
    fits: suggTot <= room + 5, cutPct: Math.round((1 - factor) * 100) };
}

// ---------------- Budget tab ----------------
function buildBudget_(tx) {
  const ss = finSS_(), S = settings_(), tz = Session.getScriptTimeZone(), now = new Date();
  const sh = ss.getSheetByName('Budget') || ss.insertSheet('Budget');
  sh.clear(); sh.getRange('A:I').breakApart(); sh.setConditionalFormatRules([]); sh.getRange('A:I').clearDataValidations();
  sh.getBandings().forEach(b => b.remove());
  uiPage_(sh); sh.setHiddenGridlines(true); sh.setFrozenRows(0);
  [190, 120, 120, 90, 118, 118, 118, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A:H').setFontFamily('Arial').setFontColor(UI.ink).setVerticalAlignment('middle').setFontSize(10);
  const B = budgetPlan_(tx, S), money = '$#,##0';
  const pb = periodBounds_(B.cur), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const totalDays = Math.round((pb[1] - pb[0]) / 864e5) + 1, dayN = Math.min(totalDays, Math.max(1, Math.round((today - pb[0]) / 864e5) + 1));
  const pace = dayN / totalDays, P = PERIOD_ || { on: false };
  const lbl = k => Utilities.formatDate(new Date(+k.slice(0, 4), +k.slice(5, 7) - 1, 15), tz, 'MMM');
  const prev = B.keys.slice(0, 2).reverse();                         // two most recent full periods, oldest first

  sh.getRange('A1:F1').merge().setValue('Budget').setFontSize(18).setFontWeight('bold');
  sh.getRange('G1:H1').merge().setValue('Updated ' + Utilities.formatDate(now, tz, 'MMM d, h:mm a')).setFontColor(UI.mute).setHorizontalAlignment('right');
  sh.getRange('A2:H2').merge().setValue((P.on ? 'Statement period ' + periodRange_(B.cur) : Utilities.formatDate(now, tz, 'MMMM yyyy')) + '  ·  day ' + dayN + ' of ' + totalDays)
    .setFontColor(UI.mute);

  const free = B.pay - B.keep - B.fixed - B.limTot;
  let r = uiTiles_(sh, 4, [
    ['EXPECTED PAY', B.pay, money, UI.ink, B.coming.n ? '$' + Math.round(B.cl.got).toLocaleString() + ' so far + ' + B.coming.n + ' paycheck' + (B.coming.n > 1 ? 's' : '') + ' coming' : 'this period'],
    ['SET ASIDE FIRST', B.keep, money, UI.accent, B.cl.parts.length ? B.cl.parts.map(x => x.label).join(' + ') : 'set a savings goal in Settings'],
    ['FIXED BILLS & OTHER', B.fixed, money, UI.ink, B.fixedCats.length ? 'usual ' + B.fixedCats.slice(0, 3).join(', ').toLowerCase() : 'none'],
    [free >= 0 ? 'UNPLANNED' : 'OVER-PLANNED', free, money + ';-$#,##0', free >= 0 ? UI.good : UI.bad, free >= 0 ? 'left after your limits' : 'limits add up to more than you have']
  ]);

  // ---- limits table
  r = uiBand_(sh, r, 'Your limits — type a new number in yellow, or tick "Use it?" to take the suggestion');
  r = uiHeader_(sh, r, ['Category', 'Your limit', 'Suggested', 'Use it?', prev[0] ? lbl(prev[0]) : '', prev[1] ? lbl(prev[1]) : '', 'This period', 'Status']);
  sh.getRange(r - 1, 2, 1, 6).setHorizontalAlignment('right'); sh.getRange(r - 1, 4).setHorizontalAlignment('center');
  const t0 = r;
  const rows = B.flex.slice().sort((a, b) => (B.st.byCat[b] || 0) + (B.avg[b] || 0) - (B.st.byCat[a] || 0) - (B.avg[a] || 0));
  const spentIn = (k, c) => tx.filter(t => t.month === k && t.cat === c && isSpend_(t) && !t.trip).reduce((s, t) => s - t.amt, 0);
  let sumLim = 0, sumSug = 0, sumP0 = 0, sumP1 = 0, sumNow = 0;
  rows.forEach((c, i) => {
    const L = B.limits[c] || 0, sg = B.sugg[c] || 0, now0 = spentIn(B.cur, c);
    const p0 = prev[0] ? spentIn(prev[0], c) : '', p1 = prev[1] ? spentIn(prev[1], c) : '';
    const st = !L ? ['—', UI.mute] : now0 > L + 0.5 ? ['Over', UI.bad] : now0 > L * pace * 1.1 + 5 && ESSENTIAL_CATS.indexOf(c) === -1 ? ['Watch', UI.warn] : ['On track', UI.good];
    sh.getRange(r, 1, 1, 8).setValues([[c, L || '', sg || '', false, p0, p1, now0, st[0]]]);
    sh.getRange(r, 4).insertCheckboxes().setValue(false).setHorizontalAlignment('center');
    sh.getRange(r, 9).setValue('limit|' + c);
    if (!sg || sg === L) sh.getRange(r, 4).removeCheckboxes().setValue('');
    sh.getRange(r, 8).setFontColor(st[1]).setFontWeight('bold').setHorizontalAlignment('right');
    [[5, p0], [6, p1]].forEach(x => { if (L && x[1] !== '') sh.getRange(r, x[0]).setFontColor(x[1] > L ? UI.bad : UI.ink); });
    if (sg && L && sg < L) sh.getRange(r, 3).setFontColor(UI.accent).setFontWeight('bold');
    sumLim += L || 0; sumSug += sg; sumP0 += Number(p0) || 0; sumP1 += Number(p1) || 0; sumNow += now0;
    r++;
  });
  if (rows.length) {
    sh.getRange(t0, 2, rows.length, 2).setNumberFormat(money);
    sh.getRange(t0, 5, rows.length, 3).setNumberFormat(money);
    sh.getRange(t0, 2, rows.length, 2).setHorizontalAlignment('right'); sh.getRange(t0, 5, rows.length, 3).setHorizontalAlignment('right');
    sh.getRange(t0, 1, rows.length, 1).setFontWeight('bold');
    uiStripes_(sh, t0, rows.length, 8);
    sh.getRange(t0, 2, rows.length, 1).setBackground(UI.edit).setFontWeight('bold');
    sh.getRange(r, 1, 1, 8).setValues([['Total', sumLim, sumSug, '', prev[0] ? sumP0 : '', prev[1] ? sumP1 : '', sumNow, '']]).setFontWeight('bold');
    sh.getRange(r, 2, 1, 6).setNumberFormat(money).setHorizontalAlignment('right');
    sh.getRange(r, 1, 1, 8).setBorder(true, false, false, false, false, false, UI.line, SpreadsheetApp.BorderStyle.SOLID);
    r++;
    sh.getRange(r, 1, 1, 8).merge().setWrap(true).setFontSize(9).setFontColor(B.fits ? UI.mute : UI.warn).setValue(
      B.factor < 1 ? (B.fits ? 'Suggestions trim your non-essential categories by about ' + B.cutPct + '% so everything fits inside what\'s left after saving. Essentials (groceries, gas, gym…) stay at your average.'
        : 'Even trimming non-essentials by half doesn\'t fit inside what\'s left after saving — look at fixed bills or the savings goal, or accept a lower target for now.')
        : 'You\'re spending within what\'s left after saving — suggestions are simply your recent averages (trips not counted).');
    sh.setRowHeight(r, 34); r += 2;
  } else {
    sh.getRange(r, 1, 1, 8).merge().setValue('Limits appear here once there\'s a full period of spending to learn from.').setFontColor(UI.mute); r += 2;
  }

  // ---- fixed bills
  if (B.fixedCats.length) {
    r = uiBand_(sh, r, 'Fixed bills and other spending — no limit needed');
    r = uiHeader_(sh, r, ['Category', 'Usual', '', '', prev[0] ? lbl(prev[0]) : '', prev[1] ? lbl(prev[1]) : '', 'This period', '']);
    sh.getRange(r - 1, 2).setHorizontalAlignment('right'); sh.getRange(r - 1, 5, 1, 3).setHorizontalAlignment('right');
    const f0 = r;
    B.fixedCats.sort((a, b) => B.avg[b] - B.avg[a]).forEach(c => {
      sh.getRange(r, 1, 1, 7).setValues([[c, B.avg[c], '', '', prev[0] ? spentIn(prev[0], c) : '', prev[1] ? spentIn(prev[1], c) : '', spentIn(B.cur, c)]]);
      r++;
    });
    sh.getRange(f0, 2, r - f0, 6).setNumberFormat(money).setHorizontalAlignment('right');
    sh.getRange(f0, 1, r - f0, 1).setFontWeight('bold');
    uiStripes_(sh, f0, r - f0, 8); r++;
  }

  // ---- yearly costs
  const yr = yearlyPerPeriod_(), up = yearlyUpcoming_(now);
  r = uiBand_(sh, r, yr.total ? 'Yearly costs — $' + Math.round(yr.per).toLocaleString() + ' set aside each period' : 'Yearly costs');
  if (!yr.total) {
    sh.getRange(r, 1, 1, 8).merge().setWrap(true).setFontColor(UI.mute).setValue('None yet. Add once-a-year bills (annual card fee, car insurance paid yearly, Prime, registration…) in the Yearly costs tab — Finance HQ then sets aside 1/12 each period so they never surprise you.');
    sh.setRowHeight(r, 40); r += 2;
  } else {
    r = uiHeader_(sh, r, ['Yearly cost', 'Per year', 'Per period', 'Due', 'Next', '', '', '']);
    sh.getRange(r - 1, 2, 1, 2).setHorizontalAlignment('right');
    const y0 = r;
    (up.length ? up : yr.list).forEach(y => {
      sh.getRange(r, 4, 1, 2).setNumberFormat('@');
      sh.getRange(r, 1, 1, 5).setValues([[y.name, y.amt, y.amt / 12, y.month >= 0 ? MON_[y.month] : '?',
        y.next ? (y.monthsAway === 0 ? 'this month' : y.monthsAway === 1 ? 'next month' : 'in ' + y.monthsAway + ' months') : 'pick a month']]);
      if (y.next && y.monthsAway <= 1) sh.getRange(r, 5).setFontColor(UI.warn).setFontWeight('bold');
      r++;
    });
    yr.list.filter(y => y.month < 0).forEach(y => { if (up.length) { sh.getRange(r, 1, 1, 5).setValues([[y.name, y.amt, y.amt / 12, '?', 'pick a month']]); r++; } });
    sh.getRange(y0, 2, r - y0, 2).setNumberFormat(money).setHorizontalAlignment('right');
    sh.getRange(y0, 1, r - y0, 1).setFontWeight('bold');
    uiStripes_(sh, y0, r - y0, 8); r++;
  }
  sh.getRange(r, 1, 1, 8).merge().setWrap(true).setFontSize(9).setFontColor(UI.mute)
    .setValue('How it works: expected pay − what you set aside first (savings goal + yearly costs) − fixed bills = what\'s left for everyday categories. Trips are kept out. Limits are per statement period and live in Settings — changing them here changes them there.');
  sh.setRowHeight(r, 34);
  sh.hideColumns(9);
  return B;
}

// ---------------- Insights tab ----------------
// Each insight: [level, headline, detail]  level: good · watch · bad · info
function insightsNow_(tx, S) {
  const out = [], now = new Date(), cur = currentPeriod_(now), st = periodStats_(cur, tx);
  const coming = payStillComing_(tx, cur), cl = checklist_(tx, S, cur, st.pay + coming.amt);
  const pb = periodBounds_(cur), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const total = Math.round((pb[1] - pb[0]) / 864e5) + 1, day = Math.min(total, Math.max(1, Math.round((today - pb[0]) / 864e5) + 1));
  const bills = billsComing_(tx, cur, S), left = cl.pay - cl.setAside - bills.amt - st.spend, daysLeft = total - day + 1;
  if (day <= 4) out.push(['info', 'Early days', '$' + (Math.round(st.spend) || 0).toLocaleString() + ' spent so far. You have $' + Math.round(left).toLocaleString() + ' to spend over ' + daysLeft + ' days' + (bills.amt ? ' after $' + Math.round(bills.amt).toLocaleString() + ' of bills still due (' + bills.list.map(x => x.cat.toLowerCase()).join(', ') + ')' : '') + ' — about $' + Math.max(0, Math.round(left / daysLeft)).toLocaleString() + '/day.']);
  else {
    const proj = st.spend / day * total + bills.amt, budget = cl.pay - cl.setAside;
    out.push([proj <= budget ? 'good' : proj <= budget * 1.1 ? 'watch' : 'bad', proj <= budget ? 'On pace' : 'Spending faster than planned',
      'At this pace you\'ll spend about $' + Math.round(proj).toLocaleString() + ' this period vs $' + Math.round(budget).toLocaleString() + ' available after saving. ' +
      (left >= 0 ? '$' + Math.round(left).toLocaleString() + ' left — about $' + Math.round(left / daysLeft).toLocaleString() + '/day.' : 'You\'re already $' + Math.round(-left).toLocaleString() + ' over.')]);
  }
  const lim = Object.keys(st.byCat).map(c => [c, st.byCat[c], explicitLimit_(c, S)]).filter(x => x[2] > 0 && NO_LIMIT_CATS.indexOf(x[0]) === -1 && FIXED_CATS.indexOf(x[0]) === -1);
  const over = lim.filter(x => x[1] >= x[2]).sort((a, b) => (b[1] - b[2]) - (a[1] - a[2]));
  const fast = lim.filter(x => x[1] < x[2] && x[1] > x[2] * (day / total) * 1.25 + 10 && day > 4);
  if (over.length) out.push(['bad', 'Over the limit', over.slice(0, 4).map(x => x[0] + ' by $' + Math.round(x[1] - x[2]).toLocaleString()).join(', ') + '.']);
  if (fast.length) out.push(['watch', 'Ahead of pace', fast.slice(0, 4).map(x => x[0] + ' ($' + Math.round(x[1]) + ' of $' + Math.round(x[2]) + ')').join(', ') + ' — already past where you\'d expect at day ' + day + '.']);
  if (!over.length && !fast.length && lim.length && day > 4) out.push(['good', 'Limits look fine', 'Every category with a limit is on track so far.']);
  cl.items.filter(i => !i.ticked && i.amt > 0.5).forEach(i => out.push([cl.got > 0 ? 'watch' : 'info', i.label + ' — $' + Math.round(i.amt).toLocaleString(),
    cl.got > 0 ? (i.progress > 0 ? '$' + Math.round(i.left).toLocaleString() + ' to go.' : 'Not seen in your accounts yet. Tick it on the Dashboard once it\'s done.') : 'Do this when your next paycheck lands.']));
  cl.items.filter(i => i.ticked && i.amt > 0.5).forEach(i => out.push(['good', i.label + ' ✓', '$' + Math.round(i.amt).toLocaleString() + ' done for this period.']));
  try {
    cardCycles_(tx, S, now).filter(c => c.configured && !c.paid && c.statement > 0 && c.daysToDue <= 10).forEach(c =>
      out.push([c.daysToDue <= 3 ? 'bad' : 'watch', c.name + ' due ' + c.due, '$' + Math.round(c.statement).toLocaleString() + ' — ' + c.status + '.']));
  } catch (e) {}
  yearlyUpcoming_(now).filter(y => y.monthsAway <= 1).forEach(y => out.push(['info', y.name + ' comes up ' + (y.monthsAway ? 'next month' : 'this month'), '$' + Math.round(y.amt).toLocaleString() + ' — paid from your yearly-costs savings.']));
  const sub = sheet_('Subscriptions');
  if (sub.getLastRow() > 1) {
    const ups = sub.getRange(2, 1, sub.getLastRow() - 1, 7).getValues().filter(x => /UP/.test(String(x[6])));
    if (ups.length) out.push(['watch', 'Price went up', ups.slice(0, 3).map(x => titleCase_(x[0]) + ' (' + String(x[6]).toLowerCase() + ')').join(', ') + '.']);
  }
  const unc = -st.tx.filter(t => t.cat === 'Uncategorized' && t.amt < 0).reduce((s, t) => s + t.amt, 0);
  if (unc > 50) out.push(['watch', '$' + Math.round(unc).toLocaleString() + ' uncategorized', 'Pick categories in Transactions (or on your phone) so limits are accurate.']);
  return out;
}
function insightsReview_(key, tx, S) {
  S = S || settings_();
  const out = [], st = periodStats_(key, tx);
  const prev = priorPeriods_(key, 3, tx).map(k => periodStats_(k, tx)).filter(p => p.tx.length);
  const goal = pctSetting_(S['Savings goal (% of pay)']) / 100, rate = st.pay ? st.left / st.pay : 0;
  out.push([rate >= Math.max(goal, 0.0001) ? 'good' : rate >= 0 ? 'watch' : 'bad', 'You kept ' + fmt$_(st.left) + (st.pay ? ' (' + Math.round(rate * 100) + '% of pay)' : ''),
    !st.pay ? 'No pay recorded this period.' : rate >= goal ? (goal ? 'That beats your ' + Math.round(goal * 100) + '% goal.' : 'Nice.') : rate >= 0 ? 'Short of the ' + Math.round(goal * 100) + '% goal by ' + fmt$_(st.pay * goal - st.left) + '.' : 'You spent more than you made.']);
  const mv = savingsMoves_(tx, key);
  if (mv.gross > 0 || goal) {
    const g = st.pay * goal;
    out.push([!goal ? 'info' : mv.into >= g - 0.5 ? 'good' : 'watch', 'Savings ' + (mv.into >= 0 ? 'grew by ' : 'shrank by ') + fmt$_(Math.abs(mv.into)),
      'In ' + fmt$_(mv.gross) + ', out ' + fmt$_(mv.takenOut) + (goal ? ' — goal ' + fmt$_(g) + ', ' + (mv.into >= g - 0.5 ? 'hit ✓' : 'missed by ' + fmt$_(g - mv.into)) : '') + '.']);
  }
  const avgPay = prev.length ? prev.reduce((s, p) => s + p.pay, 0) / prev.length : 0;
  if (avgPay && st.pay > avgPay * 1.3) out.push(['info', 'Bigger pay than usual', fmt$_(st.pay) + ' vs ~' + fmt$_(avgPay) + ' — likely an extra paycheck, so this period looks better than a normal one.']);
  const over = Object.keys(st.byCat).map(c => [c, st.byCat[c] - limitFor_(c, key, tx, S), limitFor_(c, key, tx, S)])
    .filter(x => x[2] && x[1] > 0 && NO_LIMIT_CATS.indexOf(x[0]) === -1 && FIXED_CATS.indexOf(x[0]) === -1).sort((a, b) => b[1] - a[1]);
  out.push(over.length ? ['bad', 'Over limit in ' + over.length + ' categor' + (over.length > 1 ? 'ies' : 'y'), over.slice(0, 4).map(x => x[0] + ' by ' + fmt$_(x[1])).join(', ') + '.']
    : ['good', 'Every category stayed within its limit', '']);
  if (prev.length) {
    const ch = Object.keys(Object.assign({}, st.byCat, prev[0].byCat)).filter(c => NO_LIMIT_CATS.indexOf(c) === -1).map(c =>
      [c, (st.byCat[c] || 0) - prev.reduce((s, p) => s + (p.byCat[c] || 0), 0) / prev.length]).sort((a, b) => b[1] - a[1]);
    if (ch.length && ch[0][1] > 50) out.push(['watch', 'Biggest jump: ' + ch[0][0], '+' + fmt$_(ch[0][1]) + ' vs your average.']);
    if (ch.length && ch[ch.length - 1][1] < -50) out.push(['good', 'Biggest drop: ' + ch[ch.length - 1][0], fmt$_(ch[ch.length - 1][1]) + ' vs your average.']);
  }
  const merch = {}; st.tx.filter(t => isSpend_(t) && t.amt < 0).forEach(t => { const k = displayName_(t.desc); merch[k] = (merch[k] || 0) - t.amt; });
  const topM = Object.keys(merch).sort((a, b) => merch[b] - merch[a])[0];
  if (topM) out.push(['info', 'Most money went to ' + topM, fmt$_(merch[topM]) + '.']);
  const sub = sheet_('Subscriptions');
  if (sub.getLastRow() > 1) { const tot = sub.getRange(2, 1, sub.getLastRow() - 1, 2).getValues().find(x => /total per year/i.test(x[0])); if (tot) out.push(['info', 'Recurring charges: ' + fmt$_(tot[1]) + '/yr', 'Check the Subscriptions tab for anything to cancel.']); }
  const unc = -st.tx.filter(t => t.cat === 'Uncategorized' && t.amt < 0).reduce((s, t) => s + t.amt, 0);
  if (unc > 50) out.push(['watch', fmt$_(unc) + ' still uncategorized', 'A few new Rules will sharpen next period.']);
  return out;
}
function buildInsights_(tx) {
  const ss = finSS_(), S = settings_(), tz = Session.getScriptTimeZone(), now = new Date();
  const sh = ss.getSheetByName('Insights') || ss.insertSheet('Insights');
  sh.clear(); sh.getRange('A:H').breakApart(); sh.setConditionalFormatRules([]); sh.getBandings().forEach(b => b.remove());
  uiPage_(sh); sh.setHiddenGridlines(true);
  [56, 200, 118, 118, 118, 118, 118, 118].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('A:H').setFontFamily('Arial').setFontColor(UI.ink).setVerticalAlignment('middle').setFontSize(10);
  sh.getRange('A1:F1').merge().setValue('Insights').setFontSize(18).setFontWeight('bold');
  sh.getRange('G1:H1').merge().setValue('Updated ' + Utilities.formatDate(now, tz, 'MMM d, h:mm a')).setFontColor(UI.mute).setHorizontalAlignment('right');
  sh.getRange('A2:H2').merge().setValue('Plain-English notes on your money. Green = good · amber = keep an eye on it · red = act on it · blue = good to know.').setFontColor(UI.mute);
  const LOOK = { good: ['✓', KJ.green, KJ.greenBg], watch: ['!', KJ.orange, KJ.orangeBg], bad: ['✕', KJ.red, KJ.redBg], info: ['i', KJ.blue, KJ.blueBg] };
  const block = (r, title, list) => {
    r = uiBand_(sh, r, title);
    if (!list.length) { sh.getRange(r, 1, 1, 8).merge().setValue('Nothing to report yet.').setFontColor(UI.mute); return r + 2; }
    list.forEach(x => {
      const L = LOOK[x[0]] || LOOK.info;
      sh.getRange(r, 1).setValue(L[0]).setFontColor(L[1]).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
      sh.getRange(r, 2, 1, 2).merge().setValue(x[1]).setFontWeight('bold').setFontColor(L[1]).setWrap(true);
      sh.getRange(r, 4, 1, 5).merge().setValue(x[2]).setWrap(true);
      sh.getRange(r, 1, 1, 8).setBackground(L[2]).setBorder(true, false, true, false, false, false, KJ.card, SpreadsheetApp.BorderStyle.SOLID_THICK);
      sh.setRowHeight(r, 46); r++;
    });
    return r + 1;
  };
  const cur = currentPeriod_(now), last = priorPeriods_(cur, 1, tx)[0];
  let r = block(4, 'This period so far — ' + periodRange_(cur), insightsNow_(tx, S));
  if (last) r = block(r, 'Last period in review — ' + periodRange_(last), insightsReview_(last, tx, S));
  sh.getRange(r, 1, 1, 8).merge().setValue('Updates with every refresh. The same review is emailed when each period closes (if Monthly insights email is Yes).').setFontColor(UI.mute).setFontSize(9);
}

// Friendly subscription name for display only (the saved key stays as-is): "NETFLIX COM" → Netflix, "HULU -" → Hulu.
function subName_(k) {
  const t = String(k).toUpperCase().replace(/\b(COM|NET|ORG|INC|LLC|CO|BILL|PAYMENT|AUTOPAY)\b/g, ' ').replace(/[^A-Z0-9&' ]+/g, ' ').replace(/\s+/g, ' ').trim();
  return titleCase_(t || String(k));
}
// Friendly store name: the bank's clean name after "|" when there is one (e.g. "7-Eleven"), else a tidied description.
function displayName_(desc) {
  const parts = String(desc).split('|'), clean = parts.length > 1 ? parts[parts.length - 1].trim().replace(/^(apl ?pay|tst\*|sq \*)\s*/i, '') : '';
  return clean.length >= 3 && !/^(ach|pos|debit|withdrawal|transfer)$/i.test(clean) ? clean : prettyName_(desc);
}

// ---------------- Year in review PDF ----------------
function archiveYearly() {
  const now = new Date();
  if (now.getMonth() !== 0) return;                 // runs monthly; only acts in January
  const y = now.getFullYear() - 1, tx = loadTx_();
  const pdf = archiveYear_(y, tx, false);
  const S = settings_();
  if (pdf && yes_(S['Monthly insights email (Yes/No)'])) {
    try { mail_({ to: alertEmail_(), subject: 'Finance HQ: your ' + y + ' year in review', htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px">Your ' + y + ' summary is attached and saved in the Drive folder "' + ARCHIVE_FOLDER + '".</div>', attachments: [pdf] }); } catch (e) { log_('Year email failed: ' + e); }
  }
}
function yearInReviewNow() {
  const ss = finSS_();
  toast_(ss, 'Building your year in review…', 'Finance HQ', 30);
  const y = new Date().getFullYear();
  archiveYear_(y, loadTx_(), true);
  toast_(ss, 'Saved "' + y + ' Year in Review (so far).pdf" in your Drive folder "' + ARCHIVE_FOLDER + '" ✓', 'Finance HQ', 8);
}
function ensureYearlyTrigger_() {
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'archiveYearly')) return;
  try { ScriptApp.newTrigger('archiveYearly').timeBased().onMonthDay(2).atHour(9).create(); } catch (e) { log_('Year-in-review schedule not created: ' + e); }
}
function archiveYear_(year, tx, soFar) {
  const tz = Session.getScriptTimeZone(), S = settings_(), ys = String(year);
  const $ = n => (n < 0 ? '−$' : '$') + Math.round(Math.abs(n)).toLocaleString('en-US');
  const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const cur = currentPeriod_();
  // finished, full periods only: skip the one still running and a first period your data only partly covers
  const firstTx = tx.length ? tx.reduce((a, t) => t.date < a ? t.date : a, tx[0].date) : null;
  const dataStart = firstTx ? new Date(firstTx.getFullYear(), firstTx.getMonth(), firstTx.getDate()) : null;
  const keys = [...new Set(tx.map(t => t.month))].filter(k => k.slice(0, 4) === ys && k < cur && (!dataStart || periodBounds_(k)[0] >= dataStart)).sort();
  if (!keys.length) { log_('Year in review: no finished periods yet for ' + year); toast_(finSS_(), 'No finished periods in ' + year + ' yet.', 'Finance HQ', 6); return null; }
  const P = PERIOD_ || { on: false };
  const yt = tx.filter(t => keys.indexOf(t.month) !== -1), sp = yt.filter(isSpend_);
  const stats = keys.map(k => periodStats_(k, tx));
  const pay = stats.reduce((s, x) => s + x.pay, 0), spend = stats.reduce((s, x) => s + x.spend, 0), left = pay - spend, rate = pay ? left / pay : 0;
  const goal = pctSetting_(S['Savings goal (% of pay)']) / 100;
  const moves = keys.map(k => savingsMoves_(tx, k)), saved = moves.reduce((s, m) => s + m.into, 0);
  const hit = keys.filter((k, i) => stats[i].pay && moves[i].into >= stats[i].pay * goal - 0.5).length;
  const firstDay = tx.length ? tx.reduce((a, t) => t.date < a ? t.date : a, tx[0].date) : null;
  const sumBy = (arr, f) => { const o = {}; arr.forEach(t => { const k = f(t); o[k] = (o[k] || 0) - t.amt; }); return o; };
  const cat = sumBy(sp, t => t.cat), cats = Object.keys(cat).filter(c => cat[c] > 0.5).sort((a, b) => cat[b] - cat[a]);
  const merch = sumBy(sp.filter(t => t.amt < 0), t => displayName_(t.desc)), mcount = {};
  sp.filter(t => t.amt < 0).forEach(t => { const k = displayName_(t.desc); mcount[k] = (mcount[k] || 0) + 1; });
  const topM = Object.keys(merch).sort((a, b) => merch[b] - merch[a]).slice(0, 10);
  const trips = sumBy(sp.filter(t => t.trip), t => t.trip);
  const big = sp.filter(t => t.amt < 0).sort((a, b) => a.amt - b.amt).slice(0, 10);
  const lbl = k => Utilities.formatDate(new Date(+k.slice(0, 4), +k.slice(5, 7) - 1, 15), tz, 'MMM yyyy');
  const best = keys.map((k, i) => [k, stats[i].pay ? stats[i].left / stats[i].pay : -9]).sort((a, b) => b[1] - a[1])[0];
  const priciest = keys.map((k, i) => [k, stats[i].spend]).sort((a, b) => b[1] - a[1])[0];

  // net worth + loans: first and last snapshot of the year
  const nw = sheet_('NetWorth'), nwRows = nw.getLastRow() > 1 ? nw.getRange(2, 1, nw.getLastRow() - 1, 8).getValues().filter(x => x[0] instanceof Date && x[0].getFullYear() === year) : [];
  const nw0 = nwRows[0], nw1 = nwRows[nwRows.length - 1];
  const nwTxt = nw0 && nw1 && nw0 !== nw1 ? $(nw1[3]) + ' on ' + Utilities.formatDate(nw1[0], tz, 'MMM d') + ' — ' + (nw1[3] >= nw0[3] ? 'up ' : 'down ') + $(Math.abs(nw1[3] - nw0[3])) + ' since ' + Utilities.formatDate(nw0[0], tz, 'MMM d') +
    (Number(nw0[7]) > Number(nw1[7]) ? '. Loans paid down by ' + $(nw0[7] - nw1[7]) + '.' : '.') : nw1 ? $(nw1[3]) + ' (history starts ' + Utilities.formatDate(nw1[0], tz, 'MMM d') + ')' : 'no snapshot yet';
  let subTot = 0;
  const sub = sheet_('Subscriptions');
  if (sub.getLastRow() > 1) { const t = sub.getRange(2, 1, sub.getLastRow() - 1, 2).getValues().find(x => /total per year/i.test(x[0])); if (t) subTot = Number(t[1]) || 0; }

  const tile = (l, v, s, col) => '<td class="tile"><div class="tl">' + l + '</div><div class="tv" style="color:' + (col || '#111827') + '">' + v + '</div><div class="ts">' + s + '</div></td>';
  const rows = keys.map((k, i) => { const x = stats[i]; return '<tr><td>' + lbl(k) + (k === cur ? ' (so far)' : '') + '</td><td class="n">' + $(x.pay) + '</td><td class="n">' + $(x.spend) + '</td><td class="n" style="color:' + (x.left >= 0 ? '#059669' : '#DC2626') + '">' + $(x.left) +
    '</td><td class="n">' + (x.pay ? Math.round(x.left / x.pay * 100) + '%' : '—') + '</td><td class="n">' + $(moves[i].into) + '</td></tr>'; }).join('');
  const catRows = cats.slice(0, 14).map(c => '<tr><td>' + esc(c) + '</td><td class="n">' + $(cat[c]) + '</td><td class="n">' + $(cat[c] / keys.length) + '</td><td class="n">' + Math.round(cat[c] / (spend || 1) * 100) + '%</td></tr>').join('');
  const highlights = [
    'You kept ' + $(left) + ' of ' + $(pay) + ' (' + Math.round(rate * 100) + '%)' + (goal ? ' — goal ' + Math.round(goal * 100) + '%.' : '.'),
    'Savings and investments grew by ' + $(saved) + ' from money you moved in (market gains not counted)' + (goal ? '; you hit the savings goal in ' + hit + ' of ' + keys.length + ' periods.' : '.'),
    cats.length ? 'Biggest category: ' + cats[0] + ' at ' + $(cat[cats[0]]) + ' (about ' + $(cat[cats[0]] / keys.length) + ' a period).' : '',
    best && best[1] > -9 ? 'Best period: ' + lbl(best[0]) + ' (kept ' + Math.round(best[1] * 100) + '% of pay).' : '',
    priciest ? 'Most spent: ' + lbl(priciest[0]) + ' (' + $(priciest[1]) + ').' : '',
    topM.length ? 'Most-visited: ' + Object.keys(mcount).sort((a, b) => mcount[b] - mcount[a])[0] + ' (' + mcount[Object.keys(mcount).sort((a, b) => mcount[b] - mcount[a])[0]] + ' times).' : '',
    subTot ? 'Recurring charges run about ' + $(subTot) + ' a year.' : ''
  ].filter(Boolean);

  const html = '<html><head><style>body{font-family:Arial,Helvetica,sans-serif;color:#111827;font-size:11px;margin:28px}' +
    'h1{font-size:22px;margin:0}h2{font-size:13px;margin:22px 0 6px;padding:5px 8px;background:#1F2937;color:#fff}' +
    'table{width:100%;border-collapse:collapse}td,th{padding:5px 6px;border-bottom:1px solid #E5E7EB;text-align:left}th{color:#6B7280;font-weight:normal;font-size:10px}' +
    '.n{text-align:right}.tile{background:#F3F4F6;text-align:center;border:4px solid #fff;padding:10px}.tl{font-size:9px;color:#6B7280;letter-spacing:1px}' +
    '.tv{font-size:20px;font-weight:bold;margin:4px 0}.ts{font-size:9px;color:#6B7280}.sub{color:#6B7280}</style></head><body>' +
    '<h1>' + year + ' — year in review' + (soFar ? ' (so far)' : '') + '</h1><div class="sub">Finance HQ · ' + (P.on ? 'by statement period' : 'by calendar month') + ' · ' + lbl(keys[0]) + ' – ' + lbl(keys[keys.length - 1]) +
    (firstDay && firstDay.getFullYear() === year ? ' · your data starts ' + Utilities.formatDate(firstDay, tz, 'MMM d') + ', so only full periods after that are counted' : '') + ' · generated ' + Utilities.formatDate(new Date(), tz, 'MMM d, yyyy') + '</div>' +
    '<table style="margin-top:14px"><tr>' +
    tile('PAY', $(pay), keys.length + ' period' + (keys.length > 1 ? 's' : '')) +
    tile('SPENDING', $(spend), 'about ' + $(spend / keys.length) + ' a period') +
    tile('LEFT OVER', $(left), Math.round(rate * 100) + '% of pay', left >= 0 ? '#059669' : '#DC2626') +
    tile('SAVINGS GREW BY', $(saved), goal ? 'goal hit ' + hit + ' of ' + keys.length : 'money moved in', saved >= 0 ? '#2563EB' : '#DC2626') +
    '</tr></table>' +
    '<h2>HIGHLIGHTS</h2><ul>' + highlights.map(h => '<li style="margin:4px 0">' + esc(h) + '</li>').join('') + '</ul>' +
    '<h2>PERIOD BY PERIOD</h2><table><tr><th>Period</th><th class="n">Pay</th><th class="n">Spent</th><th class="n">Left over</th><th class="n">Kept</th><th class="n">Savings grew</th></tr>' + rows + '</table>' +
    '<h2>SPENDING BY CATEGORY</h2><table><tr><th>Category</th><th class="n">Total</th><th class="n">Per period</th><th class="n">Share</th></tr>' + catRows + '</table>' +
    '<h2>WHERE THE MONEY WENT</h2><table><tr><th>Place</th><th class="n">Total</th><th class="n">Visits</th></tr>' + topM.map(m => '<tr><td>' + esc(m) + '</td><td class="n">' + $(merch[m]) + '</td><td class="n">' + mcount[m] + '</td></tr>').join('') + '</table>' +
    '<h2>BIGGEST PURCHASES</h2><table><tr><th>Date</th><th>Merchant</th><th>Category</th><th class="n">Amount</th></tr>' + big.map(t => '<tr><td>' + Utilities.formatDate(t.date, tz, 'MMM d') + '</td><td>' + esc(displayName_(t.desc)) + '</td><td>' + esc(t.cat) + '</td><td class="n">' + $(-t.amt) + '</td></tr>').join('') + '</table>' +
    (Object.keys(trips).length ? '<h2>TRIPS</h2><table><tr><th>Trip</th><th class="n">Spent</th></tr>' + Object.keys(trips).map(k => '<tr><td>' + esc(k) + '</td><td class="n">' + $(trips[k]) + '</td></tr>').join('') + '</table>' : '') +
    '<h2>NET WORTH</h2><p>' + esc(nwTxt) + '</p></body></html>';

  const name = year + ' Year in Review' + (soFar ? ' (so far)' : '') + '.pdf';
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
  const folder = archiveFolder_();
  const old = folder.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(blob);
  log_('Year in review saved: ' + name);
  return blob;
}

// ---------------- phone: change a transaction's category ----------------
// Called from the phone app. Locks the category ("manual"); optionally adds a Rule so the same store is always categorized this way.
function setTxCategory(id, cat, makeRule) {
  cat = String(cat || '').trim();
  if (!id || !cat) throw new Error('Missing transaction or category');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let ruleMsg = '';
  try {
    const sh = sheet_('Transactions'), n = sh.getLastRow() - 1;
    if (n < 1) throw new Error('No transactions');
    const ids = sh.getRange(2, 8, n, 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(id));
    if (i < 0) throw new Error('That transaction is no longer in the sheet — pull down to refresh.');
    sh.getRange(i + 2, 6, 1, 2).setValues([[cat, 'manual']]); bumpUsage_('cat');
    if (makeRule) {
      const desc = String(sh.getRange(i + 2, 4).getValue()).toUpperCase();
      let key = merchantKey_(desc);
      if (desc.indexOf(key) === -1) key = key.split(' ')[0];
      if (key.length >= 4 && desc.indexOf(key) !== -1) {
        const rs = sheet_('Rules'), m = Math.max(0, rs.getLastRow() - 1);
        const rows = m ? rs.getRange(2, 1, m, 3).getValues().filter(r => String(r[0]).trim() !== '' && String(r[0]).trim().toUpperCase() !== key) : [];
        const all = [[key, cat, 'mine']].concat(rows);
        if (m) rs.getRange(2, 1, m, 3).clearContent();
        rs.getRange(2, 1, all.length, 3).setValues(all);
        recategorizeAll_();
        ruleMsg = key;
      }
    }
  } finally { lock.releaseLock(); }
  PERIOD_ = null;
  const data = JSON.parse(getAppData()); data.ruleMade = ruleMsg;
  return JSON.stringify(data);
}
// The phone calls this right after a change so the sheet tabs catch up (runs in the background).
function refreshAfterPhoneEdit() { liveRebuild_('Transactions'); return true; }

// ===== PART 11 — Tester check-ins (usage counts only) =====
// Once a week each copy sends the Finance HQ creator a few COUNTS: is the bank syncing, how many days the phone app
// and sheet were used, whether transfers were ticked, how many categories were fixed. Never amounts, balances,
// account names or store names. Turn it off in Settings: "Share usage check-ins with the Finance HQ creator (Yes/No)".
// The creator fills CHECKIN below once (run setupTesterTracker in their own copy). Blank = check-ins are off for everyone.
const CHECKIN = {"url":"https://docs.google.com/forms/d/e/1FAIpQLScXchq1zCkaQV7h3XxZPVmPBdjLkJYV6mmv5V_pirhoTTw-Hw/formResponse","ids":{"tester":"entry.344453892","week":"entry.1595464086","synced":"entry.199959467","accounts":"entry.329170921","newTx":"entry.249597811","appDays":"entry.728590027","sheetDays":"entry.1864243060","catFixes":"entry.166478788","ticks":"entry.2128658648","limitEdits":"entry.558089935","setupDone":"entry.1823756908","savingsGoal":"entry.729450541","yearly":"entry.1270006795","problems":"entry.1275410742","version":"entry.1893017258"}};
const FHQ_VERSION = '2026.11.22.8';
const CHECKIN_FIELDS = [
  ['tester', 'Tester (Google account)'], ['week', 'Week since install'], ['synced', 'Bank synced in last 2 days'],
  ['accounts', 'Accounts linked'], ['newTx', 'New transactions this week'], ['appDays', 'Days phone app opened (7d)'],
  ['sheetDays', 'Days sheet edited (7d)'], ['catFixes', 'Categories fixed (7d)'], ['ticks', 'Transfers ticked this period'],
  ['limitEdits', 'Limits changed (7d)'], ['setupDone', 'Card dates set'], ['savingsGoal', 'Savings goal set'],
  ['yearly', 'Yearly costs listed'], ['problems', 'Health problems'], ['version', 'Version']
];

// ---- usage counters (dates only, kept 30 days, stored in your own script)
function bumpUsage_(kind) {
  try {
    const p = PropertiesService.getScriptProperties(), u = JSON.parse(p.getProperty('USAGE') || '{}');
    const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), cut = Utilities.formatDate(new Date(Date.now() - 30 * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    (u[kind] = u[kind] || []).push(day);
    Object.keys(u).forEach(k => u[k] = u[k].filter(d => d >= cut).slice(-300));
    p.setProperty('USAGE', JSON.stringify(u));
    if (!p.getProperty('INSTALLED_AT')) p.setProperty('INSTALLED_AT', String(Date.now()));
  } catch (e) {}
}
function usageCounts_(days) {
  const u = (() => { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('USAGE') || '{}'); } catch (e) { return {}; } })();
  const cut = Utilities.formatDate(new Date(Date.now() - days * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const recent = k => (u[k] || []).filter(d => d > cut);
  return { appDays: new Set(recent('app')).size, sheetDays: new Set(recent('edit')).size, catFixes: recent('cat').length, limitEdits: recent('limit').length };
}

// ---- what gets sent (counts only)
function checkinPayload_() {
  const S = settings_(), p = PropertiesService.getScriptProperties();
  if (!p.getProperty('INSTALLED_AT')) p.setProperty('INSTALLED_AT', String(Date.now()));
  const week = Math.floor((Date.now() - Number(p.getProperty('INSTALLED_AT'))) / (7 * 864e5)) + 1;
  const acc = sheet_('Accounts'), accts = acc.getLastRow() > 1 ? acc.getRange(2, 1, acc.getLastRow() - 1, 8).getValues() : [];
  const lastSync = accts.map(a => a[7]).filter(d => d instanceof Date).reduce((m, d) => !m || d > m ? d : m, null);
  const tx = sheet_('Transactions'), since = new Date(Date.now() - 7 * 864e5);
  const newTx = tx.getLastRow() > 1 ? tx.getRange(2, 1, tx.getLastRow() - 1, 1).getValues().filter(r => r[0] instanceof Date && r[0] >= since).length : 0;
  const cur = (() => { try { return currentPeriod_(); } catch (e) { return ''; } })(), ticks = getTicks_()[cur] || {};
  const cards = [...new Set(accts.filter(a => a[3] === 'Credit card').map(a => cardName_(a[2], a[1])))].filter(n => cardTracked_(n, S));
  const datesSet = cards.length ? cards.filter(n => Number(S[n + ' closes (day of month)']) >= 1).length + '/' + cards.length : 'no cards';
  let problems = '';
  try { const h = sheet_('Health'); problems = h.getLastRow() > 4 ? h.getRange(1, 1, h.getLastRow(), 1).getValues().filter(r => /Fix/.test(String(r[0]))).length : ''; } catch (e) {}
  const u = usageCounts_(7);
  return { tester: Session.getEffectiveUser().getEmail() || 'unknown', week: week, synced: lastSync && Date.now() - lastSync < 2 * 864e5 ? 'Yes' : 'No',
    accounts: accts.length, newTx: newTx, appDays: u.appDays, sheetDays: u.sheetDays, catFixes: u.catFixes,
    ticks: Object.keys(ticks).filter(k => ticks[k] === true).length, limitEdits: u.limitEdits, setupDone: datesSet,
    savingsGoal: pctSetting_(S['Savings goal (% of pay)']) > 0 ? 'Yes' : 'No', yearly: loadYearly_().length, problems: problems, version: FHQ_VERSION };
}
function checkinIfDue_() {
  if (isDemo_()) return;
  if (!CHECKIN.url || !Object.keys(CHECKIN.ids || {}).length) return;
  if (!yes_(settings_()['Share usage check-ins with the Finance HQ creator (Yes/No)'])) return;
  const p = PropertiesService.getScriptProperties(), last = Number(p.getProperty('CHECKIN_LAST') || 0);
  if (Date.now() - last < 6.5 * 864e5) return;
  const data = checkinPayload_(), payload = {};
  Object.keys(CHECKIN.ids).forEach(k => { if (data[k] !== undefined) payload[CHECKIN.ids[k]] = String(data[k]); });
  try {
    const res = UrlFetchApp.fetch(CHECKIN.url, { method: 'post', payload: payload, muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() < 400) p.setProperty('CHECKIN_LAST', String(Date.now()));
  } catch (e) {}
}
// Shows exactly what your copy would send — Finance HQ menu → "What check-ins share".
function showCheckinPreview() {
  const d = checkinPayload_(), on = yes_(settings_()['Share usage check-ins with the Finance HQ creator (Yes/No)']) && !!CHECKIN.url;
  const rows = CHECKIN_FIELDS.map(f => '<tr><td style="color:#8A857A;padding:4px 10px 4px 0">' + f[1] + '</td><td><b>' + String(d[f[0]]).replace(/[<>&]/g, '') + '</b></td></tr>').join('');
  const html = '<div style="font:14px Arial;color:#2A2A2A"><p>' + (on ? 'Once a week your copy sends these <b>counts</b> to the person who shared Finance HQ with you, so they can see if it\'s working for you:' :
    'Check-ins are <b>off</b>. If they were on, this is all that would be sent:') + '</p><table>' + rows + '</table>' +
    '<p style="color:#8A857A">Never sent: amounts, balances, account or card names, store names, transactions. Turn off in Settings → "Share usage check-ins…" (set to No).</p></div>';
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(560), 'What check-ins share'); } catch (e) { Logger.log(JSON.stringify(d)); }
}
// The creator's one-time tracker setup (setupTesterTracker) lives in their private Personal file, so shared copies never ask for Google Forms access.

// ===== PART 12 — Easy setup: copy the template → Finance HQ → Connect my bank =====
// Friends get a link that makes their own copy of a blank Finance HQ sheet (the script comes with it).
// They never open Apps Script: the menu asks for their SimpleFIN setup token, which is kept only in their own copy.

function updateLabel_() {
  try { const u = JSON.parse(PropertiesService.getScriptProperties().getProperty('UPDATE_AVAILABLE') || 'null'); if (u) return '⬆ Install update (' + u.version + ')'; } catch (e) {}
  return 'Check for updates (you have ' + FHQ_VERSION + ')';
}
function isConnected_() { try { return !!PropertiesService.getScriptProperties().getProperty(SF_ACCESS); } catch (e) { return false; } }

function showConnect() {
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(CONNECT_HTML).setWidth(520).setHeight(560), 'Connect your bank'); } catch (e) {}
}
// Called by the Connect dialog. The token goes into this copy's own script properties and is used once.
function connectBank(token) {
  if (isDemo_()) throw new Error('This copy is a demo with sample data. Make a fresh copy of Finance HQ to connect your own bank.');
  token = String(token || '').replace(/\s+/g, '');
  if (!token) throw new Error('Paste your setup token first.');
  let claim = '';
  try { claim = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString().trim(); } catch (e) {}
  if (!/^https:\/\//.test(claim)) throw new Error('That doesn\'t look like a SimpleFIN setup token. In SimpleFIN, click "New app connection" and copy the long code it shows.');
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty(SF_ACCESS)) throw new Error('Already connected. To connect different accounts, add them on the SimpleFIN website — they show up here automatically.');
  finSS_();                                            // remember which sheet this is
  P.setProperty('SIMPLEFIN_SETUP_TOKEN', token);
  try { setupFinanceHQ(); }
  catch (e) { P.deleteProperty('SIMPLEFIN_SETUP_TOKEN'); throw new Error(String(e.message || e).replace('replace SIMPLEFIN_SETUP_TOKEN', 'paste the new token here')); }
  const ss = finSS_(), w = ss.getSheetByName('Welcome');
  if (w && ss.getSheets().length > 1) ss.deleteSheet(w);
  financeMenu_();                                      // swap "Connect my bank" for the full menu
  const acc = sheet_('Accounts'), tx = sheet_('Transactions');
  return 'Connected ' + Math.max(0, acc.getLastRow() - 1) + ' accounts and ' + Math.max(0, tx.getLastRow() - 1) + ' transactions.';
}

function showPhoneSetup() {
  const id = ScriptApp.getScriptId();
  const html = '<div style="font:14px/1.5 Arial;color:#2A2A2A">' +
    '<p>The phone app is a private page only you can open. One-time setup, about 2 minutes:</p><ol>' +
    '<li><a href="https://script.google.com/home/projects/' + id + '/edit" target="_blank">Open the script</a> (opens in a new tab).</li>' +
    '<li>Top right: <b>Deploy → Test deployments</b>.</li>' +
    '<li>Copy the <b>Web app URL</b> (it ends in <b>/dev</b>).</li>' +
    '<li>Open it on your phone in Safari or Chrome, signed in to this same Google account, then <b>Share → Add to Home Screen</b>.</li></ol>' +
    '<p style="color:#8A857A">This link always runs your newest code, so updates reach the phone app by themselves — no redeploying. Only you can open it.</p></div>';
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(380), 'Set up the phone app'); } catch (e) {}
}

// ---- for the person sharing Finance HQ: run once in a NEW, BLANK sheet that has this script pasted in.
// Leaves a single friendly "Welcome" tab and no saved settings, ready to share as a "make a copy" link.
const STEP2_TEXT = [
  'A.  Wait a few seconds after the sheet opens. A "Finance HQ" menu shows up in the top menu bar, to the right of Help. Not there? Refresh the page. Still not there in Safari? Open this sheet in Chrome, or in a regular (not Private) Safari window.',
  'B.  Click  Finance HQ → Connect my bank…',
  'C.  The first time, Google asks you to allow the sheet\'s script:  OK → pick your Google account → "Google hasn\'t verified this app" → Advanced → Go to Finance HQ (unsafe) → tick Select all → Continue.',
  'D.  That first click only gave permission. Click  Finance HQ → Connect my bank…  again.',
  'E.  Paste your setup token from step 1 and press Connect. Keep the window open for 1–3 minutes while it builds your tabs.'
].join('\n');
function prepareTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Run this from a Google Sheet (Extensions → Apps Script), not a standalone script.');
  if (ss.getSheetByName('Transactions') && ss.getSheetByName('Transactions').getLastRow() > 1) throw new Error('This sheet has real data. Use a new, blank Google Sheet for the template.');
  const w = ss.getSheetByName('Welcome') || ss.insertSheet('Welcome', 0);
  ss.getSheets().filter(sh => sh.getName() !== 'Welcome').forEach(sh => ss.deleteSheet(sh));
  w.clear(); w.setHiddenGridlines(true); w.getRange('A:C').breakApart();
  w.setColumnWidth(1, 40); w.setColumnWidth(2, 620); w.setColumnWidth(3, 40);
  w.getRange(1, 1, 40, 3).setBackground(UI.page).setFontFamily('Arial').setFontColor(UI.ink).setVerticalAlignment('middle');
  const rows = [
    ['Welcome to Finance HQ', 'title'],
    ['Your own private money tracker. About 15 minutes to set up — no code, nothing to install. Do setup on a computer in Chrome, or in a regular (not Private) Safari window. The Finance HQ menu doesn\'t show in the Google Sheets phone app or in Safari Private windows. The phone app comes in step 4.', 'sub'],
    ['', ''],
    ['1 · Get SimpleFIN (links your bank and cards, about $15/year)', 'head'],
    ['Go to beta-bridge.simplefin.org → sign up → add your banks and cards → click "New app connection" → copy the setup token (a long code). Treat it like a password.', 'text'],
    ['2 · Connect your bank (about 3 minutes)', 'head'],
    [STEP2_TEXT, 'steps'],
    ['Why "unverified"? It\'s this sheet\'s own script, not an app from a store. It only reads your accounts through SimpleFIN and can never move money.', 'note'],
    ['3 · Make it yours', 'head'],
    ['The Quick setup form opens by itself: card statement dates, savings goal, debts, trips. Everything is optional.', 'text'],
    ['4 · Phone app (optional)', 'head'],
    ['Finance HQ → Set up the phone app… walks you through it.', 'text'],

    ['', ''],
    ['Stuck? Use  Finance HQ → Get help…  (it\'s in the menu even before you connect), or text the person who shared this.', 'text'],
    ['Your money data stays in your own Google account. Weekly check-ins send the person who shared this a few usage counts only (never amounts or account details) — see Start here → Terms & privacy after setup, or turn them off in Settings.', 'note']
  ];
  rows.forEach((x, i) => {
    const c = w.getRange(i + 1, 2).setValue(x[0]).setWrap(true);
    if (x[1] === 'title') { c.setFontSize(22).setFontWeight('bold'); w.setRowHeight(i + 1, 48); }
    else if (x[1] === 'sub') c.setFontColor(UI.mute).setFontSize(11);
    else if (x[1] === 'head') { c.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(UI.band); w.setRowHeight(i + 1, 30); }
    else if (x[1] === 'text') { c.setBackground('#FFFFFF'); w.setRowHeight(i + 1, 44); }
    else if (x[1] === 'steps') { c.setBackground('#FFFFFF').setVerticalAlignment('top'); w.setRowHeight(i + 1, 200); }
    else if (x[1] === 'note') { c.setFontColor(UI.mute).setFontSize(9); w.setRowHeight(i + 1, 36); }
  });
  w.setTabColor(UI.band);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('Template ready. Share this link (it makes each friend their own copy): ' + ss.getUrl().replace(/\/edit.*$/, '') + '/copy');
  return ss.getUrl().replace(/\/edit.*$/, '') + '/copy';
}

const CONNECT_HTML = `<!DOCTYPE html><html><head><base target="_top"><style>
body{font:14px/1.5 Arial,sans-serif;color:#2A2A2A;margin:0;padding:4px 8px}
h3{font-size:13px;letter-spacing:.05em;text-transform:uppercase;color:#fff;background:#2F6F4E;padding:7px 10px;border-radius:6px;margin:14px 0 8px}
.hint{color:#8A857A;font-size:12px}
textarea{width:100%;box-sizing:border-box;height:90px;font:12px monospace;padding:8px;border:1px solid #D1D5DB;border-radius:8px;background:#FFF8E1}
button{font:14px Arial;padding:9px 18px;border-radius:8px;border:1px solid #2F6F4E;background:#2F6F4E;color:#fff;font-weight:bold;cursor:pointer}
button:disabled{opacity:.5}#msg{margin-top:10px;font-size:13px}
.ok{color:#2F6F4E}.bad{color:#B8423A}
</style></head><body>
<h3>1 · Get your setup token</h3>
<p>On <b>beta-bridge.simplefin.org</b>: add your banks and cards, then click <b>New app connection</b> and copy the long code.</p>
<p class="hint">It works only once and it's like a password — don't text or email it. It's stored only in your own copy of this sheet.</p>
<h3>2 · Paste it here</h3>
<textarea id="tok" placeholder="Paste your SimpleFIN setup token"></textarea>
<p><button id="go" onclick="go()">Connect</button></p>
<div id="msg" class="hint">Connecting pulls your last few months of transactions and builds every tab — about 1–3 minutes. Keep this open.</div>
<script>
function go(){var t=document.getElementById('tok').value,b=document.getElementById('go'),m=document.getElementById('msg');
  if(!t.trim()){m.className='bad';m.textContent='Paste your setup token first.';return;}
  b.disabled=true;m.className='hint';m.textContent='Connecting and building your sheet… (1–3 minutes)';
  google.script.run.withSuccessHandler(function(r){m.className='ok';m.textContent=r+' Opening Quick setup…';
      setTimeout(function(){google.script.run.showQuickSetup();},1200);})
    .withFailureHandler(function(e){b.disabled=false;m.className='bad';m.textContent=e.message;}).connectBank(t);}
</script></body></html>`;

// ===== PART 13 — Payday nudge · feedback check-ins =====

// ---- payday nudge: when a paycheck lands, tell them what to move and what's left to spend
function paydayNudge_() {
  const S = settings_();
  if (!yes_(S['Payday nudge (Yes/No)'] === undefined ? 'Yes' : S['Payday nudge (Yes/No)'])) return;
  const P = PropertiesService.getScriptProperties(), tz = Session.getScriptTimeZone();
  const tx = loadTx_(), since = new Date(Date.now() - 3 * 864e5);
  const pays = tx.filter(t => t.cat === 'Income' && t.amt > 50);
  let seen = []; try { seen = JSON.parse(P.getProperty('PAY_SEEN') || 'null'); } catch (e) {}
  if (!Array.isArray(seen)) {                                   // first run: remember existing paychecks, don't nudge for old ones
    P.setProperty('PAY_SEEN', JSON.stringify(pays.map(t => String(t.id)).slice(-60))); return;
  }
  const fresh = pays.filter(t => t.date >= since && seen.indexOf(String(t.id)) === -1);
  if (!fresh.length) return;
  P.setProperty('PAY_SEEN', JSON.stringify(seen.concat(fresh.map(t => String(t.id))).slice(-60)));
  const check = fresh.reduce((s, t) => s + t.amt, 0);
  const cur = currentPeriod_(), st = periodStats_(cur, tx), coming = payStillComing_(tx, cur);
  const cl = checklist_(tx, S, cur, st.pay + coming.amt), bills = billsComing_(tx, cur, S);
  const pct = setAsidePct_(S), nChecks = Math.max(1, new Set(st.tx.filter(t => t.cat === 'Income' && t.amt > 50).map(t => Utilities.formatDate(t.date, tz, 'yyyyMMdd'))).size + coming.n);
  const owed = cl.items.filter(i => !i.ticked).reduce((s, i) => s + Math.max(0, i.left), 0);
  const move = Math.min(owed, check * pct + (cl.yearly || 0) / nChecks);
  const pb = periodBounds_(cur), now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const daysLeft = Math.max(1, Math.round((pb[1] - today) / 864e5) + 1);
  const left = cl.pay - cl.setAside - bills.amt - st.spend;
  const $ = n => '$' + Math.round(n).toLocaleString();
  const title = '💰 Paid ' + $(check) + (move >= 1 ? ' — move ' + $(move) + ' to savings' : ' — savings already done ✓');
  const body = 'You got paid ' + $(check) + '.\n' + (move >= 1 ? 'Move ' + $(move) + ' to savings now (' + Math.round(pct * 100) + '% of this check' + (cl.yearly ? ' + your yearly-costs share' : '') + '), then tick it on the Dashboard.\n' : 'Your savings transfers are already done this period.\n') +
    'Left to spend this period: ' + $(left) + ' — about ' + $(Math.max(0, left / daysLeft)) + '/day for ' + daysLeft + ' days' + (bills.amt ? ' (after ' + $(bills.amt) + ' of bills still due)' : '') + '.';
  let sent = false;
  if (yes_(S['Alerts: phone reminders (Yes/No)'])) { try { phoneReminder_(title, new Date(Date.now() + 2 * 60000), body); sent = true; } catch (e) { log_('Payday reminder failed: ' + e); } }
  if (!sent || yes_(S['Alerts: email (Yes/No)'])) { try { mail_(alertEmail_(), 'Finance HQ: ' + title.replace(/^💰 /, ''), body); } catch (e) { log_('Payday email failed: ' + e); } }
  log_('Payday nudge: ' + title);
}

// ---- feedback: two short forms (week 1, week 4). The creator fills FEEDBACK once (setupFeedbackForms in their Personal file).
const FEEDBACK = { week1: '', week4: '' };
function feedbackDue_() {
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('INSTALLED_AT')) P.setProperty('INSTALLED_AT', String(Date.now()));
  if (!FEEDBACK.week1 && !FEEDBACK.week4) return;
  if (!yes_(settings_()['Share usage check-ins with the Finance HQ creator (Yes/No)'])) return;
  const days = (Date.now() - Number(P.getProperty('INSTALLED_AT'))) / 864e5;
  [['week1', 7, 'FB1_SENT', 'How\'s week one going? 2 minutes'], ['week4', 28, 'FB4_SENT', 'One month in — 3 quick questions']].forEach(x => {
    if (!FEEDBACK[x[0]] || days < x[1] || P.getProperty(x[2])) return;
    const msg = 'Finance HQ check-in: ' + x[3] + '.\n\nYour answers go to the person who shared Finance HQ with you and help make it better:\n' + FEEDBACK[x[0]] + '\n\n(Also in the menu: Finance HQ → Send feedback.)';
    try { mail_(alertEmail_(), 'Finance HQ: ' + x[3], msg); } catch (e) {}
    try { if (yes_(settings_()['Alerts: phone reminders (Yes/No)'])) phoneReminder_('📝 Finance HQ: ' + x[3], new Date(Date.now() + 3 * 60000), FEEDBACK[x[0]]); } catch (e) {}
    P.setProperty(x[2], String(Date.now()));
  });
}
function sendFeedback() {
  const days = (Date.now() - Number(PropertiesService.getScriptProperties().getProperty('INSTALLED_AT') || Date.now())) / 864e5;
  const link = days >= 21 ? (FEEDBACK.week4 || FEEDBACK.week1) : (FEEDBACK.week1 || FEEDBACK.week4);
  const html = link ? '<div style="font:14px/1.5 Arial;color:#2A2A2A"><p>Thanks — honest answers help the most, especially what\'s annoying or confusing.</p><p><a href="' + link + '" target="_blank" style="display:inline-block;background:#2F6F4E;color:#fff;padding:9px 16px;border-radius:8px;text-decoration:none;font-weight:bold">Open the feedback form</a></p></div>'
    : '<div style="font:14px/1.5 Arial;color:#2A2A2A"><p>Feedback forms aren\'t set up for this copy yet. Tell us what you think through Get help instead — it goes straight to Finance HQ support.</p><p><button onclick="google.script.run.showSupport()" style="font:14px Arial;padding:9px 16px;border-radius:8px;border:1px solid #2F6F4E;background:#2F6F4E;color:#fff;font-weight:bold;cursor:pointer">Open Get help</button></p></div>';
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(180), 'Send feedback'); } catch (e) {}
}

// ===== PART 13b — Get help: message the person who shared Finance HQ, from inside the sheet =====
// Sends from the user's own Gmail, so replies come straight back to them. Technical details are optional and shown
// in full before sending: version, sync status, which health checks need attention, recent errors — dollar amounts removed.
const SUPPORT = { email: 'financehq.app@gmail.com', name: 'Finance HQ support' };

function redact_(t) {                                 // strip money, long numbers and emails from anything we send
  return String(t).replace(/-?\$\s?[\d,]+(\.\d+)?/g, '$…').replace(/\b\d{4,}\b/g, '…').replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, '(email)');
}
function supportDetails_() {
  const out = ['Finance HQ version: ' + FHQ_VERSION];
  try { out.push('Bank connected: ' + (isConnected_() ? 'yes' : 'no')); } catch (e) {}
  try { const a = sheet_('Accounts'); out.push('Accounts syncing: ' + Math.max(0, a.getLastRow() - 1)); } catch (e) {}
  try {
    const h = finSS_().getSheetByName('Health');
    if (h && h.getLastRow() > 1) {
      const bad = h.getRange(1, 1, h.getLastRow(), 2).getValues().filter(r => /Look|Fix/.test(String(r[0]))).map(r => String(r[0]).replace(/[^\w ]/g, '').trim() + ': ' + r[1]);
      out.push('Health checks needing attention: ' + (bad.length ? bad.join('; ') : 'none'));
    }
  } catch (e) {}
  try {
    const l = finSS_().getSheetByName('Log');
    if (l && l.getLastRow() > 1) {
      const n = Math.min(40, l.getLastRow() - 1), rows = l.getRange(l.getLastRow() - n + 1, 1, n, 2).getValues();
      const errs = rows.filter(r => /error|fail|problem|couldn|can.t|denied|invalid/i.test(r[1]) && !/^(Automatic update skipped|Paste this|Tester tracker|Feedback answers|Published |Help request|Running Finance HQ)/.test(String(r[1]))).slice(-6)
        .map(r => (r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'MMM d h:mm a') + ' — ' : '') + redact_(r[1]).slice(0, 200));
      out.push('Recent errors: ' + (errs.length ? '\n  ' + errs.join('\n  ') : 'none'));
    }
  } catch (e) {}
  return out.join('\n');
}
function getSupportPreview() { return JSON.stringify({ details: supportDetails_(), to: SUPPORT.name }); }
function sendSupport(message, includeDetails) {
  message = String(message || '').trim();
  if (message.length < 5) throw new Error('Tell us a little about what\'s going wrong first.');
  const P = PropertiesService.getScriptProperties(), today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sent = JSON.parse(P.getProperty('SUPPORT_SENT') || '{}');
  if ((sent[today] || 0) >= 5) throw new Error('That\'s 5 messages today — if it\'s urgent, text the person who shared Finance HQ with you.');
  let me = ''; try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  const body = message + '\n\n— Sent from Finance HQ' + (me ? ' by ' + me : '') + (includeDetails ? '\n\nTECHNICAL DETAILS\n' + supportDetails_() : '');
  MailApp.sendEmail({ to: SUPPORT.email, subject: 'Finance HQ help: ' + message.replace(/\s+/g, ' ').slice(0, 60), body: body, replyTo: me || undefined, name: 'Finance HQ' });
  sent[today] = (sent[today] || 0) + 1; P.setProperty('SUPPORT_SENT', JSON.stringify({ [today]: sent[today] }));
  try { log_('Help request sent to ' + SUPPORT.name + '.'); } catch (e) {}
  return 'Sent — ' + SUPPORT.name + ' will reply to ' + (me || 'your Gmail') + '.';
}
function showSupport() {
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(SUPPORT_HTML).setWidth(520).setHeight(560), 'Get help'); } catch (e) {}
}
const SUPPORT_HTML = `<!DOCTYPE html><html><head><base target="_top"><style>
body{font:14px/1.5 Arial,sans-serif;color:#2A2A2A;margin:0;padding:4px 8px}
h3{font-size:13px;letter-spacing:.05em;text-transform:uppercase;color:#fff;background:#2F6F4E;padding:7px 10px;border-radius:6px;margin:12px 0 8px}
.hint{color:#8A857A;font-size:12px}
textarea{width:100%;box-sizing:border-box;height:110px;font:14px Arial;padding:8px;border:1px solid #D1D5DB;border-radius:8px}
pre{white-space:pre-wrap;font:11px/1.4 monospace;background:#F7F5F0;border:1px solid #E5E1D8;border-radius:8px;padding:8px;max-height:130px;overflow:auto;margin:6px 0 0}
button{font:14px Arial;padding:9px 18px;border-radius:8px;border:1px solid #2F6F4E;background:#2F6F4E;color:#fff;font-weight:bold;cursor:pointer}
button:disabled{opacity:.5}#msg{margin-top:10px;font-size:13px}.ok{color:#2F6F4E}.bad{color:#B8423A}
</style></head><body>
<h3>What's going wrong?</h3>
<textarea id="m" placeholder="e.g. My Chase card stopped updating on Tuesday. Or: the Connect window says my token is wrong."></textarea>
<p class="hint">A screenshot helps a lot — after sending, reply to the email with one attached.</p>
<h3>Technical details</h3>
<label><input type="checkbox" id="d" checked> Include these (recommended)</label>
<pre id="pv">Loading…</pre>
<p class="hint">No balances, amounts, account numbers or transactions — exactly what's shown above.</p>
<p><button id="go" onclick="go()">Send</button></p><div id="msg"></div>
<script>
google.script.run.withSuccessHandler(function(r){var x=JSON.parse(r);document.getElementById('pv').textContent=x.details;document.getElementById('go').textContent='Send to '+x.to;})
  .withFailureHandler(function(){document.getElementById('pv').textContent='(details unavailable)';}).getSupportPreview();
function go(){var b=document.getElementById('go'),m=document.getElementById('msg');b.disabled=true;m.className='';m.textContent='Sending…';
  google.script.run.withSuccessHandler(function(r){m.className='ok';m.textContent=r+' You can close this window.';})
  .withFailureHandler(function(e){b.disabled=false;m.className='bad';m.textContent=e.message||e;})
  .sendSupport(document.getElementById('m').value,document.getElementById('d').checked);}
</script></body></html>`;

// ===== PART 15 — Demo mode: a full Finance HQ on made-up data, for showing people =====
// Finance HQ → See a demo with sample data (before connecting) fills this copy with a pretend person's last 6 months.
// The pretend bank keeps "syncing" every morning, so the demo always looks current. No emails, reminders or check-ins go out.
const DEMO_ACCESS = 'demo://finance-hq-sample';
function isDemo_() { try { return PropertiesService.getScriptProperties().getProperty(SF_ACCESS) === DEMO_ACCESS; } catch (e) { return false; } }
function mail_() { if (isDemo_()) return; return MailApp.sendEmail.apply(MailApp, arguments); }

function demoBank_(start, end) {
  const DAY_MS = 864e5, now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const from = new Date(today.getTime() - 185 * DAY_MS), lo = Math.max(start || 0, from.getTime()), hi = end || (today.getTime() + DAY_MS);
  const ymd = d => d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
  const rand = key => { let h = 2166136261; for (let i = 0; i < key.length; i++) { h ^= key.charCodeAt(i); h = Math.imul(h, 16777619); } h >>>= 0; return ((h * 9301 + 49297) % 233280) / 233280; };
  const acc = { chk: [], sav: [], sap: [], amx: [] };
  const add = (k, d, amt, desc, n) => { const t = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12); if (t.getTime() < lo || t.getTime() >= hi || t > now) return;
    acc[k].push({ id: 'DEMO-' + ymd(t) + '-' + k + '-' + (n || 0) + '-' + desc.replace(/\W/g, '').slice(0, 8), posted: Math.floor(t.getTime() / 1000), amount: amt.toFixed(2), description: desc, payee: desc }); };
  const spend = [['STARBUCKS STORE 1182', 5.5, 9.5, 0.30, 'amx'], ['CHIPOTLE 2291', 11, 16, 0.13, 'amx'], ['UBER EATS', 18, 42, 0.12, 'amx'], ['SWEETGREEN DC', 13, 17, 0.10, 'amx'],
    ['TRADER JOE S #651', 38, 92, 0.10, 'amx'], ['WHOLE FOODS MKT', 25, 70, 0.04, 'amx'], ['SHELL OIL 5741', 38, 55, 0.10, 'sap'], ['AMAZON.COM*MK2', 14, 85, 0.13, 'sap'],
    ['TARGET 00023', 22, 75, 0.07, 'sap'], ['LYFT *RIDE', 11, 27, 0.10, 'sap'], ['AMC THEATRES 0451', 16, 32, 0.03, 'sap']];
  const subs = [['NETFLIX.COM', 15.49, 3], ['SPOTIFY USA', 11.99, 7], ['APPLE.COM/BILL', 2.99, 12], ['PLANET FITNESS', 24.99, 17], ['HULU 877-8244858', 9.99, 22]];
  const charged = { sap: {}, amx: {} };                                    // per statement month, to pay the card next month
  for (let d = new Date(from); d <= today; d = new Date(d.getTime() + DAY_MS)) {
    const key = ymd(d), dom = d.getDate(), mk = d.getFullYear() * 100 + d.getMonth();
    if (Math.round((d - new Date(2026, 0, 2)) / DAY_MS) % 14 === 0) { add('chk', d, 2150, 'ACME CORP PAYROLL DIRECT DEP'); add('chk', d, -430, 'ONLINE TRANSFER TO SAVINGS 8812', 1); add('sav', d, 430, 'ONLINE TRANSFER FROM CHECKING 4421'); }
    if (dom === 1) add('chk', d, -1350, 'RENT PAYMENT OAKWOOD APTS');
    if (dom === 9) add('chk', d, -128.40, 'GEICO AUTO INSURANCE');
    if (dom === 11) add('chk', d, -65, 'VERIZON WIRELESS PAYMENT');
    if (dom === 15) add('chk', d, -212, 'NAVIENT STUDENT LOAN PMT');
    if (dom === 28) add('sav', d, 4.12, 'INTEREST PAID');
    subs.forEach((x, i) => { if (dom === x[2]) { add('sap', d, -x[1], x[0], 10 + i); charged.sap[mk] = (charged.sap[mk] || 0) + x[1]; } });
    spend.forEach((x, i) => { if (rand(key + x[0]) < x[3]) { const a = Math.round((x[1] + rand(x[0] + key) * (x[2] - x[1])) * 100) / 100; add(x[4], d, -a, x[0], i); charged[x[4]][mk] = (charged[x[4]][mk] || 0) + a; } });
    const prev = d.getMonth() === 0 ? (d.getFullYear() - 1) * 100 + 11 : mk - 1;
    if (dom === 12 && charged.sap[prev]) { const a = Math.round(charged.sap[prev] * 100) / 100; add('chk', d, -a, 'CHASE CREDIT CRD AUTOPAY', 2); add('sap', d, a, 'AUTOMATIC PAYMENT - THANK YOU', 99); }
    if (dom === 14 && charged.amx[prev]) { const a = Math.round(charged.amx[prev] * 100) / 100; add('chk', d, -a, 'AMERICAN EXPRESS EPAYMENT', 3); add('amx', d, a, 'ONLINE PAYMENT - THANK YOU', 99); }
  }
  const owe = k => -(charged[k][today.getFullYear() * 100 + today.getMonth()] || 0);
  const bd = Math.floor(now.getTime() / 1000);
  const a = (id, name, org, bal, list) => ({ id: 'ACT-demo-' + id, name: name, org: { name: org }, currency: 'USD', balance: bal.toFixed(2), 'available-balance': bal.toFixed(2), 'balance-date': bd, transactions: list.sort((x, y) => y.posted - x.posted) });
  return { errors: [], accounts: [a(1, 'Total Checking (4421)', 'Chase Bank', 3240.18, acc.chk), a(2, 'Online Savings (8812)', 'Ally Bank', 8915.40, acc.sav),
    a(3, 'Sapphire Preferred (1234)', 'Chase Bank', owe('sap'), acc.sap), a(4, 'American Express Gold Card (5678)', 'American Express', owe('amx'), acc.amx),
    a(5, 'Robinhood individual (0740)', 'Robinhood', 5212.77, []), a(6, 'Student Loan (9911)', 'Navient', -6820, [])] };
}

// Finance HQ → See a demo with sample data
function startDemo() {
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty(SF_ACCESS) && P.getProperty(SF_ACCESS) !== DEMO_ACCESS) throw new Error('This copy is connected to a real bank. Make a fresh copy of Finance HQ for the demo.');
  toast_(finSS_(), 'Building the demo — about a minute…', 'Finance HQ', 60);
  finSS_(); P.setProperty(SF_ACCESS, DEMO_ACCESS);
  setupFinanceHQ();
  try { const d = JSON.parse(getSetupData()); (d.cards || []).forEach(c => { if (/sapphire|1234/i.test(c.name + c.account)) { c.close = 19; c.due = 16; } else if (/gold|5678/i.test(c.name + c.account)) { c.close = 20; c.due = 14; } });
    (d.debts || []).forEach(x => { if (/loan|navient/i.test(x.name)) { x.rate = 5.5; x.pay = 212; } });
    d.save = 20; saveSetup(JSON.stringify(d)); } catch (e) { log_('Demo setup details skipped: ' + e.message); }
  setSetting_('Spending month (Auto / Statement / Calendar)', 'Statement');
  const ss = finSS_(), w = ss.getSheetByName('Welcome'); if (w && ss.getSheets().length > 1) ss.deleteSheet(w);
  financeMenu_(); PERIOD_ = null; rebuild_(true);
  toast_(ss, 'Demo ready — every number here is made up.', 'Finance HQ', 8);
  return 'Demo ready.';
}

// ===== PART 14 — Updates: one-minute installs =====
// The creator publishes a release (a code file + a small version file in their Google Drive, readable by link).
// Each copy checks once a day. When something newer exists it emails you once and adds "⬆ Install update" to the
// Finance HQ menu: that window copies the new code for you — paste it over this file in the script editor and save.
// (Google doesn't let a sheet's script rewrite itself without a separate Cloud project, so the paste stays a human step.)
// Only this Finance HQ file changes; your data, settings and any other script files stay.
const UPDATE_SRC = { versionFile: '1hboJqNvt9l5ER8Y7DBI1XoWO7VrK0Zez' };

function newerVersion_(a, b) {                       // true when a > b, e.g. "2026.12.1.1" > "2026.11.21.4"
  const x = String(a).split(/\D+/).map(Number), y = String(b).split(/\D+/).map(Number);
  for (let i = 0; i < Math.max(x.length, y.length); i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); }
  return false;
}
function driveText_(id) {
  const r = UrlFetchApp.fetch('https://drive.google.com/uc?export=download&id=' + id, { muteHttpExceptions: true, followRedirects: true });
  if (r.getResponseCode() !== 200) throw new Error('Could not reach the update file (' + r.getResponseCode() + ').');
  return r.getContentText();
}
function latestRelease_() {
  const info = JSON.parse(driveText_(UPDATE_SRC.versionFile));
  if (!info.version || !info.code) throw new Error('No release has been published yet.');
  return info;                                       // { version, notes, code: <Drive file id> }
}
// Daily: look for a newer release; tell you once per version.
function checkForUpdate_() {
  if (!UPDATE_SRC.versionFile) return;
  const P = PropertiesService.getScriptProperties();
  let info; try { info = latestRelease_(); } catch (e) { return; }
  if (!newerVersion_(info.version, FHQ_VERSION)) { P.deleteProperty('UPDATE_AVAILABLE'); return; }
  P.setProperty('UPDATE_AVAILABLE', JSON.stringify({ version: info.version, notes: info.notes || '' }));
  const S = settings_();
  const lastMail = Number(P.getProperty('UPDATE_TOLD_AT') || 0);           // at most one update email a week; the menu shows it right away
  if (P.getProperty('UPDATE_TOLD') !== info.version && Date.now() - lastMail > 6 * 864e5 && yes_(S['Update emails (Yes/No)'] === undefined ? 'Yes' : S['Update emails (Yes/No)'])) {
    try { mail_(alertEmail_(), 'Finance HQ: update ' + info.version + ' is ready',
      'A new version of Finance HQ is ready' + (info.notes ? ':\n\n' + info.notes : '.') +
      '\n\nTo install (about a minute): open your Finance HQ sheet → Finance HQ menu → ⬆ Install update, and follow the 3 steps.'); } catch (e) {}
    P.setProperty('UPDATE_TOLD', info.version); P.setProperty('UPDATE_TOLD_AT', String(Date.now()));
  }
}
// Called by the Install update window: the new code, checked before it's handed over.
function getUpdateCode() {
  const info = latestRelease_();
  if (!newerVersion_(info.version, FHQ_VERSION)) return JSON.stringify({ upToDate: true, version: FHQ_VERSION });
  const code = driveText_(info.code);
  if (code.indexOf("const FHQ_VERSION = '" + info.version + "'") === -1 || code.length < 50000) throw new Error('The update file looks incomplete — try again later.');
  return JSON.stringify({ version: info.version, notes: info.notes || '', code: code, editor: 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/edit' });
}
// Menu: Finance HQ → ⬆ Install update / Check for updates
function installUpdateNow() {
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(UPDATE_HTML).setWidth(520).setHeight(470), 'Update Finance HQ'); } catch (e) {}
}
// The Install update window asks which version is saved right now, so it can tell when the paste worked.
function installedVersion() { return FHQ_VERSION; }
function finishUpdate() { afterUpdate_(); try { rebuild_(true); } catch (e) { log_('Refresh after update: ' + e.message); } return FHQ_VERSION; }
// Runs once after you paste an update and open the sheet / refresh: rebuild with the new version.
function afterUpdate_() {
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty('RUNNING_VERSION') === FHQ_VERSION) return false;
  P.setProperty('RUNNING_VERSION', FHQ_VERSION); P.deleteProperty('UPDATE_AVAILABLE');
  log_('Running Finance HQ ' + FHQ_VERSION + '.');
  return true;
}

const UPDATE_HTML = `<!DOCTYPE html><html><head><base target="_top"><style>
body{font:14px/1.5 Arial,sans-serif;color:#2A2A2A;margin:0;padding:4px 8px}
h3{font-size:13px;letter-spacing:.05em;text-transform:uppercase;color:#fff;background:#2F6F4E;padding:7px 10px;border-radius:6px;margin:12px 0 8px}
.hint{color:#8A857A;font-size:12px}
button{font:15px Arial;padding:11px 20px;border-radius:8px;border:1px solid #2F6F4E;background:#2F6F4E;color:#fff;font-weight:bold;cursor:pointer}
textarea{position:absolute;left:-9999px;top:0;width:10px;height:10px}
#msg,#watch{font-size:13px;margin-top:8px}.ok{color:#2F6F4E;font-weight:bold}.bad{color:#B8423A}
kbd{font:12px Arial;background:#F3F1EC;border:1px solid #D8D3C8;border-radius:4px;padding:1px 6px}
</style></head><body>
<div id="top" class="hint">Checking for the newest version…</div>
<div id="steps" style="display:none">
<h3>1 · Click this</h3>
<button id="cp" onclick="go()">Copy the update &amp; open my script</button><div id="msg"></div>
<h3>2 · In the tab that opens</h3>
<p>Click <b>FinanceHQ.gs</b> (or <b>Code.gs</b>) on the left, then press <kbd>⌘A</kbd> <kbd>⌘V</kbd> <kbd>⌘S</kbd>.<br><span class="hint">On Windows: Ctrl+A, Ctrl+V, Ctrl+S. Only that file — leave any others alone.</span></p>
<h3>3 · Come back here</h3>
<p id="watch" class="hint">This window notices the new version by itself and refreshes your sheet. Your data and settings stay as they are.</p>
</div>
<textarea id="code"></textarea>
<script>
var D=null,ED='',timer=null,started=0;
google.script.run.withSuccessHandler(function(j){D=JSON.parse(j);var t=document.getElementById('top');
  if(D.upToDate){t.className='ok';t.textContent='You already have the newest version ('+D.version+').';return;}
  t.innerHTML='<b>Version '+D.version+'</b>'+(D.notes?' — '+D.notes.replace(/[<>&]/g,''):'');
  document.getElementById('code').value=D.code;ED=D.editor;document.getElementById('steps').style.display='block';})
 .withFailureHandler(function(e){var t=document.getElementById('top');t.className='bad';t.textContent=e.message;}).getUpdateCode();
function go(){var a=document.getElementById('code'),m=document.getElementById('msg'),ok=false;a.focus();a.select();
  try{ok=document.execCommand('copy');}catch(e){}            // copy first, while this window has focus
  var done=function(){m.className='ok';if(!timer){started=Date.now();timer=setInterval(poll,4000);}};
  if(ok){var w=window.open(ED,'_blank');done();m.innerHTML=w?'Copied ✓ — your script is open in a new tab.':'Copied ✓ — now <a href="'+ED+'" target="_blank">open your script</a>.';return;}
  if(navigator.clipboard)navigator.clipboard.writeText(a.value).then(function(){done();m.innerHTML='Copied ✓ — now <a href="'+ED+'" target="_blank">open your script</a>.';},
    function(){m.className='bad';m.textContent='Your browser blocked copying — click the button again.';});
  else{m.className='bad';m.textContent='Your browser blocked copying — click the button again.';}}
function poll(){google.script.run.withSuccessHandler(function(v){var w=document.getElementById('watch');
    if(v===D.version){clearInterval(timer);w.className='ok';w.textContent='Installed ✓ Version '+v+'. Refreshing your sheet — about a minute…';
      google.script.run.withSuccessHandler(function(){w.textContent='All done ✓ You’re on version '+v+'. You can close this window.';})
        .withFailureHandler(function(){w.textContent='Installed ✓ Version '+v+'. Use Finance HQ → Refresh now to finish.';}).finishUpdate();}
    else if(Date.now()-started>150000){w.className='bad';w.textContent='Still on the old version. Make sure you pasted into FinanceHQ.gs (or Code.gs) and pressed ⌘S to save.';}
  }).withFailureHandler(function(){}).installedVersion();}
</script></body></html>`;
