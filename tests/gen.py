# Synthetic SimpleFIN data for different pay schedules. 'Today' is passed in; history covers ~5 months.
import json, sys, random, datetime as dt
random.seed(7)
scen, today = sys.argv[1], dt.date.fromisoformat(sys.argv[2])
start = today - dt.timedelta(days=150)
tx = []; n = [0]
def add(d, amt, desc):
    if start <= d <= today:
        n[0] += 1; ts = int(dt.datetime(d.year, d.month, d.day, 12).timestamp())
        tx.append({'id': f'T{n[0]}', 'posted': ts, 'amount': f'{amt:.2f}', 'description': desc, 'payee': desc})
def days():
    d = start
    while d <= today: yield d; d += dt.timedelta(days=1)
pay = []
if scen == 'biweekly':     pay = [(d, 1800, 'ACME CORP DIR DEP') for d in days() if (d - dt.date(2026,1,2)).days % 14 == 0]
if scen == 'weekly':       pay = [(d, 900, 'BRIGHT STAR LOGISTICS') for d in days() if d.weekday() == 4]
if scen == 'semimonthly':
    import calendar
    for d in days():
        last = calendar.monthrange(d.year, d.month)[1]
        if d.day in (15, last): pay.append((d, 1950, 'FAIRFAX COUNTY GOV'))
if scen == 'monthly':      pay = [(d, 3900, 'DELOITTE LLP') for d in days() if d.day == 1]
if scen == 'irregular':
    for d in days():
        if random.random() < 0.12: pay.append((d, random.choice([250, 400, 600, 1200, 180, 900]), random.choice(['STRIPE TRANSFER', 'UPWORK ESCROW', 'FIVERR'])))
if scen == 'student':
    for d in days():
        if d.day == 3: pay.append((d, 400, 'ZELLE FROM MOM'))
if scen == 'stopped':      pay = [(d, 1800, 'ACME CORP DIR DEP') for d in days() if (d - dt.date(2026,1,2)).days % 14 == 0 and d < today - dt.timedelta(days=70)]
for d, a, desc in pay: add(d, a, desc)
for d in days():
    if d.day == 1: add(d, -1400, 'BRIGHTVIEW APTS RENT')
    if d.day == 9: add(d, -85, 'VERIZON WIRELESS')
    if random.random() < 0.5: add(d, -round(random.uniform(8, 60), 2), random.choice(['CHIPOTLE', 'STARBUCKS', 'SHELL OIL', 'GIANT FOOD', 'AMAZON MKTPLACE', 'UBER EATS']))
bal = 2500
out = {'errors': [], 'accounts': [{'id': 'ACT-CHK', 'name': 'Everyday Checking', 'org': {'name': 'Capital One'}, 'balance': f'{bal:.2f}',
        'available-balance': f'{bal:.2f}', 'balance-date': int(dt.datetime(today.year, today.month, today.day).timestamp()), 'transactions': tx}]}
json.dump(out, open(f'/tmp/pay/{scen}.json', 'w'))
# truth: pay landing in the calendar month of 'today'
truth = sum(a for d, a, _ in pay if d.month == today.month and d.year == today.year) if scen not in ('irregular',) else None
full = None
print(json.dumps({'scen': scen, 'paid_so_far': sum(a for d,a,_ in pay if d.month==today.month and d<=today)}))
