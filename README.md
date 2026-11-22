# Finance HQ — setup guide

Finance HQ is a free personal-finance tracker that lives in your own Google Sheet. It links your bank and card accounts, refreshes every morning, and gives you:

- A **Dashboard** sheet, plus a **phone app**, showing what's left this month, spending by category with color-coded limits, and card bills due
- **Statement-period budgeting**: if most of your spending is on credit cards, your "month" runs from statement close to statement close
- **Alerts** by email and as phone notifications (big charges, going over a limit, new subscriptions, bills due soon, checking running low)
- A **cash-flow forecast** for the next 45 days
- A **weekly email digest** every Sunday and a **monthly PDF summary** saved to Google Drive
- A **Subscriptions** finder, a **CardCheck** (points you missed by using the wrong card) and a **Loan planner**

**Cost:** SimpleFIN Bridge, about $15/year (it links your accounts). Everything else is free.
**Privacy:** it can only read your accounts and can never move money. Your money data stays in your own Google account — nobody else, including whoever sent you this, can see it. (Weekly check-ins share a few usage counts only; see *Terms of use & privacy* at the bottom.)

---

## 1. Connect your accounts (SimpleFIN), about 10 minutes

1. Go to **beta-bridge.simplefin.org**, create an account and subscribe.
2. Add your banks and cards. Some take a few minutes to connect.
3. Click **New app connection** and copy the **setup token**, a long string of text.
   - Treat it like a password. Don't email or text it to anyone.
   - It works only once. If setup fails, make a new one.

## 2. Install, about 5 minutes

**Easy way (if someone sent you a Finance HQ link):**

1. Open the link and click **Make a copy**. You get your own private sheet with Finance HQ already inside.
2. Reload the page, then go to **Finance HQ → Connect my bank…** in the top menu and paste your setup token.
   - Google will say "unverified app". It's the sheet's own script, so click **Advanced → Go to Finance HQ → Allow**.
   - It connects and builds everything in 1–3 minutes, then opens Quick setup (step 3).
3. For the phone app, use **Finance HQ → Set up the phone app…**

That's it, and you can skip the rest of this section.

**Manual way:**

1. Create a new, blank **Google Sheet** and name it **Finance HQ**.
2. Open **Extensions → Apps Script**. Delete the sample code, paste the whole **FinanceHQ.gs** file, and save (⌘S / Ctrl+S).
3. Click the **⚙ Project Settings** gear, scroll to **Script properties** and add a property:
   - Property: `SIMPLEFIN_SETUP_TOKEN`
   - Value: your setup token
4. Go back to the editor, choose **setupFinanceHQ** in the function menu and click **Run**.
   - Google will say "unverified app". This is your own script, so click **Advanced → Go to Finance HQ → Allow**.
   - Setup takes 1–3 minutes. You'll get a test weekly-digest email at the end.

## 3. Make it yours, about 3 minutes

Open your sheet. In the top menu there's now a **Finance HQ** menu. Choose **Finance HQ → Quick setup…** and fill in one short form:

1. **Your cards:** the day each statement closes and the day it's due, both on your statement. The due day fills in an estimate once you type the closing day. Your card's rewards are matched automatically; just check the match.
2. **Savings goal:** the % of pay to save first (20% is common), plus optional giving and tax set-aside.
3. **Budget and alerts:** your dining budget, an alert email, and whether you want phone reminders and the Sunday digest.
4. **Debts:** student loans, car loan, mortgage and so on. Debts it spotted in your accounts or payments are already listed, so you only add the balance and rate.
5. **Trips (optional):** trip dates, so travel spending is kept separate from your everyday budget.
6. **Yearly costs (optional):** bills that come once a year, like an annual card fee, car insurance paid yearly or Amazon Prime. Finance HQ sets aside 1/12 of them every period.

Everything is optional. Click **Save**, and the Dashboard and Debt plan rebuild in about a minute. Every other spending limit is suggested from your own spending, so there's nothing else you have to do.

**Savings goals and checklist.** Set "Save this % of pay" in Quick setup (20% is a common target). Each period, the Dashboard's **Transfers to make** box tells you how much to move. It ticks itself once your savings have grown by that amount, and you can tick or untick it yourself.

**Left to spend.** The Dashboard and phone show what's yours to spend this period: the pay you expect (including paychecks still coming), minus your savings goal and yearly costs, minus what you've spent.

**Budget tab.** Your limit for every category next to what you spent in the last two periods and this one, with a status (On track, Watch, Over). The **Suggested** column shows limits that fit inside what's left after saving, trimming non-essentials first. Tick **Use it?** to take a suggestion, or type your own number in the yellow column. Changes go straight to Settings and the Dashboard.

**Insights tab.** Plain-English notes in color: whether you're on pace, what's over its limit, what still needs moving to savings, bills due soon, price increases, and a review of last period.

**Year in review.** Each January a PDF of last year lands in your **Finance HQ Archive** folder (and your email, if the monthly insights email is on). Run **Finance HQ → Year in review (PDF)** any time for this year so far.

**Payday nudge.** When a paycheck lands you get a phone reminder (or an email): how much to move to savings from that check, and what's left to spend per day. Turn it off in Settings → "Payday nudge".

**Savings tracker.** In the **Subscriptions** tab, set each subscription's Status to *keep*, *could cancel* or *cancelled*. The Dashboard then shows how much you could save a year and how much you've already saved. If something you marked cancelled keeps charging you, Health warns you.

**Other handy things:**

- **Changes show up immediately.** Edit anything (Settings, debts, trips, rules or a category) and the Dashboard rebuilds itself in a few seconds. A small "Updating your dashboard…" note appears in the corner.
- **Finance HQ → Refresh now** pulls new transactions from your bank right away instead of waiting for 6 AM.
- **Paychecks are found automatically.** If your paycheck doesn't say "payroll", Finance HQ spots the same employer depositing a similar amount every 1–5 weeks and labels it Income. You'll see the rule it added in the Rules tab, marked "auto".
- **Forecast tab** answers "Will I have enough in checking?" in one sentence. Below that it shows a week-by-week view and every upcoming paycheck and card bill.
- **Health tab** (or **Finance HQ → Health check**) checks everything: the bank connection, last sync, broken accounts, missing card dates, card payments counted as spending, duplicates, uncategorized spending, missing paychecks, incomplete debts and bad dates. Each problem comes with how to fix it, and you're emailed if something breaks.
- **Feedback?** Use **Finance HQ → Send feedback**. You'll also get a short form on day 7 and day 28.
- **Wrong category?** Pick the right one from the dropdown in the **Transactions** tab, or tap the transaction in the phone app. It stays locked after that. On the phone you can also tick "Always use this for…" to make a rule for that store.
- The **Start here** tab explains every tab in plain English. If you want to fine-tune later, use **Settings** (the yellow cells) and **Rules**.

## 4. Phone app (optional), about 2 minutes

1. In Apps Script, go to **Deploy → New deployment → ⚙ → Web app**.
2. Set **Execute as: Me** and **Who has access: Only myself**, then click **Deploy**.
3. Open the URL on your phone. In Safari, use **Share → Add to Home Screen**.
4. After any code update, go to **Deploy → Manage deployments → ✏️ → New version → Deploy**. This keeps the same link.

## What runs automatically

| When | What |
|---|---|
| Daily ~6 AM | Sync accounts |
| Daily ~7 AM | Rebuild the Dashboard |
| Daily ~8 AM | Alerts and forecast |
| Sundays 9 AM | Weekly digest email |
| Day after your statement closes (or the 1st) | Monthly PDF in Drive → **Finance HQ Archive**, plus an insights email |
| January 2nd | Last year's **Year in Review** PDF in the same folder |

**Phone notifications:** alerts and bill reminders are added to your Google Calendar as events with a pop-up. Turn on Google Calendar notifications on your phone to get them.

## Troubleshooting

Start with the **Health** tab. It checks most of what can go wrong and tells you how to fix it.


- **A setup step failed:** open **Executions** in Apps Script to see the error. The most common cause is a setup token that was already used. Make a new one and replace the property.
- **Purchases have the wrong category:** add a Rule with **mine** in Notes, then run **phase2Daily**.
- **A bank is missing:** check that it's connected in SimpleFIN. The next sync will pick it up.
- **Only about 90 days of history:** that's how far back SimpleFIN goes. From here on, history keeps building up in the Transactions tab.

## Terms of use & privacy

- **Your money data stays with you.** Transactions, balances and everything else live in this Google Sheet in your own Google account. Finance HQ can only *read* your accounts through your own SimpleFIN subscription; it can never move money.
- **Weekly check-ins.** While Finance HQ is being tested, your copy sends the person who shared it with you a few usage **counts** once a week, so they can see whether it's working for you:
  - your Google account email
  - which week since install, and whether your bank synced in the last 2 days
  - how many accounts are linked, and how many new transactions came in that week
  - how many days you opened the phone app and edited the sheet
  - how many categories and limits you changed
  - whether transfers were ticked, card dates are set, a savings goal is set, and yearly costs are listed
  - how many Health problems there are, and the version of Finance HQ
- **Never sent:** amounts, balances, account or card names, store names, or any transaction.
- **See it or switch it off any time.** **Finance HQ → What check-ins share** shows exactly what your copy sends. To stop it, set **Settings → "Share usage check-ins with the Finance HQ creator (Yes/No)"** to **No**.
- **Updates.** Once a day your copy checks whether the person who shared it has published a newer version, and emails you if so. Nothing installs on its own. **Finance HQ → ⬆ Install update** copies the new code for you to paste over the old Finance HQ file, which takes about a minute. Your data and settings stay. To stop the emails, set **Settings → "Update emails"** to **No**.
- Finance HQ is provided as-is for personal use and is not financial advice.


## Sharing Finance HQ with friends (for the person sharing it)

1. Make a **new, blank Google Sheet** called **Finance HQ**. Don't use your own sheet.
2. Go to **Extensions → Apps Script** and paste FinanceHQ.gs, including your `CHECKIN` line if you use check-ins. Save it.
3. Choose **prepareTemplate** in the function menu and click **Run**. It leaves just a **Welcome** tab and clears any saved settings.
4. Click **Share → General access → Anyone with the link → Viewer**. Then send friends the link with the ending changed from `/edit…` to **`/copy`**. The Execution log shows this link.

Each friend gets their own separate copy. Their bank token and data stay in their copy only, and a copy can never reuse your bank connection.


## Updates (for the person sharing it)

Every copy checks your release files once a day. When there's something newer, the person gets one email, an **⬆ Install update** menu item, and a line on the Health tab. The update window copies the new code, links to their script, and walks them through paste, save and refresh. The phone app uses the **Test deployments** link (ending in /dev), which always runs the latest code, so nobody redeploys.

Installs can't happen silently: Google doesn't let a sheet's script rewrite itself unless each person sets up their own Google Cloud project.

**To publish an update:**
1. Paste the new FinanceHQ.gs into your own project. It must have a higher `FHQ_VERSION`.
2. Edit `RELEASE_NOTES` in your Personal file.
3. Run **publishRelease**. This needs the **Drive API** service added once: in the editor, click **+** next to **Services** and add **Drive API**.
