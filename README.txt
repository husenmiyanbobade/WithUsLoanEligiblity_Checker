WITH US Society Register - PWA

FILES
  index.html, manifest.webmanifest, sw.js, icon-*.png  -> upload all together to any HTTPS static host
  apps-script/Code.gs                                   -> shared data store (Google Apps Script)

HOSTING (HTTPS is required for install, offline and copy buttons)
  Netlify Drop, GitHub Pages, Cloudflare Pages or Firebase Hosting all work.
  Open the hosted link on a phone > browser menu > "Add to Home screen" / "Install app".

MUST DO ONCE: UPDATE THE GOOGLE SCRIPT (needed for the admin login and the Monthly tab)
  1. Open your Apps Script project at script.google.com.
  2. Replace everything with the contents of apps-script/Code.gs.
  3. The login is set at the top of Code.gs:  ADMIN_USER = 'withusadmin'  and  ADMIN_PASS = 'withusadmin'.
     Change them there any time (then deploy a new version). They are checked only by the script, never inside index.html.
  4. Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.
     The Web app URL stays the same; index.html already points to it.
  Until this is done the login screen says "The Google script needs the new Code.gs".

ADMIN LOGIN
  - The whole app is locked behind the login. Without signing in nothing is shown, and the script refuses to send data.
  - "Remember me" ON (default): the phone stays signed in until you tap Logout (menu button with three dots > Logout).
    OFF: signed in only until the browser is closed.
  - After 5 wrong tries the login pauses for 15 minutes on that phone only; other phones are not affected.
  - Logout signs that phone out, removes the saved data from it, and the script forgets the session.
  - To sign every phone out at once: Apps Script > Project Settings > Script properties > delete "sessions".
  - Old passcode (CHANGE-THIS-PASSCODE) is no longer used.

MONTHLY CONTRIBUTION TAB (Rs 500 per member)
  - Search a member and tap "Mark Rs 500" when they inform you. Tap "Rs 500 paid" to undo; an Undo bar also appears.
  - "Generate list" builds the October list in the usual WhatsApp format; "Unpaid list" builds the pending list.
    Both have Copy and Share to WhatsApp buttons.
  - Marks are saved in a separate Drive file (withus-contributions.json), so re-uploading the Excel never erases them.
  - Marking needs the admin login (you are already signed in).
  - A new month starts automatically on the 1st: everyone unpaid. Earlier months stay in "All months";
    months already in the Excel file (Jan-Sep) show as paid from the Excel columns.
  - "Import paid members from a WhatsApp list": paste an existing list; every line with 500 is marked paid.

MONTHLY LOAN REPAYMENT (inside the Monthly tab: switch "Loan repayment")
  - Pick a month: every loan due in it is listed (from the Excel loan sheet), all Pending.
  - Tap "Mark received" when a member informs you. Undo is one tap. The date received is saved quietly.
  - "All loans" / "Pending only" generate the WhatsApp list in your format, with Copy and Share buttons.
  - Marks are saved in their own Drive file (withus-repayments.json). The Excel data is never changed.
  - "Import from a WhatsApp list": paste the current list; loans marked Received are picked up.
  - Return amount = loan + fixed charge (5,000+200, 10,000+500, 25,000+1,000, 50,000+2,000, 1,00,000+4,000).

MORE CHANGES IN THIS VERSION
  - Monthly tab: Clear button next to Import; "All months" is collapsed (tap to expand, closes after you pick a month);
    when the app language is Marathi, the Pending lists (unpaid contributions, pending loans) are generated in Marathi.
    The full lists stay English so they can be imported back.
  - Overview: Fines accrued and Repayment rate cards removed.
  - Deposits: tap "Collected this year" or "New members" to see the member-wise details.
  - The top header stays fixed while scrolling on phones.

MARATHI LISTS
  - With the language set to Marathi, the Unpaid contribution list and the Pending loan list are written in Marathi,
    in the format you posted. Names come from your lists. A name never seen before is transliterated automatically,
    so please check new names the first time they appear.

UPDATING
  After changing index.html, bump CACHE ('withus-v12' -> 'withus-v13') in sw.js so installed apps refresh.


LOAN APPLICATION LIST (new, Eligibility tab -> Applications)
- Check a member as usual, then tap "Add to application list" (or "Save as rejected application" when the applicant is not eligible).
- Applicant and guarantor are judged separately; remarks are written in Marathi and can be edited before saving.
- Only accepted applications appear in the generated list. Rejected ones are kept separately, with a "Copy message" for the member.
- A warning appears if the member is already in this month's list.
- Saved in a new Drive file, withus-applications.json, via the Google script. Excel is never changed.
- IMPORTANT: re-paste apps-script/Code.gs into your Apps Script project and Deploy > Manage deployments > New version
  (keep ADMIN_USER / ADMIN_PASS as you set them). The app cache is now withus-v13.


BRIGHT TILES LOOK (cache withus-v16)
- Cream and deep green with colour tiles, floating bottom bar, "Welcome" greeting in the phone header. Light and dark themes.
- Monthly (Contribution and Loan repayment) is compact for phones: smaller rows, buttons and tiles, list flows with the page.
- The "Undo" bar after marking now disappears by itself after 5 seconds (the mark stays).
- Styling only: no change to data, login, lists or Code.gs.


LOAN RECEIVED FREES THE BORROWER AND GUARANTORS (cache withus-v17)
- When a loan is marked "Mark received" in the Monthly > Loan repayment list, the Eligibility check treats it as closed:
  the guarantor(s) are free at once (the result says "...loan was marked Received on <date>, so the guarantor is free"),
  and the borrower no longer has an open loan.
- Rule book: the borrower cannot take a new loan in the same month the loan was repaid. The check shows
  "A loan was repaid this month. A new loan can be taken from next month" until the 1st of next month.
- Undo on the repayment list ties the guarantors to the loan again. The latest mark always wins. Excel is never changed.
- (v18) The same-month rule also covers loans already shown as repaid in the Excel file: if its Loan Repayment Date is in the
  current month, the member cannot take a new loan until next month.

v19 - Calm banking look (look only; Excel and Code.gs unchanged)
- Deep green header with a Welcome greeting and the current tab name, soft grey page, white cards, thin borders.
- Quick links on Home (Deposits, Eligibility, Lists, Monthly).
- Docked white bottom bar with an indicator on the active tab.
- Deep green dark mode (moon/sun button).
- Opening logo animation (about 2.5 s). Tap to skip. Off when the phone has "reduce motion" on.
- Upload index.html, sw.js and manifest.webmanifest over the old ones.

v20 - Faint tinted cards and Lists moved into Loans (look and layout only; Excel and Code.gs unchanged)
- Cards use very faint tints with a thin hairline: green = money in, blue = loans out, amber = pending/in process.
- The Lists tab is now inside the Loans tab: Check | Applications | Lists. The tab is named "Loans" (Marathi: कर्ज).
- Bottom bar has 4 tabs: Home, Deposits, Loans, Monthly. Home quick links: Deposits, Loans, Monthly.
- Upload index.html and sw.js over the old ones.

v21 - Disbursed loans list, floating Undo bar, Deposits cards (Code.gs CHANGED: redeploy it)
- Loans > Applications: "Mark disbursed" on each accepted application (date defaults to today; return date and return amount are worked out). Disbursed loans cannot be edited, only removed (with a confirm and a 5-second Undo).
- "Loan details list" button makes the monthly message in your format (Status: Loan Amount Disbursed).
- Disbursed loans count in the eligibility checks straight away. Excel stays the master: when your Excel shows the same loan, the app copy is ignored.
- Undo bar (Contribution, Loan repayment, Disbursed) now floats above the tab bar, so the list never moves. It has a countdown line and ignores taps for the first 0.4 s.
- Deposits: the four cards and the search-result cards are two by two, in colour.
- Apps Script: open Code.gs in the Google script, replace it with apps-script/Code.gs, then Deploy > Manage deployments > Edit > New version > Deploy.
- Upload index.html and sw.js over the old ones.

v22
- Loan repayment: a received loan no longer has a big Undo button. It shows "Received on <date>" and a small red Remove link. Remove asks for confirmation (the loan goes back to Pending), then the floating Undo bar appears for 5 seconds.
- Mark received ignores a second tap within 1 second.
- Bottom bar: the selected tab now has a soft green pill behind a solid (filled) icon, bold label. Works in light and dark mode.
- Upload only index.html and sw.js. Code.gs is unchanged from v21.

v23
- Loan details list and Mark disbursed sheet: Loan Taken Date is the actual disbursement date. Loan Return Date is always the 15th of the repayment month (3, 4 or 5 months later depending on the amount).
- Upload only index.html and sw.js.

=====================================================================
v24 - FOUNDATION OF THE ADMIN MODULE (database, two admin logins)
=====================================================================
What changed
- Two admin logins instead of one shared login:
    withusadmin1  (owner, full access)     withusadmin2  (daily work)
  Passwords are stored only as salted hashes in the database Sheet, never in the app.
- withusadmin1 sets withusadmin2's password. withusadmin2 cannot change any password.
- A private Google Sheet is now the society database (members, contributions, loans, ledger,
  penalties, expenses, audit log and more). The Excel file stays the master for now; the Sheet
  is filled from it and checked against it. The rest of the app works exactly as before.
- Menu (three dots) > Admin accounts (withusadmin1 only): manage withusadmin2, import the Excel
  data into the Sheet, check totals, export everything to Excel (backup).

ONE-TIME SETUP (about 10 minutes)
1. script.google.com > your project > replace the code with apps-script/Code.gs.
2. Deploy > Manage deployments > pencil icon > Version "New version" > Deploy.
   (The Web app URL stays the same. Google will ask you to allow Sheets/Drive access: allow.)
3. Pick the function  setup  in the editor's function list and press Run.
   Then View > Logs. It prints the link of the new Sheet and a TEMPORARY password for withusadmin1.
4. Upload the new app files (index.html, sw.js, ...) as usual. Open the app, sign in as
   withusadmin1 with the temporary password. The app makes you choose your own password.
5. Menu > Admin accounts > Set new password for withusadmin2. Tell them the password yourself.
6. Same screen > Database > Import from Excel data, then read the totals check. Every line must say OK.
7. Back in the script editor run  installNightlyBackup  once: a daily copy of the Sheet is saved in the
   Drive folder "WITH US backups" (last 30 kept).
Until step 3 is done, the old shared login (withusadmin / withusadmin) keeps working, so nothing breaks
while you switch over. After step 3 it stops working.

If withusadmin1 forgets the password
- The owner of the Google account runs the function  resetAdmin1  in the script editor and reads the
  new temporary password in View > Logs.

Notes
- Everyone signs in again once after this update (old sessions are not carried over).
- Exports and the audit log never contain passwords or hashes.
- The import copies deposits, monthly contributions, loans and the fine pot. Loan payments and
  disbursements are recorded in the Sheet from go-live onwards (next build steps).


=== v25: Members and contributions (Step 2) ===
What is new
 - A new "Members" tab (second tab). It appears only after the database is set up.
 - Members list with search and filters (owes, not paid this month, paid up).
 - Collect month: tap "Mark paid" next to a name. It uses the last payment mode (Cash, Bank or UPI). "Undo" stays for 8 seconds.
 - Add member (name, mobile, joining date, date of birth, joining fee, first month now or later).
 - Record contribution for one or many months, with a receipt you can send on WhatsApp, share or copy.
 - Member profile: deposits, what is owed, loans and guarantees, receipts, edit, suspend/reactivate (withusadmin1 only).
 - Void: withusadmin1 any time; withusadmin2 only their own entry within 2 minutes. Two taps to confirm.
 - Birthday strip on the list when a member's birthday is today.

One-time steps
 1. Open the Apps Script project, replace Code.gs with the new apps-script/Code.gs, Save.
 2. Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy. (The web app address stays the same.)
 3. The Sheet columns upgrade by themselves on the first request. Nothing to do.
 4. Open the app, sign in as withusadmin1. The Members tab appears.

Important
 - Entries made in the Members tab are real and are saved to the database Sheet, not to the Excel file.
 - While you are only testing, you may empty all data and import the Excel again: in the script editor run
   wipeAllData('YES-WIPE-ALL-DATA')   then in the app: Admin accounts > Import from Excel data.
   Once real entries exist, do not wipe.


=== v26: Speed fix + Penalties and eligibility block (Step 3) ===
Speed
 - The server used to re-open the Google Sheet about 7 times for every Members request. It now opens it once and reads each tab once.
 - The app shows the last member list instantly (saved on the phone) and refreshes it in the background.
 - "Mark paid" and "Undo" in Collect month change the row at once; the server confirms in the background.
Penalties (Members tab > Penalties)
 - After a month ends, open Penalties, tap the month, review the names, and post. Each unpaid member gets a Rs 500 penalty (Settings: PenaltyAmount).
 - On a member profile: Penalties card, "Pay penalty" (part payment allowed, receipt), Void (own entry within 2 minutes, or withusadmin1), Waive (withusadmin1 only, reason required).
 - Penalty money goes to the fine pot.
Eligibility block
 - A member with unpaid past months or unpaid penalties cannot apply or be a guarantor. The profile names exactly what is owed.
 - The Eligibility tab check now has a "Contributions and penalties paid" line for the applicant and each guarantor (when the database is in use).
Update steps
 1. Replace Code.gs in Apps Script with apps-script/Code.gs, Save.
 2. Deploy > Manage deployments > edit > New version > Deploy.
 3. Open the app, close and reopen it once so the new version loads.


=== v27: Loans on the ledger (Step 4) + speed + navigation ===
Speed
 - The server now keeps recent answers (member list, month view, profile, loan book) for a few minutes. Any save clears them at once.
   If you edit the Google Sheet by hand, changes can take up to 10 minutes to show in the app.
 - The app refreshes the member list less often and after the first screen has loaded.
Navigation
 - The four cards from the old Deposits tab (total members, deposit balance, collected this year, new members) are now on Home.
 - Once the database is in use the Deposits tab is hidden: Members replaces it. Search a member in Members to see everything: deposits, what is owed,
   penalties, loans (guarantors, dates, amounts, paid, balance, status), guarantees, receipts.
Loans (Loans tab > Loan book, and on every member profile)
 - Loan book: principal out, available to lend, to collect; filters Defaulted / Overdue / Active; search; tap a loan.
 - Record payment: full or part, from borrower or guarantor. Goes to principal first, then extra charge, then late extra. Extra and late extra go to the fine pot. Receipt on WhatsApp/Share/Copy.
 - Void a payment: withusadmin1 any time; withusadmin2 only their own within 2 minutes. Two taps.
 - Overdue: from the 16th of the repayment month the loan is Overdue, late extra (a second copy of the extra charge) is added and the borrower's 6-month loan suspension is recorded.
   This is applied automatically on the first request of each day.
 - Defaulted (from the 1st of the next month): withusadmin1 can "Settle from deposits": borrower's deposit first, then guarantors equally, with a preview and a two-tap confirm.
 - withusadmin1 can waive late extra or lift a suspension, each with a reason.
 - New loan: choose borrower, amount (10,000 / 25,000 / 50,000 / 1,00,000), guarantors (1 for 25,000 and 50,000, 2 for 1,00,000). The app checks eligibility, guarantors, open loans and available funds.
Update steps
 1. Replace Code.gs in Apps Script with apps-script/Code.gs, Save.
 2. Deploy > Manage deployments > edit > New version > Deploy.
 3. Open the app and close/reopen it once.


=== v28: Fine pot and expenses (Step 5) ===
 - Home: once the database is in use, the old Excel "Fine details" card is replaced by a "Fine pot" card: balance, money in and out this financial year (April to March).
 - Open fine pot: Statement (every line with running balance: loan extra, late extra, penalties, joining fee, expenses) and Monthly (money in by source, money out by category, closing balance).
 - Add expense (withusadmin1 only): date, category (Bank charges, Stationery, Printing, Meeting, Travel, Donation, Other), amount, note. An expense cannot be more than the balance.
 - Remove expense (withusadmin1 only, reason required): it is voided, not deleted, and the balance goes back up.
 - withusadmin2 can see the fine pot but cannot add or remove expenses.
 - The 13 spending items imported from Excel appear under "Imported spending"; the Excel fine total appears as "Opening balance (imported)" on the import date.
 - To change the category list, add a Settings row in the Sheet: Key = ExpenseCategories, Value = comma separated names.
Update steps: replace Code.gs, Deploy > Manage deployments > edit > New version > Deploy, then close and reopen the app.


=== v29: Fixes after live testing ===
 - Monthly > Contribution now works on the database once it is live. A mark made there appears in Members and the member profile at once, and the reverse. Paid by (Cash / Bank / UPI) chips added. Payments imported from Excel show as paid and cannot be removed there; app payments can be removed (confirm first).
 - Members > "Collect month" removed (Monthly does the same job). "Import from WhatsApp list" hidden once the database is live.
 - Home: cards in an even 2 x 2 grid, quick links include Members, footer line removed. Fine pot card: no more cut-off amounts, with the latest entries shown.
 - Loan book: tabs fit on one row, compact summary tiles, last book saved on the phone (shows instantly) and loaded in the background right after sign-in.
 - Lists (eligible members, guarantors, application messages): "Share to WhatsApp" added next to Copy.
 - Code.gs: new light call monthsSummary. REDEPLOY: paste Code.gs, Deploy > Manage deployments > Edit > New version > Deploy.
 - Upload index.html and sw.js (cache withus-v29).


=== v30: Monthly > Loan repayment on the database ===
 - Once the database is live, the Loan repayment list shows the database loans due in the chosen month (Excel loans included). Paying here and paying in Loans are now the same thing: one ledger, one receipt series, the fine pot and deposit pot update, guarantors are freed when the balance reaches zero.
 - "Mark received Rs X" pays the full balance (borrower, today's date, last payment mode). A floating Undo bar stays 5 seconds.
 - "Part payment" opens the payment sheet (amount, date, who paid, mode, where the money goes) and gives a receipt.
 - Each payment has its own Remove link (confirm first). Admin2 can remove only own entries within 2 minutes; admin1 any time.
 - Overdue cards show the owed amount with the doubled extra and the days left. Closed-early loans say "Closed early on <date>".
 - "Generate list" and "Pending list" (English and Marathi) are built from the database. Pending list shows the balance still to pay.
 - Old Excel-era "received" ticks saved before this version are not copied. Re-enter any real October payments with Part payment / Mark received (use the real date).
 - Code.gs: new call loansDue. REDEPLOY (Deploy > Manage deployments > Edit > New version). Upload index.html and sw.js (cache withus-v30).


=== v31: Speed, look, bank check, month close, exit, reports ===
Speed
 - The app page is now about half the size (455 KB, was 865 KB): the logo is a small file instead of being embedded four times, and the two big scripts (Excel library, charts) no longer block start-up. The unused chart script is removed; the Excel library loads quietly in the background.
 - The app opens at once from the saved copy and fetches the newest copy in the background. When a new version is ready a bar "A new version is ready - Reload" appears.
 - After sign-in, everything for Home (members, fine pot, dashboard) is requested together instead of one after another. Monthly and Loans data load in the background and are kept on the phone, so those tabs open with the last data showing.
 - Google Drive is no longer opened for the old Excel-era "has anything changed" checks (the server remembers the version). The old contribution/repayment marks are not downloaded once the database is live.
 - Opening a member shows their name, deposit and dues at once; the rest fills in a moment later. A member opened before is shown from memory while it refreshes.
Look
 - Home: new "Society at a glance" card (deposit pot, available to lend, loans out, fine pot) with a "Needs attention" list (members behind, overdue loans, loans due this month), then the Reports and Bank check buttons. The old Excel figures sit below, labelled "From the Excel records".
 - Removed the empty space at the bottom of every tab. Header title on one line. Softer "Due this month" tag. Smooth sheet fade, tap feedback, loading placeholders.
 - Members tab: search on top with a "+" button to add a member, a one-line summary, one scrolling row of filters with Penalties at the end. The Record contribution button is gone (use Monthly, or Record inside the member).
 - Fine pot: imported spending now shows the real names from Excel (Jackets, T-shirt, Trip...). Home no longer lists latest entries.
Admin features (withusadmin1 acts, withusadmin2 can only look)
 - Bank check: enter bank balance and cash; the app shows the expected total (deposit pot + fine pot - loans out) and the difference. Saved as a record.
 - Month close: after a bank check dated in or after that month, and with penalties for that month decided, a month can be locked. Entries and removals in a closed month are refused; fix with a new entry in an open month.
 - Exit settlement (inside a member): blocked if the member has an open loan or is a guarantor on one. Refund = deposit - unpaid penalties - unpaid months. Joining fee, paid penalties and loan charges are not refunded. Penalties deducted go to the fine pot. Two-tap confirm.
 - Reports (both admins): monthly collection, arrears, penalties, member statement, member list, loan book, loan payments, guarantor exposure, settlements, fine pot, expenses, deposit pot, bank checks, financial year summary, audit log. Show on screen, export to Excel or PDF.
Update steps: paste Code.gs, Deploy > Manage deployments > Edit > New version > Deploy. Upload index.html, sw.js, logo-128.png (all new/changed). Open the app twice (first open loads the new files).

==== v32 (what is new) ====
1. More tab (5th tab): Reports, Bank check & month close, Audit log, Settings, Admin accounts (owner), Language, Dark/light, Sign out.
2. Settings screen (More > Settings): the owner (admin 1) can change monthly contribution, joining fee, penalty per missed month, society UPI ID and expense categories. Other rules are shown as fixed. Every change is written to the Audit log. A new contribution amount applies to every month, so change it only at the start of a financial year.
3. Clean-up import (Admin accounts > Clean-up import): choose the latest Excel file. It adds the closed accounts as Exited members (history only), removes blank loan rows, links loans to members where the name matches, and lists spelling variants that need YOUR approval one by one (or leave as they are). Run once.
4. Home shows birthdays today.
DEPLOY: paste apps-script/Code.gs into the Apps Script editor, Deploy > Manage deployments > new version; upload index.html, sw.js, logo-128.png. Reopen the app twice (first open shows the update bar).

==== v33 (performance release) ====
Mobile performance run (Android mid-range 4x CPU slowdown, Android low-end 6x slowdown on 3G-like network, iPhone-class 2x slowdown).
- Monthly tab opens about 8x lighter on low-end Android (screen freeze 1000 ms down to 127 ms): rows off-screen are no longer laid out until scrolled to.
- Repeated name matching is cached; the Monthly list is not redrawn when nothing changed.
- Reports show the first 60 rows quickly, with a "Show all rows" button. Excel and PDF always contain every row.
- Saving the data copy on the phone no longer reads it back.
IMPORTANT FOR SPEED: make sure your host serves index.html compressed (gzip). GitHub Pages and Netlify do this automatically; it cuts the first download from 560 KB to about 125 KB.
DEPLOY: upload index.html and sw.js (Code.gs is unchanged from v32).

==== v34 ====
1. Reports > Loan payments: it was empty because loans closed in the Excel history have no individual payment rows. They are now listed once each (on the closing or due date, marked "Excel history"), together with every payment recorded in the app. Date filters work on both.
2. Payment mode: Cash is removed everywhere (contribution, penalty, loan repayment, new member's first contribution). Only UPI and Bank are accepted, UPI is the default, and the server refuses anything else. Old Cash receipts stay readable.
3. Contributions from a WhatsApp list: Monthly tab > "Contributions from a WhatsApp list" (also More menu). Choose month, date, UPI or Bank, paste the list. Lines like "1. NAME - 500" or "NAME - 500/✅" count as paid; 🚫 lines, headings and unclear lines are ignored. Exact names (any word order) are matched; spelling differences are shown as choices for you to pick; unknown names and members already recorded are listed and skipped. Nothing is saved until you press Record. One receipt per member.
DEPLOY: paste apps-script/Code.gs (new deployment version), upload index.html and sw.js.

==== v35 (go-live release) ====
1. Closed accounts: every row of "Closed Account Details" is imported as its own closed account (34 rows = 34 closed accounts; people who left and joined again have a separate record per membership). Amounts and remarks are kept in a new "ClosedAccounts" tab for history; they do not touch the pots. Loans link to the live member first, to the closed account if the name has exactly one, otherwise you choose.
2. More > Check against Excel: choose the latest Excel; the app lists every difference (members, deposits, month-by-month payments, loans only in one place, loan status, totals). Nothing is changed.
3. More > Go live (owner only): checklist (Excel check clean in last 36 h, bank check with no difference in last 3 days, backup in last 24 h, closed accounts imported, no blank loan rows). When all are ticked, type GO LIVE and press the button. From then the Excel import and clean-up are locked for good.
DEPLOY: paste apps-script/Code.gs (new deployment version; the Sheet gets the new ClosedAccounts tab automatically on the next sign-in), upload index.html and sw.js.

V13 (v36, 24 Oct 2026): from the Android screen recording
1. More is now a real tab (page) with the tab bar visible and More highlighted/filled. Language, Dark/light and Sign out removed from it (they are in the header; Sign out there now asks to confirm).
2. "Contributions from a WhatsApp list" removed from Monthly; it stays in More.
3. Bank check: Cash in hand box removed (bank balance only).
4. Loans > Applications: new "Not disbursed" filter; numbering removed from cards; compact layout, Mark disbursed on top.
5. Blank space above the tab bar removed on all tabs.
6. "From the Excel records" moved off Home into More > Old Excel records.
7. Report date fields no longer cut off on narrow phones.
Redeploy is NOT needed for Code.gs (no server change); upload index.html and sw.js.

V14 (v37, 25 Oct 2026)
1. Reports now show as cards on a phone (Cards / Table switch at the top of every report; Excel and PDF unchanged).
2. Loan payments report has a Borrower column.
3. Server (Code.gs, REDEPLOY): overdue-loan updates are written in one batch (first request after the 16th of a month was slow).
4. The new master Excel holds the 3 loans dated 10/10/2026 (L-1017 to L-1019, In Process 78) and the spelling MOHAMMADALI GANI NACHARE. They come in by re-importing the Excel: menu > Choose file > Process file, then Admin accounts > Database > Replace import, then More > Check against Excel.

V15 (v38) - 25 Oct 2026
  Check totals / Check against Excel: closed accounts are counted on their own line, not as members; October entries made in the app
  are shown as a note instead of a difference; blank "Unnamed applicant" loan rows are no longer counted.
  Clean-up import now also: removes the blank loan rows, adds Excel loans the Sheet does not have (3 dated 10/10/2026), corrects a
  changed name (MOHAMMADALI GANI NACHARE), and adopts closed accounts imported by earlier builds. Safe to run twice.
  The Excel reader no longer keeps numbered rows that are empty.
  REDEPLOY Code.gs (new version) and upload index.html + sw.js.

V16 (v39) - 26 Oct 2026
  SPEED: the script asked Google for the same stored settings (Sheet id, data version) thousands of times in one request, which made
  Bank check & month close, Check against Excel and Clean-up import take minutes or fail with "Check the connection". Each is now read once
  per request (3 to 15 reads instead of about 1,300).
  Loans > Applications, Monthly > Contribution and Loan repayment: the Generate list buttons sit above the member list, and the long list
  is folded behind "Show list" (opens by itself when you search or pick a filter; remembered while the app stays open).
  REDEPLOY Code.gs (new version) and upload index.html + sw.js.

V17 (v40) - 26 Oct 2026
  Back from Check against Excel / Reports / Bank check etc. now returns to the More tab (the old menu sheet with Sign out and Language is gone).
  Home: Members / Loans / Monthly shortcut icons removed.
  Fine pot cards renamed "Fine collected" and "Fine spent" with the financial year under the name.
  Home > Needs attention > Loans due this month opens Monthly > Loan repayment with the list shown.
  Bank check & month close: the screen's data is fetched in the background when you open Home or More and kept by the script for a few
  minutes, so it opens at once; opening it still refreshes the numbers.
  REDEPLOY Code.gs (new version) and upload index.html + sw.js.
