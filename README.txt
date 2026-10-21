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
