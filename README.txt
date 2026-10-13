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


BRIGHT TILES LOOK (cache withus-v14)
- New cream/green look with colour tiles, floating bottom bar, "Welcome" greeting in the phone header (shows the tab name on other tabs).
- Styling only: no change to data, login, lists or Code.gs. Light and dark themes both updated.
