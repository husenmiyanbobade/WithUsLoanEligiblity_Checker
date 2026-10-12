WITH US Society Register - PWA

FILES
  index.html, manifest.webmanifest, sw.js, icon-*.png  -> upload all together to any HTTPS static host
  apps-script/Code.gs                                   -> shared data store (Google Apps Script)

HOSTING (HTTPS is required for install, offline and copy buttons)
  Netlify Drop, GitHub Pages, Cloudflare Pages or Firebase Hosting all work.
  Open the hosted link on a phone > browser menu > "Add to Home screen" / "Install app".

MUST DO ONCE: UPDATE THE GOOGLE SCRIPT (needed for the Monthly tab, including loan repayment)
  1. Open your Apps Script project at script.google.com.
  2. Replace everything with the contents of apps-script/Code.gs.
  3. On the line  var PASSCODE = '...';  type YOUR existing passcode (the same one you use for uploads).
  4. Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.
     The Web app URL stays the same; index.html already points to it.
  Until this is done the Monthly tab shows "The shared store needs the updated Code.gs" and changes nothing.

MONTHLY CONTRIBUTION TAB (Rs 500 per member)
  - Search a member and tap "Mark Rs 500" when they inform you. Tap "Rs 500 paid" to undo; an Undo bar also appears.
  - "Generate list" builds the October list in the usual WhatsApp format; "Unpaid list" builds the pending list.
    Both have Copy and Share to WhatsApp buttons.
  - Marks are saved in a separate Drive file (withus-contributions.json), so re-uploading the Excel never erases them.
  - Marking asks for the passcode once per device. Everyone with the link can see the marks.
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

UPDATING
  After changing index.html, bump CACHE ('withus-v9' -> 'withus-v10') in sw.js so installed apps refresh.
