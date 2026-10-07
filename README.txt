WITH US Society PWA

1) Sharing data with everyone (one-time, 5 minutes)
   - Go to script.google.com > New project. Paste apps-script/Code.gs. Change PASSCODE.
   - Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Authorize when asked.
   - Copy the Web app URL. Open index.html in a text editor, find  var SYNC_URL = '';  and paste the URL between the quotes.
   - After this, whoever uploads + processes the Excel file enters the passcode once. Everyone else sees the latest data on open.

2) Hosting (needed for install + sharing a link; must be https)
   Upload this whole folder to Netlify Drop, Cloudflare Pages or GitHub Pages. Share that link.

3) Install: open the link on the phone > browser menu > Add to Home screen / Install app.
   When you change index.html later, also bump CACHE name in sw.js (withus-v1 -> withus-v2).
