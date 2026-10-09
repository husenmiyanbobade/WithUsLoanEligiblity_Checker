WITH US Society Register - PWA

FILES
  index.html, manifest.webmanifest, sw.js, icon-*.png  -> upload all together to any HTTPS static host
  apps-script/Code.gs                                   -> shared data store (Google Apps Script)

HOSTING (HTTPS is required for install + offline)
  Netlify Drop, GitHub Pages, Cloudflare Pages or Firebase Hosting all work.
  Open the hosted URL on a phone > browser menu > "Add to Home screen" / "Install app".

SHARED DATA (so everyone sees the latest uploaded data)
  1. Go to script.google.com > New project > paste apps-script/Code.gs.
  2. Change PASSCODE. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
  3. Done: SYNC_URL in index.html already points to your deployed Web app.
  Upload the Excel once as admin (enter the passcode set in Code.gs); everyone else sees the shared data.

UPDATING
  After changing index.html, bump CACHE ('withus-v3' -> 'withus-v4') in sw.js so installed apps refresh.
