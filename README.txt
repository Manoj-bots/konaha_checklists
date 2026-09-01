MOBILE / GITHUB PAGES SETUP
1. Paste the SAME deployed Apps Script /exec URL used by the laptop into SHEETS_WEB_APP_URL in script.js.
2. Upload the contents of this folder to the ROOT of a GitHub repo.
3. GitHub repo > Settings > Pages > Deploy from branch > main > /(root).
4. Open the Pages URL on your phone and optionally Add to Home screen / Install app.
5. The app syncs from Sheets every 20 seconds; tap Sync for immediate refresh.

IMPORTANT: Apps Script must be deployable to 'Anyone'. A company-domain-only deployment generally blocks GitHub Pages/Lively requests.
Code.gs is included as a copy for convenience; use one shared Apps Script deployment for both devices.
