# EventPass Lite — GitHub Pages Migration

This is the migrated version of your 5 "before_mig" files, following the plan in
your migration guide: all HTML/CSS/JS now lives here (for GitHub Pages), and
Apps Script is left as a pure JSON API.

## What changed

**Backend (`apps-script/Code.gs`)**
- Every business-logic function (`getAllGuests`, `checkInTicket`, `generateBlankTickets`, etc.) is **unchanged**.
- `doGet`/`doPost` were rewritten into a JSON API router (`handleAction_`) — one `action` name per frontend call.
- Removed a few things that no longer apply once the frontend moves off Apps Script:
  - The old `doGet` page-serving switch (`AdminDashboard`/`StaffDashboard`/`Login` templates).
  - `include()` (an HtmlService templating helper, unused now).
  - A stray copy of client-side `openScannerPage()` that had accidentally been pasted into `Code.gs` — it referenced `sessionStorage`/`window`, which don't exist server-side, so it could never have run.
  - A duplicate `generateQrCacheForGuests()` definition — kept the later, more robust one (auto-detects the "QR Payload" column instead of assuming column 5).
- Kept the old JSONP `?page=scanner-api` endpoint as-is, since `scanner.html` already uses it successfully cross-origin.

**Frontend**
- `index.html` (was `login_before_mig.html`) — `google.script.run` → `fetch()` via `js/api.js`; redirects to `admin.html`/`staff.html` instead of Apps Script `?page=` URLs.
- `admin.html` — every `google.script.run` call converted; QR library now local (`js/qrcode.min.js`) instead of `unpkg.com`.
- `staff.html` — every `google.script.run` call converted; scanner library now local (`js/html5-qrcode.min.js`).
- `scanner.html` — barely touched (it was already written to be hosted separately and called the backend via JSONP with no CORS issues). Only the CDN script tag and default "Back" link were updated.
- `js/api.js` — the single shared helper every page calls through.

## Important: the CORS trick

Apps Script can't handle a CORS preflight (`OPTIONS`) request. `js/api.js` avoids
triggering one by sending `Content-Type: text/plain;charset=utf-8` instead of
`application/json` — the body is still a JSON string, and `doPost()` parses it
normally. **Don't change this to `application/json`** or cross-origin requests
will start failing silently.

## Deploy steps

1. **Update Apps Script**: paste `apps-script/Code.gs` over your existing `Code.gs` in the Apps Script editor, then **Deploy → Manage deployments → Edit → New version**. Copy the `/exec` URL.
2. **Set the API URL**: open `js/api.js` and replace `DEFAULT_API_URL` with that `/exec` URL.
3. **Push to GitHub Pages**: commit `index.html`, `admin.html`, `staff.html`, `scanner.html`, and the `js/` folder to your Pages repo (e.g. the same `04-graphics.github.io/EVT/` repo `scanner.html` already lives in).
4. **Test in this order** (matches your guide's staged plan):
   - Login → confirm redirect to the right dashboard.
   - Admin: guest list, ticket generation, CSV import, QR cache, printing.
   - Staff: search, manual check-in.
   - Scanner: open from Staff, scan a real ticket, confirm valid / already-used / invalid all render correctly.
   - Logs and Settings tabs.
5. Once everything above works from GitHub Pages, you can delete the Apps Script deployment's old HTML files (`Login.html`, `AdminDashboard.html`, `StaffDashboard.html`) — the backend no longer serves them.

## File map

```
index.html          ← was login_before_mig.html
admin.html           ← was admin_before_mig.html
staff.html           ← was StaffDashboard_before_mig.html
scanner.html         ← was scanner_before_mig.html (unchanged logic)
js/api.js            ← new — shared fetch() helper
js/qrcode.min.js     ← new — local build of the `qrcode` npm package (was unpkg)
js/html5-qrcode.min.js ← new — local copy of html5-qrcode@2.3.8 (was unpkg)
apps-script/Code.gs  ← was code_before_mig.html — now a JSON API
```
