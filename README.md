# Receipt Helper

A small iPhone web app for company credit card receipts.

- **During the month:** tap **+**, take a photo (or pick a screenshot/PDF), enter amount, category, cost code and a note.
- **At month-end:** open the spreadsheet AP emails you. The app matches each line to a receipt by amount and date,
  fills in Guest/Purpose, Cost Code and the category column, and builds one PDF of the receipts in spreadsheet order.
  Share both files back to AP.

Everything is stored on the phone (IndexedDB). Nothing is uploaded anywhere. Use **Settings → Save backup file** now and then.

## Install on iPhone
1. Open the site in **Safari**.
2. Tap **Share → Add to Home Screen**.
3. Open it from the home screen icon, go to **Settings → Import from Excel…** and pick your Cost Codes spreadsheet.

## Files
- `index.html`, `styles.css`, `app.js` — the app
- `xlsx.js` — reads/fills the AP spreadsheet by editing its XML directly, so AP's formatting and formulas are preserved
- `sw.js`, `manifest.webmanifest`, `icon-*.png` — offline support and home-screen icon

After changing the app, bump `CACHE` in `sw.js` so phones pick up the new version.
