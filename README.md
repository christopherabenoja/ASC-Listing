# Scannable Listing Generator

A standalone web app (installable on phone and desktop) that replaces the
"Barcode creation for TM.xlsm" macro workflow. Everything runs in the
browser — no server, no Excel, no macros.

## What it does

1. **Import** — upload a BMF-style Excel/CSV export. Columns are auto-matched
   to: FPC/Principal Code, Brand, SAP Code, SAP Description, Case/Bundle/Piece
   Barcode, Case/Piece Selling Price, Piece Dimensions, Source/Country of
   Origin, HS Code, Image Link. You can fix the mapping before importing.
   HS Code, Country of Origin, and prices usually aren't in a BMF export, so
   those stay editable per row after import.
2. **Items** — a searchable, editable grid of everything imported. Rows
   missing HS Code or Country of Origin are flagged.
3. **Open an item** — generates a customer-facing listing card with real,
   scannable barcodes rendered live (EAN-8, EAN-13, or CODE128, auto-detected
   by digit length — same logic as the old macro, but no picture-pasting).
4. **Export** — Download PDF or Excel for one item, or select several rows
   and **Export selected (ZIP)** to get a `/PDF` and `/Excel` folder with one
   file per SAP code, matching the old output structure.

Imported/edited data is saved to the device's local storage automatically,
so closing the tab or app doesn't lose work. "Clear all data" wipes it.

## Running it

No build step, no internet required. The app bundles its own libraries in
`vendor/` — nothing is fetched from a CDN, so it isn't affected by
firewalls, proxies, or being offline.

- **Just open it:** double-click `index.html`. Everything works except the
  "Install app" prompt and offline caching (those need it served over
  http/https, even locally).
- **Serve it** (recommended, enables install + offline):
  ```
  cd scannable-listing-app
  python3 -m http.server 8080
  ```
  then open `http://localhost:8080` in Chrome/Edge/Safari.

**Keep the whole folder together** — `index.html`, `app.js`, `manifest.json`,
`sw.js`, `icons/`, and `vendor/` all need to stay alongside each other. If
you see a red banner saying the app can't start and naming a missing file,
it means `vendor/` wasn't copied along with `index.html` — re-copy the
whole folder.

## Deploying for your team

Any static file host works — no backend required:

- **Company intranet / IIS / Nginx**: copy the folder as-is.
- **GitHub Pages**: push the folder to a repo, enable Pages on the `main`
  branch. Free, HTTPS by default (required for "Install app" and offline
  support on most browsers).
- **OneDrive/SharePoint site pages**: works if it can serve static files
  directly (test the install prompt there — some tenants restrict it).

Once hosted over HTTPS, open it on a phone and use "Add to Home Screen"
(iOS Safari) or the install prompt (Android Chrome/desktop) to get a
proper app icon with no browser chrome.

## Notes & limitations

- **Item data is also per-device.** Everything imported/edited in the Items
  grid is saved to that browser's local storage, not a shared server — the
  same as the photo behavior above. If two people import the same BMF file
  on their own devices, they get two independent working copies. Re-import
  the same source file on each device to catch it up, or treat one device
  as the "master" and export/share the ZIP output from there.
- **HS Code and Country of Origin are not in standard SAP/BMF exports** —
  these need to be entered once per SAP code in the Items grid (or you can
  maintain a small master list separately and re-import it; matching is by
  SAP Code, so re-importing updates existing rows instead of duplicating them).
- **Barcode format detection**: 8-digit codes try EAN-8, 13-digit codes try
  EAN-13, everything else uses CODE128 — same behavior as the original
  workbook. If an EAN checksum is invalid, it automatically falls back to
  CODE128 so nothing fails silently.
- **Excel export is data-only** (no embedded barcode images) — this keeps it
  dependency-free and fast. The PDF export has the full visual listing with
  live barcodes; the Excel export is for record-keeping in the same fields
  as your legacy "Format" sheet.
- **Output files are always named by SAP code** — `20028417.pdf`,
  `20028417.xlsx`, etc. — regardless of whether you export a single item or
  a batch. If an item genuinely has no SAP code, it falls back to
  `no-sap-code-<description>` so it's still identifiable. When exporting
  multiple selected items as a ZIP, any filename collision (duplicate or
  missing SAP code) automatically gets a `-2`, `-3`… suffix so nothing gets
  silently overwritten inside the ZIP.
- **Two places to attach a photo**: the small "Upload" button in the Items
  grid, or — easier to find — open the item (**Open** button) and use the
  **Upload photo from device** button right in the listing preview, which
  also has the paste-a-link field alongside it.
- **Team-shared vs. this-device-only.** Local storage (both the item data
  and any uploaded photo) lives in *that one browser, on that one device* —
  it never leaves it. So the recommended path for photos everyone can see is:

  1. Put product photos in your shared OneDrive/SharePoint folder.
  2. Right-click the photo → **Share** → set access to *"People in
     [company] with the link"* (view-only) → copy the link.
  3. Paste that link into the item's **Image URL** field.

  The app automatically converts OneDrive/SharePoint/Google Drive share
  links (which normally open a preview page) into the direct image URL an
  `<img>` tag needs — you don't have to hunt for the "direct link" format
  yourself. This is handled by `toDirectImageUrl()` in `app.js` if you ever
  need to adjust it for a different host.

  **The 📷 upload button and "Attach photos (bulk)" are personal, this-device
  previews only** — useful for checking how a listing will look before the
  photo is up in the shared folder, but a teammate on a different device
  will not see them. They're intentionally labelled "this device only" in
  the app so nobody assumes it's shared.

  If a pasted link shows a placeholder instead of the photo, the most common
  cause is the share permission being set to specific people instead of
  "anyone in the company/anyone with the link" — `<img>` tags can't pass
  along a login prompt.
- All processing happens in the browser; no item data is sent anywhere.

## Files

```
index.html     — app shell, layout, styles
app.js         — import/mapping, item state, barcode rendering, export logic
manifest.json  — PWA metadata (name, icons, colors)
sw.js          — service worker (offline caching)
icons/         — app icons (192px, 512px)
vendor/        — bundled libraries (SheetJS, JsBarcode, jsPDF, html2canvas,
                  JSZip) — no CDN calls, works with no internet access
```
