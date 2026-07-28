/* ==========================================================================
   Scannable Listing Generator
   Client-side only: import -> map -> edit -> generate barcoded listing card
   -> export as PDF / Excel, matching the legacy "Barcode creation for TM"
   workbook fields.
   ========================================================================== */

const COMPANY_NAME = "Abudawood Al Saffar Company";
const STORAGE_KEY = "scannables_items_v1";

// Internal field model, in the order the legacy "Format" sheet used them.
const FIELD_DEFS = [
  { key: "principalCode", label: "FPC / Principal Code", aliases: ["fpccode", "fpc", "principalcode", "fpcprincipalcode"] },
  { key: "brand", label: "Brand", aliases: ["brand", "masterbranddesc", "masterbrand"] },
  { key: "sapCode", label: "SAP Code", aliases: ["sapcode", "materialcode", "matcode", "itemcode", "material"] },
  { key: "sapDescription", label: "SAP Description", aliases: ["sapdescription", "materialdescription", "description", "itemdescription"] },
  { key: "caseBarcode", label: "Case Barcode", aliases: ["casebarcode", "maincasebarcode"] },
  { key: "outerBarcode", label: "Outer / Bundle Barcode", aliases: ["outerbarcode", "bundlebarcode", "mainbundlebarcode"] },
  { key: "pieceBarcode", label: "Piece Barcode", aliases: ["piecebarcode", "mainpiecebarcode"] },
  { key: "casePrice", label: "Case Selling Price (BHD)", aliases: ["casesellingpricebhd", "casesellingprice", "caseprice"] },
  { key: "piecePrice", label: "Piece Selling Price (BHD)", aliases: ["piecesellingpricebhd", "piecesellingprice", "pieceprice"] },
  { key: "width", label: "Width", aliases: ["width"] },
  { key: "length", label: "Length", aliases: ["length"] },
  { key: "height", label: "Height", aliases: ["height"] },
  { key: "dimUnit", label: "Dimension Unit", aliases: ["dimunit", "dim-unit", "unit", "volunit"] },
  { key: "countryOfOrigin", label: "Source / Country of Origin", aliases: ["sourcecountryoforigin", "countryoforigin", "source", "primarysource"] },
  { key: "hsCode", label: "HS Code", aliases: ["hscode", "harmonizedcode", "hscustomscode", "harmonisedcode"] },
  { key: "imageUrl", label: "Image Link", aliases: ["imagelink", "image", "imageurl", "photo"] },
];

const REQUIRED_FOR_COMPLETE = ["sapCode", "sapDescription", "hsCode", "countryOfOrigin"];
const PHOTO_MAX_DIM = 640;   // longest side, px — keeps stored photos small
const PHOTO_QUALITY = 0.82;  // JPEG quality

let state = {
  items: [],          // imported/edited items
  rawSheetData: null,  // { sheetNames, workbook } while in the import/mapping stage
  headers: [],
  rows: [],
  mapping: {},          // fieldKey -> source column index (or -1 = skip)
  activeItemId: null,
};

// ---------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------
function normalizeHeader(h) {
  return String(h || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}
function uid() {
  return "it_" + Math.random().toString(36).slice(2, 10);
}
function toast(msg) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast._h);
  toast._h = setTimeout(() => t.classList.remove("show"), 2200);
}
function isComplete(item) {
  return REQUIRED_FOR_COMPLETE.every((k) => String(item[k] || "").trim() !== "");
}
function saveToStorage() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.items)); } catch (e) { /* ignore quota errors */ }
}
function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) state.items = JSON.parse(raw);
  } catch (e) { state.items = []; }
}

// ---------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------
function showTab(name) {
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
  document.getElementById("tab-import").classList.toggle("hidden", name !== "import");
  document.getElementById("tab-items").classList.toggle("hidden", name !== "items");
}

// ---------------------------------------------------------------------
// Import: file parsing
// ---------------------------------------------------------------------
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");

dropzone.addEventListener("click", () => fileInput.click());
["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("drag"); })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove("drag"); })
);
dropzone.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f) handleFile(f);
});
fileInput.addEventListener("change", (e) => {
  const f = e.target.files[0];
  if (f) handleFile(f);
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      state.rawSheetData = wb;
      populateSheetPicker(wb);
    } catch (err) {
      toast("Could not read that file — is it a valid Excel/CSV export?");
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function populateSheetPicker(wb) {
  const picker = document.getElementById("sheetPicker");
  const select = document.getElementById("sheetSelect");
  select.innerHTML = "";
  wb.SheetNames.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name; opt.textContent = name;
    select.appendChild(opt);
  });
  // Prefer a sheet literally called "Complete" or "Data" if present (matches BMF/legacy conventions)
  const preferred = wb.SheetNames.find((n) => /complete|data/i.test(n));
  if (preferred) select.value = preferred;
  picker.classList.remove("hidden");
  loadSheet(select.value);
  select.onchange = () => loadSheet(select.value);
}

function loadSheet(sheetName) {
  const ws = state.rawSheetData.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  if (!json.length) { toast("That sheet looks empty."); return; }
  state.headers = json[0].map((h) => String(h || "").trim());
  state.rows = json.slice(1).filter((r) => r.some((c) => String(c).trim() !== ""));
  buildMapping();
  renderMappingTable();
  document.getElementById("mappingArea").classList.remove("hidden");
  document.getElementById("rowPreviewCount").textContent = `${state.rows.length} rows found`;
}

function buildMapping() {
  const normalizedHeaders = state.headers.map(normalizeHeader);
  state.mapping = {};
  FIELD_DEFS.forEach((f) => {
    let idx = -1;
    for (const alias of f.aliases) {
      const found = normalizedHeaders.indexOf(alias);
      if (found !== -1) { idx = found; break; }
    }
    // fallback: partial match
    if (idx === -1) {
      idx = normalizedHeaders.findIndex((h) => f.aliases.some((a) => h.includes(a) || a.includes(h)));
    }
    state.mapping[f.key] = idx;
  });
}

function renderMappingTable() {
  const table = document.getElementById("mapTable");
  table.innerHTML = `
    <tr><th>Target field</th><th>Source column</th></tr>
    ${FIELD_DEFS.map((f) => `
      <tr>
        <td>${f.label}</td>
        <td>
          <select data-field="${f.key}" onchange="onMapChange(this)">
            <option value="-1">— skip —</option>
            ${state.headers.map((h, i) => `<option value="${i}" ${state.mapping[f.key] === i ? "selected" : ""}>${h || "(column " + (i + 1) + ")"}</option>`).join("")}
          </select>
        </td>
      </tr>
    `).join("")}
  `;
}
function onMapChange(sel) {
  state.mapping[sel.dataset.field] = parseInt(sel.value, 10);
}

document.getElementById("confirmImportBtn").addEventListener("click", () => {
  const imported = state.rows.map((row) => {
    const item = { id: uid() };
    FIELD_DEFS.forEach((f) => {
      const idx = state.mapping[f.key];
      item[f.key] = idx >= 0 && idx < row.length ? String(row[idx] ?? "").trim() : "";
    });
    return item;
  }).filter((it) => it.sapCode || it.sapDescription);

  if (!imported.length) {
    toast("No usable rows — check the column mapping.");
    return;
  }

  // Merge on SAP Code: update existing, append new
  const bySap = new Map(state.items.map((it) => [it.sapCode, it]));
  imported.forEach((it) => {
    if (it.sapCode && bySap.has(it.sapCode)) {
      Object.assign(bySap.get(it.sapCode), it, { id: bySap.get(it.sapCode).id });
    } else {
      state.items.push(it);
      if (it.sapCode) bySap.set(it.sapCode, it);
    }
  });

  saveToStorage();
  renderItems();
  showTab("items");
  toast(`Imported ${imported.length} item${imported.length === 1 ? "" : "s"}.`);
});

// ---------------------------------------------------------------------
// Items table
// ---------------------------------------------------------------------
const itemsTbody = document.getElementById("itemsTbody");
const searchBox = document.getElementById("searchBox");
searchBox.addEventListener("input", renderItems);
document.getElementById("selectAllBtn").addEventListener("click", () => {
  const boxes = itemsTbody.querySelectorAll("input[type=checkbox]");
  const allChecked = [...boxes].every((b) => b.checked);
  boxes.forEach((b) => (b.checked = !allChecked));
});
document.getElementById("clearAllBtn").addEventListener("click", () => {
  if (!confirm("Remove all imported items from this device? This cannot be undone.")) return;
  state.items = [];
  saveToStorage();
  renderItems();
});
document.getElementById("exportSelectedBtn").addEventListener("click", exportSelectedZip);
document.getElementById("bulkPhotoBtn").addEventListener("click", () => document.getElementById("bulkPhotoInput").click());
document.getElementById("bulkPhotoInput").addEventListener("change", (e) => {
  handleBulkPhotoAttach(e.target.files);
  e.target.value = "";
});

function currentFiltered() {
  const q = searchBox.value.trim().toLowerCase();
  if (!q) return state.items;
  return state.items.filter((it) =>
    [it.sapCode, it.sapDescription, it.brand, it.principalCode].join(" ").toLowerCase().includes(q)
  );
}

function renderItems() {
  const hasItems = state.items.length > 0;
  document.getElementById("itemsEmptyState").classList.toggle("hidden", hasItems);
  document.getElementById("itemsContent").classList.toggle("hidden", !hasItems);
  document.getElementById("itemCountBadge").textContent = `(${state.items.length})`;
  if (!hasItems) return;

  const list = currentFiltered();
  document.getElementById("filteredCount").textContent = `${list.length} item${list.length === 1 ? "" : "s"}`;
  document.getElementById("incompleteCount").textContent = `${state.items.filter((it) => !isComplete(it)).length} incomplete`;

  itemsTbody.innerHTML = list.map((it) => rowHtml(it)).join("");
}

function rowHtml(it) {
  const missing = (val) => (String(val || "").trim() === "" ? "missing" : "");
  const dims = `${it.width || "—"}×${it.length || "—"}×${it.height || "—"}${it.dimUnit ? " " + it.dimUnit : ""}`;
  return `
    <tr data-id="${it.id}">
      <td><input type="checkbox" class="rowSelect"></td>
      <td class="sap-cell">${escapeHtml(it.sapCode)}</td>
      <td><input value="${escapeHtml(it.brand)}" oninput="updateField('${it.id}','brand',this.value)"></td>
      <td><input value="${escapeHtml(it.sapDescription)}" style="min-width:180px" oninput="updateField('${it.id}','sapDescription',this.value)"></td>
      <td><input value="${escapeHtml(it.caseBarcode)}" oninput="updateField('${it.id}','caseBarcode',this.value)"></td>
      <td><input value="${escapeHtml(it.outerBarcode)}" oninput="updateField('${it.id}','outerBarcode',this.value)"></td>
      <td><input value="${escapeHtml(it.pieceBarcode)}" oninput="updateField('${it.id}','pieceBarcode',this.value)"></td>
      <td><input value="${escapeHtml(it.casePrice)}" oninput="updateField('${it.id}','casePrice',this.value)"></td>
      <td><input value="${escapeHtml(it.piecePrice)}" oninput="updateField('${it.id}','piecePrice',this.value)"></td>
      <td title="Edit width/length/height in the label view">${dims}</td>
      <td><input class="${missing(it.countryOfOrigin)}" value="${escapeHtml(it.countryOfOrigin)}" oninput="updateField('${it.id}','countryOfOrigin',this.value)"></td>
      <td><input class="${missing(it.hsCode)}" value="${escapeHtml(it.hsCode)}" oninput="updateField('${it.id}','hsCode',this.value)"></td>
      <td>
        <div class="photo-cell">
          ${isDataUrlOrHttp(it.imageUrl) ? `<img class="thumb" src="${escapeHtml(toDirectImageUrl(it.imageUrl))}" alt="">` : `<div class="thumb thumb-empty">no photo</div>`}
          <input value="${escapeHtml(it.imageUrl)}" placeholder="paste a OneDrive/SharePoint link, or upload →" style="min-width:110px" oninput="updateField('${it.id}','imageUrl',this.value); refreshThumb('${it.id}')">
          <label class="upload-btn" title="Upload from this device — only visible on this device, not shared with your team">
            Upload
            <input type="file" accept="image/*" onchange="handlePhotoUpload('${it.id}', this)">
          </label>
        </div>
      </td>
      <td class="row-actions">
        <button class="btn btn-outline btn-sm" onclick="openLabel('${it.id}')">Open</button>
        <button class="btn btn-outline btn-sm" onclick="removeItem('${it.id}')" title="Remove">✕</button>
      </td>
    </tr>
  `;
}
function isDataUrlOrHttp(v) {
  return /^(data:image|https?:)/i.test(String(v || "").trim());
}
function refreshThumb(id) {
  const tr = itemsTbody.querySelector(`tr[data-id="${id}"]`);
  if (!tr) return;
  const it = state.items.find((x) => x.id === id);
  const cell = tr.querySelector(".photo-cell");
  const existing = cell.querySelector(".thumb");
  if (isDataUrlOrHttp(it.imageUrl)) {
    if (existing.tagName === "IMG") existing.src = toDirectImageUrl(it.imageUrl);
    else existing.outerHTML = `<img class="thumb" src="${escapeHtml(toDirectImageUrl(it.imageUrl))}" alt="">`;
  }
}
function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function updateField(id, key, value) {
  const it = state.items.find((x) => x.id === id);
  if (!it) return;
  it[key] = value;
  saveToStorage();
  // Keep counts fresh without a full re-render (avoids losing input focus)
  document.getElementById("incompleteCount").textContent = `${state.items.filter((x) => !isComplete(x)).length} incomplete`;
}
function removeItem(id) {
  state.items = state.items.filter((x) => x.id !== id);
  saveToStorage();
  renderItems();
}

// ---------------------------------------------------------------------
// Photos: local upload (resized to a data URL, no hosting required)
// ---------------------------------------------------------------------
function resizeImageFile(file, maxDim = PHOTO_MAX_DIM, quality = PHOTO_QUALITY) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > height) {
          if (width > maxDim) { height = Math.round(height * (maxDim / width)); width = maxDim; }
        } else if (height > maxDim) {
          width = Math.round(width * (maxDim / height)); height = maxDim;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => reject(new Error("Could not read that image file."));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error("Could not read that file."));
    reader.readAsDataURL(file);
  });
}

// Shared attach logic — used by the table row upload button and the
// (larger, easier to find) upload control in the item's label view.
async function attachPhotoFile(id, file) {
  if (!file) return;
  try {
    const dataUrl = await resizeImageFile(file);
    const it = state.items.find((x) => x.id === id);
    if (!it) return;
    it.imageUrl = dataUrl;
    saveToStorage();
    refreshThumb(id);                    // updates the table row, if visible
    if (state.activeItemId === id) openLabel(id); // refresh the open label card
    toast("Photo attached.");
  } catch (e) {
    toast(e.message || "Could not attach that photo.");
  }
}

// Row upload button — resets the <input> afterward so choosing the same
// file twice in a row still fires a change event.
function handlePhotoUpload(id, fileInput) {
  const file = fileInput.files[0];
  attachPhotoFile(id, file).finally(() => { fileInput.value = ""; });
}

// Bulk attach: pick many photo files at once, match each to an item by
// filename === SAP Code (extension stripped, case-insensitive).
async function handleBulkPhotoAttach(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  const byCode = new Map(state.items.map((it) => [String(it.sapCode).trim().toLowerCase(), it]));
  let matched = 0;
  const unmatched = [];

  toast(`Matching ${files.length} photo${files.length === 1 ? "" : "s"}…`);
  for (const file of files) {
    const stem = file.name.replace(/\.[^.]+$/, "").trim().toLowerCase();
    const it = byCode.get(stem);
    if (!it) { unmatched.push(file.name); continue; }
    try {
      it.imageUrl = await resizeImageFile(file);
      matched++;
    } catch (e) { unmatched.push(file.name); }
  }
  saveToStorage();
  renderItems();
  toast(`Attached ${matched} photo${matched === 1 ? "" : "s"}${unmatched.length ? `, ${unmatched.length} unmatched (filename ≠ SAP code)` : ""}.`);
  if (unmatched.length) console.warn("Unmatched photo filenames:", unmatched);
}

// ---------------------------------------------------------------------
// Share-link conversion: OneDrive/SharePoint/Google Drive "share" links open
// a preview page, not the raw image — <img> needs the direct file URL.
// The stored value stays exactly what the person pasted; this only affects
// what gets rendered.
// ---------------------------------------------------------------------
function toDirectImageUrl(rawUrl) {
  const url = String(rawUrl || "").trim();
  if (!url) return url;

  // Company SharePoint / OneDrive for Business share links
  // (e.g. https://yourtenant-my.sharepoint.com/:i:/g/personal/...)
  if (/\.sharepoint\.com\//i.test(url)) {
    if (/[?&]download=1\b/i.test(url)) return url;
    return url + (url.includes("?") ? "&" : "?") + "download=1";
  }

  // Personal OneDrive share links (1drv.ms or onedrive.live.com)
  if (/1drv\.ms\//i.test(url) || /onedrive\.live\.com\//i.test(url)) {
    try {
      const b64 = btoa(url).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
      return `https://api.onedrive.com/v1.0/shares/u!${b64}/root/content`;
    } catch (e) {
      return url; // btoa fails on non-Latin1 chars in rare cases — fall back to as-pasted
    }
  }

  // Google Drive share links
  const gdrive = url.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  if (gdrive) return `https://drive.google.com/uc?export=view&id=${gdrive[1]}`;

  // Already a direct URL (image host, data URL, etc.) — use as-is
  return url;
}

// ---------------------------------------------------------------------
// Barcode rendering
// ---------------------------------------------------------------------
function renderBarcodeToCanvas(canvas, rawValue) {
  const digits = String(rawValue || "").replace(/\s+/g, "");
  if (!digits) return null;
  const opts = { displayValue: false, margin: 4, height: 50, width: 2 };
  const attempts = digits.length === 8
    ? [{ format: "EAN8" }, { format: "CODE128" }]
    : digits.length === 13
    ? [{ format: "EAN13" }, { format: "CODE128" }]
    : [{ format: "CODE128" }];

  for (const a of attempts) {
    try {
      JsBarcode(canvas, digits, { ...opts, ...a });
      return a.format;
    } catch (e) { /* try next */ }
  }
  return null;
}

// ---------------------------------------------------------------------
// Label overlay
// ---------------------------------------------------------------------
const overlay = document.getElementById("labelOverlay");
function openLabel(id) {
  state.activeItemId = id;
  const it = state.items.find((x) => x.id === id);
  if (!it) return;
  document.getElementById("labelOverlayTitle").textContent = `${it.sapCode || "Item"} — ${it.sapDescription || ""}`;
  document.getElementById("labelCardHost").innerHTML = labelCardHtml(it);
  overlay.classList.remove("hidden");

  // Render barcodes after DOM insert
  [["case", it.caseBarcode], ["outer", it.outerBarcode], ["piece", it.pieceBarcode]].forEach(([key, val]) => {
    const canvas = document.getElementById(`bc-${key}`);
    if (!canvas) return;
    const fmt = renderBarcodeToCanvas(canvas, val);
    const fmtEl = document.getElementById(`fmt-${key}`);
    if (fmtEl) fmtEl.textContent = fmt ? fmt : "";
  });
}
function closeLabel() {
  overlay.classList.add("hidden");
  state.activeItemId = null;
}
overlay.addEventListener("click", (e) => { if (e.target === overlay) closeLabel(); });

function barcodeBoxHtml(key, label, value) {
  if (!String(value || "").trim()) {
    return `<div class="barcode-box empty">${label}<br>— not provided —</div>`;
  }
  return `
    <div class="barcode-box">
      <div class="label">${label}</div>
      <canvas id="bc-${key}"></canvas>
      <div class="digits">${escapeHtml(value)}</div>
      <div class="fmt" id="fmt-${key}"></div>
    </div>
  `;
}

function labelCardHtml(it) {
  const dims = (it.width || it.length || it.height)
    ? `${it.width || "—"} × ${it.length || "—"} × ${it.height || "—"}${it.dimUnit ? " " + it.dimUnit : ""}`
    : "—";
  return `
    <div class="label-card" id="labelCardPrintable">
      <div class="label-head">
        <div class="company">${escapeHtml(COMPANY_NAME)}</div>
        <div class="doc-id">SAP ${escapeHtml(it.sapCode || "—")}</div>
      </div>
      <div class="label-body">
        <div class="label-media">
          ${it.imageUrl ? `<img src="${escapeHtml(toDirectImageUrl(it.imageUrl))}" crossorigin="anonymous" onerror="this.replaceWith(Object.assign(document.createElement('div'),{className:'no-image',textContent:'Image failed to load'}))">` : `<div class="no-image">No image link provided</div>`}
          <div class="media-upload-row">
            <label class="btn btn-outline btn-sm" style="cursor:pointer;">
              Upload photo from device
              <input type="file" accept="image/*" class="hidden" onchange="attachPhotoFile('${it.id}', this.files[0])">
            </label>
          </div>
          <input class="media-link-input" value="${escapeHtml(it.imageUrl)}" placeholder="…or paste a OneDrive/SharePoint link" oninput="updateField('${it.id}','imageUrl',this.value)" onchange="openLabel('${it.id}')">
        </div>
        <div class="label-fields">
          <div class="field-row"><div class="k">FPC / Principal Code</div><div class="v">${escapeHtml(it.principalCode) || "—"}</div></div>
          <div class="field-row"><div class="k">Brand</div><div class="v">${escapeHtml(it.brand) || "—"}</div></div>
          <div class="field-row desc"><div class="k">Description</div></div>
          <div class="field-row desc" style="border-bottom:1px dotted var(--line);"><div class="v" style="font-weight:500;">${escapeHtml(it.sapDescription) || "—"}</div></div>
          <div class="field-row price-row"><div class="k">Case Price (BHD)</div><div class="v">${escapeHtml(it.casePrice) || "—"}</div></div>
          <div class="field-row price-row"><div class="k">Piece Price (BHD)</div><div class="v">${escapeHtml(it.piecePrice) || "—"}</div></div>
          <div class="field-row"><div class="k">Piece Dimensions (W×L×H)</div><div class="v">${dims}</div></div>
        </div>
      </div>
      <div class="customs-strip">
        <div class="cell">
          <div class="k">Source / Country of Origin</div>
          <div class="v">${escapeHtml(it.countryOfOrigin) || "— not provided —"}</div>
        </div>
        <div class="cell">
          <div class="k">HS Code</div>
          <div class="v">${escapeHtml(it.hsCode) || "— not provided —"}</div>
        </div>
      </div>
      <div class="barcode-strip">
        ${barcodeBoxHtml("case", "Case Barcode", it.caseBarcode)}
        ${barcodeBoxHtml("outer", "Bundle Barcode", it.outerBarcode)}
        ${barcodeBoxHtml("piece", "Piece Barcode", it.pieceBarcode)}
      </div>
    </div>
  `;
}

// ---------------------------------------------------------------------
// Export: single item PDF
// ---------------------------------------------------------------------
async function itemToPdfBlob(it) {
  // Render off-screen so we can generate PDFs for items other than the one open in the overlay
  const host = document.createElement("div");
  host.style.position = "fixed";
  host.style.left = "-9999px";
  host.style.top = "0";
  host.style.width = "760px";
  host.innerHTML = labelCardHtml(it);
  document.body.appendChild(host);

  [["case", it.caseBarcode], ["outer", it.outerBarcode], ["piece", it.pieceBarcode]].forEach(([key, val]) => {
    const canvas = host.querySelector(`#bc-${key}`);
    if (canvas) renderBarcodeToCanvas(canvas, val);
  });

  // Let images attempt to load (best-effort; falls back to placeholder on failure already wired via onerror)
  await new Promise((res) => setTimeout(res, 350));

  const cardEl = host.querySelector("#labelCardPrintable");
  const canvas = await html2canvas(cardEl, { scale: 2, backgroundColor: "#ffffff", useCORS: true });
  document.body.removeChild(host);

  const imgData = canvas.toDataURL("image/png");
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 24;
  const maxW = pageWidth - margin * 2;
  const ratio = canvas.height / canvas.width;
  let w = maxW;
  let h = w * ratio;
  if (h > pageHeight - margin * 2) {
    h = pageHeight - margin * 2;
    w = h / ratio;
  }
  pdf.addImage(imgData, "PNG", (pageWidth - w) / 2, margin, w, h);
  return pdf.output("blob");
}

document.getElementById("downloadPdfBtn").addEventListener("click", async () => {
  const it = state.items.find((x) => x.id === state.activeItemId);
  if (!it) return;
  toast("Generating PDF…");
  const blob = await itemToPdfBlob(it);
  downloadBlob(blob, `${outputFilename(it)}.pdf`);
});

// ---------------------------------------------------------------------
// Export: single item Excel (mirrors the legacy "Format" sheet fields)
// ---------------------------------------------------------------------
function itemToWorkbook(it) {
  const rows = [
    [COMPANY_NAME],
    ["FPC / Principal Code", it.principalCode || ""],
    ["Brand", it.brand || ""],
    ["SAP Code", it.sapCode || ""],
    ["SAP Description", it.sapDescription || ""],
    ["Case Barcode", it.caseBarcode || ""],
    ["Outer / Bundle Barcode", it.outerBarcode || ""],
    ["Piece Barcode", it.pieceBarcode || ""],
    ["Case Selling Price – BHD", it.casePrice || ""],
    ["Piece Selling Price – BHD", it.piecePrice || ""],
    ["Piece Dimension (W)", it.width || ""],
    ["Piece Dimension (L)", it.length || ""],
    ["Piece Dimension (H)", it.height || ""],
    ["Dimension Unit", it.dimUnit || ""],
    ["Source / Country of Origin", it.countryOfOrigin || ""],
    ["HS Code", it.hsCode || ""],
    ["Image Link", it.imageUrl || ""],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 26 }, { wch: 40 }];
  const wb = XLSX.utils.book_new();
  const sheetName = sanitizeSheetName(it.sapCode || "Listing");
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}
function sanitizeSheetName(name) {
  return String(name).replace(/[\\/?*[\]:]/g, "-").slice(0, 31) || "Listing";
}
function sanitizeFilename(name) {
  return String(name).replace(/[\\/?*:"<>|]/g, "-").trim().slice(0, 80) || "listing";
}
// The exported file's name is always the SAP code — that's what your team
// scans/searches by. Falls back only if an item genuinely has no SAP code.
function outputFilename(it) {
  const code = String(it.sapCode || "").trim();
  if (code) return sanitizeFilename(code);
  const fallback = (it.sapDescription || it.id || "listing").slice(0, 40);
  return sanitizeFilename(`no-sap-code-${fallback}`);
}

document.getElementById("downloadXlsxBtn").addEventListener("click", () => {
  const it = state.items.find((x) => x.id === state.activeItemId);
  if (!it) return;
  const wb = itemToWorkbook(it);
  XLSX.writeFile(wb, `${outputFilename(it)}.xlsx`);
});

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------
// Bulk export: selected items -> ZIP with /PDF and /Excel, one file per SAP code
// ---------------------------------------------------------------------
async function exportSelectedZip() {
  const ids = [...itemsTbody.querySelectorAll("tr")]
    .filter((tr) => tr.querySelector(".rowSelect")?.checked)
    .map((tr) => tr.dataset.id);

  if (!ids.length) { toast("Select at least one item first (checkbox on the left)."); return; }

  toast(`Generating ${ids.length} listing${ids.length === 1 ? "" : "s"}…`);
  const zip = new JSZip();
  const pdfFolder = zip.folder("PDF");
  const xlsxFolder = zip.folder("Excel");

  // Every selected item exports as its own PDF + Excel file, named by SAP
  // code. If two selected items would produce the same filename (duplicate
  // or missing SAP code), a -2, -3… suffix keeps them from overwriting
  // each other inside the ZIP.
  const usedNames = new Set();
  function uniqueName(base) {
    if (!usedNames.has(base)) { usedNames.add(base); return base; }
    let n = 2;
    while (usedNames.has(`${base}-${n}`)) n++;
    const name = `${base}-${n}`;
    usedNames.add(name);
    return name;
  }

  for (const id of ids) {
    const it = state.items.find((x) => x.id === id);
    if (!it) continue;
    const name = uniqueName(outputFilename(it));
    try {
      const pdfBlob = await itemToPdfBlob(it);
      pdfFolder.file(`${name}.pdf`, pdfBlob);
    } catch (e) { console.error("PDF failed for", name, e); }

    const wb = itemToWorkbook(it);
    const xlsxArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    xlsxFolder.file(`${name}.xlsx`, xlsxArray);
  }

  const content = await zip.generateAsync({ type: "blob" });
  downloadBlob(content, `scannable-listings-${new Date().toISOString().slice(0, 10)}.zip`);
  toast(`Download ready — ${ids.length} PDF${ids.length === 1 ? "" : "s"} + ${ids.length} Excel file${ids.length === 1 ? "" : "s"}, named by SAP code.`);
}

// ---------------------------------------------------------------------
// PWA install prompt
// ---------------------------------------------------------------------
let deferredInstallPrompt = null;
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  document.getElementById("installStrip").classList.add("show");
});
document.getElementById("installBtn").addEventListener("click", async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  document.getElementById("installStrip").classList.remove("show");
});
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch((e) => console.warn("SW registration failed", e));
  });
}

// ---------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------
(function checkDependencies() {
  const missing = [];
  if (typeof XLSX === "undefined") missing.push("xlsx.full.min.js");
  if (typeof JsBarcode === "undefined") missing.push("JsBarcode.all.min.js");
  if (typeof window.jspdf === "undefined") missing.push("jspdf.umd.min.js");
  if (typeof html2canvas === "undefined") missing.push("html2canvas.min.js");
  if (typeof JSZip === "undefined") missing.push("jszip.min.js");
  if (missing.length) {
    const banner = document.createElement("div");
    banner.style.cssText = "background:#B3341F;color:#fff;padding:12px 20px;font-family:sans-serif;font-size:13px;text-align:center;";
    banner.textContent = `This app can't start — missing file(s): ${missing.join(", ")}. Make sure the "vendor" folder was copied alongside index.html, then reload.`;
    document.body.prepend(banner);
    throw new Error("Missing dependencies: " + missing.join(", "));
  }
})();

loadFromStorage();
renderItems();
if (state.items.length) showTab("items");
