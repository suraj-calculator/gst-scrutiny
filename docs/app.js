"use strict";

// ---------------------------------------------------------------------
// Every merge/align/extract/convert call in this file runs the real,
// unmodified scripts from the repo (vendored below) inside Pyodide —
// nothing here is simulated. Uploaded files never leave this browser tab.
// ---------------------------------------------------------------------

const ICONS = {
  upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 15V4M12 4 8 8M12 4l4 4"/><path d="M4 16v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/></svg>',
  file: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7 3h7l4 4v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M14 3v4h4"/></svg>',
  check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
  checkSm: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
  warn: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4M12 17h.01M10.3 3.9 2.8 17a1.6 1.6 0 0 0 1.4 2.4h15.6a1.6 1.6 0 0 0 1.4-2.4L13.7 3.9a1.6 1.6 0 0 0-2.8 0Z"/></svg>',
  lock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="10" width="16" height="10" rx="1.5"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></svg>',
};

// ---- section config -----------------------------------------------------
const UPLOAD_SECTIONS = [
  { id: "ewb_inward", step: 1, title: "Inward E-Way Bill", required: false, py: { kind: "ewb", direction: "inward" },
    desc: "EWB MIS Report exports from the inward E-Way Bill portal folder.",
    accept: "EWB_MIS_Report_Excel (N).xls" },
  { id: "ewb_outward", step: 2, title: "Outward E-Way Bill", required: false, py: { kind: "ewb", direction: "outward" },
    desc: "EWB MIS Report exports from the outward E-Way Bill portal folder.",
    accept: "EWB_MIS_Report_Excel (N).xls" },
  { id: "einv", step: 3, title: "E-Invoice", required: false, py: { kind: "merge", mergeKind: "einv" },
    desc: "Monthly E-Invoice exports, any number of periods.",
    accept: "per-period E-Invoice .xlsx" },
  { id: "gstr1", step: 4, title: "GSTR-1", required: true, py: { kind: "merge", mergeKind: "gstr1" },
    desc: "Monthly or quarterly GSTR-1 exports for the financial year.",
    accept: "per-period GSTR-1 .xlsx" },
  { id: "gstr2a", step: 5, title: "GSTR-2A", required: false, py: { kind: "merge", mergeKind: "gstr2a" },
    desc: "Monthly GSTR-2A exports for the financial year.",
    accept: "per-period GSTR-2A .xlsx",
    note: "Not yet exercised against real GSTR-2A data in this build — the code path is identical to the other return types, just unverified. Report an issue if it misbehaves." },
  { id: "gstr2b", step: 6, title: "GSTR-2B", required: false, py: { kind: "gstr2b" },
    desc: "Monthly GSTR-2B exports — every workbook is aligned to the same set of worksheets before merging.",
    accept: "per-period GSTR-2B .xlsx" },
  { id: "gstr3b", step: 7, title: "GSTR-3B", required: true, py: { kind: "gstr3b" },
    desc: "GSTR3B_&lt;GSTIN&gt;_&lt;MMYYYY&gt;.zip bundles straight from the portal (or already-extracted .xlsx files).",
    accept: ".zip bundles or .xlsx" },
];

const LEDGER_SLOTS = [
  { id: "cash", name: "Cash Ledger", ext: "CSV" },
  { id: "credit", name: "Credit Ledger", ext: "CSV" },
  { id: "liab1", name: "Liability Register — Part I", ext: "CSV" },
  { id: "liab2", name: "Liability Ledger — Part II (DRC)", ext: "CSV" },
  { id: "comparison", name: "Tax Liability & ITC Comparison", ext: "XLSX" },
  { id: "table8a", name: "Table 8A", ext: "XLSX" },
];
const PDF_SLOTS = [
  { id: "bo_profile", name: "BO / 360° Profile", ext: "PDF" },
  { id: "gstr9", name: "GSTR-9 Annual Return", ext: "PDF" },
  { id: "gstr9c", name: "GSTR-9C Reconciliation", ext: "PDF" },
];
const MASTER_SLOTS = [
  { id: "hsn_master", name: "HSN / SAC Code Master", ext: "XLSX" },
  { id: "blocked_itc", name: "Blocked-ITC Keyword Master", ext: "XLSX" },
  { id: "machinery", name: "Machinery HSN Master", ext: "XLSX" },
];
// label -> the real BS_PL_DATA dict key from bs_pl_input.py (see
// "main gst tool/bs_pl_input.py"), so form values assemble into exactly
// the shape master_build.py already knows how to consume.
const BSPL_ROWS = [
  { label: "Total Assets", key: "total_assets" },
  { label: "Total Equity & Liabilities", key: "total_equity_liab" },
  { label: "Trade Receivables", key: "trade_receivables" },
  { label: "Trade Payables", key: "trade_payables" },
  { label: "Revenue from Operations", key: "revenue_from_operations" },
  { label: "Net Profit After Tax", key: "net_profit_after_tax" },
];

const RAIL_META = [
  ...UPLOAD_SECTIONS.map(s => ({ id: s.id, step: s.step, title: s.title, required: s.required })),
  { id: "ledgers", step: 8, title: "Ledgers & Portal Reports", required: false },
  { id: "annual_pdfs", step: 9, title: "Annual PDF Reports", required: false },
  { id: "masters", step: 10, title: "Reference Masters", required: false },
  { id: "bs_pl", step: 11, title: "Balance Sheet / P&L", required: false },
];
const REQUIRED_IDS = RAIL_META.filter(s => s.required).map(s => s.id);
const OPTIONAL_IDS = RAIL_META.filter(s => !s.required).map(s => s.id);

// Everything collected so far, keyed by section id. Not yet consumed by
// anything (main gst tool wiring is next) — this is where that will read
// from once the "Run full scrutiny" button is wired up.
const workbench = { done: new Set() };

// ---------------------------------------------------------------------
// Pyodide runtime — runs in a Web Worker (worker.js), never on the main
// thread. A multi-second merge used to freeze the tab (confirmed by
// testing the earlier main-thread version); this keeps the UI responsive
// throughout. All communication is via postMessage, matched by request id.
// ---------------------------------------------------------------------
let worker = null;
let workerReady = false;
const pendingCalls = new Map(); // id -> {resolve, reject}
let _reqSeq = 0;

// One centralized boot overlay instead of every section separately saying
// "waiting for runtime" — a single place to look, with a bit of personality
// while Pyodide + pandas/openpyxl actually load (~5-10s, once per visit).
const BOOT_MESSAGES = [
  "Waking up Python…",
  "Untangling GST spaghetti…",
  "Teaching pandas to read Excel…",
  "Politely interrogating openpyxl…",
  "Double-checking Section 17(5)…",
  "Reconciling a few imaginary invoices, just to warm up…",
  "Summoning the HSN code oracle…",
  "Convincing WebAssembly this is a good idea…",
];
let _bootMsgTimer = null;

function cycleBootMessage() {
  const el = document.getElementById("boot-message");
  let i = 0;
  el.textContent = BOOT_MESSAGES[0];
  _bootMsgTimer = setInterval(() => {
    i = (i + 1) % BOOT_MESSAGES.length;
    el.textContent = BOOT_MESSAGES[i];
  }, 1700);
}

function setRuntimeState(state, text) {
  const pill = document.getElementById("runtime-pill");
  pill.dataset.state = state;
  pill.innerHTML = `<span class="dot"></span>${text}`;
  document.getElementById("rail-hint").textContent =
    state === "ready"
      ? "Only GSTR-1 and GSTR-3B are required. Everything else is optional and the final report will say plainly what was skipped."
      : text;

  const overlay = document.getElementById("boot-overlay");
  if (state === "ready") {
    clearInterval(_bootMsgTimer);
    overlay.dataset.hidden = "true";
  } else if (state === "error") {
    clearInterval(_bootMsgTimer);
    document.getElementById("boot-spinner").dataset.state = "error";
    document.getElementById("boot-message").textContent = "Something went wrong";
    const sub = document.getElementById("boot-sub");
    sub.textContent = text;
    sub.classList.add("error");
  }
  // "loading" state: overlay already visible with rotating messages from
  // cycleBootMessage(), started once in initRuntime() — nothing to do here.
}

function initRuntime() {
  setRuntimeState("loading", "Starting Python runtime…");
  cycleBootMessage();
  worker = new Worker("worker.js");
  worker.onmessage = e => {
    const msg = e.data;
    if (msg.type === "status") {
      setRuntimeState(msg.state, msg.text);
      if (msg.state === "ready") {
        workerReady = true;
        enableAllDropzones();
        updateRunbar();
      }
    } else if (msg.type === "result") {
      const pending = pendingCalls.get(msg.id);
      if (!pending) return;
      pendingCalls.delete(msg.id);
      if (msg.ok) pending.resolve(msg.result);
      else pending.reject(new Error(msg.error));
    }
  };
  worker.onerror = err => {
    console.error(err);
    setRuntimeState("error", "Runtime failed to load — reload the page");
  };
}

function callWorker(adapter, args) {
  return new Promise((resolve, reject) => {
    const id = ++_reqSeq;
    pendingCalls.set(id, { resolve, reject });
    worker.postMessage({ type: "call", id, adapter, args });
  });
}

async function callEwb(direction, filePairs) {
  return await callWorker("ewb", { direction, filePairs });
}
async function callMerge(kind, filePairs) {
  return await callWorker("merge", { mergeKind: kind, filePairs });
}
async function callGstr3b(filePairs) {
  return await callWorker("gstr3b", { filePairs });
}
async function callGstr2b(filePairs) {
  return await callWorker("gstr2b", { filePairs });
}
async function callFullScrutiny(filePairs, bsPlData) {
  return await callWorker("full_scrutiny", { filePairs, bsPlData });
}

// ---------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------
const railList = document.getElementById("rail-list");
railList.innerHTML = RAIL_META.map(s => `
  <li class="rail-item" data-rail="${s.id}">
    <a href="#sec-${s.id}">
      <span class="rail-num">${String(s.step).padStart(2, "0")}</span>
      <span class="rail-dot" data-dot="${s.id}"></span>
      ${s.title}
      ${s.required ? '<span class="req-mark">REQ</span>' : ""}
    </a>
  </li>`).join("");

function uploadSectionHtml(cfg) {
  return `
  <section class="card" id="sec-${cfg.id}" data-section="${cfg.id}">
    <div class="card-head">
      <div class="card-head-left">
        <div class="card-title-row">
          <h2>${cfg.title}</h2>
          <span class="tag ${cfg.required ? "required" : "optional"}">${cfg.required ? "Required" : "Optional"}</span>
        </div>
        <p class="card-desc">${cfg.desc}${cfg.note ? ` <em>${cfg.note}</em>` : ""}</p>
      </div>
      <span class="status-pill" data-status="${cfg.id}" data-state="empty"><span class="dot"></span>Not started</span>
    </div>
    <div class="card-body">
      <div class="dropzone" data-dropzone="${cfg.id}" data-disabled="true" tabindex="0" role="button" aria-label="Upload files for ${cfg.title}">
        <span class="dz-icon">${ICONS.upload}</span>
        <span class="dz-text">
          <strong>Drop files here, or click to choose</strong>
          <span>${cfg.accept} &middot; multiple files</span>
        </span>
        <input type="file" data-input="${cfg.id}" multiple>
      </div>
      <div class="file-count" data-filecount="${cfg.id}"></div>
      <div data-progress="${cfg.id}"></div>
      <div data-result="${cfg.id}"></div>
    </div>
  </section>`;
}

function slotSectionHtml(id, title, desc, slots, opts) {
  opts = opts || {};
  return `
  <section class="card" id="sec-${id}" data-section="${id}" ${opts.disabled ? 'data-disabled="true"' : ""}>
    <div class="card-head">
      <div class="card-head-left">
        <div class="card-title-row">
          <h2>${title}</h2>
          <span class="tag ${opts.deferred ? "deferred" : "optional"}">${opts.deferred ? "Deferred" : "Optional"}</span>
        </div>
        <p class="card-desc">${desc}</p>
      </div>
      <span class="status-pill" data-status="${id}" data-state="empty"><span class="dot"></span>${opts.disabled ? "Not available" : `0 of ${slots.length}`}</span>
    </div>
    <div class="card-body">
      <div class="slot-grid">
        ${slots.map(sl => `
          <div class="slot" data-slot="${id}:${sl.id}" ${opts.disabled ? "" : `tabindex="0" role="button"`} aria-label="Upload ${sl.name}">
            <span class="slot-icon">${ICONS.file}</span>
            <span>
              <span class="slot-name">${sl.name}</span>
              <span class="slot-sub"> &middot; ${sl.ext}</span>
            </span>
            <span class="slot-status">${opts.disabled ? "—" : "Empty"}</span>
            ${opts.disabled ? "" : `<input type="file" data-slotinput="${id}:${sl.id}">`}
          </div>`).join("")}
      </div>
      ${opts.persist ? `<p class="persist-note">${ICONS.lock} Saved in this browser — you won't need to re-upload these next time.</p>` : ""}
      ${opts.footnote ? `<p class="card-desc" style="margin-top:12px;">${opts.footnote}</p>` : ""}
    </div>
  </section>`;
}

function bsplSectionHtml() {
  return `
  <section class="card" id="sec-bs_pl" data-section="bs_pl">
    <div class="card-head">
      <div class="card-head-left">
        <div class="card-title-row">
          <h2>Balance Sheet / P&amp;L</h2>
          <span class="tag optional">Optional</span>
        </div>
        <p class="card-desc">Typed in, not uploaded — the tool never OCRs a scanned balance sheet. Figures cross-check against the return data.</p>
      </div>
      <span class="status-pill" data-status="bs_pl" data-state="empty"><span class="dot"></span>Not started</span>
    </div>
    <div class="card-body">
      <p class="bspl-tag">Tagged to <span class="mono" id="bspl-gstin-tag">—</span> — refused automatically by master_build.py if it doesn't match the GSTIN being processed.</p>
      <div class="bspl-grid" id="bspl-grid">
        <div class="head">Line item</div><div class="head" style="text-align:right;">FY prior</div><div class="head" style="text-align:right;">FY current</div>
        ${BSPL_ROWS.map(row => `
          <div class="label">${row.label}</div>
          <div><input class="mono" data-bspl="${row.key}-prior" type="text" inputmode="decimal" placeholder="—"></div>
          <div><input class="mono" data-bspl="${row.key}-current" type="text" inputmode="decimal" placeholder="—"></div>`).join("")}
      </div>
      <p class="bspl-more">+ 10 more line items in the full form (reserves, provisions, finance costs, depreciation, and more).</p>
    </div>
  </section>`;
}

const pipeline = document.getElementById("pipeline");
pipeline.innerHTML =
  UPLOAD_SECTIONS.map(uploadSectionHtml).join("") +
  slotSectionHtml("ledgers", "Ledgers &amp; Portal Reports",
    "Straight portal exports — no conversion needed, stored as-is.", LEDGER_SLOTS) +
  slotSectionHtml("annual_pdfs", "Annual PDF Reports",
    "PDF exports from the portal, normally converted to structured Excel automatically.",
    PDF_SLOTS, {
      disabled: true, deferred: true,
      footnote: "Deferred: the PDF library this needs (pdfplumber) depends on a component with no WebAssembly build, so it can't run in the browser yet. Confirmed by testing directly — not an assumption. These 3 inputs are optional, so the rest of the tool works without them.",
    }) +
  slotSectionHtml("masters", "Reference Masters",
    "Your organisation's own reference lists — upload once, reused on every filing.",
    MASTER_SLOTS, { persist: true }) +
  bsplSectionHtml();

// ---------------------------------------------------------------------
// Shared status/workbench helpers
// ---------------------------------------------------------------------
function setStatus(id, state, text) {
  const pill = document.querySelector(`[data-status="${id}"]`);
  if (!pill) return;
  pill.dataset.state = state;
  pill.innerHTML = `<span class="dot"></span>${text}`;
  const dot = document.querySelector(`[data-dot="${id}"]`);
  if (dot) dot.dataset.state = state;
}

function markDone(id) { workbench.done.add(id); updateRunbar(); }
function markUndone(id) { workbench.done.delete(id); updateRunbar(); }

function updateRunbar() {
  const reqDone = REQUIRED_IDS.filter(id => workbench.done.has(id)).length;
  const optDone = OPTIONAL_IDS.filter(id => workbench.done.has(id)).length;
  document.getElementById("req-status").innerHTML = `Required: <span class="mono">${reqDone}/${REQUIRED_IDS.length}</span> ready`;
  document.getElementById("req-status").classList.toggle("ok", reqDone === REQUIRED_IDS.length);
  document.getElementById("opt-status").textContent = `${optDone}/${OPTIONAL_IDS.length}`;
  document.getElementById("opt-bar").style.width = `${(optDone / OPTIONAL_IDS.length) * 100}%`;
  const btn = document.getElementById("run-btn");
  // Only truly disabled while the runtime itself isn't up yet — nothing can
  // run at all then. Once it's ready, the button stays clickable even with
  // required sections missing; clicking it is what shows the "nothing
  // uploaded" alert, rather than a silently disabled button with no
  // explanation.
  btn.disabled = !workerReady;
  btn.title = workerReady ? "" : "Waiting for the Python runtime to finish loading";
}

function toast(msg) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove("show"), 3200);
}

function enableAllDropzones() {
  // Every upload section's dropzone starts with its own data-disabled="true"
  // marker (cleared here once the runtime is ready). Deferred sections
  // (Annual PDF Reports) use .slot elements, not .dropzone, so they're
  // never matched by this selector and stay off on their own.
  document.querySelectorAll(".dropzone[data-disabled]").forEach(dz => {
    dz.removeAttribute("data-disabled");
  });
  UPLOAD_SECTIONS.forEach(cfg => setStatus(cfg.id, "empty", "Not started"));
}

async function fileToBytes(file) {
  return new Uint8Array(await file.arrayBuffer());
}
async function filesToPairs(fileList) {
  const out = [];
  for (const f of fileList) out.push([f.name, await fileToBytes(f)]);
  return out;
}
// ---------------------------------------------------------------------
// Upload sections — real processing
// ---------------------------------------------------------------------
function showError(cfg, message) {
  setStatus(cfg.id, "error", "Failed");
  document.querySelector(`[data-result="${cfg.id}"]`).innerHTML = `
    <div class="result-line error">
      ${ICONS.warn}
      <p>${message}</p>
    </div>`;
  markUndone(cfg.id);
}

function showResult(cfg, r, extraLine) {
  if (!r) {
    showError(cfg, `No ${cfg.title}-shaped files were detected among what you uploaded. Double-check these are the right export type.`);
    return;
  }
  setStatus(cfg.id, "ready", "Ready");
  const kb = (r.output_bytes.length / 1024).toFixed(0);
  document.querySelector(`[data-result="${cfg.id}"]`).innerHTML = `
    <div class="result-line">
      ${ICONS.check}
      <p>${extraLine ? extraLine + "<br>" : ""}<span class="out-file">→ ${r.output_name}</span> (${kb} KB)</p>
    </div>`;
  workbench[cfg.id] = r;
  markDone(cfg.id);
}

async function processSection(cfg, filePairs) {
  if (!workerReady) return;
  // Shown immediately and left in place through processing/success/failure,
  // so the user always has a plain confirmation of how many files they
  // actually selected, independent of whatever the processing result was.
  document.querySelector(`[data-filecount="${cfg.id}"]`).textContent =
    `${filePairs.length} file${filePairs.length === 1 ? "" : "s"} uploaded`;
  setStatus(cfg.id, "processing", "Processing…");
  document.querySelector(`[data-result="${cfg.id}"]`).innerHTML = "";
  const progWrap = document.querySelector(`[data-progress="${cfg.id}"]`);
  progWrap.innerHTML = `<div class="progress-wrap"><div class="progress-track"><div class="progress-fill" data-fill></div></div><div class="progress-label">Running the real merge script on ${filePairs.length} file${filePairs.length === 1 ? "" : "s"}…</div></div>`;
  requestAnimationFrame(() => { const f = progWrap.querySelector("[data-fill]"); if (f) f.style.width = "100%"; });

  try {
    let result, extraLine;
    if (cfg.py.kind === "ewb") {
      result = await callEwb(cfg.py.direction, filePairs);
      extraLine = result ? `${result.rows} rows merged` : null;
    } else if (cfg.py.kind === "merge") {
      result = await callMerge(cfg.py.mergeKind, filePairs);
    } else if (cfg.py.kind === "gstr2b") {
      result = await callGstr2b(filePairs);
    } else if (cfg.py.kind === "gstr3b") {
      result = await callGstr3b(filePairs);
    }
    progWrap.innerHTML = "";
    showResult(cfg, result, extraLine);
  } catch (err) {
    console.error(err);
    progWrap.innerHTML = "";
    showError(cfg, `Error while processing: ${String(err.message || err).slice(0, 200)}`);
  }
}

UPLOAD_SECTIONS.forEach(cfg => {
  const dz = document.querySelector(`[data-dropzone="${cfg.id}"]`);
  const input = document.querySelector(`[data-input="${cfg.id}"]`);

  function guarded(fn) {
    return (...args) => { if (dz.dataset.disabled === "true") return; fn(...args); };
  }

  dz.addEventListener("click", guarded(() => input.click()));
  dz.addEventListener("keydown", guarded(e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } }));
  ["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, guarded(e => { e.preventDefault(); dz.classList.add("drag"); })));
  ["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, guarded(e => { e.preventDefault(); dz.classList.remove("drag"); })));
  dz.addEventListener("drop", guarded(async e => {
    const files = e.dataTransfer.files;
    if (files && files.length) processSection(cfg, await filesToPairs(files));
  }));
  input.addEventListener("change", async () => {
    if (input.files && input.files.length) await processSection(cfg, await filesToPairs(input.files));
    input.value = "";
  });
});

// ---------------------------------------------------------------------
// Ledgers / masters — real byte capture, no processing needed
// ---------------------------------------------------------------------
function wireSlotSection(sectionId, slots, opts) {
  opts = opts || {};
  const doneSlots = new Set();
  slots.forEach(sl => {
    const key = `${sectionId}:${sl.id}`;
    const el = document.querySelector(`[data-slot="${key}"]`);
    const input = document.querySelector(`[data-slotinput="${key}"]`);
    if (!el || !input) return;

    async function complete(file) {
      const bytes = await fileToBytes(file);
      workbench[key] = { name: file.name, bytes };
      doneSlots.add(sl.id);
      el.dataset.done = "true";
      el.querySelector(".slot-icon").innerHTML = ICONS.checkSm;
      el.querySelector(".slot-status").textContent = file.name.length > 22 ? file.name.slice(0, 20) + "…" : file.name;
      setStatus(sectionId, doneSlots.size === slots.length ? "ready" : "processing", `${doneSlots.size} of ${slots.length}`);
      if (doneSlots.size > 0) markDone(sectionId);
      if (opts.persist) persistMaster(sl.id, file.name, bytes);
    }

    el.addEventListener("click", () => input.click());
    el.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
    input.addEventListener("change", () => { if (input.files[0]) complete(input.files[0]); });
  });
  return { restore: (slotId, name, bytes) => {
    const el = document.querySelector(`[data-slot="${sectionId}:${slotId}"]`);
    if (!el) return;
    workbench[`${sectionId}:${slotId}`] = { name, bytes };
    doneSlots.add(slotId);
    el.dataset.done = "true";
    el.querySelector(".slot-icon").innerHTML = ICONS.checkSm;
    el.querySelector(".slot-status").textContent = name.length > 22 ? name.slice(0, 20) + "…" : name;
    setStatus(sectionId, doneSlots.size === slots.length ? "ready" : "processing", `${doneSlots.size} of ${slots.length}`);
    if (doneSlots.size > 0) markDone(sectionId);
  }};
}
wireSlotSection("ledgers", LEDGER_SLOTS);
const mastersHandle = wireSlotSection("masters", MASTER_SLOTS, { persist: true });

// Reference masters persist across visits — this is genuinely per-viewer
// browser storage, not shared with anyone; wrapped defensively since
// localStorage can throw (private browsing, quota, disabled storage).
const MASTER_STORAGE_KEY = "scrutiny-desk:masters:v1";
function persistMaster(slotId, name, bytes) {
  try {
    const store = JSON.parse(localStorage.getItem(MASTER_STORAGE_KEY) || "{}");
    let binary = "";
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    store[slotId] = { name, b64: btoa(binary) };
    localStorage.setItem(MASTER_STORAGE_KEY, JSON.stringify(store));
  } catch (err) {
    console.warn("could not persist master file locally", err);
  }
}
function loadPersistedMasters() {
  try {
    const store = JSON.parse(localStorage.getItem(MASTER_STORAGE_KEY) || "{}");
    for (const [slotId, { name, b64 }] of Object.entries(store)) {
      const binary = atob(b64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      mastersHandle.restore(slotId, name, bytes);
    }
  } catch (err) {
    console.warn("could not restore saved master files", err);
  }
}
loadPersistedMasters();

// ---------------------------------------------------------------------
// BS/PL form
// ---------------------------------------------------------------------
function parseBsplNumber(raw) {
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const n = Number(trimmed.replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}

document.getElementById("bspl-grid").addEventListener("input", () => {
  const inputs = Array.from(document.querySelectorAll("[data-bspl]"));
  const anyFilled = inputs.some(i => i.value.trim() !== "");
  setStatus("bs_pl", anyFilled ? "ready" : "empty", anyFilled ? "Ready" : "Not started");
  if (anyFilled) {
    // Shaped exactly like BS_PL_DATA in "main gst tool/bs_pl_input.py" —
    // web_adapters.process_full_scrutiny() writes this straight into a
    // real bs_pl_input.py module for master_build.py to import.
    const data = { _gstin: gstinField.value.trim() || null };
    BSPL_ROWS.forEach(row => {
      const prior = parseBsplNumber(document.querySelector(`[data-bspl="${row.key}-prior"]`).value);
      const current = parseBsplNumber(document.querySelector(`[data-bspl="${row.key}-current"]`).value);
      if (prior !== null || current !== null) data[row.key] = { fy_prior: prior, fy_current: current };
    });
    workbench.bs_pl = data;
    markDone("bs_pl");
  } else {
    delete workbench.bs_pl;
    markUndone("bs_pl");
  }
});

const gstinField = document.getElementById("gstin-field");
const fyField = document.getElementById("fy-field");
gstinField.addEventListener("input", () => {
  document.getElementById("bspl-gstin-tag").textContent = gstinField.value || "—";
});

// ---------------------------------------------------------------------
// Scroll-spy
// ---------------------------------------------------------------------
const sections = Array.from(document.querySelectorAll("[data-section]"));
const railItems = Array.from(document.querySelectorAll(".rail-item"));
const spy = new IntersectionObserver(entries => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      const id = entry.target.dataset.section;
      railItems.forEach(li => li.classList.toggle("active", li.dataset.rail === id));
    }
  });
}, { rootMargin: "-15% 0px -70% 0px" });
sections.forEach(s => spy.observe(s));

// ---------------------------------------------------------------------
// Run full scrutiny — assembles everything collected so far and runs the
// real main gst tool engine (master_build.py) against it.
// ---------------------------------------------------------------------
function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function statTile(cls, n, label) {
  return `<div class="stat-tile ${cls}"><div class="n">${n}</div><div class="l">${label}</div></div>`;
}

function renderScrutinyResult(result) {
  const results = document.getElementById("results");
  const log = result.log || "";

  // Every figure below is parsed straight out of master_build.py's own
  // printed run summary (see its main()) — nothing here is invented. If a
  // line's shape ever changes, that tile just doesn't render; the full raw
  // log underneath is always the ground truth regardless.
  const monthsMatch = log.match(/Months covered: \[(.*?)\]/);
  const hsnMatch = log.match(/HSN & Fraud Pattern Checks: (\d+) total \((\d+) FLAG, (\d+) REVIEW\)/);
  const flowMatch = log.match(/Flow \/ counterparty checks: (\d+) findings across (\d+) sheets \((\d+) FLAG, (\d+) REVIEW\)/);
  const cancelledMatch = log.match(/Cancelled e-invoices found: (\d+)/);
  const rectMatch = log.match(/Rectification pairs: (\d+)/);

  let tiles = "";
  if (monthsMatch) tiles += statTile("info", monthsMatch[1].split(",").filter(s => s.trim()).length, "Months covered");
  if (hsnMatch) tiles += statTile("flag", hsnMatch[2], "HSN/fraud flags");
  if (hsnMatch) tiles += statTile("review", hsnMatch[3], "HSN/fraud review");
  if (flowMatch) tiles += statTile("flag", flowMatch[3], "Flow flags");
  if (flowMatch) tiles += statTile("review", flowMatch[4], "Flow review");
  if (cancelledMatch) tiles += statTile("info", cancelledMatch[1], "Cancelled e-inv.");
  if (rectMatch) tiles += statTile("info", rectMatch[1], "Rectification pairs");

  const gstin = gstinField.value.trim();
  const fy = fyField.value.trim();
  const kb = (result.output_bytes.length / 1024).toFixed(0);

  results.hidden = false;
  results.innerHTML = `
    <section class="card">
      <div class="card-head">
        <div class="card-head-left">
          <div class="card-title-row"><h2>Scrutiny complete</h2></div>
          <p class="card-desc">${gstin ? `<span class="mono">${gstin}</span>` : ""}${gstin && fy ? " &middot; " : ""}${fy ? `FY ${fy}` : ""}</p>
        </div>
      </div>
      ${tiles ? `<div class="stat-row">${tiles}</div>` : ""}
      <div class="download-row">
        <div>
          <div class="fname">${result.output_name}</div>
          <div class="fmeta">${kb} KB &middot; the real master_build.py output — Master Dashboard, per-month sheets, HSN &amp; forensic checks, QA review layer</div>
        </div>
        <button class="btn btn-primary" id="download-btn" type="button">Download workbook</button>
      </div>
      <details class="run-log">
        <summary>Full run log</summary>
        <pre>${log.replace(/[<>&]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c]))}</pre>
      </details>
    </section>`;
  document.getElementById("download-btn").addEventListener("click", () => downloadBytes(result.output_bytes, result.output_name));
  results.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderScrutinyError(err) {
  const results = document.getElementById("results");
  results.hidden = false;
  results.innerHTML = `
    <section class="card">
      <div class="card-head">
        <div class="card-head-left">
          <div class="card-title-row"><h2>Scrutiny couldn't run</h2></div>
        </div>
        <span class="status-pill" data-state="error"><span class="dot"></span>Failed</span>
      </div>
      <div class="card-body">
        <div class="result-line error">${ICONS.warn}<p>${String(err.message || err).replace(/[<>&]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c]))}</p></div>
      </div>
    </section>`;
  results.scrollIntoView({ behavior: "smooth", block: "start" });
}

document.getElementById("run-btn").addEventListener("click", async function () {
  if (this.disabled) return;

  const missingRequired = RAIL_META.filter(s => s.required && !workbench.done.has(s.id));
  if (missingRequired.length > 0) {
    if (workbench.done.size === 0) {
      alert("Nothing uploaded yet. Please upload GSTR-1 and GSTR-3B — the two required inputs — before running the full scrutiny.");
    } else {
      alert(`Please also upload before running: ${missingRequired.map(s => s.title).join(", ")}.`);
    }
    return;
  }

  const originalHtml = this.innerHTML;
  this.disabled = true;
  const t0 = performance.now();
  // A full year's worth of checks genuinely takes a couple of minutes even
  // natively (measured: ~2 min for 12 months under plain CPython) — WASM
  // is slower still. Show elapsed time rather than a plain spinner, so a
  // long wait reads as "working" instead of "frozen."
  const tickHandle = setInterval(() => {
    const secs = Math.round((performance.now() - t0) / 1000);
    this.innerHTML = `${ICONS.upload} Running full scrutiny… (${secs}s — a full year typically takes 1-3 minutes, please keep this tab open)`;
  }, 1000);
  this.innerHTML = `${ICONS.upload} Running full scrutiny…`;

  const files = [];
  UPLOAD_SECTIONS.forEach(cfg => {
    const r = workbench[cfg.id];
    if (r) files.push([r.output_name, r.output_bytes]);
  });
  LEDGER_SLOTS.forEach(sl => {
    const r = workbench[`ledgers:${sl.id}`];
    if (r) files.push([r.name, r.bytes]);
  });
  MASTER_SLOTS.forEach(sl => {
    const r = workbench[`masters:${sl.id}`];
    if (r) files.push([r.name, r.bytes]);
  });

  try {
    const result = await callFullScrutiny(files, workbench.bs_pl || null);
    renderScrutinyResult(result);
  } catch (err) {
    console.error(err);
    renderScrutinyError(err);
  } finally {
    clearInterval(tickHandle);
    this.innerHTML = originalHtml;
    updateRunbar();
  }
});

// ---------------------------------------------------------------------
updateRunbar();
initRuntime();
