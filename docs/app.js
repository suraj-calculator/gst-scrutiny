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
const ANNUAL_DOC_SLOTS = [
  { id: "bo_profile", name: "BO / 360° Profile", ext: "XLSX" },
  { id: "gstr9", name: "GSTR-9 Annual Return", ext: "XLSX" },
  { id: "gstr9c", name: "GSTR-9C Reconciliation", ext: "XLSX" },
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
  { id: "annual_pdfs", step: 9, title: "Annual Reports", required: false },
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
// "waiting for runtime" — a single place to look. Pyodide + pandas/openpyxl
// take ~5-10s to load (once per visit); instead of a stream of "loading
// step N" status text, show one random GST fact for the whole wait — no
// rotation/timer, just a single pick per page load.
const GST_FACTS = [
  "GST launched at midnight on 1 July 2017, replacing 17 separate central and state taxes — excise duty, service tax, VAT and more — with one.",
  "GST runs on 4 main slabs — 0%, 5%, 12%, 18% and 28% — with an extra cess on top of 28% for luxury and \"sin\" goods like tobacco and aerated drinks.",
  "A transaction within a state splits GST into CGST + SGST; across states it's IGST — same total rate, different pockets it lands in.",
  "The GST Council — the body that sets rates and rules — is chaired by the Union Finance Minister, with a finance minister from every state as a member.",
  "GST was enabled by the 101st Constitutional Amendment Act, 2016, which inserted Article 246A giving both Parliament and state legislatures power to tax GST.",
  "Input Tax Credit (ITC) is the backbone of GST: tax paid on purchases can be set off against tax owed on sales, so tax is only ever paid on the value actually added.",
  "An e-way bill is mandatory for moving goods worth more than ₹50,000 — it's what the EWB Pattern and Timeline checks in this tool are built around.",
  "GST registration is mandatory once turnover crosses ₹40 lakh for goods or ₹20 lakh for services (lower thresholds apply in some special-category states).",
  "The Composition Scheme lets small taxpayers (turnover up to ₹1.5 crore) pay GST at a small flat rate instead of the regular slab rates — at the cost of not being able to claim ITC.",
  "Petroleum products, alcohol for human consumption, and electricity are still outside GST — they continue under the old excise/VAT regime.",
  "1 July is celebrated as GST Day in India, marking the day the country moved to a single national tax on goods and services.",
  "Gold and gold jewellery attract a special GST rate of 3% — lower than almost anything else, reflecting how price-sensitive the category is.",
  "GSTR-1 reports outward supplies (sales), GSTR-3B is the summary return with tax payment, and GSTR-9 is the annual return that reconciles the whole year.",
  "Under Reverse Charge Mechanism (RCM), the buyer — not the seller — pays the GST directly to the government for certain notified goods and services.",
  "Exports under GST are \"zero-rated\" — taxed at 0%, with exporters still allowed to claim credit for tax paid on their inputs.",
  "HSN (Harmonised System of Nomenclature) codes classify goods for GST — the same coding system used in over 200 countries for customs.",
  "E-invoicing — real-time reporting of B2B invoices to a government portal — is mandatory once a business crosses a notified turnover threshold, to curb fake invoicing.",
  "GSTIN, the 15-digit GST registration number, encodes the state code in its first two digits and the PAN of the business in the next ten.",
];
function showRandomBootFact() {
  const el = document.getElementById("boot-message");
  el.textContent = GST_FACTS[Math.floor(Math.random() * GST_FACTS.length)];
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
    overlay.dataset.hidden = "true";
  } else if (state === "error") {
    document.getElementById("boot-spinner").dataset.state = "error";
    document.getElementById("boot-message").textContent = "Something went wrong";
    const sub = document.getElementById("boot-sub");
    sub.textContent = text;
    sub.classList.add("error");
  }
  // "loading" state: overlay already visible with its one random fact from
  // showRandomBootFact(), called once in initRuntime() — nothing to do here.
}

function initRuntime() {
  setRuntimeState("loading", "Starting Python runtime…");
  showRandomBootFact();
  // Cache-busting query string -- confirmed by direct testing that a plain
  // `new Worker("worker.js")` can keep instantiating a stale cached copy of
  // the WORKER SCRIPT ITSELF indefinitely (this static host sends no
  // Cache-Control header, only Last-Modified, which the browser's own
  // heuristic caching happily reuses across reloads, hard-reload included --
  // a returning visitor can be silently stuck running old worker.js/Python
  // logic after a deploy, with no error or visible sign anything is wrong).
  // A query string makes every page load request a URL the cache has never
  // seen, without touching how the worker resolves its own relative fetch()
  // calls (those resolve against the path, not the query string).
  worker = new Worker(`worker.js?_=${Date.now()}`);
  worker.onmessage = e => {
    const msg = e.data;
    if (msg.type === "status") {
      setRuntimeState(msg.state, msg.text);
      if (msg.state === "ready") {
        workerReady = true;
        enableAllDropzones();
        updateRunbar();
      }
    } else if (msg.type === "started") {
      const pending = pendingCalls.get(msg.id);
      if (pending && pending.onStarted) pending.onStarted();
    } else if (msg.type === "result") {
      const pending = pendingCalls.get(msg.id);
      if (!pending) return; // already timed out and cleaned up — ignore
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

// Every call gets an explicit timeout — with the worker now processing one
// call at a time (see worker.js's queue), a truly hung call would otherwise
// show "Processing…" forever with zero feedback. Individual sections
// normally finish in well under a minute.
//
// full_scrutiny's budget was originally 10 min, based on an early ~117s
// native-CPython measurement from BEFORE GSTR-2A and the Annual Reports
// (BO Profile/GSTR-9/9C) paths were wired in. With those enabled on a real
// full 12-month run, native CPython alone measured 312s (~5.2 min) for a
// smaller taxpayer and ~23 min for a larger one -- and this genuinely
// legitimate, non-hung processing was hitting the old 10-minute timeout
// in-browser (WASM has no JIT and is measurably slower than native CPython
// on top of that), producing a false "genuinely stuck" error on a run that
// was simply still working. Raised well above the slowest real run
// observed so far, so an actual hang is still eventually caught without
// false-failing correct-but-slow runs.
//
// The timeout clock starts when the worker actually begins running this
// call (the "started" message), NOT when it's queued — a call queued
// behind other heavy uploads can legitimately wait a while before its turn
// even comes, and that wait must never eat into its own processing budget.
// (Bug found and fixed after exactly that happened: three 12-file uploads
// queued together each timed out because the clock had already been
// running since they were queued, not since they actually started.)
const DEFAULT_TIMEOUT_MS = 5 * 60 * 1000;
const FULL_SCRUTINY_TIMEOUT_MS = 45 * 60 * 1000;

function callWorker(adapter, args, opts) {
  opts = opts || {};
  const timeoutMs = opts.timeoutMs || DEFAULT_TIMEOUT_MS;
  return new Promise((resolve, reject) => {
    const id = ++_reqSeq;
    let timer = null;
    pendingCalls.set(id, {
      resolve: v => { clearTimeout(timer); resolve(v); },
      reject: e => { clearTimeout(timer); reject(e); },
      onStarted: () => {
        timer = setTimeout(() => {
          pendingCalls.delete(id);
          reject(new Error(
            `Timed out after ${Math.round(timeoutMs / 1000)}s of actual processing with no ` +
            `response. This means the Python code itself is genuinely stuck (not just queued) ` +
            `— reload the page and try again with fewer files, or report this.`
          ));
        }, timeoutMs);
        if (opts.onStarted) opts.onStarted();
      },
    });
    worker.postMessage({ type: "call", id, adapter, args });
  });
}

async function callEwb(direction, filePairs, onStarted) {
  return await callWorker("ewb", { direction, filePairs }, { onStarted });
}
async function callMerge(kind, filePairs, onStarted) {
  return await callWorker("merge", { mergeKind: kind, filePairs }, { onStarted });
}
async function callGstr3b(filePairs, onStarted) {
  return await callWorker("gstr3b", { filePairs }, { onStarted });
}
async function callGstr2b(filePairs, onStarted) {
  return await callWorker("gstr2b", { filePairs }, { onStarted });
}
async function callFullScrutiny(filePairs, bsPlData, onStarted) {
  return await callWorker("full_scrutiny", { filePairs, bsPlData }, { onStarted, timeoutMs: FULL_SCRUTINY_TIMEOUT_MS });
}
async function callPdfExport(xlsxBytes, onStarted) {
  // Reuses the long timeout — rendering every sheet of a full year's
  // workbook to PDF is comparable in cost to the scrutiny run itself
  // (~1 min native CPython on a real 92-sheet/127k-row workbook; budget
  // more in-browser WASM).
  return await callWorker("pdf_export", { xlsxBytes }, { onStarted, timeoutMs: FULL_SCRUTINY_TIMEOUT_MS });
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
          <div class="slot" data-slot="${id}:${sl.id}" ${opts.disabled ? 'data-disabled="true"' : `tabindex="0" role="button"`} aria-label="Upload ${sl.name}">
            <span class="slot-icon">${ICONS.file}</span>
            <span>
              <span class="slot-name">${sl.name}</span>
              <span class="slot-sub"> &middot; ${sl.ext}</span>
            </span>
            <span class="slot-status">${opts.disabled ? "—" : "Empty"}</span>
            ${opts.disabled ? "" : `<button class="slot-clear" type="button" data-slotclear="${id}:${sl.id}">Clear</button><input type="file" data-slotinput="${id}:${sl.id}">`}
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
  slotSectionHtml("annual_pdfs", "Annual Reports",
    "BO / 360° Profile, GSTR-9 and GSTR-9C — as the structured Excel export, not the PDF. (The PDF library this would need has no WebAssembly build, so it can't run in-browser — export these three from the portal as Excel first.)",
    ANNUAL_DOC_SLOTS) +
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
  // marker (cleared here once the runtime is ready). Slot-based sections
  // (Ledgers, Annual Reports, Reference Masters) use .slot elements, not
  // .dropzone — plain byte capture, no worker processing — so they're never
  // matched by this selector and are usable immediately, runtime or not.
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
function escapeHtml(s) {
  return String(s).replace(/[<>&]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c]));
}

// `detail`, when given, is the complete raw error text (often a real Python
// traceback from the actual script) — shown in full, not truncated, so the
// user can see exactly what went wrong rather than a vague summary.
function showError(cfg, message, detail) {
  setStatus(cfg.id, "error", "Failed");
  const detailHtml = detail && detail !== message
    ? `<pre class="error-detail">${escapeHtml(detail)}</pre>`
    : "";
  document.querySelector(`[data-result="${cfg.id}"]`).innerHTML = `
    <div class="result-line error">
      ${ICONS.warn}
      <p>${escapeHtml(message)}</p>
    </div>
    ${detailHtml}`;
  markUndone(cfg.id);
}

function showResult(cfg, r, extraLine) {
  if (!r) {
    showError(cfg, `No ${cfg.title}-shaped files were detected among what you uploaded. Double-check these are the right export type.`);
    return;
  }
  setStatus(cfg.id, "ready", "Ready");
  const kb = (r.output_bytes.length / 1024).toFixed(0);
  const resultBox = document.querySelector(`[data-result="${cfg.id}"]`);
  resultBox.innerHTML = `
    <div class="result-line">
      ${ICONS.check}
      <p>${extraLine ? extraLine + "<br>" : ""}<span class="out-file">→ ${r.output_name}</span> (${kb} KB)</p>
    </div>
    <div class="result-actions">
      <button class="btn small ghost" type="button" data-action="download">Download merged file</button>
      <button class="btn small danger" type="button" data-action="clear">Clear</button>
    </div>`;
  workbench[cfg.id] = r;
  // r.output_bytes/r.output_name are closed over here rather than re-read
  // from workbench[cfg.id] on click, so a later Clear (which deletes that
  // key) can never turn this button into a dead click.
  resultBox.querySelector('[data-action="download"]').addEventListener("click", () => downloadBytes(r.output_bytes, r.output_name));
  resultBox.querySelector('[data-action="clear"]').addEventListener("click", () => clearUploadSection(cfg));
  markDone(cfg.id);
}

// Resets one merge-type section back to empty — e.g. after realising the
// wrong files were dropped in — without touching any other section's
// progress or forcing a full page reload.
function clearUploadSection(cfg) {
  delete workbench[cfg.id];
  document.querySelector(`[data-result="${cfg.id}"]`).innerHTML = "";
  document.querySelector(`[data-filecount="${cfg.id}"]`).textContent = "";
  const progWrap = document.querySelector(`[data-progress="${cfg.id}"]`);
  if (progWrap) progWrap.innerHTML = "";
  const input = document.querySelector(`[data-input="${cfg.id}"]`);
  if (input) input.value = "";
  setStatus(cfg.id, "empty", "Not started");
  markUndone(cfg.id);
}

async function processSection(cfg, filePairs) {
  if (!workerReady) return;
  // Shown immediately and left in place through processing/success/failure,
  // so the user always has a plain confirmation of how many files they
  // actually selected, independent of whatever the processing result was.
  document.querySelector(`[data-filecount="${cfg.id}"]`).textContent =
    `${filePairs.length} file${filePairs.length === 1 ? "" : "s"} uploaded`;
  // "Queued…" until the worker actually starts on this call — the worker
  // processes one call at a time (see worker.js), so a section uploaded
  // while another is still running will genuinely wait its turn rather
  // than run at the same time (which used to silently corrupt both).
  setStatus(cfg.id, "processing", "Queued…");
  document.querySelector(`[data-result="${cfg.id}"]`).innerHTML = "";
  const progWrap = document.querySelector(`[data-progress="${cfg.id}"]`);
  progWrap.innerHTML = "";

  const onStarted = () => {
    setStatus(cfg.id, "processing", "Processing…");
    progWrap.innerHTML = `<div class="progress-wrap"><div class="progress-track"><div class="progress-fill" data-fill></div></div><div class="progress-label">Running the real merge script on ${filePairs.length} file${filePairs.length === 1 ? "" : "s"}…</div></div>`;
    requestAnimationFrame(() => { const f = progWrap.querySelector("[data-fill]"); if (f) f.style.width = "100%"; });
  };

  try {
    let result, extraLine;
    if (cfg.py.kind === "ewb") {
      result = await callEwb(cfg.py.direction, filePairs, onStarted);
      extraLine = result ? `${result.rows} rows merged` : null;
    } else if (cfg.py.kind === "merge") {
      result = await callMerge(cfg.py.mergeKind, filePairs, onStarted);
    } else if (cfg.py.kind === "gstr2b") {
      result = await callGstr2b(filePairs, onStarted);
    } else if (cfg.py.kind === "gstr3b") {
      result = await callGstr3b(filePairs, onStarted);
    }
    progWrap.innerHTML = "";
    showResult(cfg, result, extraLine);
  } catch (err) {
    console.error(err);
    progWrap.innerHTML = "";
    const full = String(err.message || err);
    const firstLine = full.split("\n")[0];
    const headline = firstLine.length > 160 ? firstLine.slice(0, 160) + "…" : firstLine;
    // Only attach the detail block when there's genuinely more to show —
    // a multi-line traceback or a headline that got truncated — so a
    // plain one-line error doesn't get a redundant duplicate underneath.
    showError(cfg, `Error while processing: ${headline}`, full.length > firstLine.length ? full : null);
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

    // Resets just this one slot (e.g. the wrong file was dropped in) —
    // stopPropagation so the click doesn't also bubble up to the parent
    // .slot's own click-to-reopen-the-file-picker handler below.
    function clear() {
      delete workbench[key];
      doneSlots.delete(sl.id);
      delete el.dataset.done;
      el.querySelector(".slot-icon").innerHTML = ICONS.file;
      el.querySelector(".slot-status").textContent = "Empty";
      input.value = "";
      setStatus(sectionId, doneSlots.size === 0 ? "empty" : "processing", doneSlots.size === 0 ? "Not started" : `${doneSlots.size} of ${slots.length}`);
      if (doneSlots.size === 0) markUndone(sectionId);
      if (opts.persist) removePersistedMaster(sl.id);
    }

    el.addEventListener("click", () => input.click());
    el.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
    input.addEventListener("change", () => { if (input.files[0]) complete(input.files[0]); });
    const clearBtn = document.querySelector(`[data-slotclear="${key}"]`);
    if (clearBtn) clearBtn.addEventListener("click", e => { e.stopPropagation(); clear(); });
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
wireSlotSection("annual_pdfs", ANNUAL_DOC_SLOTS);
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
function removePersistedMaster(slotId) {
  try {
    const store = JSON.parse(localStorage.getItem(MASTER_STORAGE_KEY) || "{}");
    delete store[slotId];
    localStorage.setItem(MASTER_STORAGE_KEY, JSON.stringify(store));
  } catch (err) {
    console.warn("could not remove persisted master file locally", err);
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
const MIME_TYPES = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pdf: "application/pdf",
};
function downloadBytes(bytes, filename) {
  const ext = filename.split(".").pop().toLowerCase();
  const blob = new Blob([bytes], { type: MIME_TYPES[ext] || "application/octet-stream" });
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
        <div class="download-btns">
          <button class="btn btn-primary" id="download-btn" type="button">Download Excel</button>
          <button class="btn" id="download-pdf-btn" type="button">Download PDF</button>
        </div>
      </div>
      <details class="run-log">
        <summary>Full run log</summary>
        <pre>${log.replace(/[<>&]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c]))}</pre>
      </details>
    </section>`;
  document.getElementById("download-btn").addEventListener("click", () => downloadBytes(result.output_bytes, result.output_name));

  const pdfBtn = document.getElementById("download-pdf-btn");
  const pdfName = result.output_name.replace(/\.xlsx$/i, ".pdf");
  let pdfCache = null; // generated lazily, once — re-clicking re-downloads the same bytes instead of re-rendering
  pdfBtn.addEventListener("click", async () => {
    if (pdfCache) { downloadBytes(pdfCache, pdfName); return; }
    const original = pdfBtn.innerHTML;
    pdfBtn.disabled = true;
    let tick = null;
    const onStarted = () => {
      const t0 = performance.now();
      tick = setInterval(() => {
        const secs = Math.round((performance.now() - t0) / 1000);
        pdfBtn.innerHTML = `Rendering PDF… (${secs}s)`;
      }, 1000);
    };
    pdfBtn.innerHTML = "Queued…";
    try {
      const pdfResult = await callPdfExport(result.output_bytes, onStarted);
      pdfCache = pdfResult.output_bytes;
      downloadBytes(pdfCache, pdfName);
    } catch (err) {
      console.error(err);
      toast(`PDF export failed: ${String(err.message || err).split("\n")[0]}`);
    } finally {
      if (tick) clearInterval(tick);
      pdfBtn.disabled = false;
      pdfBtn.innerHTML = original;
    }
  });
  results.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderScrutinyError(err) {
  const results = document.getElementById("results");
  const full = String(err.message || err);
  const firstLine = full.split("\n")[0];
  const headline = firstLine.length > 200 ? firstLine.slice(0, 200) + "…" : firstLine;
  const detailHtml = full.length > firstLine.length
    ? `<pre class="error-detail">${escapeHtml(full)}</pre>`
    : "";
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
        <div class="result-line error">${ICONS.warn}<p>${escapeHtml(headline)}</p></div>
        ${detailHtml}
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
  // Queued behind any still-running upload until the worker actually starts
  // on this call (see worker.js's queue) — the elapsed-time clock only
  // starts once it's genuinely executing, not while it's waiting its turn.
  this.innerHTML = `${ICONS.upload} Queued — waiting for other uploads to finish…`;
  let tickHandle = null;
  const onStarted = () => {
    const t0 = performance.now();
    // A full year's worth of checks, with GSTR-2A + Annual Reports (BO
    // Profile/GSTR-9/9C) also supplied, has measured 5-25+ minutes under
    // plain native CPython depending on data volume -- WASM (no JIT) is
    // slower still. Show elapsed time rather than a plain spinner, so a
    // long wait reads as "working" instead of "frozen," and set an honest
    // expectation instead of the old "1-3 minutes" estimate (measured before
    // GSTR-2A/Annual Reports existed, on a smaller check set -- it was
    // actively causing real, correctly-running scrutiny to look stuck).
    tickHandle = setInterval(() => {
      const secs = Math.round((performance.now() - t0) / 1000);
      this.innerHTML = `${ICONS.upload} Running full scrutiny… (${secs}s — a full year with every optional source can take 10-20+ minutes in-browser, please keep this tab open)`;
    }, 1000);
    this.innerHTML = `${ICONS.upload} Running full scrutiny…`;
  };

  const files = [];
  UPLOAD_SECTIONS.forEach(cfg => {
    const r = workbench[cfg.id];
    if (r) files.push([r.output_name, r.output_bytes]);
  });
  LEDGER_SLOTS.forEach(sl => {
    const r = workbench[`ledgers:${sl.id}`];
    if (r) files.push([r.name, r.bytes]);
  });
  ANNUAL_DOC_SLOTS.forEach(sl => {
    const r = workbench[`annual_pdfs:${sl.id}`];
    if (r) files.push([r.name, r.bytes]);
  });
  MASTER_SLOTS.forEach(sl => {
    const r = workbench[`masters:${sl.id}`];
    if (r) files.push([r.name, r.bytes]);
  });

  try {
    const result = await callFullScrutiny(files, workbench.bs_pl || null, onStarted);
    renderScrutinyResult(result);
  } catch (err) {
    console.error(err);
    renderScrutinyError(err);
  } finally {
    if (tickHandle) clearInterval(tickHandle);
    this.innerHTML = originalHtml;
    updateRunbar();
  }
});

// ---------------------------------------------------------------------
updateRunbar();
initRuntime();
