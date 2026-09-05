"use strict";
// Runs Pyodide off the main thread so a multi-second merge doesn't freeze
// the tab. Talks to app.js purely via postMessage — see the "call"/"status"/
// "result" message shapes below. No DOM access here; this is the only file
// that actually touches Pyodide.

importScripts("https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js");

let pyodide = null;
let ready = false;
let _callSeq = 0;

const PY_FILES = [
  "web_adapters.py",
  "ewb/auto_ewb_merger.py",
  "ewb/convert_ewb_files.py",
  "merge/gstr1/gst_merge_common.py", "merge/gstr1/merge_gstr1.py",
  "merge/gstr2a/gst_merge_common.py", "merge/gstr2a/merge_r2a.py",
  "merge/gstr2b/gst_merge_common.py", "merge/gstr2b/merge_gstr2b.py",
  "merge/gstr3b/gst_merge_common.py", "merge/gstr3b/merge_gstr3b.py",
  "merge/e invoice/gst_merge_common.py", "merge/e invoice/merge_einv.py",
  "extract_align/extractor/run.py",
  "extract_align/alligner/complete_workbooks.py",
  "core/bs_pl_input.py",
  "core/gst_blocked_credit.py",
  "core/gst_checks_flow.py",
  "core/gst_checks_forensic.py",
  "core/gst_checks_hsn_fraud.py",
  "core/gst_checks_monthly.py",
  "core/gst_config.py",
  "core/gst_core.py",
  "core/gst_machinery_scan.py",
  "core/gst_parsers_dept.py",
  "core/gst_parsers_returns.py",
  "core/gst_report.py",
  "core/gst_report_pdf.py",
  "core/master_build.py",
];

// Binary (non-.py) assets fetched the same way but written as raw bytes,
// not decoded as text -- currently just the Unicode font gst_report_pdf.py
// needs (fpdf2's built-in core fonts are Latin-1 only and crash on the
// Rupee sign, which is throughout this tool's output).
const BINARY_FILES = ["core/assets/DejaVuSans.ttf"];

function post(msg) { postMessage(msg); }

async function init() {
  try {
    post({ type: "status", state: "loading", text: "Starting Python runtime…" });
    pyodide = await loadPyodide();

    post({ type: "status", state: "loading", text: "Loading pandas, openpyxl, and friends…" });
    await pyodide.loadPackage(["pandas", "lxml", "html5lib", "xlrd"]);
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");
    await micropip.install("fpdf2");

    post({ type: "status", state: "loading", text: "Loading the merge/align/extract scripts…" });
    // cache: "reload" -- confirmed by direct testing (a page hard-reload, Ctrl+Shift+R
    // included, was NOT enough to pick up a fixed .py file: the browser kept serving an
    // HTTP-cache hit from an earlier visit indefinitely, since this static host sends no
    // Cache-Control header for these files, only Last-Modified, which Chrome's heuristic
    // caching happily reused across reloads). Without this, every returning visitor can
    // get silently stuck on stale Python logic after any future deploy, with no error and
    // no visible sign anything is wrong -- these files are small and fetched once per page
    // load, so forcing a real network fetch here is essentially free.
    for (const rel of PY_FILES) {
      const resp = await fetch(`py/${rel}`, { cache: "reload" });
      if (!resp.ok) throw new Error(`failed to fetch py/${rel}: HTTP ${resp.status}`);
      const full = `/site/py/${rel}`;
      pyodide.FS.mkdirTree(full.substring(0, full.lastIndexOf("/")));
      pyodide.FS.writeFile(full, await resp.text());
    }
    for (const rel of BINARY_FILES) {
      const resp = await fetch(`py/${rel}`, { cache: "reload" });
      if (!resp.ok) throw new Error(`failed to fetch py/${rel}: HTTP ${resp.status}`);
      const full = `/site/py/${rel}`;
      pyodide.FS.mkdirTree(full.substring(0, full.lastIndexOf("/")));
      pyodide.FS.writeFile(full, new Uint8Array(await resp.arrayBuffer()));
    }
    pyodide.FS.mkdirTree("/site/py/core");
    pyodide.FS.mkdirTree("/work");
    await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, "/site/py")
import web_adapters
`);

    ready = true;
    post({ type: "status", state: "ready", text: "Python runtime ready" });
  } catch (err) {
    console.error(err);
    post({ type: "status", state: "error", text: "Runtime failed to load — reload the page" });
  }
}

async function runPy(code, globalsObj) {
  for (const [k, v] of Object.entries(globalsObj || {})) {
    pyodide.globals.set(k, pyodide.toPy(v));
  }
  const resultProxy = await pyodide.runPythonAsync(code);
  if (resultProxy && typeof resultProxy.toJs === "function") {
    const val = resultProxy.toJs({ dict_converter: Object.fromEntries });
    resultProxy.destroy();
    return val;
  }
  return resultProxy;
}

async function callEwb(direction, filePairs) {
  const inward = direction === "inward" ? filePairs : [];
  const outward = direction === "outward" ? filePairs : [];
  return await runPy(`
inward = [(n, bytes(d)) for n, d in _inward]
outward = [(n, bytes(d)) for n, d in _outward]
result = web_adapters.process_ewb(inward, outward, _work_dir)
result["${direction}"]
`, { _inward: inward, _outward: outward, _work_dir: `/work/ewb_${direction}_${++_callSeq}` });
}

async function callMerge(kind, filePairs) {
  return await runPy(`
files = [(n, bytes(d)) for n, d in _files]
result = web_adapters.process_merge(_kind, files, _work_dir)
result
`, { _files: filePairs, _kind: kind, _work_dir: `/work/merge_${kind}_${++_callSeq}` });
}

async function callGstr3b(filePairs) {
  return await runPy(`
files = [(n, bytes(d)) for n, d in _files]
result = web_adapters.process_gstr3b(files, _work_dir)
result
`, { _files: filePairs, _work_dir: `/work/gstr3b_${++_callSeq}` });
}

async function callGstr2b(filePairs) {
  return await runPy(`
files = [(n, bytes(d)) for n, d in _files]
result = web_adapters.process_gstr2b(files, _work_dir)
result
`, { _files: filePairs, _work_dir: `/work/gstr2b_${++_callSeq}` });
}

async function callFullScrutiny(filePairs, bsPlData) {
  // runPy() already runs every value through pyodide.toPy(), so _bs_pl
  // arrives as a real Python dict (or None) — no further conversion needed.
  return await runPy(`
files = [(n, bytes(d)) for n, d in _files]
result = web_adapters.process_full_scrutiny(files, _bs_pl, _work_dir)
result
`, { _files: filePairs, _bs_pl: bsPlData || null, _work_dir: `/work/full_${++_callSeq}` });
}

async function callPdfExport(xlsxBytes) {
  return await runPy(`
result = web_adapters.process_pdf_export(bytes(_xlsx), _work_dir)
result
`, { _xlsx: xlsxBytes, _work_dir: `/work/pdf_${++_callSeq}` });
}

// Every call MUST run to full completion (including web_adapters.py's own
// os.chdir(work_dir) / os.chdir(prev_cwd) cleanup) before the next one's
// Python code starts. pyodide.globals is one shared mutable namespace, and
// os.chdir affects the whole process's cwd — if two calls' Python code were
// ever "in flight" at once, the second one's runPy() setup can silently
// overwrite the first one's inputs/cwd out from under it, corrupting both
// and hanging forever with no error (confirmed: a real user hit this by
// uploading three sections within a few seconds of each other). Chaining
// every call through one promise queue is what actually guarantees only
// one is ever mid-flight, regardless of how fast messages arrive.
let _queueTail = Promise.resolve();

onmessage = (e) => {
  const msg = e.data;
  if (msg.type !== "call") return;
  _queueTail = _queueTail.then(() => handleCall(msg));
};

async function handleCall(msg) {
  try {
    if (!ready) {
      post({ type: "result", id: msg.id, ok: false, error: "Python runtime isn't ready yet" });
      return;
    }
    // Lets the UI distinguish "queued behind another upload" from
    // "actually running" instead of a single ambiguous "Processing…".
    post({ type: "started", id: msg.id });
    let result;
    if (msg.adapter === "ewb") result = await callEwb(msg.args.direction, msg.args.filePairs);
    else if (msg.adapter === "merge") result = await callMerge(msg.args.mergeKind, msg.args.filePairs);
    else if (msg.adapter === "gstr2b") result = await callGstr2b(msg.args.filePairs);
    else if (msg.adapter === "gstr3b") result = await callGstr3b(msg.args.filePairs);
    else if (msg.adapter === "full_scrutiny") result = await callFullScrutiny(msg.args.filePairs, msg.args.bsPlData);
    else if (msg.adapter === "pdf_export") result = await callPdfExport(msg.args.xlsxBytes);
    else throw new Error(`unknown adapter: ${msg.adapter}`);
    post({ type: "result", id: msg.id, ok: true, result });
  } catch (err) {
    console.error(err);
    post({ type: "result", id: msg.id, ok: false, error: String((err && err.message) || err) });
  }
  // Never let a thrown error break out of this function — that would leave
  // _queueTail permanently rejected and silently stop every call queued
  // after it. The try/catch above already turns every failure into a
  // posted "result" message instead.
}

init();
