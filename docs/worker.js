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
];

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

    post({ type: "status", state: "loading", text: "Loading the merge/align/extract scripts…" });
    for (const rel of PY_FILES) {
      const resp = await fetch(`py/${rel}`);
      if (!resp.ok) throw new Error(`failed to fetch py/${rel}: HTTP ${resp.status}`);
      const full = `/site/py/${rel}`;
      pyodide.FS.mkdirTree(full.substring(0, full.lastIndexOf("/")));
      pyodide.FS.writeFile(full, await resp.text());
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

onmessage = async (e) => {
  const msg = e.data;
  if (msg.type !== "call") return;
  if (!ready) {
    post({ type: "result", id: msg.id, ok: false, error: "Python runtime isn't ready yet" });
    return;
  }
  try {
    let result;
    if (msg.adapter === "ewb") result = await callEwb(msg.args.direction, msg.args.filePairs);
    else if (msg.adapter === "merge") result = await callMerge(msg.args.mergeKind, msg.args.filePairs);
    else if (msg.adapter === "gstr2b") result = await callGstr2b(msg.args.filePairs);
    else if (msg.adapter === "gstr3b") result = await callGstr3b(msg.args.filePairs);
    else throw new Error(`unknown adapter: ${msg.adapter}`);
    post({ type: "result", id: msg.id, ok: true, result });
  } catch (err) {
    console.error(err);
    post({ type: "result", id: msg.id, ok: false, error: String((err && err.message) || err) });
  }
};

init();
