"use strict";

// ---------------------------------------------------------------------
// Team run notifications (Telegram), via the Cloudflare Worker in
// telegram-relay/. The bot token lives only in the Worker (a secret) --
// nothing here can talk to Telegram directly. Uploaded INPUT files are
// never sent; only the finished report and run details are.
//
// Every send is best-effort and silent: a relay/network failure is only
// console.warn'd and never blocks or slows a scrutiny run. While RELAY_URL
// is empty, all of this is a no-op.
// ---------------------------------------------------------------------

const RELAY_URL = "https://gst-telegram-relay.suraj-gst-tool.workers.dev";

const TG_DEVICE_KEY = "gst_device_id_v1";

function tgDeviceId() {
  try {
    let id = localStorage.getItem(TG_DEVICE_KEY);
    if (!id) {
      id = (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2)).slice(0, 8).toUpperCase();
      localStorage.setItem(TG_DEVICE_KEY, id);
    }
    return id;
  } catch (e) { return "unknown"; }
}

function tgBrowserInfo() {
  const ua = navigator.userAgent || "";
  const browser =
    /Edg\//.test(ua) ? "Edge" : /OPR\//.test(ua) ? "Opera" : /Chrome\//.test(ua) ? "Chrome" :
    /Firefox\//.test(ua) ? "Firefox" : /Safari\//.test(ua) ? "Safari" : "Unknown";
  const os =
    /Windows/.test(ua) ? "Windows" : /Android/.test(ua) ? "Android" : /iPhone|iPad/.test(ua) ? "iOS" :
    /Mac OS X/.test(ua) ? "macOS" : /Linux/.test(ua) ? "Linux" : "Unknown";
  let tz = "";
  try { tz = Intl.DateTimeFormat().resolvedOptions().timeZone; } catch (e) { /* ignore */ }
  return `${browser} on ${os} · ${screen.width}x${screen.height} · ${tz || "tz?"}`;
}

async function tgPost(path, body) {
  if (!RELAY_URL) return;
  const r = await fetch(RELAY_URL.replace(/\/$/, "") + path, { method: "POST", body });
  if (!r.ok) throw new Error(`relay ${path} -> HTTP ${r.status}`);
}

function tgEventForm(event, reportBytes, reportName) {
  const fd = new FormData();
  fd.append("event", JSON.stringify({ ...event, deviceId: tgDeviceId(), browser: tgBrowserInfo() }));
  if (reportBytes && reportBytes.length) fd.append("report", new Blob([reportBytes]), reportName || "GST_report.xlsx");
  return fd;
}

// Finished run: status card, then the finished report as a reply to it.
async function sendGstResultReport({ taxpayerName, gstin, fy, reportBytes, reportName, durationSecs, issues }) {
  try {
    await tgPost("/notify", tgEventForm({ kind: "run", taxpayerName, gstin, fy, durationSecs, issues, reportName }, reportBytes, reportName));
  } catch (e) { console.warn("[telemetry] result notification failed:", e); }
}

// Crashed / failed run: alert with the exception message/traceback.
async function sendTelegramErrorLog({ taxpayerName, gstin, fy, error, stage }) {
  try {
    const detail = String((error && error.message) || error || "unknown error");
    await tgPost("/notify", tgEventForm({ kind: "error", taxpayerName, gstin, fy, error: detail, stage }));
  } catch (e) { console.warn("[telemetry] error notification failed:", e); }
}

// On-demand summary of ALL runs the relay has recorded (the scheduled daily
// and monthly posts happen server-side). From the console:
//   sendTelegramSummaryReport("daily")  or  sendTelegramSummaryReport("monthly")
async function sendTelegramSummaryReport(period) {
  try {
    await tgPost("/summary", JSON.stringify({ period }));
  } catch (e) { console.warn("[telemetry] summary failed:", e); }
}

window.sendGstResultReport = sendGstResultReport;
window.sendTelegramErrorLog = sendTelegramErrorLog;
window.sendTelegramSummaryReport = sendTelegramSummaryReport;
