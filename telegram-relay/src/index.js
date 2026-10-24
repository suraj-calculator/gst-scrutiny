// Telegram relay for the GST scrutiny site. Holds the bot token (a Worker
// secret) so it never appears in the public repo, formats the messages,
// records runs for the daily/monthly summaries, and posts them on a cron.

const ALLOWED_ORIGINS = [
  "https://suraj-calculator.github.io",
  "http://localhost:8000",
  "http://127.0.0.1:8000",
];
const MAX_REPORT_BYTES = 49 * 1024 * 1024;
const MAX_EVENTS_PER_DEVICE_PER_HOUR = 60;
const IST_OFFSET_MS = 5.5 * 3600 * 1000;

const esc = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const clip = (s, n) => String(s == null ? "" : s).slice(0, n);

function cors(origin) {
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Vary": "Origin",
  };
}
const json = (obj, status, origin) =>
  new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json", ...cors(origin) } });

async function tg(env, method, body) {
  const r = await fetch(`https://api.telegram.org/bot${env.BOT_TOKEN}/${method}`, {
    method: "POST",
    ...(body instanceof FormData ? { body } : { headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }),
  });
  const j = await r.json().catch(() => ({}));
  if (!j.ok) throw new Error(`Telegram ${method}: ${j.description || r.status}`);
  return j.result;
}
const sendMessage = (env, text) =>
  tg(env, "sendMessage", { chat_id: env.CHAT_ID, text: text.slice(0, 4000), parse_mode: "HTML", disable_web_page_preview: true });

function nowIst() {
  return new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata", day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: true }) + " IST";
}
const fmtDur = s => { s = Math.round(+s || 0); return s >= 60 ? `${Math.floor(s / 60)}m ${s % 60}s` : `${s}s`; };

function deviceBlock(ev, request) {
  const cf = request.cf || {};
  const ip = request.headers.get("CF-Connecting-IP") || "?";
  const where = [cf.city, cf.region, cf.country].filter(Boolean).join(", ") || "?";
  return {
    ip,
    text:
      `🖥 <b>Device:</b> <code>${esc(clip(ev.deviceId, 16))}</code> · ${esc(clip(ev.browser, 120))}\n` +
      `🌐 <b>IP:</b> <code>${esc(ip)}</code>${cf.asOrganization ? ` (${esc(cf.asOrganization)})` : ""}\n` +
      `📍 <b>Location:</b> ${esc(where)}`,
  };
}

async function handleNotify(request, env, origin) {
  const form = await request.formData();
  const ev = JSON.parse(form.get("event") || "{}");
  const report = form.get("report");
  if (ev.kind !== "run" && ev.kind !== "error") return json({ ok: false, error: "bad kind" }, 400, origin);

  const device = clip(ev.deviceId, 16);
  const since = Date.now() - 3600 * 1000;
  const recent = await env.DB.prepare("SELECT COUNT(*) AS n FROM runs WHERE device = ? AND ts > ?").bind(device, since).first();
  if (recent && recent.n >= MAX_EVENTS_PER_DEVICE_PER_HOUR) return json({ ok: false, error: "rate limited" }, 429, origin);

  const ok = ev.kind === "run";
  const dev = deviceBlock(ev, request);
  await env.DB.prepare("INSERT INTO runs (ts, ok, gstin, name, device, ip) VALUES (?, ?, ?, ?, ?, ?)")
    .bind(Date.now(), ok ? 1 : 0, clip(ev.gstin, 15), clip(ev.taxpayerName, 200), device, dev.ip).run();

  const head =
    `🕒 <b>Time:</b> ${esc(nowIst())}\n` +
    `🏢 <b>Taxpayer:</b> ${esc(clip(ev.taxpayerName, 200) || "—")}\n` +
    `🔢 <b>GSTIN:</b> <code>${esc(clip(ev.gstin, 15) || "—")}</code>\n` +
    `📅 <b>FY:</b> ${esc(clip(ev.fy, 9) || "—")}\n`;

  if (!ok) {
    await sendMessage(env,
      `🚨 <b>GST scrutiny FAILED</b>\n\n${head}📊 <b>Status:</b> Unsuccessful${ev.stage ? ` (${esc(clip(ev.stage, 60))})` : ""}\n\n${dev.text}\n\n` +
      `<b>Error:</b>\n<pre>${esc(clip(ev.error, 6000).slice(-2200))}</pre>`);
    return json({ ok: true }, 200, origin);
  }

  const msg = await sendMessage(env,
    `✅ <b>GST scrutiny completed</b>\n\n${head}📊 <b>Status:</b> Successful${ev.issues ? ` · ${esc(clip(ev.issues, 120))}` : ""}\n` +
    `⏱ <b>Duration:</b> ${esc(fmtDur(ev.durationSecs))}\n\n${dev.text}`);
  if (report && typeof report !== "string" && report.size) {
    if (report.size > MAX_REPORT_BYTES) {
      await sendMessage(env, `⚠️ Report <code>${esc(clip(ev.reportName, 120))}</code> is over Telegram's 50 MB limit and was not attached.`);
    } else {
      const fd = new FormData();
      fd.append("chat_id", env.CHAT_ID);
      fd.append("reply_to_message_id", String(msg.message_id));
      fd.append("caption", `📎 ${clip(ev.taxpayerName || ev.gstin || "", 200)} — scrutiny report`);
      fd.append("document", report, clip(ev.reportName, 150) || "GST_report.xlsx");
      await tg(env, "sendDocument", fd);
    }
  }
  return json({ ok: true }, 200, origin);
}

// period: "daily" | "monthly" (current, so far) | "monthly-prev" (cron on the 1st)
async function sendSummary(env, period) {
  const ist = new Date(Date.now() + IST_OFFSET_MS);
  const y = ist.getUTCFullYear(), m = ist.getUTCMonth(), d = ist.getUTCDate();
  let start, end, label;
  if (period === "daily") {
    start = Date.UTC(y, m, d) - IST_OFFSET_MS; end = start + 86400000;
    label = new Date(start + IST_OFFSET_MS).toLocaleDateString("en-IN", { timeZone: "UTC", day: "2-digit", month: "short", year: "numeric" });
  } else {
    const mm = period === "monthly-prev" ? m - 1 : m;
    start = Date.UTC(y, mm, 1) - IST_OFFSET_MS; end = period === "monthly-prev" ? Date.UTC(y, m, 1) - IST_OFFSET_MS : Date.now() + 1;
    label = new Date(start + IST_OFFSET_MS).toLocaleString("en-IN", { timeZone: "UTC", month: "long", year: "numeric" });
  }
  const { results } = await env.DB.prepare("SELECT ok, gstin, name FROM runs WHERE ts >= ? AND ts < ?").bind(start, end).all();
  const total = results.length, ok = results.filter(r => r.ok).length;
  const taxpayers = [...new Set(results.filter(r => r.ok).map(r => (r.name ? `${r.name} (${r.gstin})` : r.gstin)).filter(Boolean))];
  await sendMessage(env,
    `📈 <b>${period === "daily" ? "Daily" : "Monthly"} summary — ${esc(label)}</b>\n\n` +
    `▶️ <b>Total runs:</b> ${total}\n✅ <b>Successful:</b> ${ok}\n❌ <b>Failed:</b> ${total - ok}\n` +
    `🎯 <b>Success rate:</b> ${total ? Math.round((ok / total) * 100) : 0}%` +
    (taxpayers.length ? `\n\n<b>Taxpayers done (${taxpayers.length}):</b>\n${taxpayers.slice(0, 30).map(t => "• " + esc(t)).join("\n")}` : ""));
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    if (!ALLOWED_ORIGINS.includes(origin)) return new Response("forbidden", { status: 403 });
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(origin) });
    if (request.method !== "POST") return json({ ok: false }, 405, origin);
    const path = new URL(request.url).pathname;
    try {
      if (path === "/notify") return await handleNotify(request, env, origin);
      if (path === "/summary") {
        const { period } = await request.json();
        if (period !== "daily" && period !== "monthly") return json({ ok: false, error: "bad period" }, 400, origin);
        await sendSummary(env, period);
        return json({ ok: true }, 200, origin);
      }
      return json({ ok: false }, 404, origin);
    } catch (e) {
      return json({ ok: false, error: String(e.message || e) }, 500, origin);
    }
  },
  async scheduled(event, env, ctx) {
    ctx.waitUntil(sendSummary(env, event.cron === "30 3 1 * *" ? "monthly-prev" : "daily"));
  },
};
