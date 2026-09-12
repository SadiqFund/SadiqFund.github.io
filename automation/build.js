// ============================================================================
// خودکارساز داشبورد صندوق
// هر بار اجرا: پیام‌های تازه‌ی کانال تلگرام را از بات می‌گیرد و (اگر کلید پُرس‌لاین تنظیم شده باشد)
// پاسخ‌های فرم درخواست وام را مستقیم از پُرس‌لاین می‌خواند. اگر اکسل صندوق تازه‌ای آمده بود یا پاسخ‌های
// فرم تغییر کرده بود، همان کنترل‌های سازنده را اجرا می‌کند و در صورت درست بودن، index.html را از نو می‌سازد.
//
// حریم خصوصی:
//   - اکسل‌ها فقط در حافظه‌ی همین اجرا پردازش می‌شوند و هیچ‌جا ذخیره یا commit نمی‌شوند.
//   - لاگ GitHub Actions برای مخزن عمومی، برای همه قابل‌دیدن است؛ پس در لاگ فقط پیام‌های کلی
//     و بدون داده چاپ می‌شود. جزئیات کنترل‌ها (که ممکن است نام داشته باشد) فقط با بات برای شما فرستاده می‌شود.
// ============================================================================
"use strict";
const fs = require("fs");
const path = require("path");
global.XLSX = require("xlsx");
const D = require("./dashboard.js");

const ROOT = path.resolve(__dirname, "..");
const STATE_FILE = path.join(__dirname, "state.json");
const OUT_FILE = path.join(ROOT, "index.html");
const FONTS = JSON.parse(fs.readFileSync(path.join(__dirname, "fonts.json"), "utf8"));

const TOKEN = process.env.TELEGRAM_BOT_TOKEN || "";
const CHAT_ID = (process.env.TELEGRAM_CHAT_ID || "").trim(); // شناسه‌ی کانال (مثل -100…)
const NOTIFY_ID = (process.env.NOTIFY_CHAT_ID || "").trim() || CHAT_ID; // پیام نتیجه به کجا برود
const SITE_URL = (process.env.SITE_URL || "").trim();
const API = (process.env.TELEGRAM_API_BASE || "https://api.telegram.org").replace(/\/$/, "");
const FORCE = process.env.FORCE_REBUILD === "true"; // ساخت دوباره با آخرین فایل‌ها، حتی بدون فایل تازه
// پُرس‌لاین (اختیاری): اگر کلید و شناسه‌ی پرسش‌نامه باشد، گزینه‌های سؤال کشویی ضامن خودکار به‌روز می‌شود
const PL_KEY = (process.env.PORSLINE_API_KEY || "").trim();
const PL_SURVEY = (process.env.PORSLINE_SURVEY_ID || "").trim();
const PL_QUESTION = (process.env.PORSLINE_QUESTION_ID || "").trim();
const PL_API = (process.env.PORSLINE_API_BASE || "https://survey.porsline.ir").replace(/\/$/, "");
// خواندن مستقیم پاسخ‌های فرم از پُرس‌لاین؛ با همان کلید و شناسه روشن است (PORSLINE_READ_RESPONSES=false خاموشش می‌کند)
const PL_READ = !!(PL_KEY && PL_SURVEY) && process.env.PORSLINE_READ_RESPONSES !== "false";
const PL_WAIT_MS = Number(process.env.PORSLINE_EXPORT_WAIT_MS || 4000); // فاصله‌ی تلاش‌ها تا آماده شدن فایل خروجی
// هر چند ساعت یک بار پاسخ‌های فرم خوانده شود. با هر اکسل تازه‌ی صندوق یا اجرای دستی هم خوانده می‌شود.
const PL_EVERY_H = Number(process.env.PORSLINE_CHECK_HOURS || 48);
const crypto = require("crypto");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
// اثر انگشت کوتاه (برگشت‌ناپذیر) برای فهمیدن تغییر؛ خود داده هیچ‌جا ذخیره نمی‌شود
const fingerprint = (lines) => crypto.createHash("sha256").update([...lines].sort().join("\n")).digest("hex").slice(0, 16);

const log = (msg) => console.log(`[dashboard] ${msg}`); // فقط پیام کلی، هرگز داده

// نام درست اعضا از Secret با نام NAME_FIXES: خط‌های «نام در اکسل : نام درست» و «نام | سال ورود | موبایل | موبایل ۲»
const NAMES = D.loadNameFixes(process.env.NAME_FIXES || "");
if (NAMES.fixes || NAMES.roster) log(`name fixes loaded (${NAMES.fixes} manual, ${NAMES.roster} roster lines).`);

async function tg(method, params) {
  const res = await fetch(`${API}/bot${TOKEN}/${method}`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(params || {}),
  });
  const j = await res.json().catch(() => ({}));
  if (!j.ok) throw new Error(`Telegram ${method} failed (${res.status}${j.description ? ": " + j.description : ""})`);
  return j.result;
}

async function download(fileId) {
  const f = await tg("getFile", { file_id: fileId });
  const res = await fetch(`${API}/file/bot${TOKEN}/${f.file_path}`);
  if (!res.ok) throw new Error(`file download failed (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

// گزارش ضامن‌ها برای مدیر: یک پیام خلاصه، جزئیات در بخش‌های جمع‌شونده‌ی تلگرام (با لمس باز می‌شوند)
// plStatus: وضعیت گزینه‌های فرم پُرس‌لاین که در خط خلاصه می‌آید
const escH = (x) => String(x == null ? "" : x).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
async function sendGuarantors(r, plStatus) {
  const g = r.guarantors;
  if (!g || !NOTIFY_ID) return;
  const n = (x) => D.fmtInt(x);
  const rs = (g.reasons || []).filter((x) => x.names.length);
  const outN = rs.reduce((s, x) => s + x.names.length, 0);
  const days = (d) => `${n(d)} روز`;
  const money = (v) => D.fmtMoney(v, true);
  const detail = {
    late: (a) => `${n(a.late)} قسط معوق`,
    history: (a) => `بیشترین دیرکرد ${days(a.maxDelay)}`,
    inactive: (a) => (a.lastTxDays == null ? "هیچ پرداختی ندارد" : `آخرین پرداخت ${days(a.lastTxDays)} پیش از اکسل`),
    tenure: (a) => `${(a.tenureDays / 30.44).toLocaleString("fa-IR", { maximumFractionDigits: 1 })} ماه عضویت`,
    capital: (a) => `سرمایه ${money(a.capital)}`,
    debt: (a) => `مانده‌ی وام ${money(a.activeDebt)}، سرمایه ${money(a.capital)}`,
    cap: (a) => `ضامن ${n(a.guarantees)} وام از سقف ${n(a.maxGuarantees)}`,
  };
  const order = { late: (a) => -a.late, history: (a) => -a.maxDelay, inactive: (a) => -(a.lastTxDays == null ? 1e9 : a.lastTxDays),
    tenure: (a) => a.tenureDays, capital: (a) => a.capital, debt: (a) => -(a.activeDebt / Math.max(a.capital, 1)), cap: (a) => -(a.guarantees - a.maxGuarantees) };
  const uList = g.unknownGList || [];
  const dormant = uList.filter((u) => u.roster), unknown = uList.filter((u) => !u.roster);

  // خلاصه
  const gaps = [];
  if (g.noReq) gaps.push(`${n(g.noReq)} وام بی‌درخواست در فرم`);
  if (dormant.length) gaps.push(`${n(dormant.length)} ضامن بدون حساب فعال`);
  if (unknown.length) gaps.push(`${n(unknown.length)} ضامن ناشناس`);
  const head =
    `👥 <b>ضامن‌ها، اکسل ${escH(D.fmtDate(r.asOf))}</b>\n` +
    `✅ مجاز: <b>${n(g.eligible.length)} نفر</b>${plStatus ? ` · ${escH(plStatus)}` : ""}\n` +
    `⛔ خارج از فهرست: <b>${n(outN)} نفر</b>` +
    (gaps.length ? `\n⚠️ ضامن نامعلوم برای شمارش سقف: ${escH(gaps.join("، "))}` : "");

  // بخش‌های جمع‌شونده: [عنوان، خط‌ها]
  const names = [...new Set(g.eligible.map((e) => e.name))];
  const sections = [];
  sections.push([`ضامن‌های مجاز (${n(names.length)})`, names.map((nm, i) => `${n(i + 1)}. ${nm}`)]);
  for (const x of rs) {
    const lines = [...x.people].sort((p, q) => (order[x.key] ? order[x.key](p) - order[x.key](q) : 0))
      .map((a) => { const d = detail[x.key] ? detail[x.key](a) : ""; return `• ${a.name}${d ? `: ${d}` : ""}`; });
    sections.push([`⛔ ${x.label} (${n(x.names.length)})`, lines]);
  }
  if (dormant.length) sections.push([`ضامن بدون حساب فعال در صندوق (${n(dormant.length)} وام)`,
    [...dormant.map((u) => `• وام ${u.borrower}: ضامن ${u.roster}`), "این ضامن‌ها از صندوق رفته‌اند یا راکدند؛ اگر لازم است، از گیرنده‌ی وام ضامن تازه بخواهید."]]);
  if (unknown.length) sections.push([`ضامن ناشناس (${n(unknown.length)} وام)`,
    [...unknown.map((u) => `• وام ${u.borrower}: ضامن ${u.typed ? `«${u.typed}»` : "خالی"}`), "اگر صاحب شماره را می‌شناسید، یک خط «نام | سال ورود | شماره» به فهرست NAME_FIXES اضافه کنید."]]);
  sections.push(["شرط‌ها", (g.reasons || []).map((x) => "• " + x.rule)]);

  // چیدن در پیام‌ها (سقف تلگرام ۴۰۹۶ نویسه؛ با حاشیه)
  const LIMIT = 3600;
  const block = (title, lines) => `\n\n<b>${escH(title)}</b>\n<blockquote expandable>${lines.map(escH).join("\n")}</blockquote>`;
  const plainLen = (h) => h.replace(/<[^>]+>/g, "").replace(/&(amp|lt|gt);/g, "x").length;
  const msgs = [head];
  for (const [title, lines] of sections) {
    // بخش خیلی بلند به چند تکه تقسیم می‌شود
    let part = [], k = 0;
    const flushPart = () => {
      if (!part.length) return;
      const b = block(k++ ? `${title} (ادامه)` : title, part);
      if (plainLen(msgs[msgs.length - 1] + b) > LIMIT) msgs.push(b.replace(/^\n\n/, ""));
      else msgs[msgs.length - 1] += b;
      part = [];
    };
    for (const l of lines) { if (plainLen(part.concat(l).join("\n")) > LIMIT - 200) flushPart(); part.push(l); }
    flushPart();
  }
  for (const m of msgs) await notify(m, true);
}

// ---------- پُرس‌لاین: جایگزینی گزینه‌های سؤال کشویی «ضامن» با فهرست مجاز
async function pl(method, pathname, body) {
  const res = await fetch(`${PL_API}${pathname}`, {
    method,
    headers: { Authorization: `API-Key ${PL_KEY}`, "content-type": "application/json", accept: "application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let j = null; try { j = JSON.parse(text); } catch (e) {}
  if (!res.ok) { const err = new Error(`Porsline ${method} ${res.status}`); err.status = res.status; err.body = j != null ? j : text.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().slice(0, 200); throw err; }
  return j;
}

async function findGuarantorQuestion(quiet) {
  if (PL_QUESTION) return Number(PL_QUESTION);
  const survey = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/`);
  const qs = (survey && survey.questions) || [];
  // سؤال‌هایی که «ضامن» در عنوانشان هست و گزینه دارند (کشویی/چندگزینه‌ای)
  const cands = qs.filter((q) => Array.isArray(q.choices) && /ضامن/.test(String(q.title || q.html_title || "")));
  if (cands.length === 1) {
    if (!quiet) await notify(`ℹ️ سؤال ضامن در پُرس‌لاین پیدا شد: «${String(cands[0].title || "").slice(0, 80)}» (شناسه ${cands[0].id}). اگر درست نیست، شناسه‌ی درست را در Secret با نام PORSLINE_QUESTION_ID بگذارید.`);
    return cands[0].id;
  }
  const list = qs.filter((q) => Array.isArray(q.choices)).map((q) => `• ${q.id}: ${String(q.title || "").slice(0, 60)}`).join("\n");
  await notify(`⚠️ ${cands.length ? "چند" : "هیچ"} سؤال گزینه‌دار با کلمه‌ی «ضامن» در عنوان پیدا شد. شناسه‌ی سؤال کشویی ضامن را از این فهرست در Secret با نام PORSLINE_QUESTION_ID بگذارید:\n${list || "(سؤال گزینه‌داری پیدا نشد)"}`);
  return null;
}

async function updatePorsline(r, quiet) {
  if (!PL_KEY || !PL_SURVEY || !r.guarantors) return null;
  const names = [...new Set(r.guarantors.eligible.map((e) => e.name))];
  if (names.length < 3) { // محافظ: فهرست خیلی کوتاه احتمالاً یعنی داده‌ی ناقص؛ فرم را خالی نمی‌کنیم
    return { changed: false, text: `⚠️ فهرست فقط ${D.fmtInt(names.length)} نفر است؛ برای احتیاط فرم پُرس‌لاین عوض نشد` };
  }
  try {
    const qid = await findGuarantorQuestion(quiet);
    if (!qid) return { changed: false, text: "⚠️ سؤال ضامن در فرم پیدا نشد" };
    const q = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/questions/${qid}/`);
    const old = Array.isArray(q.choices) ? q.choices : [];
    const key = (x) => String(x || "").replace(/[يى]/g, "ی").replace(/ك/g, "ک").replace(/[\u200c\u200e\u200f\s]/g, "");
    const byName = new Map(old.map((c) => [key(c.name), c]));
    // نام‌هایی که از قبل بوده‌اند شناسه‌ی خودشان را نگه می‌دارند؛ نام‌های تازه بدون شناسه ساخته می‌شوند
    const choices = names.map((n) => { const c = byName.get(key(n)); return c && c.id != null ? { id: c.id, name: n } : { name: n }; });
    const same = old.length === choices.length && choices.every((c) => c.id != null) && old.every((c) => names.some((n) => key(n) === key(c.name)));
    if (same) { log("porsline: guarantor choices already up to date."); return { changed: false, text: "فرم پُرس‌لاین به‌روز است" }; }
    await pl("PATCH", `/api/v2/surveys/${PL_SURVEY}/questions/${qid}/`, { choices });
    const added = choices.filter((c) => c.id == null).length;
    const removed = old.filter((c) => !names.some((n) => key(n) === key(c.name))).length;
    log(`porsline: guarantor choices updated.`);
    return { changed: true, text: `فرم پُرس‌لاین به‌روز شد (${D.fmtInt(added)} اضافه، ${D.fmtInt(removed)} حذف)` };
  } catch (e) {
    log(`porsline: update failed (${e.status || e.name}).`);
    const why = e.status === 401 || e.status === 403 ? "کلید API پذیرفته نشد یا اجازه‌ی ویرایش ندارد (شاید این کار اشتراک لازم دارد)."
      : e.status === 404 ? "پرسش‌نامه یا سؤال با این شناسه پیدا نشد."
      : e.status ? `پاسخ پُرس‌لاین: خطای ${e.status}.` : "به پُرس‌لاین وصل نشد (شاید از بیرون ایران در دسترس نیست).";
    return { changed: false, error: true, text: `⚠️ فرم پُرس‌لاین به‌روز نشد: ${why}` };
  }
}

// ---------- پُرس‌لاین: خواندن پاسخ‌های فرم درخواست وام (فقط در حافظه‌ی همین اجرا)
// شمار پاسخ‌ها؛ درخواست سبکی که هر نیم ساعت نشان می‌دهد لازم است کل پاسخ‌ها دوباره گرفته شود یا نه
async function plCount() {
  const j = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/responses/results-table/?page_size=1`);
  const count = Number(j && j.responders_count);
  return Number.isFinite(count) ? { count, invisible: Number(j.invisible_responders_count) || 0 } : null;
}

// خروجی اکسل همه‌ی پاسخ‌ها؛ همان فایلی که از دکمه‌ی خروجی اکسل پُرس‌لاین گرفته می‌شود
async function plExport() {
  // مقدار export_format در مستندات «1/2» است، ولی پُرس‌لاین واقعی «1» را نپذیرفت؛ به‌ترتیب امتحان می‌کنیم
  let j = null, fmt = null, lastErr = null;
  for (const f of ["xlsx", "csv", "1"]) {
    try { j = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/responses/export/?export_format=${f}`); fmt = f; break; }
    catch (e) { lastErr = e; if (!(e.status === 400 && e.body && typeof e.body === "object" && "export_format" in e.body)) throw e; }
  }
  if (!fmt) throw lastErr;
  let url = j && j.export;
  if (!url) { const e = new Error("no export url"); e.kind = "format"; throw e; }
  if (url.startsWith("/")) url = PL_API + url;
  // کلید فقط برای نشانی خود پُرس‌لاین فرستاده می‌شود، نه برای میزبان دیگری که فایل را نگه می‌دارد
  const sameHost = (() => { try { return new URL(url).host === new URL(PL_API).host; } catch (e) { return false; } })();
  const nameCol = D.CONFIG.sheets.porsline.columns.name;
  let last = 0;
  for (let i = 0; i < 8; i++) {
    for (const auth of sameHost ? [false, true] : [false]) {
      const res = await fetch(url, auth ? { headers: { Authorization: `API-Key ${PL_KEY}` } } : undefined).catch(() => null);
      if (!res) continue;
      last = res.status;
      if (res.ok) {
        const buf = new Uint8Array(await res.arrayBuffer());
        if (buf[0] === 0x50 && buf[1] === 0x4b) return XLSX.read(buf, { type: "array" }); // فایل اکسل (zip)
        const text = new TextDecoder("utf-8").decode(buf).replace(/^﻿/, "");
        if (text.includes(nameCol)) { // فایل csv
          const wb = XLSX.read(text, { type: "string", raw: true });
          const name = D.CONFIG.sheets.porsline.name;
          if (wb.SheetNames[0] !== name) { wb.Sheets[name] = wb.Sheets[wb.SheetNames[0]]; delete wb.Sheets[wb.SheetNames[0]]; wb.SheetNames[0] = name; }
          return wb;
        }
        break; // هنوز آماده نیست
      }
      if (res.status !== 401 && res.status !== 403) break;
    }
    await sleep(PL_WAIT_MS);
  }
  const e = new Error(`export download ${last}`); e.kind = "download"; throw e;
}

// جدول نتایج (JSON)؛ راه دوم وقتی خروجی اکسل پُرس‌لاین خطا بدهد. به همان شکل خروجی اکسل درمی‌آید
// تا همان خواننده‌ی سازنده استفاده شود.
const PERSIAN_DT = new Intl.DateTimeFormat("en-u-ca-persian-nu-latn", { timeZone: "Asia/Tehran", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" });
function toJalaliText(v) {
  // تاریخ میلادی (مثل 2026-09-10T20:45:10+03:30) ← 1405/06/20-00:15:10 به وقت تهران
  if (typeof v !== "string" || !/^(19|20)\d{2}-\d{2}-\d{2}[T ]\d{2}:\d{2}/.test(v)) return v;
  const d = new Date(/[zZ]|[+-]\d{2}:?\d{2}$/.test(v) ? v : v.replace(" ", "T") + "+03:30");
  if (isNaN(d)) return v;
  const p = Object.fromEntries(PERSIAN_DT.formatToParts(d).map((x) => [x.type, x.value]));
  return `${p.year}/${p.month}/${p.day}-${p.hour}:${p.minute}:${p.second}`;
}
const cellText = (c) => c == null ? null
  : Array.isArray(c) ? c.map(cellText).filter((x) => x != null && x !== "").join("، ")
  : typeof c === "object" ? cellText(c.name ?? c.title ?? c.value ?? c.text ?? null)
  : toJalaliText(c);

async function plTable() {
  let header = null, total = null;
  const rows = [];
  for (let page = 1; page <= 50; page++) {
    const j = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/responses/results-table/?page=${page}&page_size=1000`);
    if (!header) header = (j.header || []).map((h) => String(cellText(h) ?? ""));
    total = Number(j.responders_count);
    let body = Array.isArray(j.body) ? j.body : [];
    if (body.length && !Array.isArray(body[0]) && (typeof body[0] !== "object" || body[0] === null)) {
      const flat = body; body = []; // ردیف‌ها پشت سر هم در یک آرایه
      for (let i = 0; i < flat.length; i += header.length) body.push(flat.slice(i, i + header.length));
    } else if (body.length && !Array.isArray(body[0])) {
      body = body.map((o) => { const arr = Object.values(o).find(Array.isArray); return arr || header.map((h, i) => o[h] ?? o[i] ?? null); });
    }
    rows.push(...body.map((r) => r.map(cellText)));
    if (!body.length || (Number.isFinite(total) && rows.length >= total)) break;
  }
  // ستون زمان پایان، اگر عنوانش با خروجی اکسل فرق داشت
  const P = D.CONFIG.sheets.porsline.columns;
  const nh = header.map((h) => h.replace(/[يى]/g, "ی").replace(/[‌\s]/g, ""));
  if (!nh.includes(P.date.replace(/\s/g, ""))) {
    const dateLike = (i) => rows.some((r) => typeof r[i] === "string" && /^1[34]\d{2}\/\d{2}\/\d{2}/.test(r[i]));
    let i = nh.findIndex((h, k) => /اتمام|پایان|ارسال|ثبت/.test(h) && dateLike(k));
    if (i === -1) i = nh.findIndex((h, k) => /تاریخ|زمان|date|time|submit/i.test(h) && dateLike(k));
    if (i !== -1) header[i] = P.date;
  }
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, D.CONFIG.sheets.porsline.name);
  wb.plHeader = header; // فقط عنوان سؤال‌ها، برای پیام خطا
  return wb;
}

// اول خروجی اکسل، اگر نشد جدول نتایج
async function plResponses() {
  const tryOne = async (via, fn) => {
    const wb = await fn();
    if (D.detectKind(wb) !== "porsline") {
      const e = new Error("format"); e.kind = "format"; e.via = via;
      if (wb.plHeader) e.header = wb.plHeader;
      throw e;
    }
    return { wb, via };
  };
  try { return await tryOne("export", plExport); }
  catch (e1) {
    e1.via = e1.via || "export";
    try { const r = await tryOne("table", plTable); r.exportErr = e1; return r; }
    catch (e2) { e2.via = e2.via || "table"; e2.first = e1; throw e2; }
  }
}

const plDetail = (e) => {
  if (!e) return "";
  const b = e.body;
  const msg = b == null ? "" : typeof b === "string" ? b : JSON.stringify(b);
  return msg ? ` پیام پُرس‌لاین: ${msg.slice(0, 250)}` : "";
};
function plWhyOne(e) {
  const where = e.via === "table" ? "جدول نتایج" : "خروجی اکسل";
  const why = e.status === 401 || e.status === 403 ? "کلید API پذیرفته نشد یا اجازه‌ی خواندن پاسخ‌ها را ندارد (شاید این کار اشتراک لازم دارد)."
    : e.status === 404 ? "پرسش‌نامه با این شناسه پیدا نشد."
    : e.kind === "format" ? "ستون‌های فرم درخواست وام پیدا نشد." + (e.header ? ` عنوان ستون‌ها: ${e.header.filter(Boolean).join(" | ").slice(0, 700)}` : "")
    : e.kind === "download" ? "فایل خروجی پاسخ‌ها آماده یا دانلود نشد."
    : e instanceof D.ReportError ? "پاسخ‌ها خوانده نشد: " + e.message
    : e.status ? `خطای ${e.status}.` : "به پُرس‌لاین وصل نشد.";
  return `${where}: ${why}${plDetail(e)}`;
}
function plWhy(e) {
  return e.first ? `${plWhyOne(e.first)} / ${plWhyOne(e)}` : plWhyOne(e);
}

// ---------- کارت بررسی هر درخواست تازه (فقط به کانال خصوصی مدیر)
const dec = (x, d = 1) => x.toLocaleString("fa-IR", { maximumFractionDigits: d });
// نام کوتاه نوع تراکنش برای کارت بررسی
const TX_SHORT = { repay: "قسط", fee: "حق عضویت", deposit: "افزایش موجودی" };
// «۶ شهریور»؛ اگر مال سال دیگری باشد، سال هم می‌آید
const txDay = (jdn, asOf) => (D.fmtDate(jdn).split(" ").pop() === D.fmtDate(asOf).split(" ").pop()
  ? D.fmtDate(jdn).split(" ").slice(0, 2).join(" ") : D.fmtDate(jdn));
const monthsOf = (days) => dec(days / 30.44);
function reviewCard(r, x) {
  const R = r.review;
  const who = R.resolve(x.mobile, x.name);
  const a = who ? R.assess(who) : null;
  const lines = [`📝 درخواست وام تازه، ثبت ${D.fmtDate(x.jdn)}`];
  const rf = !a && R.rosterFind ? R.rosterFind(x.mobile, x.name) : null;
  lines.push(`متقاضی: ${a ? a.name
    : rf ? `${rf.name} ⚠️ حساب فعالی در صندوق ندارد (در فهرست اعضا هست${rf.code ? `، ورودی ${D.fmtNum(rf.code)}` : ""})`
    : `${x.nameRaw || "(بی‌نام)"} ⚠️ با هیچ عضوی جور نشد (نه موبایل، نه نام)`}`);
  const amt = x.amount > 0 ? D.fmtMoney(x.amount, true) : x.amountRaw || "";
  if (amt || x.count) lines.push(`درخواست: ${amt || "مبلغ نامعلوم"}${x.count ? ` در ${D.fmtInt(x.count)} قسط` : ""}` +
    (a && x.amount > 0 && a.capital > 0 ? ` (${dec(x.amount / a.capital)} برابر سرمایه‌اش)` : ""));
  if (a) {
    lines.push("", `وضع متقاضی (بر اساس اکسل ${D.fmtDate(r.asOf)}):`);
    lines.push(`• سرمایه: ${D.fmtMoney(a.capital, true)}`);
    lines.push(`• عضویت: ${monthsOf(a.tenureDays)} ماه`);
    lines.push(`• قسط معوق: ${a.late ? `⚠️ ${D.fmtInt(a.late)} قسط` : "ندارد"}`);
    lines.push(`• بیشترین تأخیر قسط در ${D.fmtInt(D.CONFIG.guarantor.lateLookbackMonths)} ماه اخیر: ${a.loansN ? `${a.maxDelay > D.CONFIG.guarantor.lateDaysLimit ? "⚠️ " : ""}${D.fmtInt(a.maxDelay)} روز` : "وامی نگرفته"}`);
    lines.push(`• وام در جریان: ${a.activeLoansN ? `${D.fmtInt(a.activeLoansN)} وام، مانده ${D.fmtMoney(a.activeDebt, true)}` : "ندارد"}${a.loansN ? ` (تا حالا ${D.fmtInt(a.loansN)} وام)` : ""}`);
    // آخرین حرکت مالی خود عضو (دریافت وام و تسویه با مدیر حساب نمی‌شود)
    lines.push(`• آخرین تراکنش پرداختی: ${a.lastTxDays == null ? "⚠️ ندارد" : `${txDay(a.lastTxJdn, r.asOf)} (${TX_SHORT[a.lastTxType] || "پرداخت"})`}`);
    // برداشت از سرمایه دیگر شرط رد نیست؛ فقط یادداشت اطلاعاتی است
    if (a.withdraw) lines.push(`• ℹ️ در ${D.fmtInt(D.CONFIG.guarantor.withdrawNoteMonths)} ماه اخیر از سرمایه‌اش برداشت کرده`);
  }
  // ضامن
  if (R.hasGuarantor) {
    const g = R.resolve(x.gMobile, x.gName);
    const ga = g ? R.assess(g) : null;
    lines.push("", "ضامن: " + (!x.gName && !x.gMobile ? "⚠️ وارد نشده"
      : !ga ? (() => { const gf = R.rosterFind ? R.rosterFind(x.gMobile, x.gName) : null;
          return gf ? `${gf.name} ⛔ حساب فعالی در صندوق ندارد` : `${x.gNameRaw || "(بی‌نام)"} ⚠️ با هیچ عضوی جور نشد`; })()
      : g === who ? `${ga.name} ⛔ خود متقاضی است`
      : `${ga.name} ${ga.why ? `⛔ ${ga.whyLabel}` : "✅ مجاز"}${ga.guarantees ? ` (الان ضامن ${D.fmtInt(ga.guarantees)} وام در جریان از سقف ${D.fmtInt(ga.maxGuarantees)})` : ""}`));
  }
  // صف
  const i = R.queue.findIndex((q) => q.r === x);
  if (i !== -1) {
    lines.push("", `صف: نفر ${D.fmtInt(i + 1)} از ${D.fmtInt(R.queue.length)}` +
      (R.waitMedian != null ? `؛ زمان انتظار معمول حدود ${monthsOf(R.waitMedian)} ماه` : ""));
    if (R.queue[i].n > 1) lines.push(`ℹ️ این نفر ${D.fmtInt(R.queue[i].n)} درخواست در صف داشت؛ فقط همین آخری حساب می‌شود.`);
  } else if (who) lines.push("", "صف: این درخواست در صف حساب نشد (درخواست تازه‌تری از همین نفر هست، یا منقضی شده).");
  return lines.join("\n");
}

async function sendReviews(r) {
  const R = r.review;
  if (!R || !NOTIFY_ID) return;
  const keyOf = (x) => fingerprint([[x.mobile || x.name, x.jdn, x.gMobile || x.gName].join("|")]).slice(0, 10);
  const keys = [...new Set(R.list.map(keyOf))].sort();
  if (!Array.isArray(state.seenReq)) { // بار اول: درخواست‌های موجود دیده‌شده حساب می‌شوند، کارتی فرستاده نمی‌شود
    state.seenReq = keys;
    await notify(`ℹ️ کارت بررسی درخواست‌ها روشن شد. ${D.fmtInt(R.list.length)} درخواست قبلی دیده‌شده حساب شد؛ از این به بعد برای هر درخواست تازه کارت می‌آید.`);
    return;
  }
  const seen = new Set(state.seenReq);
  const fresh = R.list.filter((x) => !seen.has(keyOf(x))).sort((a, b) => a.jdn - b.jdn);
  state.seenReq = keys; // فقط اثر انگشت‌های کوتاه؛ نام یا شماره ذخیره نمی‌شود
  if (!fresh.length) return;
  const MAX = 10;
  for (const x of fresh.slice(-MAX)) await notify(reviewCard(r, x));
  if (fresh.length > MAX) await notify(`ℹ️ ${D.fmtInt(fresh.length - MAX)} درخواست تازه‌ی قدیمی‌تر هم بود که کارتش فرستاده نشد.`);
  log(`review: ${fresh.length} new request card(s).`);
}

async function notify(text, html) {
  if (!NOTIFY_ID) return;
  if (html) {
    try { await tg("sendMessage", { chat_id: NOTIFY_ID, text, parse_mode: "HTML", disable_web_page_preview: true }); return; }
    catch (e) { log("formatted notification rejected; sending plain text."); text = text.replace(/<[^>]+>/g, "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&"); }
  }
  const t = text.length > 3900 ? text.slice(0, 3900) + "\n…" : text;
  try { await tg("sendMessage", { chat_id: NOTIFY_ID, text: t, disable_web_page_preview: true }); }
  catch (e) { log("could not send Telegram notification"); }
}

const readState = () => { try { return JSON.parse(fs.readFileSync(STATE_FILE, "utf8")); } catch (e) { return {}; } };
const writeState = (s) => fs.writeFileSync(STATE_FILE, JSON.stringify(s, null, 2) + "\n");

let state = {};
async function main() {
  if (!TOKEN) { log("TELEGRAM_BOT_TOKEN is not set; nothing to do."); return; }
  state = readState();

  // ---------- ۱. پیام‌های تازه
  const me = await tg("getMe").catch(() => null);
  if (!me) { log("the bot token was rejected by Telegram — check the TELEGRAM_BOT_TOKEN secret."); throw new Error("bad token"); }
  const hook = await tg("getWebhookInfo").catch(() => null);
  if (hook && hook.url) { await tg("deleteWebhook").catch(() => {}); log("a webhook was set on this bot; removed it so updates can be read."); }
  const updates = await tg("getUpdates", { offset: state.offset || 0, timeout: 0, allowed_updates: ["channel_post", "message"] });
  if (updates.length) state.offset = updates[updates.length - 1].update_id + 1;
  log(`${updates.length} new update(s).`);

  // حالت کشف: هنوز شناسه‌ی کانال تنظیم نشده؛ شناسه را در همان کانال/گفت‌وگو می‌فرستیم
  if (!CHAT_ID) {
    const chats = new Map();
    for (const u of updates) { const m = u.channel_post || u.message; if (m && m.chat) chats.set(String(m.chat.id), m.chat); }
    let okCount = 0;
    for (const [id] of chats) {
      try { await tg("sendMessage", { chat_id: id, text: `شناسه‌ی این گفت‌وگو برای تنظیم خودکارساز:\n${id}\n\nاین عدد را در گیت‌هاب با نام TELEGRAM_CHAT_ID ذخیره کنید.` }); okCount++; }
      catch (e) { log("discovery: could not post the id — the bot probably lacks the 'Post Messages' admin permission."); }
    }
    if (!chats.size) log("discovery: the bot saw no messages. Make sure the bot is an ADMIN of the channel, post a NEW message in the channel after adding it, then run again.");
    log(`discovery mode: found ${chats.size} chat(s), posted the id to ${okCount}.`);
    writeState(state);
    return;
  }

  // ---------- ۲. اکسل‌های تازه‌ی همان کانال
  const docs = [];
  for (const u of updates) {
    const m = u.channel_post || u.message;
    if (!m || String(m.chat.id) !== CHAT_ID || !m.document) continue;
    const name = m.document.file_name || "";
    if (!/\.xlsx?$/i.test(name)) continue;
    docs.push({ file_id: m.document.file_id, file_unique_id: m.document.file_unique_id, name, date: m.date, size: m.document.file_size || 0 });
  }
  log(`${docs.length} new Excel file(s) in the channel.`);

  const fresh = { fund: null, req: null };
  const skipped = [];
  for (const d of docs) {
    if (d.size > 20 * 1024 * 1024) { skipped.push(`«${d.name}» بزرگ‌تر از ۲۰ مگابایت است و بات نمی‌تواند آن را دریافت کند.`); continue; }
    try {
      const wb = XLSX.read(await download(d.file_id), { type: "array" });
      const kind = D.detectKind(wb);
      if (kind === "fund") { if (!fresh.fund || d.date >= fresh.fund.meta.date) fresh.fund = { meta: d, wb }; }
      else if (kind === "requests" || kind === "porsline") { if (!fresh.req || d.date >= fresh.req.meta.date) fresh.req = { meta: d, wb }; }
      else skipped.push(`«${d.name}» نه خروجی نرم‌افزار صندوق است و نه فایل درخواست‌ها؛ نادیده گرفته شد.`);
    } catch (e) {
      skipped.push(`«${d.name}» خوانده نشد.`);
    }
  }
  // اکسل صندوقی که قبلاً رسیده ولی چون پُرس‌لاین در دسترس نبود، منتظر مانده است
  if (!fresh.fund && state.pendingFund) {
    try { fresh.fund = { meta: state.pendingFund, wb: XLSX.read(await download(state.pendingFund.file_id), { type: "array" }), pending: true }; }
    catch (e) { delete state.pendingFund; }
  }
  if (fresh.req) state.req = { file_id: fresh.req.meta.file_id, file_unique_id: fresh.req.meta.file_unique_id, date: fresh.req.meta.date };

  // ---------- ۳. پاسخ‌های فرم درخواست وام، مستقیم از پُرس‌لاین
  let api = null, apiErr = null, plChanged = false;
  const plDue = FORCE || !!fresh.fund || !state.plCheckedAt || Date.now() - Date.parse(state.plCheckedAt) >= PL_EVERY_H * 3600e3 - 10 * 60e3;
  if (PL_READ && !plDue) log("porsline: not due yet.");
  if (PL_READ && plDue) {
    try {
      let cnt = null;
      try { cnt = await plCount(); } catch (e) { cnt = null; } // اگر این یکی نشد، مستقیم کل پاسخ‌ها را می‌گیریم
      if (cnt && cnt.invisible !== (state.plInvisible || 0)) {
        if (cnt.invisible > 0) await notify(`⚠️ پُرس‌لاین ${D.fmtInt(cnt.invisible)} پاسخ فرم را به‌خاطر محدودیت پلن نشان نمی‌دهد؛ این درخواست‌ها در داشبورد و محاسبه‌ی ضامن نمی‌آیند.`);
        state.plInvisible = cnt.invisible;
      }
      const needAll = FORCE || !!fresh.fund || !cnt || cnt.count !== state.plCount || !state.plHash;
      if (needAll) {
        const got = await plResponses();
        const wb = got.wb;
        let req;
        try { req = D.parseRequests(wb); } catch (e) { e.via = got.via; if (wb.plHeader) { e.kind = "format"; e.header = wb.plHeader; } throw e; }
        if (req.badDate && !req.list.length) { const e = new Error("dates"); e.kind = "format"; e.via = got.via; e.header = wb.plHeader; throw e; }
        // اگر راه خواندن عوض شد (مثلاً خروجی اکسل خطا داد و جدول نتایج جواب داد)، یک بار خبر می‌دهیم
        if (got.via !== (state.plVia || "export")) {
          await notify(got.via === "table"
            ? `ℹ️ خروجی اکسل پُرس‌لاین جواب نداد (${plWhyOne(got.exportErr)})؛ پاسخ‌ها از جدول نتایج خوانده شد: ${D.fmtInt(req.list.length)} درخواست.`
            : "ℹ️ پاسخ‌ها دوباره از خروجی اکسل پُرس‌لاین خوانده می‌شود.");
          state.plVia = got.via;
        }
        const hash = fingerprint(req.list.map((x) => [x.mobile, x.name, x.jdn, x.gMobile, x.gName].join("|")));
        api = { wb, req };
        plChanged = hash !== state.plHash;
        state.plHash = hash; // فقط اثر انگشت؛ هیچ نام یا شماره‌ای ذخیره نمی‌شود
      }
      if (cnt) state.plCount = cnt.count;
      state.plCheckedAt = new Date().toISOString(); // اگر خطا بدهد ثبت نمی‌شود تا اجرای بعدی دوباره امتحان شود
      if (state.plFail) { delete state.plFail; await notify("✅ خواندن پاسخ‌های فرم از پُرس‌لاین دوباره برقرار شد."); }
      log(`porsline: responses ${needAll ? (plChanged ? "changed" : "unchanged") : "count unchanged"}.`);
    } catch (e) {
      apiErr = e;
      log(`porsline: reading responses failed (${e.status || e.kind || e.name}).`);
    }
    if (!apiErr && fresh.req) skipped.push(`«${fresh.req.meta.name}» استفاده نشد؛ پاسخ‌های فرم مستقیم از پُرس‌لاین خوانده می‌شود.`);
  }

  // اگر پُرس‌لاین در دسترس نبود، فایل درخواستی که در کانال گذاشته‌اید جایگزینش می‌شود
  const reqFresh = PL_READ ? plChanged || (!!apiErr && !!fresh.req) : !!fresh.req;
  if (!fresh.fund && !reqFresh && !FORCE) {
    if (apiErr && !state.plFail) {
      state.plFail = true; // فقط یک بار خبر می‌دهیم، تا وقتی دوباره برقرار شود
      await notify(`⚠️ پاسخ‌های فرم از پُرس‌لاین خوانده نشد: ${plWhy(apiErr)}\nتا وقتی درست شود، درخواست‌های تازه در داشبورد نمی‌آید. هر نیم ساعت دوباره امتحان می‌شود و وقتی برقرار شد خبر می‌دهم.`);
    }
    if (skipped.length) await notify("ℹ️ داشبورد به‌روز نشد:\n\n" + skipped.map((x) => "• " + x).join("\n"));
    writeState(state);
    log("no new dashboard input; done.");
    return;
  }

  // پُرس‌لاین در دسترس نیست و اکسل صندوق تازه آمده: به‌جای ساختن داشبورد بدون درخواست‌ها، اجرای بعدی دوباره امتحان می‌کنیم
  if (PL_READ && apiErr && !fresh.req && !FORCE) {
    if (!fresh.fund.pending) {
      await notify(`📥 اکسل صندوق رسید، ولی پاسخ‌های فرم از پُرس‌لاین خوانده نشد: ${plWhy(apiErr)}\n` +
        "داشبورد در اجرای بعدی (حدود نیم ساعت دیگر) دوباره امتحان می‌شود. اگر نمی‌خواهید صبر کنید، خروجی اکسل پاسخ‌های فرم را در کانال بگذارید، یا خودکارساز را با تیک force اجرا کنید (آخرین فایل درخواست‌هایی که در کانال بوده استفاده می‌شود).");
      state.plFail = true;
    }
    state.pendingFund = fresh.fund.meta;
    writeState(state);
    log("fund export kept pending until Porsline responses can be read.");
    return;
  }
  delete state.pendingFund;

  // ---------- ۴. آخرین نسخه‌ی هر فایل (تازه، یا آخرین فایلی که قبلاً دیده شده)
  let fundWb = fresh.fund && fresh.fund.wb;
  if (!fundWb && state.fund) {
    try { fundWb = XLSX.read(await download(state.fund.file_id), { type: "array" }); } catch (e) { fundWb = null; }
  }
  let reqWb = null, reqNote = "";
  if (api) reqWb = api.wb;
  else {
    reqWb = fresh.req && fresh.req.wb;
    if (!reqWb && state.req) {
      try { reqWb = XLSX.read(await download(state.req.file_id), { type: "array" }); } catch (e) { reqWb = null; }
    }
    if (PL_READ && apiErr) reqNote = `پاسخ‌های فرم از پُرس‌لاین خوانده نشد: ${plWhy(apiErr)} ` + (reqWb ? "به‌جایش آخرین فایل درخواست‌هایی که در کانال بود استفاده شد." : "فایل درخواستی هم در کانال نبود.");
  }

  if (!fundWb) {
    if (fresh.req) {
      await notify("⚠️ فایل درخواست‌ها دریافت شد، ولی هنوز خروجی نرم‌افزار صندوق در دسترس نیست. داشبورد با اولین اکسل صندوق ساخته می‌شود." +
        (skipped.length ? "\n\n" + skipped.map((x) => "• " + x).join("\n") : ""));
    }
    writeState(state);
    log("waiting for a fund export.");
    return;
  }

  // ---------- ۵. همان کنترل‌های سازنده
  let r, reqError = null;
  try {
    let req = api ? api.req : null;
    if (!req && reqWb) { try { req = D.parseRequests(reqWb); } catch (e) { reqError = e; } }
    r = D.compute(D.parseWorkbook(fundWb), req);
    if (reqError) r.checks.unshift({ level: "error", text: "فایل درخواست‌ها خوانده نشد: " + reqError.message });
    if (reqNote) r.checks.push({ level: "warn", text: reqNote });
  } catch (e) {
    await notify("⛔ داشبورد منتشر نشد.\n\n" + (e instanceof D.ReportError ? e.message : "فایل صندوق خوانده نشد.") + "\n\nنسخه‌ی قبلی داشبورد بدون تغییر ماند.");
    writeState(state);
    log("build failed while reading the fund export; nothing published.");
    process.exitCode = 0;
    return;
  }
  const errors = r.checks.filter((c) => c.level === "error");
  const warns = r.checks.filter((c) => c.level === "warn");

  if (errors.length || reqError) {
    await notify("⛔ داشبورد منتشر نشد؛ یکی از کنترل‌ها نخواند:\n\n" + errors.map((c) => "• " + c.text).join("\n") +
      (warns.length ? "\n\nهشدارها:\n" + warns.map((c) => "• " + c.text).join("\n") : "") +
      "\n\nنسخه‌ی قبلی داشبورد بدون تغییر ماند. اگر مطمئنید داده درست است، با سازنده‌ی روی کامپیوترتان آگاهانه منتشر کنید.");
    // فایل صندوق را به‌عنوان «آخرین فایل» ثبت نمی‌کنیم تا دفعه‌ی بعد دوباره استفاده نشود
    writeState(state);
    log(`checks failed (${errors.length} error(s)); nothing published.`);
    return;
  }

  // ---------- ۶. ساخت و ذخیره
  const html = D.renderReport(r, FONTS);
  const prev = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, "utf8") : "";
  if (fresh.fund) state.fund = { file_id: fresh.fund.meta.file_id, file_unique_id: fresh.fund.meta.file_unique_id, date: fresh.fund.meta.date };
  // گزارش کامل ضامن‌ها فقط با اکسل صندوق تازه یا وقتی فهرست مجاز عوض شده باشد؛ وگرنه یک خط در پیام داشبورد
  const gHash = r.guarantors ? fingerprint(r.guarantors.eligible.map((e) => e.name)) : null;
  const loud = !!fresh.fund || FORCE; // اجرایی که شما شروعش کرده‌اید؛ پیام کامل
  const sendG = !!r.guarantors && (!!fresh.fund || gHash !== state.gHash);
  const gLine = r.guarantors && !sendG ? `\n👥 فهرست ضامن‌ها تغییری نکرد (${D.fmtInt(r.guarantors.eligible.length)} نفر مجاز).` : "";
  if (gHash) state.gHash = gHash;
  const changed = html !== prev;
  if (changed) state.lastBuild = { reportDate: D.fmtDate(r.asOf), at: new Date().toISOString() };
  writeState(state);

  const reqCount = api ? `\n(اکنون ${D.fmtInt(api.req.list.length)} درخواست در فرم)` : "";
  if (!changed) {
    log("dashboard unchanged.");
    if (loud || fresh.req) await notify(`ℹ️ ${fresh.fund || fresh.req ? "فایل دریافت شد" : "ساخت دوباره انجام شد"}، ولی داشبورد تغییری نکرد.\nتاریخ گزارش: ${D.fmtDate(r.asOf)}` + gLine +
      (warns.length ? "\n\nهشدارها:\n" + warns.map((c) => "• " + c.text).join("\n") : ""));
    else await notify(`📝 پاسخ‌های فرم درخواست وام در پُرس‌لاین تغییر کرد؛ داشبورد تغییری نکرد.${reqCount}`);
  } else {
    fs.writeFileSync(OUT_FILE, html);
    log("index.html rebuilt.");
    await notify(
      `✅ داشبورد به‌روز شد${!fresh.fund && plChanged ? " (پاسخ تازه در فرم پُرس‌لاین)" : ""}\nتاریخ گزارش: ${D.fmtDate(r.asOf)}\nدارایی کل: ${D.fmtMoney(r.capital, true)}` + gLine +
      (!fresh.fund && plChanged ? reqCount : "") +
      (r.wait ? "" : "\n(فایل درخواست‌ها نبود؛ بخش زمان انتظار ساخته نشد)") +
      (warns.length ? "\n\nهشدارها:\n" + warns.map((c) => "• " + c.text).join("\n") : "") +
      (skipped.length ? "\n\n" + skipped.map((x) => "• " + x).join("\n") : "") +
      (SITE_URL ? `\n\nچند دقیقه‌ی دیگر روی لینک دیده می‌شود:\n${SITE_URL}` : "")
    );
  }
  await sendReviews(r);
  writeState(state);
  // اول فرم پُرس‌لاین به‌روز می‌شود تا وضعیتش در خلاصه‌ی گزارش ضامن‌ها بیاید
  const plStatus = r.guarantors ? await updatePorsline(r, !loud) : null;
  if (sendG) await sendGuarantors(r, plStatus && plStatus.text);
  else if (plStatus && (plStatus.changed || plStatus.error)) await notify(`👥 ${plStatus.text}`);
}

main().catch(async (e) => {
  // پیام خطا ممکن است جزئی از داده داشته باشد؛ در لاگ فقط نوع خطا را می‌نویسیم
  log(`unexpected error: ${e && e.name ? e.name : "Error"}`);
  try { if (state.offset) writeState(state); } catch (x) {}
  await notify("⛔ خودکارساز داشبورد با خطای پیش‌بینی‌نشده متوقف شد. نسخه‌ی قبلی داشبورد بدون تغییر ماند.");
  process.exitCode = 1;
});
