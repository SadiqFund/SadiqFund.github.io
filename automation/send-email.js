// ============================================================================
// ایمیل ماهانه‌ی صندوق
//
// اول هر ماه شمسی، همان ایمیلی را که سازنده‌ی محلی می‌سازد، از روی آخرین اکسل صندوق
// (همان فایلی که در کانال تلگرام است) می‌سازد و می‌فرستد.
//
// دو حالت دارد:
//   MODE=test  → فقط به ایمیل خود صندوق، با موضوعِ «[آزمایشی]». حالت پیش‌فرضِ اجرای خودکار.
//   MODE=all   → به همه‌ی اعضا در Bcc. فقط با اجرای دستی از تب Actions.
//
// حریم خصوصی:
//   - فهرست ایمیل اعضا فقط در Secret با نام MEMBER_EMAILS است و هیچ‌جا چاپ یا ذخیره نمی‌شود.
//   - همه‌ی گیرنده‌ها در Bcc می‌روند، پس هیچ عضوی ایمیل بقیه را نمی‌بیند.
//   - اکسل فقط در حافظه‌ی همین اجرا خوانده می‌شود.
//   - در لاگ فقط تعداد چاپ می‌شود، نه نشانی.
// ============================================================================
"use strict";
const fs = require("fs");
const path = require("path");
global.XLSX = require("xlsx");
const D = require("./dashboard.js");
const nodemailer = require("nodemailer");

const STATE_FILE = path.join(__dirname, "state.json"); // فقط خوانده می‌شود (شناسه‌ی آخرین اکسل)
const MAIL_STATE = path.join(__dirname, "email-state.json"); // حافظه‌ی خودِ ایمیل

const TOKEN = process.env.TELEGRAM_BOT_TOKEN || "";
const NOTIFY_ID = (process.env.NOTIFY_CHAT_ID || process.env.TELEGRAM_CHAT_ID || "").trim();
const API = (process.env.TELEGRAM_API_BASE || "https://api.telegram.org").replace(/\/$/, "");
const SITE_URL = (process.env.SITE_URL || "").trim().replace(/\/$/, "");
// نشانی صفحه‌ی اجرای همین workflow، برای دکمه‌ی پیام تلگرام (گیت‌هاب خودش این دو را می‌دهد)
const REPO = (process.env.GITHUB_REPOSITORY || "").trim();
const SERVER = (process.env.GITHUB_SERVER_URL || "https://github.com").replace(/\/$/, "");
const RUN_URL = REPO ? `${SERVER}/${REPO}/actions/workflows/monthly-email.yml` : "";

const GMAIL_USER = (process.env.GMAIL_USER || "").trim();
const GMAIL_PASS = (process.env.GMAIL_APP_PASSWORD || "").replace(/\s+/g, ""); // App Password گوگل با فاصله نشان داده می‌شود
const FROM_NAME = (process.env.EMAIL_FROM_NAME || "صندوق صادق‌اندیشان").trim();

const MODE = (process.env.MODE || "test").trim() === "all" ? "all" : "test";
const IGNORE_DATE = process.env.IGNORE_DATE === "true"; // اجرای دستی: بدون توجه به اول ماه بودن

const log = (m) => console.log(`[email] ${m}`); // فقط پیام کلی، هرگز داده
let told = false; // اگر خطا را با پیام روشن‌تری خبر داده‌ایم، دوباره خبر نمی‌دهیم

const readJSON = (f) => { try { return JSON.parse(fs.readFileSync(f, "utf8")); } catch (e) { return {}; } };
const writeJSON = (f, o) => fs.writeFileSync(f, JSON.stringify(o, null, 2) + "\n");

async function tg(method, params) {
  const res = await fetch(`${API}/bot${TOKEN}/${method}`, {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(params || {}),
  });
  const j = await res.json().catch(() => ({}));
  if (!j.ok) throw new Error(`Telegram ${method} failed (${res.status}${j.description ? ": " + j.description : ""})`);
  return j.result;
}
async function notify(text, buttonUrl, buttonText) {
  if (!TOKEN || !NOTIFY_ID) return;
  const msg = { chat_id: NOTIFY_ID, text, disable_web_page_preview: true };
  if (buttonUrl) msg.reply_markup = { inline_keyboard: [[{ text: buttonText, url: buttonUrl }]] };
  try { await tg("sendMessage", msg); }
  catch (e) { // اگر دکمه به هر دلیل قبول نشد، دست‌کم خود پیام برود
    if (!msg.reply_markup) return log("notify failed.");
    delete msg.reply_markup;
    try { await tg("sendMessage", msg); } catch (e2) { log("notify failed."); }
  }
}
async function download(fileId) {
  const f = await tg("getFile", { file_id: fileId });
  const res = await fetch(`${API}/file/bot${TOKEN}/${f.file_path}`);
  if (!res.ok) throw new Error(`file download failed (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

// امروز به وقت تهران، در تقویم شمسی
function todayJalali() {
  const f = new Intl.DateTimeFormat("en-u-ca-persian-nu-latn", {
    timeZone: "Asia/Tehran", year: "numeric", month: "numeric", day: "numeric",
  });
  const p = Object.fromEntries(f.formatToParts(new Date()).map((x) => [x.type, x.value]));
  return { jy: Number(String(p.year).replace(/\D/g, "")), jm: Number(p.month), jd: Number(p.day) };
}

// فهرست ایمیل اعضا: هر نشانی در یک خط. خط‌های خالی، خط‌های با # و متن‌های اضافه نادیده گرفته می‌شوند.
function readRecipients(text) {
  const seen = new Set();
  for (const raw of String(text || "").split(/[\r\n,;]+/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const m = line.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    if (m) seen.add(m[0].toLowerCase());
  }
  return [...seen];
}

async function main() {
  if (!GMAIL_USER || !GMAIL_PASS) throw new Error("GMAIL_USER یا GMAIL_APP_PASSWORD تنظیم نشده است.");
  if (!TOKEN) throw new Error("TELEGRAM_BOT_TOKEN تنظیم نشده است.");

  const today = todayJalali();
  const tag = `${today.jy}-${String(today.jm).padStart(2, "0")}`;
  const mail = readJSON(MAIL_STATE);

  // زمان‌بندی: فقط روز اول ماه شمسی، و هر ماه فقط یک بار
  if (!IGNORE_DATE) {
    if (today.jd !== 1) { log(`not the 1st of the month (day ${today.jd}); nothing to do.`); return; }
    if (mail.lastSent === tag) { log(`already sent for ${tag}.`); return; }
  }

  // آخرین اکسل صندوق، از همان فایلی که خودکارساز داشبورد در کانال دیده است
  const state = readJSON(STATE_FILE);
  if (!state.fund || !state.fund.file_id) throw new Error("هنوز هیچ اکسل صندوقی در کانال دیده نشده است.");
  const wb = XLSX.read(await download(state.fund.file_id), { type: "array" });
  const r = D.compute(D.parseWorkbook(wb), null);
  log(`workbook read; report date ${D.fmtDate(r.asOf)}.`);

  // کنترل‌های حسابداری: اگر یکی نخواند، ایمیل نمی‌رود
  const errs = (r.checks || []).filter((c) => c.level === "error");
  if (errs.length) {
    await notify(`⛔️ ایمیل ماهانه فرستاده نشد: ${errs.length} کنترل حسابداری در آخرین اکسل نخواند.\nاکسل تازه‌ای در کانال بگذارید و بعد از تب Actions دستی اجرا کنید.`);
    told = true;
    throw new Error(`${errs.length} accounting check(s) failed; refusing to send.`);
  }

  // هشدار داده‌ی کهنه (فقط هشدار؛ جلوی ارسال را نمی‌گیرد)
  const ageDays = Math.max(0, (Math.floor(Date.now() / 86400000) + 2440588) - r.asOf);
  const stale = ageDays > 45 ? `\n⚠️ آخرین اکسل صندوق ${ageDays} روز پیش است.` : "";

  const html = D.renderEmail(r, { siteUrl: SITE_URL });
  const text = D.renderEmailText(r, { siteUrl: SITE_URL });
  // موضوع از تاریخ گزارش ساخته می‌شود، نه از تقویم امروز: ایمیلِ اول مهر خلاصه‌ی شهریور را دارد
  const subject = `${MODE === "test" ? "[آزمایشی] " : ""}${r.fundName} — گزارش تا ${D.fmtDate(r.asOf)}`;

  const all = readRecipients(process.env.MEMBER_EMAILS);
  const bcc = MODE === "all" ? all.filter((e) => e !== GMAIL_USER.toLowerCase()) : [];
  if (MODE === "all" && !bcc.length) throw new Error("فهرست MEMBER_EMAILS خالی است.");
  log(`mode=${MODE}; recipients in bcc: ${bcc.length} (list has ${all.length}).`);

  // پیش‌فرض Gmail است؛ SMTP_* فقط برای آزمون محلی
  const port = Number(process.env.SMTP_PORT || 465);
  const tx = nodemailer.createTransport({
    host: process.env.SMTP_HOST || "smtp.gmail.com",
    port, secure: process.env.SMTP_SECURE ? process.env.SMTP_SECURE === "true" : port === 465,
    auth: { user: GMAIL_USER, pass: GMAIL_PASS },
    tls: { rejectUnauthorized: process.env.SMTP_HOST ? false : true },
  });
  await tx.sendMail({
    from: { name: FROM_NAME, address: GMAIL_USER },
    to: GMAIL_USER,       // همیشه یک نسخه به خود صندوق
    bcc,                  // اعضا؛ هیچ‌کس نشانی بقیه را نمی‌بیند
    subject, html, text,
  });

  if (!IGNORE_DATE || MODE === "all") { mail.lastSent = tag; mail.lastMode = MODE; mail.lastAt = new Date().toISOString(); writeJSON(MAIL_STATE, mail); }

  if (MODE === "test") {
    await notify(
      `📧 نسخه‌ی آزمایشی ایمیل به ایمیل صندوق فرستاده شد.` +
      `\nگزارش تا ${D.fmtDate(r.asOf)}.${stale}` +
      (RUN_URL
        ? `\n\nاگر خوب بود، دکمه‌ی زیر را بزنید و در صفحه‌ای که باز می‌شود:\nRun workflow ← mode = all ← Run workflow`
        : `\n\nاگر خوب بود، برای فرستادن به اعضا:\nتب Actions ← Monthly email ← Run workflow ← گزینه‌ی all`),
      RUN_URL, `📨 فرستادن به ${D.fmtNum(all.length)} عضو`
    );
  } else {
    await notify(`✅ ایمیل برای ${D.fmtNum(bcc.length)} عضو فرستاده شد.\nگزارش تا ${D.fmtDate(r.asOf)}.${stale}`);
  }
  log("done.");
}

main().catch(async (e) => {
  log("failed: " + (e && e.message ? e.message : e));
  if (!told) await notify(`⛔️ ایمیل ماهانه انجام نشد: ${e && e.message ? e.message : "خطای نامشخص"}`);
  process.exit(1);
});
