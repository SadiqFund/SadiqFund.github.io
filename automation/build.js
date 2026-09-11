// ============================================================================
// خودکارساز داشبورد صندوق
// هر بار اجرا: پیام‌های تازه‌ی کانال تلگرام را از بات می‌گیرد، اگر اکسل تازه‌ای آمده بود
// همان کنترل‌های سازنده را اجرا می‌کند و در صورت درست بودن، index.html را از نو می‌سازد.
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

const log = (msg) => console.log(`[dashboard] ${msg}`); // فقط پیام کلی، هرگز داده

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

// فهرست ضامن‌های مجاز: یک پیام خلاصه + فهرست خالص نام‌ها برای کپی در گزینه‌های فرم پُرس‌لاین
async function sendGuarantors(r) {
  const g = r.guarantors;
  if (!g || !NOTIFY_ID) return;
  const G = D.CONFIG.guarantor;
  const n = (x) => D.fmtInt(x);
  const list = (arr) => (arr.length ? arr.join("، ") : "—");
  await notify(
    `👥 ضامن‌های مجاز بر اساس اکسل ${D.fmtDate(r.asOf)}\n` +
    `شرط‌ها: ${G.maxLateInstallments ? `حداکثر ${n(G.maxLateInstallments)} قسط معوق` : "بدون هیچ قسط معوق"}، سرمایه‌ی شخصی دست‌کم ${D.fmtMoney(G.minCapital, true)}، کمتر از ${n(G.maxActiveGuarantees)} ضمانت وام در جریان.\n\n` +
    `✅ مجاز: ${n(g.eligible.length)} نفر (فهرست در پیام بعد)\n\n` +
    `خارج از فهرست:\n• قسط معوق (${n(g.out.late.length)}): ${list(g.out.late)}\n• سرمایه‌ی کمتر از حد (${n(g.out.capital.length)}): ${list(g.out.capital)}\n• به سقف ضمانت رسیده (${n(g.out.cap.length)}): ${list(g.out.cap)}` +
    (g.noReq ? `\n\n⚠️ برای ${n(g.noReq)} وام در جریان درخواستی در فرم پیدا نشد، پس ضامنشان معلوم نیست و در شمارش سقف ضمانت نیامده‌اند.` : "") +
    (g.unknownG ? `\n⚠️ ضامنِ ${n(g.unknownG)} وام در جریان با هیچ عضوی جور نشد (نه از روی موبایل، نه نام).` : "")
  );
  // فهرست خالص، هر نام در یک خط؛ اگر طولانی بود، در چند پیام
  const names = [...new Set(g.eligible.map((e) => e.name))]; // نام‌های تکراری (دو حساب هم‌نام) یک بار
  let chunk = [];
  const flush = async () => { if (chunk.length) { await notify(chunk.join("\n")); chunk = []; } };
  for (const nm of names) { if ((chunk.join("\n") + "\n" + nm).length > 3500) await flush(); chunk.push(nm); }
  await flush();
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
  if (!res.ok) { const err = new Error(`Porsline ${method} ${res.status}`); err.status = res.status; err.body = j; throw err; }
  return j;
}

async function findGuarantorQuestion() {
  if (PL_QUESTION) return Number(PL_QUESTION);
  const survey = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/`);
  const qs = (survey && survey.questions) || [];
  // سؤال‌هایی که «ضامن» در عنوانشان هست و گزینه دارند (کشویی/چندگزینه‌ای)
  const cands = qs.filter((q) => Array.isArray(q.choices) && /ضامن/.test(String(q.title || q.html_title || "")));
  if (cands.length === 1) {
    await notify(`ℹ️ سؤال ضامن در پُرس‌لاین پیدا شد: «${String(cands[0].title || "").slice(0, 80)}» (شناسه ${cands[0].id}). اگر درست نیست، شناسه‌ی درست را در Secret با نام PORSLINE_QUESTION_ID بگذارید.`);
    return cands[0].id;
  }
  const list = qs.filter((q) => Array.isArray(q.choices)).map((q) => `• ${q.id}: ${String(q.title || "").slice(0, 60)}`).join("\n");
  await notify(`⚠️ ${cands.length ? "چند" : "هیچ"} سؤال گزینه‌دار با کلمه‌ی «ضامن» در عنوان پیدا شد. شناسه‌ی سؤال کشویی ضامن را از این فهرست در Secret با نام PORSLINE_QUESTION_ID بگذارید:\n${list || "(سؤال گزینه‌داری پیدا نشد)"}`);
  return null;
}

async function updatePorsline(r) {
  if (!PL_KEY || !PL_SURVEY || !r.guarantors) return;
  const names = [...new Set(r.guarantors.eligible.map((e) => e.name))];
  if (names.length < 3) { // محافظ: فهرست خیلی کوتاه احتمالاً یعنی داده‌ی ناقص؛ فرم را خالی نمی‌کنیم
    await notify(`⚠️ فهرست ضامن‌های مجاز فقط ${D.fmtInt(names.length)} نفر است؛ برای احتیاط گزینه‌های فرم پُرس‌لاین عوض نشد.`);
    return;
  }
  try {
    const qid = await findGuarantorQuestion();
    if (!qid) return;
    const q = await pl("GET", `/api/v2/surveys/${PL_SURVEY}/questions/${qid}/`);
    const old = Array.isArray(q.choices) ? q.choices : [];
    const key = (x) => String(x || "").replace(/[يى]/g, "ی").replace(/ك/g, "ک").replace(/[\u200c\u200e\u200f\s]/g, "");
    const byName = new Map(old.map((c) => [key(c.name), c]));
    // نام‌هایی که از قبل بوده‌اند شناسه‌ی خودشان را نگه می‌دارند؛ نام‌های تازه بدون شناسه ساخته می‌شوند
    const choices = names.map((n) => { const c = byName.get(key(n)); return c && c.id != null ? { id: c.id, name: n } : { name: n }; });
    const same = old.length === choices.length && choices.every((c) => c.id != null) && old.every((c) => names.some((n) => key(n) === key(c.name)));
    if (same) { log("porsline: guarantor choices already up to date."); await notify("ℹ️ گزینه‌های ضامن در فرم پُرس‌لاین از قبل به‌روز بود."); return; }
    await pl("PATCH", `/api/v2/surveys/${PL_SURVEY}/questions/${qid}/`, { choices });
    const added = choices.filter((c) => c.id == null).length;
    const removed = old.filter((c) => !names.some((n) => key(n) === key(c.name))).length;
    log(`porsline: guarantor choices updated.`);
    await notify(`✅ گزینه‌های ضامن در فرم پُرس‌لاین به‌روز شد: ${D.fmtInt(names.length)} نفر (${D.fmtInt(added)} اضافه، ${D.fmtInt(removed)} حذف).`);
  } catch (e) {
    log(`porsline: update failed (${e.status || e.name}).`);
    const why = e.status === 401 || e.status === 403 ? "کلید API پذیرفته نشد یا اجازه‌ی ویرایش ندارد (شاید این کار اشتراک لازم دارد)."
      : e.status === 404 ? "پرسش‌نامه یا سؤال با این شناسه پیدا نشد."
      : e.status ? `پاسخ پُرس‌لاین: خطای ${e.status}.` : "به پُرس‌لاین وصل نشد (شاید از بیرون ایران در دسترس نیست).";
    await notify(`⚠️ گزینه‌های ضامن در فرم پُرس‌لاین به‌روز نشد: ${why}\nفهرست بالا را دستی در فرم بچسبانید.`);
  }
}

async function notify(text) {
  if (!NOTIFY_ID) return;
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

  if (!fresh.fund && !fresh.req && !FORCE) {
    if (skipped.length) await notify("⚠️ فایل تازه‌ای برای داشبورد پیدا نشد.\n\n" + skipped.map((x) => "• " + x).join("\n"));
    writeState(state);
    log("no new dashboard input; done.");
    return;
  }

  // ---------- ۳. آخرین نسخه‌ی هر فایل (تازه، یا آخرین فایلی که قبلاً دیده شده)
  let fundWb = fresh.fund && fresh.fund.wb;
  if (!fundWb && state.fund) {
    try { fundWb = XLSX.read(await download(state.fund.file_id), { type: "array" }); } catch (e) { fundWb = null; }
  }
  let reqWb = fresh.req && fresh.req.wb;
  if (!reqWb && state.req) {
    try { reqWb = XLSX.read(await download(state.req.file_id), { type: "array" }); } catch (e) { reqWb = null; }
  }
  if (fresh.req) state.req = { file_id: fresh.req.meta.file_id, file_unique_id: fresh.req.meta.file_unique_id, date: fresh.req.meta.date };

  if (!fundWb) {
    await notify("⚠️ فایل درخواست‌ها دریافت شد، ولی هنوز خروجی نرم‌افزار صندوق در دسترس نیست. داشبورد با اولین اکسل صندوق ساخته می‌شود." +
      (skipped.length ? "\n\n" + skipped.map((x) => "• " + x).join("\n") : ""));
    writeState(state);
    log("requests file stored; waiting for a fund export.");
    return;
  }

  // ---------- ۴. همان کنترل‌های سازنده
  let r, reqError = null;
  try {
    let req = null;
    if (reqWb) { try { req = D.parseRequests(reqWb); } catch (e) { reqError = e; } }
    r = D.compute(D.parseWorkbook(fundWb), req);
    if (reqError) r.checks.unshift({ level: "error", text: "فایل درخواست‌ها خوانده نشد: " + reqError.message });
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

  // ---------- ۵. ساخت و ذخیره
  const html = D.renderReport(r, FONTS);
  const prev = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, "utf8") : "";
  if (fresh.fund) state.fund = { file_id: fresh.fund.meta.file_id, file_unique_id: fresh.fund.meta.file_unique_id, date: fresh.fund.meta.date };
  state.lastBuild = { reportDate: D.fmtDate(r.asOf), at: new Date().toISOString() };
  writeState(state);

  if (html === prev) {
    log("dashboard unchanged.");
    await notify(`ℹ️ ${fresh.fund || fresh.req ? "فایل دریافت شد" : "ساخت دوباره انجام شد"}، ولی داشبورد تغییری نکرد.\nتاریخ گزارش: ${D.fmtDate(r.asOf)}`);
    if (r.guarantors) { await sendGuarantors(r); await updatePorsline(r); }
    return;
  }
  fs.writeFileSync(OUT_FILE, html);
  log("index.html rebuilt.");
  await notify(
    `✅ داشبورد به‌روز شد\nتاریخ گزارش: ${D.fmtDate(r.asOf)}\nدارایی کل: ${D.fmtMoney(r.capital, true)}` +
    (r.wait ? "" : "\n(فایل درخواست‌ها نبود؛ بخش زمان انتظار ساخته نشد)") +
    (warns.length ? "\n\nهشدارها:\n" + warns.map((c) => "• " + c.text).join("\n") : "") +
    (skipped.length ? "\n\n" + skipped.map((x) => "• " + x).join("\n") : "") +
    (SITE_URL ? `\n\nچند دقیقه‌ی دیگر روی لینک دیده می‌شود:\n${SITE_URL}` : "")
  );
  if (r.guarantors) { await sendGuarantors(r); await updatePorsline(r); }
}

main().catch(async (e) => {
  // پیام خطا ممکن است جزئی از داده داشته باشد؛ در لاگ فقط نوع خطا را می‌نویسیم
  log(`unexpected error: ${e && e.name ? e.name : "Error"}`);
  try { if (state.offset) writeState(state); } catch (x) {}
  await notify("⛔ خودکارساز داشبورد با خطای پیش‌بینی‌نشده متوقف شد. نسخه‌ی قبلی داشبورد بدون تغییر ماند.");
  process.exitCode = 1;
});
