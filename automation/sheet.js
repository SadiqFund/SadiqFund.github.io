// ============================================================================
// پل به سند گوگل «داشبورد اطلاعات صندوق» (از راه Apps Script Web App)
//
// سه کار:
//   loanGuarantors() → دفتر وام‌های سند: نام گیرنده + تاریخ پرداخت + ضامن.
//                      برای وام‌هایی که درخواستشان در فرم پُرس‌لاین نیست (وام‌های قدیمی‌تر).
//   history(r)       → یک ردیف در تب «تاریخچه»: ارقام کلی همان اکسل. سند پشتیبان صندوق.
//   queue(rows)      → درخواست‌های تازه‌ی فرم را به تب «صف وام» اضافه می‌کند.
//
// اگر SHEET_URL یا SHEET_TOKEN تنظیم نشده باشد، همه‌ی این‌ها بی‌صدا رد می‌شوند و
// خودکارساز مثل قبل کار می‌کند. هیچ خطایی در این فایل نباید ساخت داشبورد را متوقف کند.
// ============================================================================
"use strict";

const URL_ = (process.env.SHEET_URL || "").trim();
const TOKEN = (process.env.SHEET_TOKEN || "").trim();
const ON = !!(URL_ && TOKEN);

const log = (m) => console.log(`[sheet] ${m}`); // فقط پیام کلی، هرگز داده

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const TRIES = 3;
const WAIT = [1500, 4000]; // فاصله‌ی تلاش‌ها
const faN = (n) => String(n).replace(/\d/g, (d) => "۰۱۲۳۴۵۶۷۸۹"[d]);

// یک تماس با سند.
//
// چرا تلاش دوباره؟ Apps Script پاسخ را با یک redirect به script.googleusercontent.com می‌دهد و
// همان مرحله گاهی بی‌دلیل ۴۰۴ (یا صفحه‌ی HTML) برمی‌گرداند — گاهی برای یک تماس و نه بقیه، در
// همان اجرا و با همان نشانی. این خطا گذراست، پس تا سه بار تلاش می‌کنیم.
//
// تلاش دوباره امن است چون هر سه نوشتن «یک‌بار-اثر»اند: تاریخچه با تاریخ گزارش همان ردیف را
// به‌روز می‌کند، و وام و صف با «نام + تاریخ» تکراری نمی‌سازند. پس اگر تماس اول در واقع نوشته
// باشد و فقط پاسخش گم شده باشد، تلاش دوم چیزی دوباره اضافه نمی‌کند.
async function call(action, extra) {
  if (!ON) return null;
  const body = JSON.stringify({ token: TOKEN, action, ...(extra || {}) });
  let last = null;
  for (let i = 0; i < TRIES; i++) {
    if (i) await sleep(WAIT[i - 1] || 4000);
    let res, text;
    try {
      res = await fetch(URL_, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body,
        redirect: "follow", // Apps Script یک بار به script.googleusercontent.com هدایت می‌کند
      });
      text = await res.text();
    } catch (e) { last = { status: 0, msg: "به سند وصل نشد" }; continue; } // شبکه؛ دوباره امتحان می‌کنیم
    let j = null;
    try { j = JSON.parse(text); } catch (e) { j = null; }
    if (j) {
      if (j.error) throw new Error(j.error); // پاسخ واقعی اسکریپت؛ تلاش دوباره فایده ندارد
      if (i) log(`call "${action}" succeeded on attempt ${i + 1}.`);
      return j;
    }
    last = { status: res.status, msg: null };
    // ۴۰۴ و خطاهای سرور گذرا هستند؛ ۴۰۱/۴۰۳ یعنی دسترسی، تلاش دوباره جوابش را عوض نمی‌کند
    if (res.status === 401 || res.status === 403) break;
    log(`call "${action}" got ${res.status} instead of JSON (attempt ${i + 1} of ${TRIES}).`);
  }
  const s = last ? last.status : 0;
  const why = s === 401 || s === 403
    ? "دسترسی Deployment روی «Anyone» نیست."
    : s === 404
      ? `بعد از ${faN(TRIES)} تلاش هم پاسخ نداد. اگر فقط گاهی پیش می‌آید، ایراد گذرای خود گوگل است و اجرای بعدی درست می‌شود. ` +
        "اگر همیشه است، نشانی Web app زنده نیست: نشانی SHEET_URL را در مرورگر باز کنید؛ اگر درست باشد باید {\"ok\":true,\"alive\":true,…} ببینید. " +
        "وگرنه در Apps Script: Deploy ← Manage deployments ← مداد روی Deployment فعال ← Version: New version ← Deploy، و همان نشانی را در Secret به نام SHEET_URL بگذارید."
      : "Deployment را روی «Anyone» تنظیم کرده‌اید؟";
  throw new Error(`پاسخ سند JSON نبود (${s}). ${why}`);
}

// نام ماه‌های شمسی برای ساختن «۱۴۰۵/۰۶/۲۰» از jdn
const pad2 = (n) => String(n).padStart(2, "0");

// دفتر وام‌ها ← [{ name, jdn, guarantor }]
async function loanGuarantors(D) {
  if (!ON) return [];
  const j = await call("read");
  const rows = (j && j.loans) || [];
  const outRows = [];
  for (const [name, date, guarantor] of rows) {
    const m = String(date).replace(/[۰-۹]/g, (d) => "۰۱۲۳۴۵۶۷۸۹".indexOf(d)).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
    if (!m) continue;
    outRows.push({ name, jdn: D.J.j2d(+m[1], +m[2], +m[3]), guarantor });
  }
  log(`loan register read: ${outRows.length} rows with a guarantor.`);
  return outRows;
}

// یک ردیف در تب «تاریخچه»
async function history(D, r, header) {
  if (!ON) return null;
  const d = D.J.d2j(r.asOf);
  const h = header || {};
  const j = await call("history", {
    row: {
      date: `${d.jy}/${pad2(d.jm)}/${pad2(d.jd)}`,
      capital: r.capital,
      cash: h.cash == null ? "" : h.cash,
      membersCapital: h.membersCapital == null ? "" : h.membersCapital,
      members: r.memberCount,
      activeLoans: r.active.n,
      remaining: r.active.remaining,
      late: (r.status.late1 || 0) + (r.status.late2 || 0),
      queue: r.wait ? r.wait.queue : "",
      waitDays: r.wait && r.wait.median != null ? Math.round(r.wait.median) : "",
    },
  });
  log(`history row ${j && j.action === "updated" ? "updated" : "appended"}.`);
  return j;
}

// وام‌های اکسل صندوق ← تب «وام‌ها»
// برمی‌گرداند: { added, filled: [...], missing: [...] }
//   filled  = ردیف‌هایی که «ضامن پیدا نشد» داشتند و حالا ضامنشان پیدا شد
//   missing = ردیف‌های تازه‌ای که ضامنشان پیدا نشد
async function loans(D, rows) {
  if (!ON || !rows.length) return null;
  const j = await call("loans", { rows });
  log(`loan rows added: ${(j && j.added) || 0}; guarantor backfilled: ${(j && j.filled ? j.filled.length : 0)}; without a guarantor: ${(j && j.missing ? j.missing.length : 0)}.`);
  return j;
}

// درخواست‌های فرم ← تب «صف وام»
async function queue(D, rows) {
  if (!ON || !rows.length) return null;
  const j = await call("queue", { rows });
  log(`queue rows added: ${(j && j.added) || 0} (skipped ${(j && j.skipped) || 0} already there).`);
  return j;
}

module.exports = { ON, loanGuarantors, history, loans, queue };
