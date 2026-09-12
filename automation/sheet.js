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

async function call(action, extra) {
  if (!ON) return null;
  const res = await fetch(URL_, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ token: TOKEN, action, ...(extra || {}) }),
    redirect: "follow", // Apps Script یک بار به script.googleusercontent.com هدایت می‌کند
  });
  const text = await res.text();
  let j;
  try { j = JSON.parse(text); } catch (e) {
    // اگر HTML برگشت، یعنی نشانی یا دسترسی Web app درست نیست
    throw new Error(`پاسخ سند JSON نبود (${res.status}). Deployment را روی «Anyone» تنظیم کرده‌اید؟`);
  }
  if (j.error) throw new Error(j.error);
  return j;
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
