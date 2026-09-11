// ساخته‌شده از همان کد سازنده‌ی داشبورد (fund-dashboard-builder.html). دستی ویرایش نکنید؛ هر تغییری در سازنده، این فایل را هم باید از نو ساخت.
/* global XLSX */
// ---------------------------------------------------------------------------- متن و عدد
const norm = (s) =>
  String(s == null ? "" : s)
    .replace(/[يى]/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/[\u200c\u200e\u200f\s]/g, "")
    .trim();

const latinDigits = (s) =>
  String(s)
    .replace(/[۰-۹]/g, (d) => "۰۱۲۳۴۵۶۷۸۹".indexOf(d))
    .replace(/[٠-٩]/g, (d) => "٠١٢٣٤٥٦٧٨٩".indexOf(d));

function toNum(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return isFinite(v) ? v : null;
  const s = latinDigits(v).replace(/[,٬\s]/g, "").replace(/[−–]/g, "-").replace("٫", ".");
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Number(s);
}

// ---------------------------------------------------------------------------- تقویم شمسی
// الگوریتم استاندارد jalaali-js (MIT). jdn = شماره‌ی روز ژولینی
const J = (() => {
  const breaks = [-61, 9, 38, 199, 426, 686, 756, 818, 1111, 1181, 1210, 1635, 2060, 2097, 2192, 2262, 2324, 2394, 2456, 3178];
  const div = (a, b) => ~~(a / b);
  const mod = (a, b) => a - ~~(a / b) * b;
  function jalCal(jy) {
    let gy = jy + 621, leapJ = -14, jp = breaks[0], jm, jump, leap, n, i;
    for (i = 1; i < breaks.length; i++) {
      jm = breaks[i];
      jump = jm - jp;
      if (jy < jm) break;
      leapJ += div(jump, 33) * 8 + div(mod(jump, 33), 4);
      jp = jm;
    }
    n = jy - jp;
    leapJ += div(n, 33) * 8 + div(mod(n, 33) + 3, 4);
    if (mod(jump, 33) === 4 && jump - n === 4) leapJ += 1;
    const leapG = div(gy, 4) - div((div(gy, 100) + 1) * 3, 4) - 150;
    const march = 20 + leapJ - leapG;
    if (jump - n < 6) n = n - jump + div(jump + 4, 33) * 33;
    leap = mod(mod(n + 1, 33) - 1, 4);
    if (leap === -1) leap = 4;
    return { leap, gy, march };
  }
  function g2d(gy, gm, gd) {
    let d = div((gy + div(gm - 8, 6) + 100100) * 1461, 4) + div(153 * mod(gm + 9, 12) + 2, 5) + gd - 34840408;
    return d - div(div(gy + 100100 + div(gm - 8, 6), 100) * 3, 4) + 752;
  }
  function d2g(jdn) {
    let j = 4 * jdn + 139361631;
    j = j + div(div(4 * jdn + 183187720, 146097) * 3, 4) * 4 - 3908;
    const i = div(mod(j, 1461), 4) * 5 + 308;
    const gd = div(mod(i, 153), 5) + 1;
    const gm = mod(div(i, 153), 12) + 1;
    const gy = div(j, 1461) - 100100 + div(8 - gm, 6);
    return { gy, gm, gd };
  }
  function j2d(jy, jm, jd) {
    const r = jalCal(jy);
    return g2d(r.gy, 3, r.march) + (jm - 1) * 31 - div(jm, 7) * (jm - 7) + jd - 1;
  }
  function d2j(jdn) {
    const gy = d2g(jdn).gy;
    let jy = gy - 621;
    const r = jalCal(jy);
    let k = jdn - g2d(gy, 3, r.march);
    if (k >= 0) {
      if (k <= 185) return { jy, jm: 1 + div(k, 31), jd: mod(k, 31) + 1 };
      k -= 186;
    } else {
      jy -= 1;
      k += 179;
      if (r.leap === 1) k += 1;
    }
    return { jy, jm: 7 + div(k, 30), jd: mod(k, 30) + 1 };
  }
  const isLeap = (jy) => jalCal(jy).leap === 0;
  const monthLength = (jy, jm) => (jm <= 6 ? 31 : jm <= 11 ? 30 : isLeap(jy) ? 30 : 29);
  // شنبه = ۰ ... جمعه = ۶
  const satIndex = (jdn) => (((jdn + 2) % 7) + 7) % 7;
  function addMonths(jy, jm, jd, k) {
    let m = jm - 1 + k;
    const y = jy + Math.floor(m / 12);
    m = (((m % 12) + 12) % 12) + 1;
    return j2d(y, m, Math.min(jd, monthLength(y, m)));
  }
  return { j2d, d2j, g2d, d2g, satIndex, addMonths, monthLength };
})();

const MONTHS = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"];
const WEEKDAYS = ["شنبه", "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه"];

// تاریخ متنی شمسی/میلادی یا عدد سریال اکسل ← jdn
function parseDate(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number" && v > 20000 && v < 80000) {
    // سریال اکسل (روز ۱ = ۱۹۰۰/۰۱/۰۱ با باگ سال کبیسه‌ی ۱۹۰۰)
    const ms = Math.round((v - 25569) * 86400000);
    const d = new Date(ms);
    return J.g2d(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }
  const s = latinDigits(v);
  const m = s.match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return y < 1700 ? J.j2d(y, mo, d) : J.g2d(y, mo, d);
}
const timeOf = (v) => {
  const m = latinDigits(v || "").match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? +m[1] * 3600 + +m[2] * 60 + +(m[3] || 0) : 0;
};

// ---------------------------------------------------------------------------- خواندن فایل
class ReportError extends Error {}

function findSheet(wb, spec) {
  const target = norm(spec.name);
  let name = wb.SheetNames.find((n) => norm(n) === target);
  if (!name) {
    // اگر نام شیت عوض شده، شیتی را پیدا کن که همه‌ی ستون‌های لازم را دارد
    name = wb.SheetNames.find((n) => locateHeader(sheetRows(wb.Sheets[n]), spec.columns).ok);
  }
  if (!name) throw new ReportError(`شیت «${spec.name}» در فایل پیدا نشد. شیت‌های موجود: ${wb.SheetNames.join("، ")}`);
  return { name, rows: sheetRows(wb.Sheets[name]) };
}

function sheetRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
}

function locateHeader(rows, columns) {
  const wanted = Object.entries(columns);
  let best = { missing: wanted.map((w) => w[1]), ok: false };
  for (let r = 0; r < Math.min(rows.length, 20); r++) {
    const cells = (rows[r] || []).map(norm);
    const idx = {};
    const missing = [];
    for (const [key, label] of wanted) {
      const i = cells.indexOf(norm(label));
      if (i === -1) missing.push(label);
      else idx[key] = i;
    }
    if (!missing.length) return { ok: true, row: r, idx };
    if (missing.length < best.missing.length) best = { ok: false, row: r, missing };
  }
  return best;
}

function readTable(wb, key) {
  const spec = CONFIG.sheets[key];
  const sheet = findSheet(wb, spec);
  const h = locateHeader(sheet.rows, spec.columns);
  if (!h.ok) {
    throw new ReportError(
      `در شیت «${sheet.name}» این ستون‌ها پیدا نشد: ${h.missing.map((m) => `«${m}»`).join("، ")}. ` +
        `اگر نرم‌افزار صندوق نام ستون‌ها را عوض کرده، بخش CONFIG در بالای کد را اصلاح کنید.`
    );
  }
  // ستون‌های اختیاری: اگر باشند خوانده می‌شوند، اگر نباشند خطا نمی‌دهند
  const headCells = (sheet.rows[h.row] || []).map(norm);
  const idx = { ...h.idx };
  const found = {};
  for (const [k, label] of Object.entries(spec.optional || {})) {
    const i = headCells.indexOf(norm(label));
    found[k] = i !== -1;
    if (i !== -1) idx[k] = i;
  }
  const body = sheet.rows.slice(h.row + 1).map((r) => {
    const o = {};
    for (const k in idx) o[k] = r[idx[k]];
    return o;
  });
  return { sheetName: sheet.name, headerRow: h.row, body, found, top: sheet.rows.slice(0, h.row) };
}

function readHeaderValues(topRows) {
  const out = {};
  for (const [key, label] of Object.entries(CONFIG.headerLabels)) {
    const target = norm(label);
    for (const row of topRows) {
      const i = (row || []).findIndex((c) => norm(c) === target);
      if (i === -1) continue;
      const v = row.slice(i + 1).find((c) => c != null && c !== "");
      out[key] = key === "reportDate" ? parseDate(v) : toNum(v);
      break;
    }
  }
  const first = topRows.flat().find((c) => typeof c === "string" && c.trim());
  const labels = Object.values(CONFIG.headerLabels).map(norm);
  out.fundName = first && !labels.includes(norm(first)) ? first.trim() : "صندوق";
  return out;
}

// ============================================================================
// تنظیمات: اگر نرم‌افزار صندوق نام شیت‌ها یا ستون‌ها را عوض کرد، فقط همین‌جا را اصلاح کنید.
// ============================================================================
const CONFIG = {
  sheets: {
    members: {
      name: "اعضا",
      columns: { row: "ردیف", name: "نام و نام خانوادگی", capital: "سرمایه شخصی" },
    },
    tx: {
      name: "تراکنش‌ها",
      columns: {
        date: "تاریخ",
        who: "مربوط به",
        type: "نوع تراکنش",
        manager: "تغییر در موجودی نزد مدیر",
        online: "تغییر در موجودی نزد همیان",
      },
    },
    loans: {
      name: "وام‌ها",
      columns: {
        date: "تاریخ اعطای وام",
        count: "تعداد اقساط وام",
        paidCount: "تعداد اقساط پرداخت‌شده",
        amount: "مبلغ وام",
        paid: "مبلغ پرداخت‌شده از وام",
        remaining: "مبلغ باقی‌مانده از وام",
      },
      optional: { who: "گیرنده‌ی وام" }, // برای وصل کردن درخواست‌ها به وام‌ها لازم است
    },
    // فایل درخواست‌های وام (loan-requests.xlsx) که مدیر صندوق خودش نگه می‌دارد
    requests: {
      name: "درخواست‌ها",
      columns: { name: "نام متقاضی", date: "تاریخ درخواست" },
      optional: { status: "وضعیت", received: "تاریخ دریافت" },
    },
  },
  // برچسب‌های بالای شیت اعضا؛ مقدار هر کدام در اولین خانه‌ی پر بعد از برچسب است
  headerLabels: {
    fundCapital: "سرمایه صندوق",
    cash: "موجودی فعلی",
    reportDate: "تاریخ صدور گزارش",
    membersCapital: "سرمایه اعضا",
    incomeExpense: "هزینه و درآمد صندوق",
  },
  // انواع تراکنش. «تسویه با مدیر» جابه‌جایی داخلی است و در آمار ورودی و خروجی نمی‌آید.
  txTypes: {
    repay: "بازپرداخت قسط",
    fee: "پرداخت حق عضویت",
    deposit: "افزایش موجودی",
    loan: "اعطای وام",
    withdraw: "برداشت متفرقه",
    settle: "تسویه با مدیر",
  },
  // حساب‌هایی که مال خود صندوق‌اند، نه اعضا. در شمارش اعضا نمی‌آیند و نامشان هیچ‌جا منتشر نمی‌شود.
  fundAccounts: ["حساب هیئت‌مدیره"],
  cancelWord: "لغو", // مقدار ستون «وضعیت» برای درخواست لغوشده
  waitMonths: 12, // میانه‌ی زمان انتظار روی وام‌های پرداخت‌شده در این چند ماه اخیر
  staleQueueDays: 365, // هشدار برای درخواستی که بیش از این تعداد روز در صف مانده
  chartRanges: [6, 12, 24], // دکمه‌های بازه‌ی نمودارهای ماهانه
  chartDefault: 6,
  growthStart: "1403/01/01", // نقطه‌ی شروع نمودار رشد دارایی
};

// ---------------------------------------------------------------------------- خواندن اکسل صندوق
function parseWorkbook(wb) {
  const members = readTable(wb, "members");
  const txT = readTable(wb, "tx");
  const loansT = readTable(wb, "loans");
  const header = readHeaderValues(members.top);
  const notes = [];

  const fundNames = CONFIG.fundAccounts.map(norm);
  const rows = members.body.filter((r) => toNum(r.row) != null && String(r.name || "").trim());
  const fundRows = rows.filter((r) => fundNames.includes(norm(r.name)));
  const memberCount = rows.length - fundRows.length;
  const personalCapitalSum = rows.reduce((s, r) => s + (toNum(r.capital) || 0), 0);
  const names = new Set(rows.map((r) => norm(r.name)));
  if (CONFIG.fundAccounts.length && fundRows.length < CONFIG.fundAccounts.length) {
    const found = fundRows.map((r) => norm(r.name));
    const missing = CONFIG.fundAccounts.filter((n) => !found.includes(norm(n)));
    notes.push({ level: "warn", text: `حساب صندوق ${missing.map((m) => `«${m}»`).join("، ")} در شیت اعضا پیدا نشد؛ اگر نامش عوض شده، fundAccounts را در CONFIG اصلاح کنید.` });
  }

  const tx = [];
  let skipped = 0;
  txT.body.forEach((r) => {
    const jdn = parseDate(r.date);
    if (jdn == null) {
      if (norm(r.date) !== norm("مجموع") && (r.type || r.manager || r.online)) skipped++;
      return;
    }
    tx.push({ jdn, t: timeOf(r.date), type: norm(r.type), amount: (toNum(r.manager) || 0) + (toNum(r.online) || 0) });
  });
  if (skipped) notes.push({ level: "warn", text: `${skipped} ردیف از شیت تراکنش‌ها تاریخ قابل‌خواندن نداشت و کنار گذاشته شد.` });
  tx.sort((a, b) => a.jdn - b.jdn || a.t - b.t);

  const loans = [];
  loansT.body.forEach((r) => {
    const jdn = parseDate(r.date);
    const count = toNum(r.count);
    if (jdn == null || !count) return;
    const paid = toNum(r.paid) || 0, remaining = toNum(r.remaining) || 0;
    loans.push({
      who: norm(r.who),
      whoRaw: String(r.who || "").trim(),
      jdn,
      count,
      paidCount: toNum(r.paidCount) || 0,
      amount: toNum(r.amount) || 0,
      paid,
      remaining,
      gross: paid + remaining, // کل مبلغی که باید برگردد؛ اگر کارمزدی ثبت نشده باشد همان «مبلغ وام» است
    });
  });
  loans.sort((a, b) => a.jdn - b.jdn);
  loans.forEach((l) => names.add(l.who));

  return { header, memberCount, personalCapitalSum, names, tx, loans, notes, hasLoanNames: loansT.found.who };
}

// ---------------------------------------------------------------------------- خواندن فایل درخواست‌ها
function parseRequests(wb) {
  const t = readTable(wb, "requests");
  const list = [];
  let badDate = 0;
  t.body.forEach((r) => {
    const nameRaw = String(r.name == null ? "" : r.name).trim();
    if (!nameRaw) return;
    const jdn = parseDate(r.date);
    if (jdn == null) { badDate++; return; }
    const status = norm(r.status);
    const received = parseDate(r.received);
    list.push({ name: norm(nameRaw), nameRaw, jdn, cancelled: status === norm(CONFIG.cancelWord), received });
  });
  return { list, badDate };
}

// تشخیص نوع فایل از روی شیت‌هایش
function detectKind(wb) {
  const has = (key) => {
    try { const s = findSheet(wb, CONFIG.sheets[key]); return locateHeader(s.rows, CONFIG.sheets[key].columns).ok; } catch (e) { return false; }
  };
  if (has("tx")) return "fund";
  if (has("requests")) return "requests";
  return null;
}

// ---------------------------------------------------------------------------- محاسبه‌ی شاخص‌ها
const median = (arr) => {
  if (!arr.length) return null;
  const a = [...arr].sort((x, y) => x - y), m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
};
const monthKey = (jdn) => { const d = J.d2j(jdn); return d.jy * 12 + d.jm - 1; };
const monthStartOf = (key) => J.j2d(Math.floor(key / 12), (key % 12) + 1, 1);

function compute(data, req) {
  const { header, tx, loans } = data;
  const T = Object.fromEntries(Object.entries(CONFIG.txTypes).map(([k, v]) => [norm(v), k]));
  const checks = [...data.notes];
  if (!tx.length) throw new ReportError("شیت تراکنش‌ها هیچ تراکنشی ندارد.");

  const firstDay = Math.min(tx[0].jdn, loans.length ? loans[0].jdn : Infinity);
  const lastTx = tx[tx.length - 1].jdn;
  let asOf = header.reportDate;
  if (asOf == null) {
    asOf = lastTx;
    checks.push({ level: "warn", text: "تاریخ صدور گزارش در فایل پیدا نشد؛ تاریخ آخرین تراکنش مبنا قرار گرفت." });
  }
  const curKey = monthKey(asOf);
  const monthStart = monthStartOf(curKey);

  // --- ماه‌ها: وام‌های پرداخت‌شده و وصولی هر ماه
  const firstKey = monthKey(firstDay);
  const months = [];
  for (let k = firstKey; k <= curKey; k++) {
    months.push({ y: Math.floor(k / 12), m: (k % 12) + 1, loanN: 0, loanAmt: 0, repay: 0, fee: 0, deposit: 0, other: 0 });
  }
  const unknown = new Set();
  let cash = 0;
  const dayDelta = new Map(); // تغییر روزانه‌ی دارایی کل
  const addDay = (j, v) => dayDelta.set(j, (dayDelta.get(j) || 0) + v);
  let repaySum = 0;
  for (const x of tx) {
    cash += x.amount;
    let key = T[x.type];
    if (key === "settle") continue;
    if (!key) {
      unknown.add(x.type);
      key = x.amount >= 0 ? "otherIn" : "otherOut";
    }
    // دارایی کل = موجودی نقد + مانده‌ی وام‌ها. پرداخت وام نقد را کم و مانده را زیاد می‌کند؛ قسط برعکس.
    addDay(x.jdn, x.amount);
    if (key === "repay") { addDay(x.jdn, -x.amount); repaySum += x.amount; }
    if (x.jdn > asOf) continue;
    const mo = months[monthKey(x.jdn) - firstKey];
    if (!mo) continue;
    if (key === "repay") mo.repay += x.amount;
    else if (key === "fee") mo.fee += x.amount;
    else if (key === "deposit") mo.deposit += x.amount;
    else if (key === "otherIn") mo.other += x.amount;
  }
  for (const l of loans) {
    addDay(l.jdn, l.gross);
    if (l.jdn > asOf) continue;
    const mo = months[monthKey(l.jdn) - firstKey];
    if (mo) { mo.loanN++; mo.loanAmt += l.amount; }
  }
  if (unknown.size) checks.push({ level: "warn", text: `نوع تراکنش ناشناخته: ${[...unknown].join("، ")}. اگر مبلغش مثبت بود، در نمودار وصولی با عنوان «سایر» آمد.` });

  // --- سری روزانه‌ی دارایی کل، از اولین روز داده تا تاریخ اکسل
  const outstanding = loans.reduce((s, l) => s + Math.max(0, l.remaining), 0);
  const capital = header.fundCapital != null ? header.fundCapital : outstanding + cash;
  const days = [...dayDelta.keys()].filter((j) => j <= asOf).sort((a, b) => a - b);
  let run = 0, minOutstanding = 0, runOut = 0;
  const outDelta = new Map();
  for (const l of loans) outDelta.set(l.jdn, (outDelta.get(l.jdn) || 0) + l.gross);
  for (const x of tx) if (T[x.type] === "repay") outDelta.set(x.jdn, (outDelta.get(x.jdn) || 0) - x.amount);
  const capAt = new Map();
  for (const j of days) {
    run += dayDelta.get(j);
    runOut += outDelta.get(j) || 0;
    minOutstanding = Math.min(minOutstanding, runOut);
    capAt.set(j, run);
  }
  const valueOn = (j) => { // مقدار در پایان روز j
    let v = 0;
    for (const d of days) { if (d > j) break; v = capAt.get(d); }
    return v;
  };
  // شروع نمودار: growthStart در CONFIG (اگر داده دیرتر شروع شده باشد، از اولین روز داده)
  const g0 = Math.max(firstDay, parseDate(CONFIG.growthStart) || firstDay);
  const series = [];
  const step = 7;
  for (let j = g0; j < asOf; j += step) series.push([j - g0, Math.round(valueOn(j))]);
  series.push([asOf - g0, Math.round(valueOn(asOf))]);
  // نقطه‌های قلاب: اول هر ماه و تاریخ گزارش
  const snaps = [];
  for (let k = monthKey(g0); k <= curKey; k++) {
    const j = monthStartOf(k);
    if (j >= g0 && j < asOf) snaps.push({ d: j - g0, v: Math.round(valueOn(j)), label: fmtDate(j) });
  }
  snaps.push({ d: asOf - g0, v: Math.round(valueOn(asOf)), label: fmtDate(asOf) });
  const years = [];
  for (let y = J.d2j(g0).jy + 1; y <= J.d2j(asOf).jy; y++) years.push({ d: J.j2d(y, 1, 1) - g0, label: fmtNum(y) });

  // --- وام‌های در جریان و وضعیت بازپرداخت (قسط ماهانه، اولین سررسید یک ماه پس از دریافت)
  const active = loans.filter((l) => l.remaining > 0 && l.paidCount < l.count);
  const status = { ahead: 0, ok: 0, late1: 0, late2: 0 };
  const dots = [];
  let dueThisMonth = 0, paidThisMonth = 0;
  for (const l of loans) {
    const d = J.d2j(l.jdn);
    let due = 0;
    for (let k = 1; k <= l.count; k++) {
      const dd = J.addMonths(d.jy, d.jm, d.jd, k);
      if (dd < asOf) due++;
      if (dd >= monthStart && dd < asOf) { dueThisMonth++; if (k <= l.paidCount) paidThisMonth++; }
    }
    if (!(l.remaining > 0 && l.paidCount < l.count)) continue;
    const delta = l.paidCount - due;
    const s = delta > 0 ? "ahead" : delta === 0 ? "ok" : delta === -1 ? "late1" : "late2";
    status[s]++;
    dots.push(s);
  }
  const order = { ok: 0, ahead: 1, late1: 2, late2: 3 };
  dots.sort((a, b) => order[a] - order[b]);
  const monthLoans = loans.filter((l) => l.jdn >= monthStart && l.jdn <= asOf);

  // --- زمان انتظار و صف (از فایل درخواست‌ها)
  let wait = null;
  if (req && !data.hasLoanNames) {
    checks.push({ level: "error", text: `ستون «${CONFIG.sheets.loans.optional.who}» در شیت وام‌ها پیدا نشد؛ درخواست‌ها به وام‌ها وصل نمی‌شوند و بخش «انتظار برای وام» ساخته نشد.` });
  } else if (req) {
    const byName = new Map();
    loans.forEach((l) => { if (!byName.has(l.who)) byName.set(l.who, []); byName.get(l.who).push({ jdn: l.jdn, used: false }); });
    const since = (() => { const d = J.d2j(asOf); return J.addMonths(d.jy, d.jm, d.jd, -CONFIG.waitMonths); })();
    const waits = [], queue = [], stale = [], unknownNames = new Set(), manualMiss = [];
    let future = 0, cancelled = 0;
    const list = [...req.list].sort((a, b) => a.jdn - b.jdn);
    for (const r of list) {
      if (r.jdn > asOf) { future++; continue; }
      if (r.cancelled) { cancelled++; continue; }
      if (!data.names.has(r.name)) unknownNames.add(r.nameRaw);
      const cands = byName.get(r.name) || [];
      let got = null;
      if (r.received != null) {
        got = r.received;
        const c = cands.find((c) => !c.used && c.jdn === r.received);
        if (c) c.used = true; else manualMiss.push(r.nameRaw);
      } else {
        const c = cands.find((c) => !c.used && c.jdn >= r.jdn);
        if (c) { c.used = true; got = c.jdn; }
      }
      if (got != null && got <= asOf) {
        if (got > since) waits.push(got - r.jdn);
      } else {
        queue.push(r);
        if (asOf - r.jdn > CONFIG.staleQueueDays) stale.push(r.nameRaw);
      }
    }
    wait = {
      n: waits.length,
      median: median(waits), // میانه: نیمی از وام‌ها در همین مدت یا کمتر پرداخت شده‌اند
      queue: queue.length,
    };
    checks.push({ level: "ok", text: `فایل درخواست‌ها: ${fmtInt(req.list.length)} درخواست خوانده شد؛ ${fmtInt(waits.length)} وام در ${fmtInt(CONFIG.waitMonths)} ماه اخیر به درخواستش وصل شد، ${fmtInt(queue.length)} درخواست در صف است${cancelled ? ` و ${fmtInt(cancelled)} درخواست لغو شده` : ""}.` });
    if (req.badDate) checks.push({ level: "warn", text: `${fmtInt(req.badDate)} ردیف از فایل درخواست‌ها تاریخ درخواست قابل‌خواندن نداشت و کنار گذاشته شد.` });
    if (unknownNames.size) checks.push({ level: "warn", text: `این نام‌ها در فایل درخواست‌ها با هیچ عضوی جور نشد (باید دقیقاً مثل نرم‌افزار نوشته شوند): ${[...unknownNames].join("، ")}` });
    if (manualMiss.length) checks.push({ level: "warn", text: `برای این درخواست‌ها «تاریخ دریافت» دستی پر شده ولی وامی با همان تاریخ به همان نام پیدا نشد؛ تاریخ دستی مبنا قرار گرفت: ${manualMiss.join("، ")}` });
    if (stale.length) checks.push({ level: "warn", text: `این درخواست‌ها بیش از ${fmtInt(CONFIG.staleQueueDays)} روز در صف مانده‌اند؛ اگر لغو شده‌اند، «${CONFIG.cancelWord}» بزنید: ${stale.join("، ")}` });
    if (future) checks.push({ level: "warn", text: `${fmtInt(future)} درخواست تاریخی بعد از تاریخ اکسل صندوق دارد و حساب نشد.` });
  } else {
    checks.push({ level: "warn", text: "فایل درخواست‌ها داده نشده؛ بخش «انتظار برای وام» در داشبورد نمی‌آید." });
  }

  // --- کنترل‌ها
  if (header.cash != null) {
    const diff = cash - header.cash;
    checks.push(diff === 0
      ? { level: "ok", text: "جمع تراکنش‌ها با «موجودی فعلی» فایل یکی است." }
      : { level: "error", text: `جمع تراکنش‌ها ${fmtInt(cash)} است ولی «موجودی فعلی» فایل ${fmtInt(header.cash)}. اختلاف: ${fmtInt(diff)} تومان.` });
  } else checks.push({ level: "warn", text: "«موجودی فعلی» در بالای شیت اعضا پیدا نشد؛ کنترل موجودی انجام نشد." });
  if (header.fundCapital != null) {
    const diff = outstanding + cash - header.fundCapital;
    checks.push(diff === 0
      ? { level: "ok", text: "مانده‌ی وام‌ها به‌اضافه‌ی موجودی نقد با «سرمایه صندوق» یکی است." }
      : { level: "error", text: `مانده‌ی وام‌ها + موجودی نقد با «سرمایه صندوق» ${fmtInt(diff)} تومان اختلاف دارد.` });
  }
  if (header.membersCapital != null) {
    const diff = data.personalCapitalSum - header.membersCapital;
    checks.push(diff === 0
      ? { level: "ok", text: "جمع سرمایه‌ی شخصی حساب‌ها با «سرمایه اعضا» یکی است." }
      : { level: "error", text: `جمع ستون «سرمایه شخصی» با «سرمایه اعضا» ${fmtInt(diff)} تومان اختلاف دارد.` });
  }
  const capToday = valueOn(asOf);
  checks.push(Math.round(capToday) === Math.round(capital)
    ? { level: "ok", text: "فرمول دارایی تاریخی در تاریخ اکسل با «سرمایه صندوق» یکی درمی‌آید." }
    : { level: "error", text: `فرمول دارایی تاریخی در تاریخ اکسل ${fmtInt(capToday)} می‌دهد ولی «سرمایه صندوق» ${fmtInt(capital)} است؛ نمودار رشد قابل‌اعتماد نیست.` });
  if (Math.round(repaySum) !== Math.round(loans.reduce((s, l) => s + l.paid, 0)))
    checks.push({ level: "warn", text: "جمع تراکنش‌های «بازپرداخت قسط» با جمع «مبلغ پرداخت‌شده از وام» یکی نیست؛ نمودار رشد ممکن است در گذشته کمی جابه‌جا باشد." });
  if (minOutstanding < 0)
    checks.push({ level: "warn", text: "در بخشی از گذشته، مانده‌ی وام‌ها منفی درمی‌آید (قسطی قبل از ثبت وامش). نمودار رشد در آن بازه تقریبی است." });
  checks.push({ level: "ok", text: `${fmtInt(data.memberCount)} عضو، ${fmtInt(tx.length)} تراکنش و ${fmtInt(loans.length)} وام خوانده شد. تاریخ داده‌ها: ${fmtDate(asOf)}.` });

  return {
    fundName: header.fundName,
    asOf,
    firstDay,
    capital,
    memberCount: data.memberCount,
    loansTotalAmount: loans.reduce((s, l) => s + l.amount, 0),
    active: {
      n: active.length,
      total: active.reduce((s, l) => s + l.gross, 0),
      paid: active.reduce((s, l) => s + l.paid, 0),
      remaining: active.reduce((s, l) => s + l.remaining, 0),
    },
    status,
    dots,
    month: {
      start: monthStart, loanN: monthLoans.length, loanAmt: monthLoans.reduce((s, l) => s + l.amount, 0), due: dueThisMonth, paid: paidThisMonth,
      list: monthLoans.map((l) => ({ amount: l.amount, count: l.count })), // فقط مبلغ و تعداد قسط هر وام؛ بدون نام
      avg12: (() => { const m = months.slice(-CONFIG.waitMonths); return m.length ? m.reduce((s, x) => s + x.loanAmt, 0) / m.length : null; })(),
    },
    wait,
    growth: { series, snaps, years, start: fmtMonthYear(g0) },
    months: months.slice(-Math.max(...CONFIG.chartRanges)),
    checks,
    hasError: checks.some((c) => c.level === "error"),
  };
}

// ---------------------------------------------------------------------------- قالب‌بندی
const fmtInt = (n) => Math.round(n).toLocaleString("fa-IR");
function fmtMoney(n, withUnit = false) {
  const a = Math.abs(n);
  let s;
  const f = (x, d) => x.toLocaleString("fa-IR", { maximumFractionDigits: d, minimumFractionDigits: 0 });
  if (a >= 1e9) { const d = a >= 1e10 ? 2 : 3; s = `${(a / 1e9).toLocaleString("fa-IR", { maximumFractionDigits: d, minimumFractionDigits: d })} میلیارد`; } // میلیاردها همیشه با ۳ رقم اعشار
  else if (a >= 1e6) s = `${f(a / 1e6, a >= 1e8 ? 0 : 1)} میلیون`;
  else if (a >= 1e3) s = `${f(a / 1e3, 0)} هزار`;
  else s = f(a, 0);
  return (n < 0 ? "−" : "") + s + (withUnit ? " تومان" : "");
}
const fmtNum = (n) => String(n).replace(/\d/g, (x) => "۰۱۲۳۴۵۶۷۸۹"[x]);
const fmtDay = (jdn) => { const d = J.d2j(jdn); return `${fmtNum(d.jd)} ${MONTHS[d.jm - 1]}`; };
const fmtDate = (jdn) => { const d = J.d2j(jdn); return `${fmtNum(d.jd)} ${MONTHS[d.jm - 1]} ${fmtNum(d.jy)}`; };
const fmtMonthYear = (jdn) => { const d = J.d2j(jdn); return `${MONTHS[d.jm - 1]} ${fmtNum(d.jy)}`; };
// ============================================================================
// ساخت فایل داشبورد عمومی. خروجی فقط ارقام کلی دارد؛ هیچ نام، شماره یا توضیح تراکنشی وارد آن نمی‌شود.
// ظاهر: زمینه‌ی آبی روشن، کارت‌های سفید با گوشه‌ی ۶ پیکسل، رنگ اصلی آبی، فونت دانا.
// ============================================================================
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// رنگ‌ها و اندازه‌های اصلی؛ برای تغییر ظاهر فقط همین‌ها را عوض کنید
const THEME = {
  bg: "#F4F6FB", // زمینه‌ی صفحه
  card: "#FFFFFF",
  border: "#E6EAF2",
  text: "#222222",
  muted: "#667791",
  primary: "#0049E8", // آبی اصلی
  primaryDark: "#0037AD",
  primarySoft: "#BFDBFB",
  accent: "#03045E", // سرمه‌ای
  good: "#0049E8", late1: "#F4A6AE", late2: "#DC3545", ahead: "#9CC0FF",
  sRepay: "#0049E8", sFee: "#0EA5A0", sDeposit: "#D97A06", sOther: "#8A96A0", // رنگ‌های نمودار وصولی (آزموده برای کوررنگی)
  radius: "16px", // گوشه‌ی کارت‌ها؛ کارت اصلی ۲۰، جعبه‌های داخلی و آیکن‌ها ۱۲
  heroGradient: "radial-gradient(420px 220px at 8% 0%,rgba(255,255,255,.14),transparent 70%),radial-gradient(300px 300px at 100% 100%,rgba(56,189,248,.25),transparent 70%),linear-gradient(135deg,#0A1A4F 0%,#15338F 55%,#2463EB 100%)",
  shadow: "0 1px 2px rgba(16,24,40,.04),0 8px 24px rgba(16,24,40,.06)",
};

const REPORT_CSS = `
:root{--bg:${THEME.bg};--card:${THEME.card};--border:${THEME.border};--ink:${THEME.text};--muted:${THEME.muted};--primary:${THEME.primary};--primary-dark:${THEME.primaryDark};--primary-soft:${THEME.primarySoft};--accent:${THEME.accent};
--good:${THEME.good};--ahead:${THEME.ahead};--late1:${THEME.late1};--late2:${THEME.late2};--s-repay:${THEME.sRepay};--s-fee:${THEME.sFee};--s-deposit:${THEME.sDeposit};--s-other:${THEME.sOther};--r:${THEME.radius};--hero:${THEME.heroGradient};--shadow:${THEME.shadow}}
*{box-sizing:border-box;margin:0;padding:0}
html{background:var(--bg);-webkit-text-size-adjust:100%}
body{font-family:"Dana FaNum",Tahoma,sans-serif;color:var(--ink);font-size:15px;line-height:1.8;font-variant-numeric:tabular-nums}
.page{max-width:1120px;margin:0 auto;padding:0 16px 32px}
.top{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;padding:18px 0 16px}
.brand{display:flex;align-items:center;gap:12px}
.logo{width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,var(--primary-dark),var(--primary));display:grid;place-items:center;flex:none}
.logo svg{width:24px;height:24px;fill:none;stroke:#fff;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.brand .fund{font-size:13px;color:var(--muted);line-height:1.5}
.brand h1{font-size:20px;font-weight:800;line-height:1.4;color:var(--ink)}
.pill{background:var(--card);border:0;box-shadow:0 1px 2px rgba(16,24,40,.06);border-radius:999px;padding:5px 14px;font-size:13px;color:var(--muted);white-space:nowrap}
.pill b{color:var(--ink);font-weight:600}
.grid-main{display:grid;grid-template-columns:1fr;gap:16px}
.card{background:var(--card);border:0;border-radius:var(--r);padding:18px;box-shadow:var(--shadow)}
.card h2{font-size:16px;font-weight:600;line-height:1.5;display:flex;align-items:center;gap:8px}
.card h2::before{content:"";width:4px;height:16px;border-radius:2px;background:var(--primary);flex:none}
.card .sub{font-size:13px;color:var(--muted);margin-bottom:12px;line-height:1.6}
.hero{background:var(--hero);border:0;border-radius:20px;color:#fff;padding:20px 20px 14px;overflow:hidden}
.hero .label{font-size:14px;font-weight:600;color:rgba(255,255,255,.85)}
.hero .big{font-size:36px;font-weight:800;line-height:1.35;margin:2px 0 2px;letter-spacing:-.5px}
.hero .big small{font-size:.45em;font-weight:600;color:rgba(255,255,255,.8);margin-right:4px}
.hero .readout{color:#fff}
.hero .hint{color:rgba(255,255,255,.75);margin-top:14px}
.readout{font-size:14px;margin-top:8px;min-height:1.9em}
.readout b{font-weight:800}
.hint{font-size:12px;color:var(--muted);line-height:1.7}
/* قاعده‌ی پر کردن کارت: هر کارت که بلندتر از محتوایش شد، بخش اصلی‌اش (نمودار یا جعبه‌ها) کش می‌آید */
.hero{display:flex;flex-direction:column}
.growth{margin-top:4px;touch-action:pan-y;user-select:none;-webkit-user-select:none;position:relative;flex:1 1 auto;min-height:190px}
.growth svg{position:absolute;inset:0;display:block;width:100%;overflow:visible}
.growth svg:focus{outline:none}.growth svg:focus-visible{outline:2px solid #fff;outline-offset:4px;border-radius:4px}
.growth text{font-family:inherit;font-size:11px;fill:rgba(255,255,255,.8)}
.stats{display:grid;grid-template-columns:2fr 3fr;gap:16px}
.stats .s-wait{grid-column:1/-1}
.stat{padding:16px 18px;display:flex;flex-direction:column;gap:6px}
.shead{display:flex;align-items:center;gap:10px}
.ico{width:40px;height:40px;border-radius:12px;display:grid;place-items:center;flex:none}
.ico svg{width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.i-blue{background:#E6EEFF;color:var(--primary)} .i-teal{background:#E0F5F3;color:#0B8480} .i-violet{background:#EFEAFE;color:#6A4FC7} .i-amber{background:#FFF1DE;color:#B86200}
.stat .k{font-size:13px;color:var(--muted);line-height:1.5}
.stat .v{font-size:28px;font-weight:800;line-height:1.3;white-space:nowrap;margin-top:4px}
.stat .v small{font-size:13px;font-weight:600;color:var(--muted);margin-right:4px}
.stat .c{font-size:12px;color:var(--muted);line-height:1.6}
.duo{display:grid;grid-template-columns:1fr auto 1fr;gap:16px;align-items:start}
.duo .sep{width:1px;align-self:stretch;background:var(--border)}
.duo .kk{font-size:13px;font-weight:600;color:var(--ink);margin-top:2px}
.big2{font-size:28px;font-weight:800;line-height:1.4}
.big2 small{font-size:14px;font-weight:600;color:var(--muted);margin-right:4px}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-left:5px;vertical-align:0}
.divider{border-top:1px solid var(--border);margin:18px 0 14px}
h3{font-size:15px;font-weight:600}
.dots{display:flex;flex-wrap:wrap;gap:7px;margin:10px 0 14px}
.dot{width:14px;height:14px;border-radius:50%}
.s-ok{background:var(--good)} .s-ahead{background:var(--ahead)} .s-late1{background:var(--late1)} .s-late2{background:var(--late2)}
ul.key{list-style:none;display:grid;grid-template-columns:1fr 1fr;gap:4px 16px;font-size:14px}
ul.key .dot{display:inline-block;width:10px;height:10px;margin-left:7px;vertical-align:0}
ul.key b{font-weight:800}
.note{font-size:12px;color:var(--muted);margin-top:12px;line-height:1.7}
/* کارت نمودار: نمودار پهن + ستون کناری «این ماه» */
.chartcard{display:grid;grid-template-columns:1fr;gap:16px}
.cc-main{display:flex;flex-direction:column;min-width:0}
.side{background:#F6F8FC;border-radius:12px;padding:16px;display:flex;flex-direction:column;gap:6px}
.side-h{font-size:14px;font-weight:600;color:var(--ink);margin-bottom:4px}
.side-h span{display:block;font-size:12px;font-weight:400;color:var(--muted)}
.side .v{font-size:30px;font-weight:800;line-height:1.35}
.side .v small{font-size:13px;font-weight:600;color:var(--muted);margin-right:4px}
.side .k{font-size:13px;color:var(--muted);line-height:1.6}
.side .k b{color:var(--ink)}
.side-ring{display:grid;place-items:center;margin:4px 0}
.cmp{font-size:12px;color:var(--muted);margin-top:auto;padding-top:10px;border-top:1px dashed var(--border)}
.cmp b{color:var(--ink)}
.chips{display:flex;flex-direction:column;gap:6px;margin-top:6px}
.chips span{display:flex;justify-content:space-between;align-items:center;gap:10px;background:#fff;border-radius:10px;padding:6px 10px;font-size:12px;color:var(--muted);box-shadow:0 1px 2px rgba(16,24,40,.05);border-right:3px solid #0B8480}
.chips b{color:var(--ink);font-size:14px;font-weight:800;white-space:nowrap}
.chips .more{justify-content:center;border-right:0;background:transparent;box-shadow:none}
/* وام‌های در جریان: در دسکتاپ دو نیمه کنار هم */
.lip{display:grid;grid-template-columns:1fr;gap:18px}
.lip-a{container-type:inline-size}
.lip-b{border-top:1px solid var(--border);padding-top:16px}
.ring{display:flex;align-items:center;gap:18px;margin-top:10px}
.ring svg{flex:none}
.ring .rt{display:grid;gap:8px;font-size:12px;color:var(--muted);line-height:1.5}
.ring .rt b{display:block;color:var(--ink);font-size:15px;font-weight:800}
.ring text{font-family:inherit}
.ring .rt{flex:1;min-width:0}
.ring .rt>div{background:#F6F8FC;border-radius:12px;padding:10px 14px}
.ring .rt b{white-space:nowrap}
/* وقتی کارت جا دارد: هر عدد در یک ردیف، عنوان راست و مبلغ چپ؛ هیچ‌وقت از کارت بیرون نمی‌زند */
@container (min-width:380px){.ring{gap:24px}.ring .rt>div{display:flex;justify-content:space-between;align-items:center;gap:10px}.ring .rt b{order:2;font-size:17px}}
.chead{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap}
.ranges{display:inline-flex;background:var(--bg);border-radius:999px;padding:3px;gap:2px}
.ranges button{font:inherit;font-size:12px;font-weight:600;border:0;background:transparent;color:var(--muted);border-radius:999px;padding:3px 12px;cursor:pointer}
.ranges button[aria-pressed="true"]{background:var(--primary);color:#fff}
.ranges button:focus-visible,.cols button:focus-visible{outline:2px solid var(--primary);outline-offset:2px}
.legend{display:flex;flex-wrap:wrap;gap:4px 14px;font-size:12px;color:var(--muted);margin:6px 0 10px}
.legend b{color:var(--ink)}
.legend .dash{display:inline-block;width:14px;border-top:2px dashed #8A96A0;margin-left:6px;vertical-align:3px}
.slegend{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:12px;color:var(--muted);margin-top:2px}.slegend:empty{display:none}.slegend span{white-space:nowrap}
.tip{position:absolute;z-index:5;background:#16213E;color:#fff;font-size:12px;line-height:1.6;padding:4px 9px;border-radius:4px;white-space:nowrap;pointer-events:none;transform:translate(-50%,calc(-100% - 6px));direction:rtl;box-shadow:0 2px 8px rgba(0,0,0,.18)}
.tip b{font-weight:800}.tip .sw{margin-left:4px}
.cols i{transition:filter .12s}.cols i.hot{filter:brightness(1.18) saturate(1.1)}
.chart{direction:ltr;padding-left:60px;flex:1 1 auto;display:flex;flex-direction:column}
.plot{position:relative;height:190px;flex:1 1 auto;min-height:190px}
.gl{position:absolute;left:0;right:0;border-top:1px solid #E7ECF3}
.gl span{position:absolute;right:calc(100% + 8px);top:-.85em;font-size:11px;color:var(--muted);white-space:nowrap;direction:rtl;line-height:1.6}
.gl.base{border-top-color:#C8D2E0}
.avg{position:absolute;left:0;right:0;border-top:2px dashed #8A96A0;pointer-events:none;z-index:2}
.cols{position:absolute;inset:0;display:flex}
.cols button{flex:1;min-width:0;height:100%;border:0;background:transparent;cursor:pointer;display:flex;flex-direction:column-reverse;align-items:center;gap:2px;padding:0;border-radius:4px 4px 0 0}
.cols button.sel{background:rgba(0,73,232,.07)}
.cols i{display:block;width:min(36px,60%);flex:none}
.cols i.top{border-radius:6px 6px 0 0}
.xl{display:flex;height:1.9em;margin-top:4px}
.xl span{flex:1;min-width:0;text-align:center;font-size:11px;color:var(--muted);white-space:nowrap;direction:rtl}
.xl span.sel{color:var(--ink);font-weight:600}
footer{padding:20px 4px 0;font-size:12px;color:var(--muted);text-align:center}
@media (min-width:900px){
  .page{padding:0 24px 40px}
  .grid-main{grid-template-columns:repeat(12,1fr);gap:20px}
  .span12{grid-column:span 12}.span7{grid-column:span 7}.span5{grid-column:span 5}.span6{grid-column:span 6}
  /* چینش نامتقارن: کارت اصلی دوسوم عرض، کارت‌های آماری در ستونی کنارش */
  .hero.span12{grid-column:span 8}
  .stats.span12{grid-column:span 4;grid-template-columns:1fr;gap:20px;align-content:start}
  .stats .s-wait{grid-column:auto}
  .stat .v{font-size:32px}
  .hero{padding:24px 26px 16px}.hero .big{font-size:44px}
  .plot{min-height:210px}
  .chartcard{grid-template-columns:minmax(0,1fr) 300px;gap:24px}
  .lip{grid-template-columns:minmax(0,5fr) minmax(0,7fr);gap:32px}
  .lip-b{border-top:0;padding-top:0;border-right:1px solid var(--border);padding-right:32px}
  .side .v{font-size:34px}
  .growth{min-height:230px}
}
@media (max-width:899px){.grid-main>*{grid-column:auto}}
@media (max-width:599px){ul.key{grid-template-columns:1fr}.stat{padding:14px}.stat .v{font-size:24px}.chead .ranges{margin-bottom:4px}}
@media (max-width:380px){.stat .v{font-size:21px}.hero .big{font-size:30px}ul.key{grid-template-columns:1fr}}
@media print{.ranges,.hint{display:none}}
`;

// آیکن‌های ساده (خطی)
const ICONS = {
  users: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.6-3.6 3.2-5.5 6.5-5.5s5.9 1.9 6.5 5.5"/><circle cx="17" cy="9" r="2.7"/><path d="M16.5 14.6c2.6.2 4.4 1.9 5 4.9"/></svg>',
  coins: '<svg viewBox="0 0 24 24"><ellipse cx="9" cy="6.5" rx="6" ry="2.8"/><path d="M3 6.5v4c0 1.5 2.7 2.8 6 2.8s6-1.3 6-2.8v-4"/><path d="M9 17.3c-3.3 0-6-1.3-6-2.8v-4"/><ellipse cx="15.5" cy="14" rx="5.5" ry="2.6"/><path d="M10 14v3.6c0 1.4 2.5 2.6 5.5 2.6s5.5-1.2 5.5-2.6V14"/></svg>',
  clock: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/></svg>',
  queue: '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M8 9h8M8 12h8M8 15h5"/></svg>',
  wallet: '<svg viewBox="0 0 24 24" style="stroke:#fff"><path d="M4 7.5A2.5 2.5 0 0 1 6.5 5H17v3"/><rect x="4" y="8" width="16" height="11" rx="2.5"/><circle cx="16" cy="13.5" r="1.3" style="fill:#fff"/></svg>',
  loan: '<svg viewBox="0 0 24 24"><path d="M3 12h18M12 3v18"/><circle cx="12" cy="12" r="8.5"/></svg>',
  pay: '<svg viewBox="0 0 24 24"><rect x="3" y="6" width="18" height="12" rx="2.5"/><path d="M3 10h18M7 15h3"/></svg>',
};

// کد تعاملی داخل فایل داشبورد (نمودار رشد و نمودارهای ماهانه). فقط روی داده‌ی کلی window.__D کار می‌کند.
function dashRuntime() {
  const D = window.__D;
  const MONTHS = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"];
  const fa = (s) => String(s).replace(/\d/g, (x) => "۰۱۲۳۴۵۶۷۸۹"[x]);
  const f = (x, d) => x.toLocaleString("fa-IR", { maximumFractionDigits: d, minimumFractionDigits: 0 });
  const money = (n, unit) => {
    const a = Math.abs(n);
    let s;
    if (a >= 1e9) { const d = a >= 1e10 ? 2 : 3; s = `${(a / 1e9).toLocaleString("fa-IR", { maximumFractionDigits: d, minimumFractionDigits: unit ? d : 0 })} میلیارد`; } // در برچسب محور بدون صفر اضافه
    else if (a >= 1e6) s = `${f(a / 1e6, a >= 1e8 ? 0 : 1)} میلیون`;
    else if (a >= 1e3) s = `${f(a / 1e3, 0)} هزار`;
    else s = f(a, 0);
    return s + (unit ? " تومان" : "");
  };
  const nice = (max, ticks) => {
    const raw = max / ticks || 1;
    const p = Math.pow(10, Math.floor(Math.log10(raw)));
    return [1, 2, 2.5, 5, 10].find((k) => k * p >= raw) * p;
  };
  const monthName = (mo) => `${MONTHS[mo.m - 1]} ${fa(mo.y)}`;
  const $ = (id) => document.getElementById(id);

  // ---------------- نمودار رشد دارایی کل (روی کارت آبی)
  function growth() {
    const box = $("growth");
    if (!box) return;
    const G = D.growth;
    const W = box.clientWidth, H = Math.max(box.clientHeight, 150); // ارتفاع نمودار = فضای باقی‌مانده‌ی کارت
    const ml = 58, mr = 10, mt = 12, mb = 26;
    const pw = W - ml - mr, ph = H - mt - mb;
    const dmax = G.series[G.series.length - 1][0] || 1;
    const vmax = Math.max(...G.series.map((p) => p[1]), 1);
    const step = nice(vmax, 4), top = Math.ceil(vmax / step) * step;
    const X = (d) => ml + (d / dmax) * pw, Y = (v) => mt + ph - (Math.max(0, v) / top) * ph;
    const pts = G.series.map((p) => `${X(p[0]).toFixed(1)},${Y(p[1]).toFixed(1)}`);
    let g = `<defs><linearGradient id="gFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff" stop-opacity=".28"/><stop offset="1" stop-color="#fff" stop-opacity=".02"/></linearGradient></defs>`;
    for (let v = 0; v <= top + 1; v += step) {
      g += `<line x1="${ml}" x2="${W - mr}" y1="${Y(v)}" y2="${Y(v)}" stroke="#fff" stroke-opacity="${v ? 0.16 : 0.4}" stroke-width="1"/>`;
      g += `<text x="${ml - 8}" y="${Y(v) + 4}" text-anchor="start">${v ? money(v) : "۰"}</text>`; // در متن راست‌به‌چپ، start یعنی لبه‌ی راست
    }
    if (!G.years.length || X(G.years[0].d) - ml > 90) g += `<text x="${ml}" y="${H - 5}" text-anchor="end">${G.start}</text>`;
    G.years.forEach((y) => {
      g += `<line x1="${X(y.d)}" x2="${X(y.d)}" y1="${mt + ph}" y2="${mt + ph + 5}" stroke="#fff" stroke-opacity=".5"/><text x="${X(y.d)}" y="${H - 5}" text-anchor="middle">${y.label}</text>`;
    });
    g += `<path d="M${X(0)},${Y(0)}L${pts.join("L")}L${X(dmax)},${Y(0)}Z" fill="url(#gFill)"/>`;
    g += `<path d="M${pts.join("L")}" fill="none" stroke="#fff" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    g += `<line id="gGuide" y1="${mt}" y2="${mt + ph}" stroke="#fff" stroke-opacity=".7" stroke-width="1" stroke-dasharray="3 3"/>`;
    const rr = pw / Math.max(1, G.snaps.length - 1) < 14 ? 2.5 : 3.5; // در صفحه‌ی باریک، نقطه‌ها کوچک‌تر
    G.snaps.forEach((s, i) => { g += `<circle data-i="${i}" cx="${X(s.d)}" cy="${Y(s.v)}" r="${rr}" fill="#fff" stroke="#0049E8" stroke-width="1.5"/>`; });
    g += `<rect x="${ml}" y="0" width="${pw}" height="${H}" fill="transparent"/>`;
    box.innerHTML = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" tabindex="0" role="img" aria-label="نمودار رشد دارایی کل صندوق از ${G.start}. با کلیدهای چپ و راست بین تاریخ‌ها جابه‌جا شوید.">${g}</svg>`;
    const svg = box.firstChild;
    const set = (i) => {
      growth.sel = i;
      const s = G.snaps[i];
      svg.querySelectorAll("circle").forEach((c) => c.setAttribute("r", +c.dataset.i === i ? 6.5 : rr));
      const gl = svg.querySelector("#gGuide");
      gl.setAttribute("x1", X(s.d)); gl.setAttribute("x2", X(s.d));
      $("growthRead").innerHTML = `${s.label}: <b>${money(s.v, true)}</b>`;
    };
    const pick = (e) => {
      const r = svg.getBoundingClientRect();
      const x = e.clientX - r.left;
      let best = 0, bd = Infinity;
      G.snaps.forEach((s, i) => { const dd = Math.abs(X(s.d) - x); if (dd < bd) { bd = dd; best = i; } });
      set(best);
    };
    svg.addEventListener("pointermove", pick);
    svg.addEventListener("pointerdown", pick);
    svg.addEventListener("keydown", (e) => {
      if (e.key === "ArrowLeft") { set(Math.max(0, growth.sel - 1)); e.preventDefault(); }
      if (e.key === "ArrowRight") { set(Math.min(G.snaps.length - 1, growth.sel + 1)); e.preventDefault(); }
    });
    set(growth.sel == null ? G.snaps.length - 1 : growth.sel);
  }

  // ---------------- نمودارهای ماهانه
  const SERIES = {
    loans: [{ k: "loanAmt", color: "var(--primary)", label: "وام پرداخت‌شده" }],
    coll: [
      { k: "repay", color: "var(--s-repay)", label: "قسط وام" },
      { k: "fee", color: "var(--s-fee)", label: "حق عضویت" },
      { k: "deposit", color: "var(--s-deposit)", label: "افزایش موجودی" },
      { k: "other", color: "var(--s-other)", label: "سایر" },
    ],
  };
  const state = { loans: { n: D.chartDefault, sel: null }, coll: { n: D.chartDefault, sel: null } };

  function bars(kind) {
    const root = $("ch-" + kind);
    if (!root) return;
    const st = state[kind];
    const series = SERIES[kind].filter((s) => s.k !== "other" || D.months.some((m) => m.other > 0));
    const ms = D.months.slice(-st.n);
    const tot = (m) => series.reduce((a, s) => a + m[s.k], 0);
    const totals = ms.map(tot);
    const avg = totals.reduce((a, b) => a + b, 0) / (totals.length || 1);
    const step = nice(Math.max(...totals, avg, 1), 4);
    const top = Math.ceil(Math.max(...totals, avg, 1) / step) * step;
    const pct = (v) => (v / top) * 100;
    if (st.sel == null || st.sel >= ms.length) st.sel = ms.length - 1;
    const every = st.n <= 6 ? 2 : st.n <= 12 ? 3 : 6;

    let grid = "";
    for (let v = 0; v <= top + 1; v += step) grid += `<div class="gl${v ? "" : " base"}" style="bottom:${pct(v)}%"><span>${v ? money(v) : "۰"}</span></div>`;
    const cols = ms.map((m, i) => {
      const segs = series.filter((s) => m[s.k] > 0);
      const h = segs.map((s, j) => `<i data-k="${s.k}" class="${j === segs.length - 1 ? "top" : ""}" style="height:calc(${pct(m[s.k]).toFixed(2)}% - ${j ? 2 : 0}px);background:${s.color}"></i>`).join("");
      return `<button type="button" data-i="${i}" class="${i === st.sel ? "sel" : ""}" aria-label="${monthName(m)}: ${money(totals[i], true)}">${h}</button>`;
    }).join("");
    const xl = ms.map((m, i) => `<span class="${i === st.sel ? "sel" : ""}">${(ms.length - 1 - i) % every === 0 ? MONTHS[m.m - 1] : ""}</span>`).join("");
    root.querySelector(".legend").innerHTML = `<span><i class="dash"></i>میانگین ماهانه: <b>${money(avg, true)}</b></span>`;
    // راهنمای رنگ‌ها در یک سطر، زیر نمودار
    root.querySelector(".slegend").innerHTML = series.length > 1 ? series.map((s) => `<span><i class="sw" style="background:${s.color}"></i>${s.label}</span>`).join("") : "";
    root.querySelector(".chart").innerHTML =
      `<div class="plot">${grid}<div class="avg" style="bottom:${pct(avg)}%"></div><div class="cols">${cols}</div><div class="tip" hidden></div></div><div class="xl">${xl}</div>`;
    root.querySelectorAll(".ranges button").forEach((b) => b.setAttribute("aria-pressed", +b.dataset.n === st.n));

    const read = (i) => {
      const m = ms[i];
      let t;
      if (kind === "loans") t = m.loanN ? `${monthName(m)}: <b>${fa(m.loanN)} وام، ${money(m.loanAmt, true)}</b>` : `${monthName(m)}: وامی پرداخت نشد`;
      else t = `${monthName(m)}: <b>${money(totals[i], true)}</b>` + (totals[i] ? `<br><span class="hint">${series.filter((s) => m[s.k] > 0).map((s) => `${s.label} ${money(m[s.k])}`).join("، ")}</span>` : "");
      root.querySelector(".readout").innerHTML = t;
    };
    const set = (i) => {
      st.sel = i;
      root.querySelectorAll(".cols button").forEach((b) => b.classList.toggle("sel", +b.dataset.i === i));
      root.querySelectorAll(".xl span").forEach((s, j) => s.classList.toggle("sel", j === i));
      read(i);
    };
    root.querySelectorAll(".cols button").forEach((b) => {
      b.addEventListener("click", () => set(+b.dataset.i));
      b.addEventListener("mouseenter", () => set(+b.dataset.i));
      b.addEventListener("focus", () => set(+b.dataset.i));
    });
    // راهنمای شناور: با بردن نشانگر یا زدن روی هر رنگ، مبلغ همان بخش نشان داده می‌شود
    const plot = root.querySelector(".plot"), tip = root.querySelector(".tip");
    const showTip = (seg) => {
      const i = +seg.parentNode.dataset.i, m = ms[i], s = series.find((x) => x.k === seg.dataset.k);
      root.querySelectorAll(".cols i.hot").forEach((x) => x.classList.remove("hot"));
      seg.classList.add("hot");
      tip.innerHTML = kind === "loans"
        ? `${MONTHS[m.m - 1]}: <b>${fa(m.loanN)} وام، ${money(m.loanAmt, true)}</b>`
        : `<i class="sw" style="background:${s.color}"></i>${s.label}: <b>${money(m[s.k], true)}</b>`;
      const pr = plot.getBoundingClientRect(), r = seg.getBoundingClientRect();
      tip.hidden = false;
      const half = tip.offsetWidth / 2;
      const x = Math.min(Math.max(r.left + r.width / 2 - pr.left, half), pr.width - half);
      tip.style.left = x + "px";
      tip.style.top = r.top - pr.top + "px";
    };
    const hideTip = () => { tip.hidden = true; root.querySelectorAll(".cols i.hot").forEach((x) => x.classList.remove("hot")); };
    root.querySelectorAll(".cols i").forEach((seg) => {
      seg.addEventListener("pointerenter", () => showTip(seg));
      seg.addEventListener("click", () => showTip(seg));
    });
    plot.addEventListener("pointerleave", hideTip);
    read(st.sel);
  }

  document.querySelectorAll(".ranges button").forEach((b) => b.addEventListener("click", () => {
    const kind = b.closest("[data-kind]").dataset.kind;
    state[kind].n = +b.dataset.n;
    state[kind].sel = null;
    bars(kind);
  }));
  const all = () => { growth(); bars("loans"); bars("coll"); };
  let rt;
  window.addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(growth, 120); });
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(all);
  all();
}

// @font-face برای فونت دانا (جاسازی‌شده در فایل)
const fontFaces = (fonts) =>
  [["regular", 400], ["semibold", 600], ["extrabold", 800]]
    .map(([k, w]) => `@font-face{font-family:"Dana FaNum";src:url(data:font/woff2;base64,${fonts[k]}) format("woff2");font-weight:${w};font-display:swap}`)
    .join("");

// حلقه‌ی پیشرفت (SVG)
function ringSVG(pct, label, sub, size, stroke = 10) {
  const c = size / 2, rad = c - stroke / 2 - 2, len = 2 * Math.PI * rad, p = Math.max(0, Math.min(1, pct));
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" role="img" aria-label="${esc(label)} ${esc(sub)}">` +
    `<circle cx="${c}" cy="${c}" r="${rad}" fill="none" stroke="#DCE8FE" stroke-width="${stroke}"/>` +
    (p > 0 ? `<circle cx="${c}" cy="${c}" r="${rad}" fill="none" stroke="${THEME.primary}" stroke-width="${stroke}" stroke-linecap="round" stroke-dasharray="${(len * p).toFixed(2)} ${len.toFixed(2)}" transform="rotate(-90 ${c} ${c})"/>` : "") +
    `<text x="50%" y="${sub ? "47%" : "54%"}" text-anchor="middle" dominant-baseline="middle" font-weight="800" font-size="${Math.round(size / 5)}" fill="${THEME.text}">${label}</text>` +
    (sub ? `<text x="50%" y="68%" text-anchor="middle" dominant-baseline="middle" font-size="11" fill="${THEME.muted}">${sub}</text>` : "") +
    `</svg>`;
}

function renderReport(r, fonts) {
  const title = `داشبورد ${r.fundName}`;
  const upd = fmtDate(r.asOf);
  const moneyParts = (n) => { const s = fmtMoney(n); const i = s.lastIndexOf(" "); return i > 0 ? `${s.slice(0, i)} <small>${s.slice(i + 1)} تومان</small>` : `${s} <small>تومان</small>`; };
  const s = r.status;
  const a = r.active;
  const paidPct = a.total ? (a.paid / a.total) * 100 : 0;

  const D = {
    chartDefault: CONFIG.chartDefault,
    growth: r.growth,
    months: r.months.map((m) => ({ y: m.y, m: m.m, loanN: m.loanN, loanAmt: m.loanAmt, repay: m.repay, fee: m.fee, deposit: m.deposit, other: m.other })),
  };

  const ranges = () => `<div class="ranges" role="group" aria-label="بازه‌ی نمودار">${CONFIG.chartRanges.map((n) => `<button type="button" data-n="${n}" aria-pressed="${n === CONFIG.chartDefault}">${fmtNum(n)} ماه</button>`).join("")}</div>`;

  // کارت‌های آماری با پهنای متفاوت: اعضا (باریک)، وام از ابتدا (متوسط)، انتظار برای وام (پهن، دو عدد)
  const head = (icon, cls, label) => `<div class="shead"><div class="ico ${cls}">${ICONS[icon]}</div><div class="k">${label}</div></div>`;
  const stats = [
    `<div class="card stat s-mem">${head("users", "i-blue", "اعضای فعال")}<div class="v">${fmtInt(r.memberCount)} <small>عضو</small></div></div>`,
    `<div class="card stat s-loan">${head("coins", "i-teal", "ارزش وام‌ها از ابتدای فعالیت صندوق")}<div class="v">${moneyParts(r.loansTotalAmount)}</div></div>`,
  ];
  if (r.wait) {
    const w = r.wait;
    stats.push(`<div class="card stat s-wait">${head("clock", "i-violet", "زمان انتظار برای وام")}
      <div class="duo">
        <div><div class="v">${w.median != null ? `${fmtInt(w.median)} <small>روز</small>` : "—"}</div><div class="kk">زمان انتظار</div>
          <div class="c">${w.n ? `نیمی از ${fmtInt(w.n)} وام ${fmtInt(CONFIG.waitMonths)} ماه اخیر در همین مدت یا کمتر پرداخت شده‌اند` : `وامی با درخواست ثبت‌شده در ${fmtInt(CONFIG.waitMonths)} ماه اخیر نبود`}</div></div>
        <div class="sep"></div>
        <div><div class="v">${fmtInt(w.queue)} <small>درخواست</small></div><div class="kk">در صف</div><div class="c">هنوز وامشان پرداخت نشده</div></div>
      </div></div>`);
  }

  const mo = r.month;
  const monthRange = mo.start === r.asOf ? fmtDate(r.asOf) : `${fmtNum(J.d2j(mo.start).jd)} تا ${fmtDate(r.asOf)}`;
  const collectPct = mo.due ? (mo.paid / mo.due) * 100 : 0;
  const cm = r.months[r.months.length - 1] || {};
  const curColl = (cm.repay || 0) + (cm.fee || 0) + (cm.deposit || 0) + (cm.other || 0);

  const ogDesc = `دارایی کل ${fmtMoney(r.capital, true)}، ${fmtInt(a.n)} وام در جریان. تاریخ گزارش: ${upd}.`;

  return `<!doctype html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<meta name="theme-color" content="${THEME.primary}">
<title>${esc(title)}</title>
<meta property="og:title" content="${esc(title)}">
<meta property="og:description" content="${esc(ogDesc)}">
<style>
/* Dana FaNum © fontiran.com — استفاده با مجوز وب مدیر صندوق */
${fontFaces(fonts)}
${REPORT_CSS}
</style>
</head>
<body>
<div class="page">
<header class="top">
  <div class="brand">
    <div class="logo" aria-hidden="true">${ICONS.wallet}</div>
    <div><div class="fund">${esc(r.fundName)}</div><h1>داشبورد صندوق</h1></div>
  </div>
  <div class="pill">تاریخ گزارش: <b>${upd}</b></div>
</header>

<main class="grid-main">
  <section class="card hero span12">
    <div class="label">دارایی کل صندوق</div>
    <div class="big">${moneyParts(r.capital)}</div>
    <div class="readout" id="growthRead" aria-live="polite"></div>
    <div class="growth" id="growth"></div>
    <p class="hint">رشد دارایی کل از ${r.growth.start}. روی نمودار بزنید یا نشانگر را ببرید تا دارایی صندوق در اول هر ماه را ببینید.</p>
  </section>

  <div class="stats span12${r.wait ? " has-wait" : ""}">${stats.join("")}</div>

  <section class="card span12 lip">
    <div class="lip-a">
      <h2>وام‌های در جریان</h2>
      <div class="big2">${fmtInt(a.n)} <small>وام</small></div>
      <div class="ring">
        ${ringSVG(paidPct / 100, `${fmtInt(Math.round(paidPct))}٪`, "بازپرداخت‌شده", 132)}
        <div class="rt">
          <div><b>${fmtMoney(a.total)}</b><span>مبلغ کل</span></div>
          <div><b>${fmtMoney(a.paid)}</b><span><i class="sw" style="background:var(--primary)"></i>پرداخت‌شده</span></div>
          <div><b>${fmtMoney(a.remaining)}</b><span><i class="sw" style="background:#BFD3FB"></i>مانده</span></div>
        </div>
      </div>
    </div>
    <div class="lip-b">
      <h3>وضعیت بازپرداخت</h3>
      <div class="sub">هر دایره یک وام در جریان است</div>
      <div class="dots" aria-hidden="true">${r.dots.map((d) => `<span class="dot s-${d}"></span>`).join("")}</div>
      <ul class="key">
        <li><span class="dot s-ok"></span><b>${fmtInt(s.ok)}</b> وام طبق برنامه</li>
        <li><span class="dot s-ahead"></span><b>${fmtInt(s.ahead)}</b> وام جلوتر از برنامه</li>
        <li><span class="dot s-late1"></span><b>${fmtInt(s.late1)}</b> وام یک قسط معوق</li>
        <li><span class="dot s-late2"></span><b>${fmtInt(s.late2)}</b> وام دو قسط یا بیشتر معوق</li>
      </ul>
      <p class="note">اقساط ماهانه فرض شده‌اند و اولین سررسید هر وام یک ماه پس از دریافت آن است.</p>
    </div>
  </section>

  <section class="card span12 chartcard" data-kind="loans" id="ch-loans">
    <div class="cc-main">
      <div class="chead"><div><h2>وام پرداخت‌شده</h2><div class="sub">مبلغ وام‌هایی که صندوق در هر ماه به اعضا داده است</div></div>${ranges()}</div>
      <div class="legend"></div>
      <div class="chart"></div>
      <div class="slegend"></div>
      <div class="readout" aria-live="polite"></div>
    </div>
    <aside class="side">
      <div class="side-h">این ماه <span>${esc(monthRange)}</span></div>
      <div class="v">${fmtInt(mo.loanN)} <small>وام</small></div>
      <div class="k">${mo.loanN ? `به ارزش <b>${fmtMoney(mo.loanAmt, true)}</b>` : "هنوز وامی پرداخت نشده"}</div>
      ${mo.list.length ? `<div class="chips" aria-label="وام‌های این ماه">${mo.list.slice(0, 4).map((l) => `<span><b>${fmtMoney(l.amount)}</b>${fmtInt(l.count)} قسط</span>`).join("")}${mo.list.length > 4 ? `<span class="more">و ${fmtInt(mo.list.length - 4)} وام دیگر</span>` : ""}</div>` : ""}
      ${mo.avg12 ? `<div class="cmp">میانگین ماهانه‌ی ${fmtInt(CONFIG.waitMonths)} ماه اخیر: <b>${fmtMoney(mo.avg12, true)}</b></div>` : ""}
    </aside>
  </section>

  <section class="card span12 chartcard" data-kind="coll" id="ch-coll">
    <div class="cc-main">
      <div class="chead"><div><h2>مجموع وصولی</h2><div class="sub">هر چه در هر ماه وارد صندوق شده، به تفکیک نوع واریز</div></div>${ranges()}</div>
      <div class="legend"></div>
      <div class="chart"></div>
      <div class="slegend"></div>
      <div class="readout" aria-live="polite"></div>
    </div>
    <aside class="side">
      <div class="side-h">این ماه <span>${esc(monthRange)}</span></div>
      <div class="k">وصول اقساط</div>
      ${mo.due
        ? `<div class="side-ring">${ringSVG(collectPct / 100, `${fmtInt(Math.round(collectPct))}٪`, "وصول", 120, 11)}</div>
      <div class="v">${fmtInt(mo.paid)} <small>قسط از ${fmtInt(mo.due)} قسط</small></div>
      <div class="k">قسطی که تا امروز سررسید شده، پرداخت شده</div>`
        : `<div class="v">—</div><div class="k">هنوز قسطی در این ماه سررسید نشده</div>`}
      <div class="cmp">وصولی این ماه تا امروز: <b>${fmtMoney(curColl, true)}</b></div>
    </aside>
  </section>
</main>

<footer>این داشبورد از خروجی نرم‌افزار صندوق در ${upd} ساخته شده و فقط ارقام کلی را نشان می‌دهد؛ اطلاعات هیچ عضوی در آن نیست.</footer>
</div>
<script>
window.__D = ${JSON.stringify(D)};
(${dashRuntime.toString()})();
<\/script>
</body>
</html>`;
}

module.exports = { CONFIG, THEME, parseWorkbook, parseRequests, detectKind, compute, renderReport, fmtDate, fmtNum, fmtInt, fmtMoney, ReportError };
