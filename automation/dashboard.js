// ساخته‌شده از همان کد سازنده (helpers.js + core.js + render.js)؛ دستی ویرایش نکنید.
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
  const rows = sheet.rows.slice(h.row + 1);
  const body = rows.map((r) => {
    const o = {};
    for (const k in idx) o[k] = r[idx[k]];
    return o;
  });
  return { sheetName: sheet.name, headerRow: h.row, body, found, idx, headCells, rows, top: sheet.rows.slice(0, h.row) };
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
      optional: { mobile: "شماره موبایل" }, // برای تطبیق متقاضی و ضامنِ فرم پُرس‌لاین با اعضا
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
    // خروجی پاسخ‌های فرم درخواست وام در پُرس‌لاین (شیت Results). فقط همین ستون‌ها خوانده می‌شوند؛
    // ستون‌های حساس فرم (کد ملی، کارت، شبا، ایمیل) هرگز خوانده نمی‌شوند.
    porsline: {
      name: "Results",
      columns: { name: "نام و نام خانوادگی", mobile: "شماره تلفن همراه", date: "تاریخ اتمام" },
      optional: { gName: "نام و نام خانوادگی ضامن:", gMobile: "شماره تلفن همراه ضامن:", amount: "مبلغ وام مورد نیاز (به تومان):", count: "تعداد اقساط پیشنهادی:" },
      // اگر سؤال نام ضامن عنوان دیگری دارد (مثل سؤال کشویی ضامن)، ستونی که «ضامن» در عنوانش هست و
      // هیچ‌کدام از این کلمه‌ها را ندارد هم نام ضامن حساب می‌شود. ستون‌های حساس به این شکل کنار می‌مانند.
      guarantorNameHint: { has: "ضامن", not: ["تلفن", "موبایل", "همراه", "شماره", "ملی", "کد", "کارت", "شبا", "حساب", "ایمیل", "آدرس", "نشانی", "سال", "تاریخ", "دانشگاه", "رشته", "تحصیل", "نسبت", "شغل"] },
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
  requestExpiryMonths: 9, // درخواستی که این مدت به وام نرسد، منقضی حساب می‌شود و در صف نمی‌ماند
  // شرط‌های ضامن مجاز (فهرست با هر اکسل تازه برای مدیر صندوق فرستاده می‌شود؛ در داشبورد عمومی نمی‌آید)
  guarantor: {
    maxLateInstallments: 0, // اگر روی وام‌های خودش بیش از این تعداد قسط معوق داشته باشد، ضامن نمی‌شود
    lateDaysLimit: 30, lateLookbackMonths: 12, // در این چند ماه هیچ قسطی بیش از این چند روز دیر پرداخت نشده باشد
    inactiveMonths: 3, // در این چند ماه دست‌کم یک تراکنش (حق عضویت، قسط، واریز، …) داشته باشد
    minTenureMonths: 6, // دست‌کم این چند ماه از اولین تراکنشش گذشته باشد
    noWithdrawMonths: 12, // در این چند ماه «برداشت متفرقه» نداشته باشد
    minCapital: 10000000, // حداقل سرمایه‌ی شخصی (تومان)
    maxOwnDebtRatio: 2, // مانده‌ی وام‌های در جریان خودش بیش از این چند برابر سرمایه‌اش نباشد
    maxActiveGuarantees: 2, // حداکثر ضمانت وام‌های در جریان (پایه)
    // استثنا برای سرمایه‌ی زیاد: به ازای هر «پله» سرمایه بیش از «از»، یک ضمانت بیشتر، تا «سقف»
    // (۶۰ میلیون ← ۳، ۸۰ میلیون ← ۴، ۱۰۰ میلیون و بیشتر ← ۵)
    extraGuarantees: { from: 40000000, step: 20000000, max: 5 },
  },
  // اصلاح نام‌های اکسل صندوق: { "نام در اکسل": "نام درست" }. در مخزن عمومی خالی می‌ماند؛
  // خودکارساز آن را از Secret با نام NAME_FIXES پر می‌کند (نام‌ها نباید در کد عمومی بیایند).
  nameFixes: {},
  roster: [],
  chartRanges: [6, 12, 24], // دکمه‌های بازه‌ی نمودارهای ماهانه
  chartDefault: 6,
  growthStart: "1403/01/01", // نقطه‌ی شروع نمودار رشد دارایی
};

// ---------------------------------------------------------------------------- نام درست اعضا
// دو منبع، هر دو از Secret با نام NAME_FIXES (در کد عمومی خالی می‌مانند):
//   - CONFIG.nameFixes: اصلاح دستی { "نام در اکسل": "نام درست" }  ← خط‌های «قدیمی : درست»
//   - CONFIG.roster: فهرست اعضا [{ name, code (سال ورود), mobiles }] ← خط‌های «نام | کد | موبایل | موبایل ۲»
// نام درست هر عضو اکسل: اصلاح دستی، وگرنه نام فهرست اعضا (از روی موبایل، یا نام + پسوند سال ورود).
// اگر دو عضو به یک نام برسند، سال ورود در پرانتز می‌آید؛ اگر سال ورود معلوم نبود، اصلاح انجام نمی‌شود.
function loadNameFixes(text) {
  let fixes = 0, roster = 0;
  for (const raw of String(text || "").split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    if (line.includes("|")) {
      const [name, code, ...mobiles] = line.split("|").map((x) => x.trim());
      if (name) { CONFIG.roster.push({ name, code: latinDigits(code || ""), mobiles: mobiles.filter(Boolean) }); roster++; }
    } else if (line.includes(":")) {
      const i = line.indexOf(":"), from = line.slice(0, i).trim(), to = line.slice(i + 1).trim();
      if (from && to) { CONFIG.nameFixes[from] = to; fixes++; }
    }
  }
  return { fixes, roster };
}

function nameFixMap(rows, notes) {
  const roster = (CONFIG.roster || []).map((e) => ({ ...e, key: norm(e.name), mobiles: (e.mobiles || []).map(normMobile).filter(Boolean) }));
  const rByMobile = new Map(), rByNameCode = new Map(), rByName = new Map();
  for (const e of roster) {
    e.mobiles.forEach((m) => { if (!rByMobile.has(m)) rByMobile.set(m, e); });
    rByNameCode.set(e.key + "|" + e.code, e);
    rByName.set(e.key, rByName.has(e.key) ? null : e); // نام تکراری در فهرست ← با نام تنها پیدا نمی‌شود
  }
  const manual = new Map(Object.entries(CONFIG.nameFixes || {}).map(([f, t]) => [norm(f), String(t).trim()]));
  const prop = rows.filter((r) => r.name != null && String(r.name).trim()).map((r) => {
    const raw = String(r.name).trim(), k = norm(raw);
    const suffix = raw.match(/^(.*?)[\s\-_]*([0-9۰-۹]{2,4})$/);
    const e = rByMobile.get(normMobile(r.mobile)) || (suffix && rByNameCode.get(norm(suffix[1]) + "|" + latinDigits(suffix[2]))) || rByName.get(k) || null;
    r._roster = e;
    const to = manual.has(k) ? manual.get(k) : e ? e.name.trim() : raw;
    return { raw, k, to, code: e && e.code, manual: manual.has(k), changed: norm(to) !== k };
  });
  const skipped = [];
  // دو ردیف با نام کاملاً یکسان در اکسل را نمی‌شود جدا کرد (تراکنش‌ها با نام وصل‌اند)؛ دست نمی‌زنیم
  const sameRaw = new Map(); prop.forEach((p) => sameRaw.set(p.k, (sameRaw.get(p.k) || 0) + 1));
  prop.forEach((p) => { if (sameRaw.get(p.k) > 1) { p.to = p.raw; p.changed = false; } });
  const dupOf = () => { const c = new Map(); prop.forEach((p) => c.set(norm(p.to), (c.get(norm(p.to)) || 0) + 1)); return c; };
  let c = dupOf();
  for (const p of prop) if (p.changed && c.get(norm(p.to)) > 1) {
    if (p.code && !p.manual) p.to = `${p.to} (${fmtNum(p.code)})`; // دو نفر هم‌نام در فهرست اعضا
    else { skipped.push(`«${p.raw}» ← «${p.to}»`); p.to = p.raw; p.changed = false; }
  }
  c = dupOf();
  for (const p of prop) if (p.changed && c.get(norm(p.to)) > 1) { skipped.push(`«${p.raw}» ← «${p.to}»`); p.to = p.raw; p.changed = false; }
  if (skipped.length) notes.push({ level: "warn", text: `این اصلاح نام‌ها انجام نشد، چون نام درست با عضو دیگری یکی می‌شد (اگر هم‌نام‌اند، اصلاح دستی را بردارید تا سال ورود از فهرست اعضا کنارش بیاید): ${skipped.join("، ")}` });
  const map = new Map();
  prop.forEach((p) => { if (p.changed) map.set(p.k, p.to); });
  // جست‌وجو در فهرست اعضا برای کسانی که حساب صندوق ندارند (کارت بررسی)
  const find = (mobile, nameKey) => (mobile && rByMobile.get(mobile)) || (nameKey && rByName.get(nameKey)) || null;
  return { map, find };
}

// ---------------------------------------------------------------------------- خواندن اکسل صندوق
function parseWorkbook(wb) {
  const members = readTable(wb, "members");
  const txT = readTable(wb, "tx");
  const loansT = readTable(wb, "loans");
  const header = readHeaderValues(members.top);
  const notes = [];
  const { map: fixMap, find: rosterFind } = nameFixMap(members.body.filter((r) => toNum(r.row) != null), notes);
  const fix = (raw) => (raw != null && fixMap.has(norm(raw)) ? fixMap.get(norm(raw)) : raw);
  members.body.forEach((r) => { r.name = fix(r.name); });
  txT.body.forEach((r) => { r.who = fix(r.who); });
  loansT.body.forEach((r) => { r.who = fix(r.who); });

  const fundNames = CONFIG.fundAccounts.map(norm);
  const rows = members.body.filter((r) => toNum(r.row) != null && String(r.name || "").trim());
  const fundRows = rows.filter((r) => fundNames.includes(norm(r.name)));
  const memberCount = rows.length - fundRows.length;
  const personalCapitalSum = rows.reduce((s, r) => s + (toNum(r.capital) || 0), 0);
  const names = new Set(rows.map((r) => norm(r.name)));
  if (names.size < rows.length) {
    const seen = new Set(), dup = new Set();
    rows.forEach((r) => { const k = norm(r.name); if (seen.has(k)) dup.add(String(r.name).trim()); seen.add(k); });
    notes.push({ level: "warn", text: `این نام‌ها در شیت اعضا بیش از یک بار آمده‌اند و یک نفر حساب شدند (سرمایه‌شان جمع شد): ${[...dup].join("، ")}. اگر دو نفر جدا هستند، در نرم‌افزار نامشان را متمایز کنید.` });
  }
  // اعضای واقعی (بدون حساب‌های صندوق) با موبایل و سرمایه؛ فقط برای تطبیق و فهرست ضامن، هرگز در داشبورد عمومی
  // اعضا به تفکیک نام؛ چند حساب با نام یکسان یک نفر حساب می‌شوند (تراکنش‌ها و وام‌ها هم با نام وصل‌اند)
  const personBy = new Map();
  for (const r of rows) {
    if (fundNames.includes(norm(r.name))) continue;
    const key = norm(r.name);
    const mobiles = [normMobile(r.mobile), ...((r._roster && r._roster.mobiles) || [])].filter(Boolean); // موبایل‌های دیگر از فهرست اعضا
    const p = personBy.get(key);
    if (p) { p.capital += toNum(r.capital) || 0; p.mobiles = [...new Set([...p.mobiles, ...mobiles])]; if (!p.mobile) p.mobile = normMobile(r.mobile); p.accounts++; continue; }
    personBy.set(key, {
      name: key, nameRaw: String(r.name).replace(/^[\s\u200c\u200e\u200f]+|[\s\u200c\u200e\u200f]+$/g, ""),
      mobile: normMobile(r.mobile), capital: toNum(r.capital) || 0, mobiles: [...new Set(mobiles)], code: (r._roster && r._roster.code) || "", accounts: 1,
    });
  }
  const people = [...personBy.values()];
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
    tx.push({ jdn, t: timeOf(r.date), who: norm(r.who), type: norm(r.type), amount: (toNum(r.manager) || 0) + (toNum(r.online) || 0) });
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

  const fixKeys = new Map([...fixMap].map(([f, t]) => [f, norm(t)])); // نام قدیمی ← نام درست (هر دو نرمال)
  return { header, memberCount, personalCapitalSum, names, fixKeys, rosterFind, people, hasMobile: members.found.mobile, tx, loans, notes, hasLoanNames: loansT.found.who };
}

// ---------------------------------------------------------------------------- خواندن فایل درخواست‌ها
// شماره‌ی موبایل ← ۱۰ رقم آخر (۹۱۲…)؛ برای تطبیق مطمئن‌تر از نام
function normMobile(v) {
  const d = latinDigits(v == null ? "" : v).replace(/\D/g, "");
  return d.length >= 10 ? d.slice(-10) : "";
}

// دو قالب پذیرفته می‌شود: خروجی پاسخ‌های پُرس‌لاین (پیشنهادی) یا loan-requests.xlsx قدیمی
function parseRequests(wb) {
  const kind = detectKind(wb);
  const list = [];
  let badDate = 0;
  if (kind === "porsline") {
    const t = readTable(wb, "porsline");
    // ستون‌های نام ضامن: ستون با عنوان دقیق، به‌علاوه‌ی هر ستون دیگری که «ضامن» دارد و حساس نیست
    const H = CONFIG.sheets.porsline.guarantorNameHint;
    const gCols = t.headCells
      .map((c, i) => i)
      .filter((i) => i === t.idx.gName || i !== t.idx.gMobile && t.headCells[i].includes(norm(H.has)) && !H.not.some((w) => t.headCells[i].includes(norm(w))))
      .sort((a, b) => (b === t.idx.gName) - (a === t.idx.gName)); // ستون با عنوان دقیق اول
    const isName = (v) => v != null && /[\u0621-\u064A\u067E\u0686\u0698\u06A9\u06AF\u06CCA-Za-z]/.test(String(v));
    t.body.forEach((r, k) => {
      const nameRaw = String(r.name == null ? "" : r.name).trim();
      if (!nameRaw && !r.mobile) return; // ردیف خالی
      const jdn = parseDate(r.date);
      if (jdn == null) { badDate++; return; }
      const row = t.rows[k] || [];
      const gv = gCols.map((i) => row[i]).find(isName);
      const gNameRaw = String(gv == null ? "" : gv).trim();
      list.push({
        name: norm(nameRaw), nameRaw, mobile: normMobile(r.mobile), jdn, cancelled: false, received: null,
        gName: norm(gNameRaw), gNameRaw, gMobile: normMobile(r.gMobile),
        amount: toNum(r.amount), amountRaw: r.amount == null ? "" : String(r.amount).trim(), count: toNum(r.count),
      });
    });
    return { list, badDate, source: "porsline", hasGuarantor: gCols.length > 0 || !!t.found.gMobile };
  }
  const t = readTable(wb, "requests");
  t.body.forEach((r) => {
    const nameRaw = String(r.name == null ? "" : r.name).trim();
    if (!nameRaw) return;
    const jdn = parseDate(r.date);
    if (jdn == null) { badDate++; return; }
    const status = norm(r.status);
    const received = parseDate(r.received);
    list.push({ name: norm(nameRaw), nameRaw, mobile: "", jdn, cancelled: status === norm(CONFIG.cancelWord), received, gName: "", gMobile: "" });
  });
  return { list, badDate, source: "manual", hasGuarantor: false };
}

// تشخیص نوع فایل از روی شیت‌هایش
function detectKind(wb) {
  const has = (key) => {
    try { const s = findSheet(wb, CONFIG.sheets[key]); return locateHeader(s.rows, CONFIG.sheets[key].columns).ok; } catch (e) { return false; }
  };
  if (has("tx")) return "fund";
  if (has("porsline")) return "porsline";
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
  const lateBy = new Map(); // تعداد قسط معوق هر گیرنده، روی وام‌های در جریانش
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
    if (delta < 0) lateBy.set(l.who, (lateBy.get(l.who) || 0) - delta);
    const s = delta > 0 ? "ahead" : delta === 0 ? "ok" : delta === -1 ? "late1" : "late2";
    status[s]++;
    dots.push(s);
  }
  const order = { ok: 0, ahead: 1, late1: 2, late2: 3 };
  dots.sort((a, b) => order[a] - order[b]);
  const monthLoans = loans.filter((l) => l.jdn >= monthStart && l.jdn <= asOf);

  // --- زمان انتظار و صف (از فایل درخواست‌ها)
  let wait = null, guarantors = null, review = null;
  if (req && !data.hasLoanNames) {
    checks.push({ level: "error", text: `ستون «${CONFIG.sheets.loans.optional.who}» در شیت وام‌ها پیدا نشد؛ درخواست‌ها به وام‌ها وصل نمی‌شوند و بخش «انتظار برای وام» ساخته نشد.` });
  } else if (req) {
    // هر درخواست ← نام عضو در نرم‌افزار: اول از روی موبایل، اگر نشد از روی نام
    const byMobile = new Map();
    data.people.forEach((p) => { if (p.mobile && !byMobile.has(p.mobile)) byMobile.set(p.mobile, p.name); });
    data.people.forEach((p) => (p.mobiles || []).forEach((m) => { if (!byMobile.has(m)) byMobile.set(m, p.name); }));
    const byName = (name) => (!name ? null : data.names.has(name) ? name : data.fixKeys && data.fixKeys.has(name) && data.names.has(data.fixKeys.get(name)) ? data.fixKeys.get(name) : null);
    const resolve = (mobile, name) => (mobile && byMobile.get(mobile)) || byName(name);

    const loansBy = new Map();
    loans.forEach((l) => { if (!loansBy.has(l.who)) loansBy.set(l.who, []); loansBy.get(l.who).push(l); });
    const since = (() => { const d = J.d2j(asOf); return J.addMonths(d.jy, d.jm, d.jd, -CONFIG.waitMonths); })();
    const expiry = (() => { const d = J.d2j(asOf); return J.addMonths(d.jy, d.jm, d.jd, -CONFIG.requestExpiryMonths); })();
    const waits = [], unknownNames = new Set(), manualMiss = [];
    let future = 0, cancelled = 0, queue = 0, expired = 0, superseded = 0;
    const reqOfLoan = new Map(); // وام ← درخواستی که به آن رسیده (برای پیدا کردن ضامن)
    const queueList = []; // صف: آخرین درخواستِ هر نفر که هنوز به وام نرسیده و منقضی نشده

    // درخواست‌ها به تفکیک شخص
    const reqBy = new Map();
    for (const r of req.list) {
      if (r.jdn > asOf) future++; // بعد از تاریخ اکسل صندوق ثبت شده؛ هنوز وامی به آن نرسیده، پس در صف حساب می‌شود
      if (r.cancelled) { cancelled++; continue; }
      const who = resolve(r.mobile, r.name);
      if (!who) { unknownNames.add(r.nameRaw || r.mobile); continue; }
      if (r.received != null) { // تاریخ دستی (فقط در فایل قدیمی)
        const l = (loansBy.get(who) || []).find((x) => x.jdn === r.received);
        if (l) reqOfLoan.set(l, r); else manualMiss.push(r.nameRaw);
        if (r.received <= asOf && r.received > since) waits.push(r.received - r.jdn);
        continue;
      }
      if (!reqBy.has(who)) reqBy.set(who, []);
      reqBy.get(who).push(r);
    }
    // برای هر وام: آخرین درخواستِ همان شخص بین وام قبلی‌اش و این وام
    for (const [who, rs] of reqBy) {
      rs.sort((a, b) => a.jdn - b.jdn);
      const ls = (loansBy.get(who) || []).filter((l) => l.jdn <= asOf);
      let prev = -Infinity;
      for (const l of ls) {
        const win = rs.filter((r) => r.jdn > prev && r.jdn <= l.jdn);
        if (win.length) {
          const r = win[win.length - 1];
          superseded += win.length - 1;
          reqOfLoan.set(l, r);
          if (l.jdn > since) waits.push(l.jdn - r.jdn);
        }
        prev = l.jdn;
      }
      const pending = rs.filter((r) => r.jdn > prev);
      if (pending.length) {
        superseded += pending.length - 1;
        const lastReq = pending[pending.length - 1];
        if (lastReq.jdn >= expiry) { queue++; queueList.push({ who, r: lastReq, n: pending.length }); } else expired++;
      }
    }
    wait = {
      n: waits.length,
      median: median(waits), // میانه: نیمی از وام‌ها در همین مدت یا کمتر پرداخت شده‌اند
      queue,
    };
    checks.push({ level: "ok", text: `فایل درخواست‌ها${req.source === "porsline" ? " (پُرس‌لاین)" : ""}: ${fmtInt(req.list.length)} درخواست خوانده شد؛ ${fmtInt(waits.length)} وام در ${fmtInt(CONFIG.waitMonths)} ماه اخیر به درخواستش وصل شد و ${fmtInt(queue)} نفر در صف‌اند.` });
    if (superseded) checks.push({ level: "ok", text: `${fmtInt(superseded)} درخواست تکراری بود؛ از هر نفر آخرین درخواستش حساب شد.` });
    if (expired) checks.push({ level: "warn", text: `${fmtInt(expired)} نفر درخواستی دارند که بیش از ${fmtInt(CONFIG.requestExpiryMonths)} ماه به وام نرسیده؛ منقضی حساب شد و در صف نیامد.` });
    if (req.source === "porsline" && !data.hasMobile) checks.push({ level: "warn", text: "ستون «شماره موبایل» در شیت اعضا پیدا نشد؛ تطبیق فقط از روی نام انجام شد." });
    if (req.badDate) checks.push({ level: "warn", text: `${fmtInt(req.badDate)} ردیف از فایل درخواست‌ها تاریخ قابل‌خواندن نداشت و کنار گذاشته شد.` });
    if (unknownNames.size) checks.push({ level: "warn", text: `این متقاضی‌ها با هیچ عضوی جور نشدند (نه از روی موبایل، نه نام): ${[...unknownNames].join("، ")}` });
    if (manualMiss.length) checks.push({ level: "warn", text: `برای این درخواست‌ها «تاریخ دریافت» دستی پر شده ولی وامی با همان تاریخ به همان نام پیدا نشد؛ تاریخ دستی مبنا قرار گرفت: ${manualMiss.join("، ")}` });
    if (future) checks.push({ level: "ok", text: `${fmtInt(future)} درخواست بعد از تاریخ اکسل صندوق ثبت شده و در صف حساب شد.` });
    queueList.sort((a, b) => a.r.jdn - b.r.jdn);

    // --- سنجش اعضا: فهرست ضامن‌های مجاز و کارت بررسی درخواست‌ها (فقط برای مدیر صندوق، هرگز در داشبورد)
    {
      const G = CONFIG.guarantor;
      const load = new Map(), unknownGNames = []; let unknownG = 0, noReq = 0;
      if (req.hasGuarantor) for (const l of active) {
        const r = reqOfLoan.get(l);
        if (!r) { noReq++; continue; }
        const g = resolve(r.gMobile, r.gName);
        if (!g) { unknownG++; unknownGNames.push(r.gNameRaw || r.gMobile || "(خالی)"); continue; }
        load.set(g, (load.get(g) || 0) + 1);
      }
      const ago = (months) => { const d = J.d2j(asOf); return J.addMonths(d.jy, d.jm, d.jd, -months); };
      const txBy = new Map();
      for (const t of tx) { if (!txBy.has(t.who)) txBy.set(t.who, []); txBy.get(t.who).push(t); }
      const activeBy = new Map();
      for (const l of active) activeBy.set(l.who, (activeBy.get(l.who) || 0) + l.remaining);
      // بیشترین تأخیر قسط‌های سررسیدشده در بازه‌ی اخیر: برنامه‌ی قسط‌ها (ماهانه، از یک ماه پس از وام) در برابر بازپرداخت‌های تجمعی
      const lookback = ago(G.lateLookbackMonths);
      const maxDelay = (who) => {
        const sched = [];
        for (const l of loans) {
          if (l.who !== who) continue;
          const d = J.d2j(l.jdn), per = l.gross / l.count;
          for (let k = 1; k <= l.count; k++) sched.push({ due: J.addMonths(d.jy, d.jm, d.jd, k), a: per });
        }
        if (!sched.length) return 0;
        sched.sort((a, b) => a.due - b.due);
        const rep = (txBy.get(who) || []).filter((t) => T[t.type] === "repay");
        let cumDue = 0, cumPaid = 0, j = 0, worst = 0;
        for (const s of sched) {
          if (s.due >= asOf) break;
          cumDue += s.a;
          const need = cumDue - s.a * 0.05; // تا ۵٪ اختلاف گرد کردن نادیده
          while (cumPaid < need && j < rep.length) cumPaid += Math.abs(rep[j++].amount);
          const paidAt = cumPaid >= need ? rep[j - 1].jdn : asOf;
          if (s.due > lookback) worst = Math.max(worst, paidAt - s.due);
        }
        return worst;
      };
      // سقف ضمانت هر نفر: پایه، به‌علاوه‌ی پله‌های سرمایه‌ی زیاد
      const X = G.extraGuarantees;
      const capOf = (capital) => !X ? G.maxActiveGuarantees
        : Math.min(Math.max(X.max, G.maxActiveGuarantees), G.maxActiveGuarantees + Math.max(0, Math.floor((capital - X.from) / X.step)));
      const reasons = [
        { key: "late", label: G.maxLateInstallments ? `بیش از ${fmtInt(G.maxLateInstallments)} قسط معوق` : "قسط معوق دارد", rule: G.maxLateInstallments ? `حداکثر ${fmtInt(G.maxLateInstallments)} قسط معوق` : "هیچ قسط معوق نداشته باشد" },
        { key: "history", label: `قسطی بیش از ${fmtInt(G.lateDaysLimit)} روز دیر در ${fmtInt(G.lateLookbackMonths)} ماه اخیر`, rule: `در ${fmtInt(G.lateLookbackMonths)} ماه اخیر هیچ قسطی را بیش از ${fmtInt(G.lateDaysLimit)} روز دیر نداده باشد` },
        { key: "inactive", label: `بدون تراکنش در ${fmtInt(G.inactiveMonths)} ماه اخیر`, rule: `در ${fmtInt(G.inactiveMonths)} ماه اخیر دست‌کم یک تراکنش داشته باشد` },
        { key: "tenure", label: `کمتر از ${fmtInt(G.minTenureMonths)} ماه عضویت`, rule: `دست‌کم ${fmtInt(G.minTenureMonths)} ماه عضو باشد` },
        { key: "withdraw", label: `برداشت از سرمایه در ${fmtInt(G.noWithdrawMonths)} ماه اخیر`, rule: `در ${fmtInt(G.noWithdrawMonths)} ماه اخیر از سرمایه‌اش برداشت نکرده باشد` },
        { key: "capital", label: `سرمایه‌ی کمتر از ${fmtMoney(G.minCapital, true)}`, rule: `سرمایه‌ی شخصی دست‌کم ${fmtMoney(G.minCapital, true)}` },
        { key: "debt", label: `مانده‌ی وام خودش بیش از ${fmtInt(G.maxOwnDebtRatio)} برابر سرمایه`, rule: `مانده‌ی وام‌های خودش حداکثر ${fmtInt(G.maxOwnDebtRatio)} برابر سرمایه‌اش` },
        { key: "cap", label: "به سقف ضمانتش رسیده",
          rule: `کمتر از ${fmtInt(G.maxActiveGuarantees)} ضمانت وام در جریان` + (X ? `؛ به ازای هر ${fmtMoney(X.step, true)} سرمایه بیش از ${fmtMoney(X.from, true)}، یک ضمانت بیشتر (حداکثر ${fmtInt(X.max)})` : "") },
      ];
      const labelOf = Object.fromEntries(reasons.map((x) => [x.key, x.label]));
      const personBy = new Map(data.people.map((p) => [p.name, p]));
      // همه‌ی آنچه درباره‌ی یک عضو لازم است؛ why = اولین شرط ضامنی که ندارد (یا null)
      const assess = (name) => {
        const p = personBy.get(name);
        if (!p) return null;
        const all = txBy.get(name) || [];
        const late = lateBy.get(name) || 0, n = load.get(name) || 0, delay = maxDelay(name);
        const mine = all.filter((t) => T[t.type] !== "loan" && T[t.type] !== "settle");
        const first = all[0], lastMine = mine[mine.length - 1];
        const withdraw = mine.some((t) => T[t.type] === "withdraw" && t.jdn > ago(G.noWithdrawMonths));
        const debt = activeBy.get(name) || 0;
        const why =
          late > G.maxLateInstallments ? "late"
          : delay > G.lateDaysLimit ? "history"
          : !lastMine || lastMine.jdn <= ago(G.inactiveMonths) ? "inactive"
          : !first || first.jdn > ago(G.minTenureMonths) ? "tenure"
          : withdraw ? "withdraw"
          : p.capital < G.minCapital ? "capital"
          : debt > G.maxOwnDebtRatio * Math.max(p.capital, 0) ? "debt"
          : n >= capOf(p.capital) ? "cap"
          : null;
        const myLoans = loansBy.get(name) || [];
        return {
          name: p.nameRaw, capital: p.capital, late, maxDelay: delay, guarantees: n, maxGuarantees: capOf(p.capital), withdraw,
          tenureDays: first ? asOf - first.jdn : 0, lastTxDays: lastMine ? asOf - lastMine.jdn : null,
          loansN: myLoans.length, activeDebt: debt, activeLoansN: active.filter((l) => l.who === name).length,
          why, whyLabel: why ? labelOf[why] : null,
        };
      };
      if (req.hasGuarantor) {
        const eligible = [], out = Object.fromEntries(reasons.map((x) => [x.key, []]));
        for (const p of data.people) {
          const a = assess(p.name);
          if (a.why) out[a.why].push(p.nameRaw);
          else eligible.push({ name: p.nameRaw, free: a.maxGuarantees - a.guarantees });
        }
        const fa = (a, b) => a.localeCompare(b, "fa");
        eligible.sort((a, b) => fa(a.name, b.name));
        guarantors = { eligible, out, reasons: reasons.map((x) => ({ ...x, names: out[x.key] })), noReq, unknownG, unknownGNames, activeN: active.length };
      }
      review = { assess, resolve, rosterFind: data.rosterFind, queue: queueList, list: req.list, waitMedian: wait.median, hasGuarantor: req.hasGuarantor };
    }
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
    guarantors, // فقط برای مدیر صندوق؛ در renderReport استفاده نمی‌شود
    review, // سنجش اعضا و صف برای کارت بررسی درخواست‌ها؛ فقط برای مدیر صندوق
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

module.exports = { CONFIG, THEME, parseWorkbook, parseRequests, detectKind, compute, renderReport, loadNameFixes, fmtDate, fmtNum, fmtInt, fmtMoney, ReportError };
