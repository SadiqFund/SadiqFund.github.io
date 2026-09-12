// ساخته‌شده از همان کد سازنده (helpers.js + core.js + render.js + email.js)؛ دستی ویرایش نکنید.
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
// راه‌های ارتباط با صندوق (پایین داشبورد و در ایمیل). برای تغییر فقط همین‌جا را اصلاح کنید.
const CONTACT = {
  bale: { handle: "@sadiqanidishan", url: "https://ble.ir/sadiqanidishan" },
  email: "sadiqandishan@gmail.com",
};

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
    maxLateInstallments: 1, // اگر روی وام‌های خودش بیش از این تعداد قسط معوق داشته باشد، ضامن نمی‌شود
    lateDaysLimit: 45, lateLookbackMonths: 12, // در این چند ماه هیچ قسطی بیش از این چند روز دیر پرداخت نشده باشد
    inactiveMonths: 3, // در این چند ماه دست‌کم یک تراکنش (حق عضویت، قسط، واریز، …) داشته باشد
    minTenureMonths: 6, // دست‌کم این چند ماه از اولین تراکنشش گذشته باشد
    minCapital: 10000000, // حداقل سرمایه‌ی شخصی (تومان)
    maxOwnDebtRatio: 2.5, // مانده‌ی وام‌های در جریان خودش بیش از این چند برابر سرمایه‌اش نباشد
    maxActiveGuarantees: 2, // حداکثر ضمانت وام‌های در جریان (پایه)
    // استثنا برای سرمایه‌ی زیاد: از «از» به بالا یک ضمانت بیشتر، و به ازای هر «پله» بیشتر یکی دیگر، تا «سقف»
    // (۵۰ میلیون ← ۳، ۷۰ میلیون ← ۴، ۹۰ میلیون و بیشتر ← ۵)
    extraGuarantees: { from: 50000000, step: 20000000, max: 5 },
    // فقط یادداشت اطلاعاتی در کارت بررسی؛ شرط رد شدن نیست
    withdrawNoteMonths: 12,
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
      const [name, code0, ...mobiles] = line.split("|").map((x) => x.trim());
      if (!name) continue;
      const code = latinDigits(code0 || "").replace(/\D/g, "");
      // همان نفر (نام یکسان و سال ورود یکسان یا خالی) ← شماره‌ها به همان ردیف اضافه می‌شوند؛ مثلاً شماره‌ی دومی که دستی اضافه شده
      const same = CONFIG.roster.filter((e) => norm(e.name) === norm(name) && (!code || !e.code || e.code === code));
      if (same.length === 1) {
        same[0].mobiles = [...new Set([...same[0].mobiles, ...mobiles.filter(Boolean)])];
        if (!same[0].code) same[0].code = code;
      } else CONFIG.roster.push({ name, code, mobiles: mobiles.filter(Boolean) });
      roster++;
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
  // ردیف‌هایی از فهرست که به هیچ عضوی وصل نشدند ولی نامشان با نام (قدیم یا درست) عضوی یکی است ← شماره‌هایشان مال همان عضو است
  const used = new Set(rows.map((r) => r._roster).filter(Boolean));
  const extra = new Map();
  for (const e of roster) if (!used.has(e)) extra.set(e.key, [...(extra.get(e.key) || []), ...e.mobiles]);
  rows.forEach((r, i) => {
    const p = prop.find((x) => x.k === norm(r.name));
    const more = [...(extra.get(norm(r.name)) || []), ...(p ? extra.get(norm(p.to)) || [] : [])];
    if (more.length) r._extraMobiles = more;
  });
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
    const mobiles = [normMobile(r.mobile), ...((r._roster && r._roster.mobiles) || []), ...(r._extraMobiles || [])].filter(Boolean); // موبایل‌های دیگر از فهرست اعضا
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
      const load = new Map(), unknownGList = []; let unknownG = 0, noReq = 0;
      if (req.hasGuarantor) for (const l of active) {
        const r = reqOfLoan.get(l);
        if (!r) { noReq++; continue; }
        const g = resolve(r.gMobile, r.gName);
        if (!g) {
          unknownG++;
          const rf = data.rosterFind ? data.rosterFind(r.gMobile, r.gName) : null;
          unknownGList.push({ borrower: l.whoRaw, typed: r.gNameRaw || (r.gMobile ? "0" + r.gMobile : ""), roster: rf ? rf.name : null });
          continue;
        }
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
      const capOf = (capital) => !X || capital < X.from ? G.maxActiveGuarantees
        : Math.min(Math.max(X.max, G.maxActiveGuarantees), G.maxActiveGuarantees + 1 + Math.floor((capital - X.from) / X.step));
      const reasons = [
        { key: "late", label: G.maxLateInstallments ? `بیش از ${fmtInt(G.maxLateInstallments)} قسط معوق` : "قسط معوق دارد", rule: G.maxLateInstallments ? `حداکثر ${fmtInt(G.maxLateInstallments)} قسط معوق` : "هیچ قسط معوق نداشته باشد" },
        { key: "history", label: `قسطی بیش از ${fmtInt(G.lateDaysLimit)} روز دیر در ${fmtInt(G.lateLookbackMonths)} ماه اخیر`, rule: `در ${fmtInt(G.lateLookbackMonths)} ماه اخیر هیچ قسطی را بیش از ${fmtInt(G.lateDaysLimit)} روز دیر نداده باشد` },
        { key: "inactive", label: `بدون پرداخت در ${fmtInt(G.inactiveMonths)} ماه اخیر`, rule: `در ${fmtInt(G.inactiveMonths)} ماه اخیر دست‌کم یک پرداخت داشته باشد: قسط، حق عضویت یا افزایش موجودی` },
        { key: "tenure", label: `کمتر از ${fmtInt(G.minTenureMonths)} ماه عضویت`, rule: `دست‌کم ${fmtInt(G.minTenureMonths)} ماه عضو باشد` },
        { key: "capital", label: `سرمایه‌ی کمتر از ${fmtMoney(G.minCapital, true)}`, rule: `سرمایه‌ی شخصی دست‌کم ${fmtMoney(G.minCapital, true)}` },
        { key: "debt", label: `مانده‌ی وام خودش بیش از ${fmtRatio(G.maxOwnDebtRatio)} برابر سرمایه`, rule: `مانده‌ی وام‌های خودش حداکثر ${fmtRatio(G.maxOwnDebtRatio)} برابر سرمایه‌اش` },
        { key: "cap", label: "به سقف ضمانتش رسیده",
          rule: `کمتر از ${fmtInt(G.maxActiveGuarantees)} ضمانت وام در جریان` + (X ? `؛ از ${fmtMoney(X.from, true)} سرمایه به بالا ${fmtInt(G.maxActiveGuarantees + 1)} وام، و به ازای هر ${fmtMoney(X.step, true)} بیشتر یک وام دیگر (حداکثر ${fmtInt(X.max)})` : "") },
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
        // «فعال بودن» با پرداخت‌های خود عضو سنجیده می‌شود: قسط، حق عضویت، افزایش موجودی.
        // دریافت وام و تسویه با مدیر پرداخت او نیستند، و برداشت هم پول گرفتن است نه پرداخت.
        const paid = all.filter((t) => ["repay", "fee", "deposit"].includes(T[t.type]));
        const first = all[0], lastMine = paid[paid.length - 1];
        // برداشت از سرمایه دیگر شرط رد نیست؛ فقط برای یادداشت در کارت بررسی نگه داشته می‌شود
        const withdraw = mine.some((t) => T[t.type] === "withdraw" && t.jdn > ago(G.withdrawNoteMonths));
        const debt = activeBy.get(name) || 0;
        const why =
          late > G.maxLateInstallments ? "late"
          : delay > G.lateDaysLimit ? "history"
          : !lastMine || lastMine.jdn <= ago(G.inactiveMonths) ? "inactive"
          : !first || first.jdn > ago(G.minTenureMonths) ? "tenure"
          : p.capital < G.minCapital ? "capital"
          : debt > G.maxOwnDebtRatio * Math.max(p.capital, 0) ? "debt"
          : n >= capOf(p.capital) ? "cap"
          : null;
        const myLoans = loansBy.get(name) || [];
        return {
          name: p.nameRaw, capital: p.capital, late, maxDelay: delay, guarantees: n, maxGuarantees: capOf(p.capital), withdraw,
          tenureDays: first ? asOf - first.jdn : 0, lastTxDays: lastMine ? asOf - lastMine.jdn : null, lastTxJdn: lastMine ? lastMine.jdn : null, lastTxType: lastMine ? T[lastMine.type] : null,
          loansN: myLoans.length, activeDebt: debt, activeLoansN: active.filter((l) => l.who === name).length,
          why, whyLabel: why ? labelOf[why] : null,
        };
      };
      if (req.hasGuarantor) {
        const eligible = [], out = Object.fromEntries(reasons.map((x) => [x.key, []]));
        for (const p of data.people) {
          const a = assess(p.name);
          if (a.why) out[a.why].push(a);
          else eligible.push({ name: p.nameRaw, free: a.maxGuarantees - a.guarantees });
        }
        const fa = (a, b) => a.localeCompare(b, "fa");
        eligible.sort((a, b) => fa(a.name, b.name));
        // هر دلیل: فهرست افراد با جزئیات سنجششان (برای پیام تفکیک‌شده‌ی مدیر)
        guarantors = { eligible, out, reasons: reasons.map((x) => ({ ...x, names: out[x.key].map((a) => a.name), people: out[x.key] })), noReq, unknownG, unknownGList, activeN: active.length };
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
// عدد با اعشار کم (مثل ۲٫۵ برابر سرمایه)؛ اگر رُند باشد اعشار نمی‌گیرد
const fmtRatio = (n) => Number(n).toLocaleString("fa-IR", { maximumFractionDigits: 2 });
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
--good:${THEME.good};--ahead:${THEME.ahead};--late1:${THEME.late1};--late2:${THEME.late2};--s-repay:${THEME.sRepay};--s-fee:${THEME.sFee};--s-deposit:${THEME.sDeposit};--s-other:${THEME.sOther};--r:${THEME.radius};--hero:${THEME.heroGradient};--shadow:${THEME.shadow};
--soft:#F6F8FC;--grid:#E7ECF3;--grid2:#C8D2E0;--selbg:rgba(0,73,232,.07);--icobg:#E6EEFF;--track:#DCE8FE;--chip:#fff;--logo:#0A1A4F;color-scheme:light}
/* حالت تیره: وقتی گوشی یا رایانه روی حالت شب است */
@media (prefers-color-scheme:dark){:root{--bg:#0B1220;--card:#131C2E;--border:#243049;--ink:#E7ECF5;--muted:#93A1B8;--primary:#4F86FF;--primary-dark:#2F6BF0;--good:#4F86FF;--ahead:#8FB4FF;--late1:#F0A3AC;--late2:#FF5A67;--s-repay:#4F86FF;--s-fee:#22B8B2;--s-deposit:#E88A1A;
--soft:#1A2439;--grid:#1F2A40;--grid2:#34425F;--selbg:rgba(79,134,255,.13);--icobg:#1E2A44;--track:#26324C;--chip:#131C2E;--shadow:none;--logo:#E7ECF5;color-scheme:dark}
.hero{box-shadow:0 0 0 1px rgba(255,255,255,.06)}.tip{background:#E7ECF5;color:#0B1220}.pill{box-shadow:none}}
*{box-sizing:border-box;margin:0;padding:0}
html{background:var(--bg);-webkit-text-size-adjust:100%}
body{font-family:"Dana FaNum",Tahoma,sans-serif;color:var(--ink);background:var(--bg);font-size:15px;line-height:1.8;font-variant-numeric:tabular-nums}
.page{max-width:1120px;margin:0 auto;padding:0 16px 32px}
.top{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;padding:18px 0 16px}
.brand{display:flex;align-items:center;gap:12px}
.logo{width:46px;height:46px;display:grid;place-items:center;flex:none;color:var(--logo)}
.logo svg{width:100%;height:100%;display:block}
.brand h1{font-size:20px;font-weight:800;line-height:1.4;color:var(--ink)}
.brand h1.wm{height:24px;line-height:0;color:var(--logo);margin:0}
.brand h1.wm svg{height:100%;width:auto;display:block}
.sr{position:absolute;width:1px;height:1px;padding:0;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0}
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
.i-blue,.i-teal,.i-violet,.i-amber{background:var(--icobg);color:var(--primary)} /* آیکن‌ها یکدست؛ رنگ‌های دیگر فقط در نمودارها معنا دارند */
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
.side{background:var(--soft);border-radius:12px;padding:16px;display:flex;flex-direction:column;gap:6px}
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
.chips span{display:flex;justify-content:space-between;align-items:center;gap:10px;background:var(--chip);border-radius:10px;padding:6px 10px;font-size:12px;color:var(--muted);box-shadow:0 1px 2px rgba(16,24,40,.05);border-right:3px solid #0B8480}
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
.ring .rt>div{background:var(--soft);border-radius:12px;padding:10px 14px}
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
.gl{position:absolute;left:0;right:0;border-top:1px solid var(--grid)}
.gl span{position:absolute;right:calc(100% + 8px);top:-.85em;font-size:11px;color:var(--muted);white-space:nowrap;direction:rtl;line-height:1.6}
.gl.base{border-top-color:var(--grid2)}
.avg{position:absolute;left:0;right:0;border-top:2px dashed #8A96A0;pointer-events:none;z-index:2}
.cols{position:absolute;inset:0;display:flex}
.cols button{flex:1;min-width:0;height:100%;border:0;background:transparent;cursor:pointer;display:flex;flex-direction:column-reverse;align-items:center;gap:2px;padding:0;border-radius:4px 4px 0 0}
.cols button.sel{background:var(--selbg)}
.cols i{display:block;width:min(36px,60%);flex:none}
.cols i.top{border-radius:6px 6px 0 0}
.xl{display:flex;height:1.9em;margin-top:4px}
.xl span{flex:1;min-width:0;text-align:center;font-size:11px;color:var(--muted);white-space:nowrap;direction:rtl}
.xl span.sel{color:var(--ink);font-weight:600}
.contact{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;padding:26px 0 2px}
.contact a{display:inline-flex;align-items:center;gap:8px;background:var(--card);border-radius:999px;padding:8px 16px;font-size:13px;color:var(--muted);text-decoration:none;box-shadow:var(--shadow);transition:color .15s ease,transform .15s ease}
.contact a:hover{color:var(--ink);transform:translateY(-1px)}
.contact a b{color:var(--ink);font-weight:600;direction:ltr;unicode-bidi:isolate}
.contact svg{width:17px;height:17px;flex:none;fill:none;stroke:currentColor;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round;color:var(--primary)}
.contact svg.bico{fill:#0CBC8D;stroke:none;width:18px;height:18px}
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
/* حباب شناور روی نمودار رشد */
.gtip{position:absolute;z-index:4;background:#fff;color:var(--accent);font-size:12px;line-height:1.55;padding:5px 10px;border-radius:8px;white-space:nowrap;pointer-events:none;box-shadow:0 4px 14px rgba(3,4,94,.28);transform:translate(-50%,calc(-100% - 14px));text-align:center;direction:rtl}
.gtip b{display:block;font-size:13px;font-weight:800;color:var(--primary-dark)}
.gtip::after{content:"";position:absolute;left:50%;bottom:-5px;width:10px;height:10px;background:#fff;transform:translateX(-50%) rotate(45deg);border-radius:2px}
.gtip.below{transform:translate(-50%,14px)}.gtip.below::after{bottom:auto;top:-5px}
/* کم‌رنگ شدن بقیه‌ی ستون‌ها وقتی روی یکی هستید */
.cols i{transition:filter .12s,opacity .15s}
.plot.dim .cols button:not(.hov) i{opacity:.32}
/* رنگ حلقه‌ها از متغیرها (برای حالت تیره) */
.rtrack{stroke:var(--track)}.rp{stroke:var(--primary)}.rl{fill:var(--ink)}.rs{fill:var(--muted)}
/* برچسب رشد ۱۲ ماه زیر عدد دارایی */
.gbadge{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.16);border-radius:999px;padding:2px 12px;font-size:13px;font-weight:600;color:#fff}
.gbadge b{font-weight:800}
/* نبض آرام آخرین نقطه‌ی نمودار رشد («امروز») */
@keyframes pulse{0%{transform:scale(1);opacity:.55}70%{transform:scale(3.2);opacity:0}100%{transform:scale(3.2);opacity:0}}
.gpulse{transform-box:fill-box;transform-origin:center;animation:pulse 2.6s ease-out infinite;pointer-events:none}
@media (prefers-reduced-motion:reduce){.gpulse{animation:none;opacity:0}}
/* عنوان گروه‌ها */
.grp{grid-column:1/-1;font-size:15px;font-weight:800;color:var(--ink);display:flex;align-items:center;gap:10px;margin:8px 2px -6px}
.grp::after{content:"";flex:1;border-top:1px solid var(--border)}
/* وضعیت بازپرداخت: دایره‌ها در دسکتاپ، نوار چندرنگ در موبایل */
.sbar{display:none;height:22px;border-radius:8px;overflow:hidden;gap:2px;margin:12px 0 14px}
.sbar i{display:block;transform-origin:100% 50%}
.sub-m{display:none}
/* «این ماه» فشرده بالای نمودار، فقط در موبایل و تبلت */
.mstrip{display:none;align-items:center;gap:12px;background:var(--soft);border-radius:12px;padding:10px 14px;margin:2px 0 10px;font-size:12px;color:var(--muted);line-height:1.7}
.mstrip svg{flex:none}.mstrip .mh{font-weight:600;color:var(--ink)}.mstrip b{color:var(--ink)}
.mstrip .mv{font-size:24px;font-weight:800;color:var(--ink);line-height:1.25;white-space:nowrap}.mstrip .mv small{font-size:12px;font-weight:600;color:var(--muted);margin-right:3px}
@media (max-width:899px){.mstrip{display:flex}.chartcard .side{display:none}}
@media (max-width:599px){.dots{display:none}.sbar{display:flex}.sub-d{display:none}.sub-m{display:inline}}
/* انیمیشن ورود (فقط بار اول هر بخش؛ اگر گوشی «کاهش حرکت» را خواسته باشد، خاموش است) */
@keyframes growx{from{transform:scaleX(0)}}
@keyframes rise{from{transform:scaleY(0)}}
@keyframes draw{from{stroke-dashoffset:1}}
@keyframes fade{from{opacity:0}}
@keyframes pop{from{opacity:0;transform:scale(.3)}}
@keyframes ringin{from{stroke-dasharray:0 1000}}
.anim [data-anim]:not(.in) .cols button{transform:scaleY(0)}
.anim [data-anim]:not(.in) .gline{stroke-dashoffset:1}
.anim [data-anim]:not(.in) .garea,.anim [data-anim]:not(.in) .gdot{opacity:0}
.anim [data-anim]:not(.in) .dot,.anim [data-anim]:not(.in) .rp,.anim [data-anim]:not(.in) .sbar i,.anim [data-anim]:not(.in) .gpulse{opacity:0}
.cols button{transform-origin:50% 100%}
.anim .in:not(.done) .cols button{animation:rise 1.15s cubic-bezier(.22,.61,.36,1) both;animation-delay:calc(var(--i,0)*90ms)}
.gline{stroke-dasharray:1}
.anim .in:not(.done) .gline{animation:draw 1.6s cubic-bezier(.33,.1,.25,1) both}
.anim .in:not(.done) .garea{animation:fade 1.3s .35s both}
.anim .in:not(.done) .gdot{animation:fade .5s 1.35s both}
.anim .in:not(.done) .gpulse{animation:fade .5s 1.6s both}
.anim .in:not(.done) .dots .dot{animation:pop .5s cubic-bezier(.2,.8,.3,1.15) both;animation-delay:calc(var(--i,0)*16ms)}
.anim .in:not(.done) .sbar i{animation:growx 1.2s cubic-bezier(.22,.61,.36,1) both;animation-delay:calc(var(--i,0)*150ms)}
.anim .in:not(.done) .rp{animation:ringin 1.5s cubic-bezier(.25,.1,.25,1) both}
`;

// نشان صندوق (بازرسم‌شده از لوگوی صندوق؛ رنگش را از محیط می‌گیرد)
// تایپ «صندوق صادق‌اندیشان» با فونت دوران، تبدیل‌شده به مسیر (فقط برای همین صندوق)
const WORDMARK_NAME = "صندوق صادق‌اندیشان";
const WORDMARK = `<svg viewBox="0 0 947.05 132.5" fill="currentColor" aria-hidden="true"><g transform="translate(-3.25,100) scale(1)"><path d="M32.25 25.00Q22.95 25.00 16.45 21.58Q9.95 18.15 6.60 11.70Q3.25 5.25 3.25 -3.75Q3.25 -12.40 5.15 -21.12Q7.05 -29.85 11.00 -39.00H12.25Q9.00 -29.25 9.00 -22.00Q9.00 -13.15 15.15 -8.82Q21.30 -4.50 34.00 -4.50Q43.10 -4.50 50.88 -6.33Q58.65 -8.15 63.20 -10.95Q67.75 -13.75 67.75 -16.50Q67.75 -18.00 63.23 -21.43Q58.70 -24.85 48.75 -30.25Q51.65 -34.35 55.33 -41.67Q59.00 -49.00 61.00 -55.00H61.75Q64.70 -47.95 66.85 -39.10Q69.00 -30.25 69.00 -18.00Q69.00 -4.30 64.72 5.33Q60.45 14.95 52.20 19.98Q43.95 25.00 32.25 25.00ZM19.25 -37.25V-38.00L30.75 -52.00Q34.35 -50.30 38.53 -46.72Q42.70 -43.15 44.75 -40.00L32.50 -27.00Q30.35 -29.70 26.18 -32.90Q22.00 -36.10 19.25 -37.25Z M80.00 -34.50Q79.25 -50.00 77.65 -60.60Q76.05 -71.20 73.00 -81.50Q76.90 -84.85 82.10 -90.43Q87.30 -96.00 90.00 -100.00H90.75Q90.75 -98.10 91.15 -92.10Q91.75 -82.60 91.75 -77.50Q91.75 -71.25 91.50 -65.95Q91.25 -60.65 90.80 -53.90Q90.25 -45.70 90.25 -42.50Q90.25 -37.70 91.72 -35.00Q93.20 -32.30 96.55 -31.15Q99.90 -30.00 105.75 -30.00L103.25 0.00Q94.90 0.00 90.20 -2.85Q85.50 -5.70 83.10 -13.08Q80.70 -20.45 80.00 -34.50Z M182.60 -18.55Q182.60 -20.20 182.75 -22.00H182.00Q181.20 -11.50 177.78 -5.75Q174.35 0.00 168.00 0.00Q161.85 0.00 158.72 -4.62Q155.60 -9.25 155.60 -18.00Q155.60 -19.90 155.75 -22.00H155.00Q153.65 -11.25 149.70 -5.62Q145.75 0.00 138.50 0.00Q132.50 0.00 128.93 -3.93Q125.35 -7.85 125.35 -15.30Q125.35 -18.40 126.00 -22.00H125.25Q120.10 0.00 103.25 0.00Q102.10 0.00 101.42 -2.65Q100.75 -5.30 100.75 -9.75Q100.75 -13.75 101.58 -18.55Q102.40 -23.35 103.60 -26.68Q104.80 -30.00 105.75 -30.00Q113.20 -30.00 117.88 -31.32Q122.55 -32.65 125.15 -35.55Q127.75 -38.45 128.75 -43.25H130.00L128.75 -32.50Q130.50 -31.40 133.70 -30.70Q136.90 -30.00 141.00 -30.00Q145.80 -30.00 148.98 -31.30Q152.15 -32.60 154.12 -35.48Q156.10 -38.35 157.25 -43.25H158.50L157.20 -32.25Q158.95 -31.15 161.95 -30.58Q164.95 -30.00 168.75 -30.00Q172.60 -30.00 175.45 -31.18Q178.30 -32.35 179.75 -34.75Q179.75 -36.40 178.68 -42.20Q177.60 -48.00 177.00 -49.75Q178.10 -50.10 179.93 -51.15Q181.75 -52.20 182.75 -53.00V-39.50Q182.75 -35.70 184.38 -33.68Q186.00 -31.65 189.82 -30.83Q193.65 -30.00 200.75 -30.00L198.25 0.00Q190.35 0.00 186.48 -4.80Q182.60 -9.60 182.60 -18.55ZM143.50 -84.50V-85.25L152.25 -96.25Q155.30 -94.75 158.95 -91.88Q162.60 -89.00 164.25 -86.75L155.25 -75.75Q153.10 -78.10 149.40 -80.80Q145.70 -83.50 143.50 -84.50ZM134.75 -62.35Q134.75 -63.55 135.07 -64.92Q135.40 -66.30 136.10 -68.50Q137.40 -72.65 137.75 -75.00H138.50Q138.50 -73.45 139.45 -72.60Q140.40 -71.75 142.60 -71.38Q144.80 -71.00 148.75 -71.00H149.00Q154.00 -71.00 161.45 -71.30Q168.90 -71.60 173.50 -72.00Q172.40 -69.95 170.70 -64.95Q169.00 -59.95 168.25 -56.50Q162.85 -56.10 156.35 -55.80Q149.85 -55.50 145.75 -55.50Q139.95 -55.50 137.35 -57.10Q134.75 -58.70 134.75 -62.35Z M203.25 0.00H198.25Q197.10 0.00 196.43 -2.65Q195.75 -5.30 195.75 -9.75Q195.75 -13.75 196.57 -18.55Q197.40 -23.35 198.60 -26.68Q199.80 -30.00 200.75 -30.00H205.75Q211.45 -30.00 217.30 -30.58Q223.15 -31.15 226.95 -32.02Q230.75 -32.90 230.75 -33.75Q230.75 -34.75 228.22 -37.10Q225.70 -39.45 221.47 -42.38Q217.25 -45.30 212.50 -48.00Q215.45 -51.90 219.18 -58.83Q222.90 -65.75 224.25 -69.75H225.00Q225.20 -69.05 226.00 -66.65Q228.50 -59.35 230.25 -52.00Q232.00 -44.65 232.00 -36.25Q232.00 -17.35 224.82 -8.68Q217.65 0.00 203.25 0.00ZM225.25 15.85Q225.25 17.05 224.88 18.52Q224.50 20.00 223.75 22.40Q222.35 26.90 222.00 29.25H221.25Q221.25 27.70 220.30 26.85Q219.35 26.00 217.15 25.62Q214.95 25.25 211.00 25.25H210.75Q205.75 25.25 198.30 25.55Q190.85 25.85 186.25 26.25Q187.35 24.20 189.18 18.83Q191.00 13.45 191.75 10.00Q197.15 9.60 203.65 9.30Q210.15 9.00 214.25 9.00Q220.05 9.00 222.65 10.60Q225.25 12.20 225.25 15.85ZM213.40 13.35 214.25 9.00Z M276.25 -37.75H275.50Q274.95 -18.45 269.38 -9.22Q263.80 0.00 252.75 0.00Q245.60 0.00 241.93 -4.70Q238.25 -9.40 238.25 -18.75Q238.25 -23.30 238.78 -27.33Q239.30 -31.35 240.60 -38.30L241.25 -42.00H242.00Q242.00 -37.45 243.30 -34.85Q244.60 -32.25 247.50 -31.12Q250.40 -30.00 255.50 -30.00Q260.30 -30.00 264.50 -31.10Q268.70 -32.20 271.23 -34.05Q273.75 -35.90 273.75 -38.00Q273.75 -39.00 272.27 -43.20Q270.80 -47.40 268.85 -51.95Q266.90 -56.50 265.50 -58.75Q267.15 -60.40 269.92 -64.30Q272.70 -68.20 275.45 -72.75Q278.20 -77.30 279.75 -80.75H280.50Q281.80 -73.70 283.25 -66.90Q284.70 -60.10 288.25 -45.50Q290.20 -37.30 294.55 -33.65Q298.90 -30.00 305.25 -30.00L302.75 0.00Q294.20 0.00 289.17 -3.58Q284.15 -7.15 281.25 -15.15Q278.35 -23.15 276.25 -37.75Z M307.75 0.00H302.75Q301.60 0.00 300.93 -2.65Q300.25 -5.30 300.25 -9.75Q300.25 -13.75 301.07 -18.55Q301.90 -23.35 303.10 -26.68Q304.30 -30.00 305.25 -30.00H310.25Q315.95 -30.00 321.80 -30.58Q327.65 -31.15 331.45 -32.02Q335.25 -32.90 335.25 -33.75Q335.25 -34.75 332.73 -37.10Q330.20 -39.45 325.98 -42.38Q321.75 -45.30 317.00 -48.00Q319.95 -51.90 323.67 -58.83Q327.40 -65.75 328.75 -69.75H329.50Q329.70 -69.05 330.50 -66.65Q333.00 -59.35 334.75 -52.00Q336.50 -44.65 336.50 -36.25Q336.50 -17.35 329.32 -8.68Q322.15 0.00 307.75 0.00ZM307.50 -84.75V-85.50L319.00 -99.50Q322.60 -97.80 326.77 -94.22Q330.95 -90.65 333.00 -87.50L320.75 -74.50Q318.60 -77.20 314.43 -80.40Q310.25 -83.60 307.50 -84.75Z M358.50 -87.45Q359.25 -55.45 359.25 -40.00Q359.25 -31.00 358.30 -24.80Q357.35 -18.60 355.98 -14.78Q354.60 -10.95 352.35 -6.30Q351.15 -3.75 350.40 -2.00Q349.65 -0.25 349.00 1.75H348.25L348.30 -0.50Q348.30 -8.75 346.98 -24.58Q345.65 -40.40 343.80 -56.60Q341.95 -72.80 340.50 -81.50Q344.40 -84.85 349.60 -90.43Q354.80 -96.00 357.50 -100.00H358.25Q358.25 -98.20 358.50 -87.45Z M432.80 -22.00Q432.80 25.00 395.30 25.00Q386.35 25.00 380.02 21.58Q373.70 18.15 370.38 11.68Q367.05 5.20 367.05 -3.75Q367.05 -12.40 368.95 -21.12Q370.85 -29.85 374.80 -39.00H376.05Q372.80 -29.25 372.80 -22.00Q372.80 -13.20 378.83 -8.85Q384.85 -4.50 397.05 -4.50Q408.50 -4.50 416.23 -6.33Q423.95 -8.15 427.75 -11.40Q431.55 -14.65 431.55 -19.00Q431.55 -20.55 430.48 -24.60Q429.40 -28.65 428.30 -31.00H427.55Q426.60 -20.40 421.60 -15.45Q416.60 -10.50 406.55 -10.50Q398.00 -10.50 393.77 -15.53Q389.55 -20.55 389.55 -30.25Q389.55 -41.45 392.25 -50.23Q394.95 -59.00 399.73 -63.88Q404.50 -68.75 410.55 -68.75Q421.80 -68.75 427.30 -57.12Q432.80 -45.50 432.80 -22.00ZM405.55 -47.25Q402.00 -47.25 398.83 -46.45Q395.65 -45.65 393.73 -44.53Q391.80 -43.40 391.80 -42.50Q391.80 -41.75 393.55 -41.05Q395.30 -40.35 398.50 -39.92Q401.70 -39.50 405.80 -39.50Q418.10 -39.50 425.55 -41.25Q422.65 -43.95 417.38 -45.60Q412.10 -47.25 405.55 -47.25ZM391.80 -82.35Q391.80 -83.55 392.18 -85.03Q392.55 -86.50 393.30 -88.90Q394.70 -93.40 395.05 -95.75H395.80Q395.80 -94.20 396.75 -93.35Q397.70 -92.50 399.90 -92.12Q402.10 -91.75 406.05 -91.75H406.30Q411.30 -91.75 418.75 -92.05Q426.20 -92.35 430.80 -92.75Q429.70 -90.70 427.88 -85.33Q426.05 -79.95 425.30 -76.50Q419.90 -76.10 413.40 -75.80Q406.90 -75.50 402.80 -75.50Q397.00 -75.50 394.40 -77.10Q391.80 -78.70 391.80 -82.35Z M453.55 0.00Q446.40 0.00 442.73 -4.70Q439.05 -9.40 439.05 -18.75Q439.05 -23.30 439.58 -27.33Q440.10 -31.35 441.40 -38.30L442.05 -42.00H442.80Q442.80 -37.45 444.10 -34.85Q445.40 -32.25 448.30 -31.12Q451.20 -30.00 456.30 -30.00Q467.70 -30.00 474.00 -31.18Q480.30 -32.35 480.30 -34.25Q480.30 -35.70 476.35 -38.68Q472.40 -41.65 466.15 -45.18Q459.90 -48.70 453.30 -51.75Q455.20 -55.30 457.85 -63.42Q460.50 -71.55 461.05 -75.50H461.80Q462.50 -74.65 465.00 -72.50Q469.50 -68.65 472.88 -64.73Q476.25 -60.80 478.90 -54.18Q481.55 -47.55 481.55 -38.50Q481.55 -23.60 479.05 -15.40Q476.55 -7.20 470.55 -3.60Q464.55 0.00 453.55 0.00Z M492.55 -34.50Q491.80 -50.00 490.20 -60.60Q488.60 -71.20 485.55 -81.50Q489.45 -84.85 494.65 -90.43Q499.85 -96.00 502.55 -100.00H503.30Q503.30 -98.10 503.70 -92.10Q504.30 -82.60 504.30 -77.50Q504.30 -71.25 504.05 -65.95Q503.80 -60.65 503.35 -53.90Q502.80 -45.70 502.80 -42.50Q502.80 -37.70 504.27 -35.00Q505.75 -32.30 509.10 -31.15Q512.45 -30.00 518.30 -30.00L515.80 0.00Q507.45 0.00 502.75 -2.85Q498.05 -5.70 495.65 -13.08Q493.25 -20.45 492.55 -34.50Z M603.55 -38.25Q603.55 -26.90 599.05 -18.23Q594.55 -9.55 586.45 -4.78Q578.35 0.00 567.80 0.00Q554.10 0.00 547.03 -5.33Q539.95 -10.65 538.55 -22.00H537.80Q532.65 0.00 515.80 0.00Q514.65 0.00 513.98 -2.65Q513.30 -5.30 513.30 -9.75Q513.30 -13.75 514.12 -18.55Q514.95 -23.35 516.15 -26.68Q517.35 -30.00 518.30 -30.00Q525.75 -30.00 530.43 -31.32Q535.10 -32.65 537.70 -35.55Q540.30 -38.45 541.30 -43.25H542.55Q542.85 -39.95 543.65 -37.78Q544.45 -35.60 546.15 -34.10Q547.85 -32.60 550.85 -31.70Q557.40 -41.90 562.95 -48.10Q568.50 -54.30 573.90 -57.28Q579.30 -60.25 585.30 -60.25Q590.95 -60.25 595.08 -57.45Q599.20 -54.65 601.38 -49.65Q603.55 -44.65 603.55 -38.25ZM570.05 -30.00Q581.20 -30.00 588.38 -30.50Q595.55 -31.00 598.80 -31.78Q602.05 -32.55 602.05 -33.50Q602.05 -34.55 599.70 -35.78Q597.35 -37.00 593.43 -37.88Q589.50 -38.75 585.05 -38.75Q576.00 -38.75 567.30 -36.90Q558.60 -35.05 552.25 -31.35Q557.85 -30.00 570.05 -30.00Z M691.80 -22.00Q691.80 25.00 654.30 25.00Q645.35 25.00 639.03 21.58Q632.70 18.15 629.38 11.68Q626.05 5.20 626.05 -3.75Q626.05 -12.40 627.95 -21.12Q629.85 -29.85 633.80 -39.00H635.05Q631.80 -29.25 631.80 -22.00Q631.80 -13.20 637.83 -8.85Q643.85 -4.50 656.05 -4.50Q667.50 -4.50 675.23 -6.33Q682.95 -8.15 686.75 -11.40Q690.55 -14.65 690.55 -19.00Q690.55 -20.55 689.48 -24.60Q688.40 -28.65 687.30 -31.00H686.55Q685.60 -20.40 680.60 -15.45Q675.60 -10.50 665.55 -10.50Q657.00 -10.50 652.78 -15.53Q648.55 -20.55 648.55 -30.25Q648.55 -41.45 651.25 -50.23Q653.95 -59.00 658.73 -63.88Q663.50 -68.75 669.55 -68.75Q680.80 -68.75 686.30 -57.12Q691.80 -45.50 691.80 -22.00ZM664.55 -47.25Q661.00 -47.25 657.83 -46.45Q654.65 -45.65 652.73 -44.53Q650.80 -43.40 650.80 -42.50Q650.80 -41.75 652.55 -41.05Q654.30 -40.35 657.50 -39.92Q660.70 -39.50 664.80 -39.50Q677.10 -39.50 684.55 -41.25Q681.65 -43.95 676.38 -45.60Q671.10 -47.25 664.55 -47.25ZM650.80 -82.35Q650.80 -83.55 651.18 -85.03Q651.55 -86.50 652.30 -88.90Q653.70 -93.40 654.05 -95.75H654.80Q654.80 -94.20 655.75 -93.35Q656.70 -92.50 658.90 -92.12Q661.10 -91.75 665.05 -91.75H665.30Q670.30 -91.75 677.75 -92.05Q685.20 -92.35 689.80 -92.75Q688.70 -90.70 686.88 -85.33Q685.05 -79.95 684.30 -76.50Q678.90 -76.10 672.40 -75.80Q665.90 -75.50 661.80 -75.50Q656.00 -75.50 653.40 -77.10Q650.80 -78.70 650.80 -82.35Z M742.30 -12.00Q742.30 2.40 738.28 12.48Q734.25 22.55 728.10 27.52Q721.95 32.50 715.55 32.50Q711.95 32.50 708.78 31.08Q705.60 29.65 703.12 27.60Q700.65 25.55 697.15 22.15Q694.05 19.10 691.93 17.33Q689.80 15.55 687.55 14.50V13.75Q701.20 11.25 713.60 6.97Q726.00 2.70 733.53 -1.73Q741.05 -6.15 741.05 -9.00Q741.05 -10.55 739.98 -14.60Q738.90 -18.65 737.80 -21.00H737.05Q736.10 -10.40 731.10 -5.45Q726.10 -0.50 716.05 -0.50Q707.55 -0.50 703.30 -5.75Q699.05 -11.00 699.05 -20.75Q699.05 -31.95 701.73 -40.58Q704.40 -49.20 709.18 -53.98Q713.95 -58.75 720.05 -58.75Q731.30 -58.75 736.80 -47.12Q742.30 -35.50 742.30 -12.00ZM715.05 -37.25Q711.50 -37.25 708.33 -36.45Q705.15 -35.65 703.23 -34.52Q701.30 -33.40 701.30 -32.50Q701.30 -31.75 703.05 -31.05Q704.80 -30.35 708.00 -29.93Q711.20 -29.50 715.30 -29.50Q727.60 -29.50 735.05 -31.25Q732.15 -33.95 726.88 -35.60Q721.60 -37.25 715.05 -37.25Z M786.55 -37.75H785.80Q785.25 -18.45 779.68 -9.22Q774.10 0.00 763.05 0.00Q755.90 0.00 752.23 -4.70Q748.55 -9.40 748.55 -18.75Q748.55 -23.30 749.08 -27.33Q749.60 -31.35 750.90 -38.30L751.55 -42.00H752.30Q752.30 -37.45 753.60 -34.85Q754.90 -32.25 757.80 -31.12Q760.70 -30.00 765.80 -30.00Q770.60 -30.00 774.80 -31.10Q779.00 -32.20 781.53 -34.05Q784.05 -35.90 784.05 -38.00Q784.05 -39.00 782.58 -43.20Q781.10 -47.40 779.15 -51.95Q777.20 -56.50 775.80 -58.75Q777.45 -60.40 780.23 -64.30Q783.00 -68.20 785.75 -72.75Q788.50 -77.30 790.05 -80.75H790.80Q792.10 -73.70 793.55 -66.90Q795.00 -60.10 798.55 -45.50Q800.50 -37.30 804.85 -33.65Q809.20 -30.00 815.55 -30.00L813.05 0.00Q804.50 0.00 799.48 -3.58Q794.45 -7.15 791.55 -15.15Q788.65 -23.15 786.55 -37.75Z M844.45 -19.55Q844.45 -22.60 844.80 -26.25H844.05Q842.80 -16.25 839.43 -10.53Q836.05 -4.80 830.28 -2.40Q824.50 0.00 815.55 0.00H813.05Q811.90 0.00 811.23 -2.65Q810.55 -5.30 810.55 -9.75Q810.55 -13.75 811.38 -18.55Q812.20 -23.35 813.40 -26.68Q814.60 -30.00 815.55 -30.00H818.55Q830.60 -30.00 836.25 -31.12Q841.90 -32.25 844.62 -35.62Q847.35 -39.00 849.05 -46.50H850.30V-39.75Q850.30 -30.00 865.05 -30.00L862.55 0.00Q856.10 0.00 852.12 -2.05Q848.15 -4.10 846.30 -8.38Q844.45 -12.65 844.45 -19.55ZM827.80 -62.25V-63.00L839.30 -77.00Q842.90 -75.30 847.08 -71.72Q851.25 -68.15 853.30 -65.00L841.05 -52.00Q838.90 -54.70 834.73 -57.90Q830.55 -61.10 827.80 -62.25Z M950.30 -38.25Q950.30 -26.90 945.80 -18.23Q941.30 -9.55 933.20 -4.78Q925.10 0.00 914.55 0.00Q900.85 0.00 893.78 -5.33Q886.70 -10.65 885.30 -22.00H884.55Q879.40 0.00 862.55 0.00Q861.40 0.00 860.73 -2.65Q860.05 -5.30 860.05 -9.75Q860.05 -13.75 860.88 -18.55Q861.70 -23.35 862.90 -26.68Q864.10 -30.00 865.05 -30.00Q872.50 -30.00 877.18 -31.32Q881.85 -32.65 884.45 -35.55Q887.05 -38.45 888.05 -43.25H889.30Q889.60 -39.95 890.40 -37.78Q891.20 -35.60 892.90 -34.10Q894.60 -32.60 897.60 -31.70Q904.15 -41.90 909.70 -48.10Q915.25 -54.30 920.65 -57.28Q926.05 -60.25 932.05 -60.25Q937.70 -60.25 941.83 -57.45Q945.95 -54.65 948.12 -49.65Q950.30 -44.65 950.30 -38.25ZM916.80 -30.00Q927.95 -30.00 935.12 -30.50Q942.30 -31.00 945.55 -31.78Q948.80 -32.55 948.80 -33.50Q948.80 -34.55 946.45 -35.78Q944.10 -37.00 940.18 -37.88Q936.25 -38.75 931.80 -38.75Q922.75 -38.75 914.05 -36.90Q905.35 -35.05 899.00 -31.35Q904.60 -30.00 916.80 -30.00Z" transform="translate(0.00,0)"/></g></svg>`;
const LOGO = `<svg viewBox="0 0 172 172" fill="currentColor" aria-hidden="true"><g transform="translate(-14,-14) scale(1)"><path d="M125.80,100.00L155.90,116.00L169.14,114.72L186.00,100.00L169.14,85.28L155.90,84.00ZM122.34,112.90L140.41,141.81L152.52,147.32L174.48,143.00L167.24,121.82L156.41,114.09ZM112.90,122.34L114.09,156.41L121.82,167.24L143.00,174.48L147.32,152.52L141.81,140.41ZM100.00,125.80L84.00,155.90L85.28,169.14L100.00,186.00L114.72,169.14L116.00,155.90ZM87.10,122.34L58.19,140.41L52.68,152.52L57.00,174.48L78.18,167.24L85.91,156.41ZM77.66,112.90L43.59,114.09L32.76,121.82L25.52,143.00L47.48,147.32L59.59,141.81ZM74.20,100.00L44.10,84.00L30.86,85.28L14.00,100.00L30.86,114.72L44.10,116.00ZM77.66,87.10L59.59,58.19L47.48,52.68L25.52,57.00L32.76,78.18L43.59,85.91ZM87.10,77.66L85.91,43.59L78.18,32.76L57.00,25.52L52.68,47.48L58.19,59.59ZM100.00,74.20L116.00,44.10L114.72,30.86L100.00,14.00L85.28,30.86L84.00,44.10ZM112.90,77.66L141.81,59.59L147.32,47.48L143.00,25.52L121.82,32.76L114.09,43.59ZM122.34,87.10L156.41,85.91L167.24,78.18L174.48,57.00L152.52,52.68L140.41,58.19Z"/><circle cx="100" cy="100" r="15.996"/></g></svg>`;

// آیکن‌های ساده (خطی)
const ICONS = {
  users: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20c.6-3.6 3.2-5.5 6.5-5.5s5.9 1.9 6.5 5.5"/><circle cx="17" cy="9" r="2.7"/><path d="M16.5 14.6c2.6.2 4.4 1.9 5 4.9"/></svg>',
  coins: '<svg viewBox="0 0 24 24"><ellipse cx="9" cy="6.5" rx="6" ry="2.8"/><path d="M3 6.5v4c0 1.5 2.7 2.8 6 2.8s6-1.3 6-2.8v-4"/><path d="M9 17.3c-3.3 0-6-1.3-6-2.8v-4"/><ellipse cx="15.5" cy="14" rx="5.5" ry="2.6"/><path d="M10 14v3.6c0 1.4 2.5 2.6 5.5 2.6s5.5-1.2 5.5-2.6V14"/></svg>',
  mail: '<svg viewBox="0 0 24 24"><rect x="2.5" y="4.5" width="19" height="15" rx="2.5"/><path d="M3.5 6.5 12 13l8.5-6.5"/></svg>',
  // نشان رسمی پیام‌رسان بله (فایل SVG خودشان، فقط دقت اعشار کم شده)
  bale: '<svg class="bico" viewBox="0 0 601.5 600.1" aria-hidden="true"><g transform="translate(59.0625,9.375)"><path d="M0 0 C10.2 6.6 20.1 13.6 29.9 20.6 C30.7 21.2 31.5 21.7 32.3 22.3 C34.6 23.9 36.9 25.6 39.1 27.2 C39.9 27.7 40.6 28.2 41.3 28.8 C45.2 31.6 49 34.4 52.7 37.4 C53.4 38 54.2 38.6 55 39.3 C56.5 40.4 57.9 41.6 59.4 42.8 C60 43.4 60.7 43.9 61.4 44.4 C61.9 44.9 62.5 45.4 63.1 45.9 C65 46.9 65 46.9 67.3 46 C70.1 44.5 72.5 42.9 75 40.9 C80.7 36.7 86.6 33.2 92.8 29.8 C94.4 28.8 94.4 28.8 96 27.9 C167.6 -11.2 250.8 -19.1 328.9 3.5 C382.1 19.3 434.8 52 469.7 95.8 C472 98.7 474.4 101.5 476.8 104.2 C487.7 117.4 496.7 131.7 504.9 146.6 C505.4 147.4 505.8 148.1 506.2 148.9 C544.2 216.5 551.2 299.1 530.8 373.3 C515.1 428.2 482.9 480.6 438.9 517.6 C438.4 518.1 437.9 518.5 437.3 519 C391.4 557.4 335.7 582.9 275.9 589.6 C274.7 589.8 274.7 589.8 273.5 589.9 C263.1 590.9 252.7 590.9 242.2 590.9 C240.7 590.9 240.7 590.9 239.2 590.9 C223.9 590.9 209 590.3 193.9 587.6 C192.9 587.5 191.9 587.3 190.9 587.1 C133.2 577.1 81.1 551 37.9 511.6 C37 510.8 36 509.9 35 509 C8.2 485 -13.8 454.3 -29.1 421.7 C-29.8 420.3 -30.4 418.8 -31.1 417.4 C-44.2 390.5 -52 361.8 -56.4 332.2 C-56.5 331.2 -56.7 330.1 -56.9 329 C-59.5 309.8 -59.2 290.4 -59.2 271.1 C-59.2 268.4 -59.2 265.7 -59.2 263 C-59.3 255.7 -59.3 248.4 -59.3 241.1 C-59.3 236.5 -59.3 231.9 -59.3 227.3 C-59.3 214.6 -59.3 202 -59.3 189.3 C-59.3 188.4 -59.3 187.6 -59.3 186.8 C-59.3 186 -59.3 185.2 -59.3 184.4 C-59.3 182.7 -59.3 181.1 -59.3 179.4 C-59.3 178.6 -59.3 177.8 -59.3 177 C-59.3 163.7 -59.3 150.5 -59.3 137.3 C-59.4 123.7 -59.4 110.1 -59.4 96.5 C-59.4 88.8 -59.4 81.2 -59.4 73.6 C-59.4 67.1 -59.4 60.6 -59.4 54.1 C-59.4 50.8 -59.4 47.5 -59.4 44.2 C-59.4 40.6 -59.4 37 -59.4 33.4 C-59.4 32.4 -59.4 31.4 -59.5 30.3 C-59.4 19.5 -57.7 9.7 -50.1 1.6 C-49.5 1 -48.9 0.4 -48.4 -0.3 C-34 -14.6 -15.2 -9.6 0 0 Z M318.9 184.6 C317.6 185.8 316.3 187.1 315.1 188.4 C312.3 191.1 309.6 193.9 306.9 196.6 C302.5 201 298.2 205.3 293.9 209.6 C284.7 218.8 275.6 227.9 266.4 237.1 C256.5 247.1 246.6 257 236.7 266.9 C232.4 271.2 228.1 275.5 223.8 279.8 C221.1 282.5 218.5 285.2 215.8 287.8 C214.6 289.1 213.3 290.3 212.1 291.6 C210.4 293.3 208.7 294.9 207 296.6 C206.6 297.1 206.1 297.6 205.6 298.1 C202.2 301.5 202.2 301.5 199.9 302.6 C199.7 302.1 199.5 301.5 199.2 300.9 C197.6 298.1 195.7 296.2 193.5 293.9 C192.1 292.5 192.1 292.5 190.6 291.1 C189.6 290 188.6 289 187.6 288 C186.5 287 185.5 285.9 184.4 284.9 C182.2 282.7 180 280.5 177.8 278.3 C175 275.5 172.2 272.7 169.4 269.9 C167.3 267.7 165.1 265.5 162.9 263.3 C161.9 262.3 160.9 261.3 159.8 260.2 C144.9 245.3 129.7 233.8 107.8 233.2 C91.1 233.9 76.9 239.9 64.9 251.6 C52.3 266.2 48.9 281.4 49.7 300.1 C50.2 305.7 51.7 310.6 53.9 315.6 C54.3 316.5 54.7 317.5 55.1 318.4 C62.9 334.5 79.7 347.5 92.1 360 C93.1 361 94 362 95 363 C99.5 367.5 104.1 372 108.6 376.6 C113.8 381.8 119 387 124.1 392.2 C128.2 396.3 132.2 400.3 136.3 404.4 C138.7 406.8 141.1 409.2 143.4 411.6 C160.2 429.3 160.2 429.3 181.4 440.6 C182.6 441 183.7 441.4 184.8 441.8 C200 446 215.6 444 229.6 437.1 C242.7 429.1 253.4 416.3 264.2 405.6 C265.9 403.9 267.6 402.2 269.3 400.5 C274.4 395.4 279.4 390.4 284.5 385.3 C286.4 383.4 288.3 381.5 290.2 379.6 C298.7 371.1 307.2 362.7 315.6 354.2 C317.7 352.2 319.7 350.1 321.8 348.1 C322.3 347.5 322.8 347 323.3 346.5 C331.7 338.2 340 329.9 348.3 321.6 C356.8 313.1 365.4 304.6 374 296 C378.7 291.2 383.5 286.5 388.3 281.7 C392.9 277.2 397.4 272.7 401.9 268.2 C403.5 266.5 405.1 264.9 406.8 263.3 C421.5 248.7 421.5 248.7 430.8 230.4 C431.2 229.3 431.6 228.1 432 227 C436.2 212 434 195.9 426.6 182.3 C417.2 167 404.3 157.9 386.9 153.6 C357.1 147.9 338.3 165.1 318.9 184.6 Z"/></g></svg>',
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
    g += `<path class="garea" d="M${X(0)},${Y(0)}L${pts.join("L")}L${X(dmax)},${Y(0)}Z" fill="url(#gFill)"/>`;
    g += `<path class="gline" pathLength="1" d="M${pts.join("L")}" fill="none" stroke="#fff" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    g += `<line id="gGuide" y1="${mt}" y2="${mt + ph}" stroke="#fff" stroke-opacity=".7" stroke-width="1" stroke-dasharray="3 3"/>`;
    const rr = pw / Math.max(1, G.snaps.length - 1) < 14 ? 2.5 : 3.5; // در صفحه‌ی باریک، نقطه‌ها کوچک‌تر
    const ls = G.snaps[G.snaps.length - 1];
    g += `<circle class="gpulse" cx="${X(ls.d)}" cy="${Y(ls.v)}" r="5" fill="#fff"/>`; // نبض «امروز»
    G.snaps.forEach((s, i) => { g += `<circle class="gdot" data-i="${i}" cx="${X(s.d)}" cy="${Y(s.v)}" r="${rr}" fill="#fff" stroke="#0049E8" stroke-width="1.5"/>`; });
    g += `<rect x="${ml}" y="0" width="${pw}" height="${H}" fill="transparent"/>`;
    box.innerHTML = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" tabindex="0" role="img" aria-label="نمودار رشد دارایی کل صندوق از ${G.start}. با کلیدهای چپ و راست بین تاریخ‌ها جابه‌جا شوید.">${g}</svg><div class="gtip" hidden></div>`;
    const svg = box.firstChild, tip = box.lastChild;
    // حباب کنار نقطه‌ی انتخاب‌شده؛ فقط وقتی انگشت یا نشانگر روی نمودار است
    const bubble = (s) => {
      tip.innerHTML = `${s.label}<b>${money(s.v, true)}</b>`;
      tip.hidden = false;
      const half = tip.offsetWidth / 2, x = Math.min(Math.max(X(s.d), half), W - half), y = Y(s.v);
      tip.classList.toggle("below", y - tip.offsetHeight - 16 < 0);
      tip.style.left = x + "px"; tip.style.top = y + "px";
    };
    const set = (i, show) => {
      growth.sel = i;
      const s = G.snaps[i];
      svg.querySelectorAll("circle.gdot").forEach((c) => c.setAttribute("r", +c.dataset.i === i ? 6.5 : rr));
      const gl = svg.querySelector("#gGuide");
      gl.setAttribute("x1", X(s.d)); gl.setAttribute("x2", X(s.d));
      if (show) bubble(s); // مقدار هر ماه در حباب؛ زیر عدد دارایی برچسب ثابت رشد ۱۲ ماه می‌ماند
    };
    const pick = (e) => {
      const r = svg.getBoundingClientRect();
      const x = e.clientX - r.left;
      let best = 0, bd = Infinity;
      G.snaps.forEach((s, i) => { const dd = Math.abs(X(s.d) - x); if (dd < bd) { bd = dd; best = i; } });
      set(best, true);
    };
    svg.addEventListener("pointermove", pick);
    svg.addEventListener("pointerdown", pick);
    svg.addEventListener("pointerleave", () => { tip.hidden = true; });
    svg.addEventListener("blur", () => { tip.hidden = true; });
    svg.addEventListener("keydown", (e) => {
      if (e.key === "ArrowLeft") { set(Math.max(0, growth.sel - 1), true); e.preventDefault(); }
      if (e.key === "ArrowRight") { set(Math.min(G.snaps.length - 1, growth.sel + 1), true); e.preventDefault(); }
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
      return `<button type="button" data-i="${i}" style="--i:${i}" class="${i === st.sel ? "sel" : ""}" aria-label=""${monthName(m)}: ${money(totals[i], true)}">${h}</button>`;
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
    // راهنمای شناور: با بردن نشانگر یا زدن روی هر رنگ، مبلغ همان بخش نشان داده می‌شود
    const plot = root.querySelector(".plot"), tip = root.querySelector(".tip");
    // ستونی که روی آن هستید پررنگ می‌ماند و بقیه کم‌رنگ می‌شوند
    const hover = (b) => { plot.classList.add("dim"); root.querySelectorAll(".cols button").forEach((x) => x.classList.toggle("hov", x === b)); };
    root.querySelectorAll(".cols button").forEach((b) => {
      b.addEventListener("click", () => set(+b.dataset.i));
      b.addEventListener("pointerenter", () => { set(+b.dataset.i); hover(b); });
      b.addEventListener("pointerdown", () => hover(b));
      b.addEventListener("focus", () => set(+b.dataset.i));
    });
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
    const hideTip = () => { tip.hidden = true; plot.classList.remove("dim"); root.querySelectorAll(".cols i.hot").forEach((x) => x.classList.remove("hot")); };
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
  // انیمیشن ورود: هر بخش وقتی دیده شد یک بار؛ عددها از صفر می‌شمارند
  const anim = document.documentElement.classList.contains("anim");
  const nums = [...document.querySelectorAll(".num[data-to]")];
  if (anim) nums.forEach((el) => { el.dataset.txt = el.textContent; el.textContent = (0).toLocaleString("fa-IR", { minimumFractionDigits: +el.dataset.dec, maximumFractionDigits: +el.dataset.dec }); });
  const count = (el) => {
    const to = +el.dataset.to, dec = +el.dataset.dec, t0 = performance.now(), dur = 1500;
    const tick = (t) => {
      const k = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - k, 4); // پایان نرم
      el.textContent = k < 1 ? (to * e).toLocaleString("fa-IR", { minimumFractionDigits: dec, maximumFractionDigits: dec }) : el.dataset.txt;
      if (k < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  };
  const reveal = (sec) => {
    if (sec.classList.contains("in")) return;
    sec.classList.add("in");
    if (anim) sec.querySelectorAll(".num[data-to]").forEach(count);
    setTimeout(() => sec.classList.add("done"), 3200); // بعد از این، تغییر بازه یا اندازه‌ی صفحه دیگر انیمیشن ندارد
  };
  const secs = document.querySelectorAll("[data-anim]");
  if (anim && "IntersectionObserver" in window) {
    const io = new IntersectionObserver((es) => es.forEach((e) => { if (e.isIntersecting) { reveal(e.target); io.unobserve(e.target); } }), { threshold: 0.25 });
    secs.forEach((s) => io.observe(s));
  } else secs.forEach(reveal);
  const all = () => { growth(); bars("loans"); bars("coll"); };
  let rt;
  window.addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(growth, 120); });
  // یک بار رسم، بعد از آماده شدن فونت (تا اندازه‌ها درست باشد و انیمیشن دو بار شروع نشود)
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(all, all); else all();
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
    `<circle class="rtrack" cx="${c}" cy="${c}" r="${rad}" fill="none" stroke="#DCE8FE" stroke-width="${stroke}"/>` +
    (p > 0 ? `<circle class="rp" cx="${c}" cy="${c}" r="${rad}" fill="none" stroke="${THEME.primary}" stroke-width="${stroke}" stroke-linecap="round" stroke-dasharray="${(len * p).toFixed(2)} ${len.toFixed(2)}" transform="rotate(-90 ${c} ${c})"/>` : "") +
    `<text x="50%" y="${sub ? "47%" : "54%"}" text-anchor="middle" dominant-baseline="middle" font-weight="800" font-size="${Math.round(size / 5)}" fill="${THEME.text}" class="rl">${label}</text>` +
    (sub ? `<text x="50%" y="68%" text-anchor="middle" dominant-baseline="middle" font-size="11" fill="${THEME.muted}" class="rs">${sub}</text>` : "") +
    `</svg>`;
}

// عدد فارسی داخل span تا در انیمیشن ورود از صفر شمرده شود (متن اصلی دست‌نخورده می‌ماند)
function numSpan(text) {
  const t = String(text), m = t.replace(/[۰-۹]/g, (d) => "۰۱۲۳۴۵۶۷۸۹".indexOf(d)).replace(/٬/g, "");
  if (!/^\d+(٫\d+)?$/.test(m)) return esc(t);
  const dec = m.includes("٫") ? m.split("٫")[1].length : 0;
  return `<span class="num" data-to="${m.replace("٫", ".")}" data-dec="${dec}">${esc(t)}</span>`;
}

function renderReport(r, fonts) {
  const title = `داشبورد ${r.fundName}`;
  const upd = fmtDate(r.asOf);
  const moneyParts = (n) => { const s = fmtMoney(n); const i = s.lastIndexOf(" "); return i > 0 ? `${numSpan(s.slice(0, i))} <small>${s.slice(i + 1)} تومان</small>` : `${numSpan(s)} <small>تومان</small>`; };
  const intN = (n) => numSpan(fmtInt(n));
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
    `<div class="card stat s-mem" data-anim>${head("users", "i-blue", "اعضای فعال")}<div class="v">${intN(r.memberCount)} <small>عضو</small></div></div>`,
    `<div class="card stat s-loan" data-anim>${head("coins", "i-teal", "ارزش وام‌ها از ابتدای فعالیت صندوق")}<div class="v">${moneyParts(r.loansTotalAmount)}</div></div>`,
  ];
  if (r.wait) {
    const w = r.wait;
    stats.push(`<div class="card stat s-wait" data-anim>${head("clock", "i-violet", "زمان انتظار برای وام")}
      <div class="duo">
        <div><div class="v">${w.median == null ? "—" : w.median <= 60 ? `${intN(w.median)} <small>روز</small>` : `${numSpan((w.median / 30.44).toLocaleString("fa-IR", { maximumFractionDigits: 1 }))} <small>ماه</small>`}</div>
          <div class="c">${w.n ? `نیمی از ${fmtInt(w.n)} وام ${fmtInt(CONFIG.waitMonths)} ماه اخیر در همین مدت یا کمتر پرداخت شده‌اند` : `وامی با درخواست ثبت‌شده در ${fmtInt(CONFIG.waitMonths)} ماه اخیر نبود`}</div></div>
        <div class="sep"></div>
        <div><div class="v">${intN(w.queue)} <small>درخواست</small></div><div class="kk">در صف</div><div class="c">هنوز وامشان پرداخت نشده</div></div>
      </div></div>`);
  }

  const mo = r.month;
  const monthRange = mo.start === r.asOf ? fmtDate(r.asOf) : `${fmtNum(J.d2j(mo.start).jd)} تا ${fmtDate(r.asOf)}`;
  const collectPct = mo.due ? (mo.paid / mo.due) * 100 : 0;
  const cm = r.months[r.months.length - 1] || {};
  const curColl = (cm.repay || 0) + (cm.fee || 0) + (cm.deposit || 0) + (cm.other || 0);

  // برچسب رشد ۱۲ ماه اخیر زیر عدد دارایی
  const growthBadge = (() => {
    const G = r.growth, last = G.snaps[G.snaps.length - 1];
    if (!last || !G.series.length) return "";
    const g0 = r.asOf - last.d, d = J.d2j(r.asOf), target = J.addMonths(d.jy, d.jm, d.jd, -12) - g0;
    if (target < 0) return "";
    let base = null;
    for (const p of G.series) { if (p[0] <= target) base = p[1]; else break; }
    if (!base) return "";
    const pct = Math.round((last.v / base - 1) * 100);
    return `<span class="gbadge">${pct >= 0 ? "▲" : "▼"} <b>${fmtInt(Math.abs(pct))}٪</b> ${pct >= 0 ? "رشد" : "کاهش"} در ۱۲ ماه اخیر</span>`;
  })();

  // نشانی سایت برای تصویر پیش‌نمایش لینک و آیکن گوشی (در خودکارساز از SITE_URL می‌آید)
  const SITE = (typeof SITE_URL === "string" ? SITE_URL : (typeof process !== "undefined" && process.env && process.env.SITE_URL) || "").replace(/\/$/, "");

  // پیش‌نمایش لینک (تلگرام و …): عنوان + یک خط خلاصه
  const ogDesc = `دارایی کل: ${fmtMoney(r.capital, true)}، به‌روز تا ${upd}`;

  return `<!doctype html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<meta name="theme-color" content="${THEME.primary}">
<meta name="color-scheme" content="light dark">
<link rel="icon" href="data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="${'0 0 172 172'}" fill="#0A1A4F">${'<g transform="translate(-14,-14) scale(1)"><path d="M125.80,100.00L155.90,116.00L169.14,114.72L186.00,100.00L169.14,85.28L155.90,84.00ZM122.34,112.90L140.41,141.81L152.52,147.32L174.48,143.00L167.24,121.82L156.41,114.09ZM112.90,122.34L114.09,156.41L121.82,167.24L143.00,174.48L147.32,152.52L141.81,140.41ZM100.00,125.80L84.00,155.90L85.28,169.14L100.00,186.00L114.72,169.14L116.00,155.90ZM87.10,122.34L58.19,140.41L52.68,152.52L57.00,174.48L78.18,167.24L85.91,156.41ZM77.66,112.90L43.59,114.09L32.76,121.82L25.52,143.00L47.48,147.32L59.59,141.81ZM74.20,100.00L44.10,84.00L30.86,85.28L14.00,100.00L30.86,114.72L44.10,116.00ZM77.66,87.10L59.59,58.19L47.48,52.68L25.52,57.00L32.76,78.18L43.59,85.91ZM87.10,77.66L85.91,43.59L78.18,32.76L57.00,25.52L52.68,47.48L58.19,59.59ZM100.00,74.20L116.00,44.10L114.72,30.86L100.00,14.00L85.28,30.86L84.00,44.10ZM112.90,77.66L141.81,59.59L147.32,47.48L143.00,25.52L121.82,32.76L114.09,43.59ZM122.34,87.10L156.41,85.91L167.24,78.18L174.48,57.00L152.52,52.68L140.41,58.19Z"/><circle cx="100" cy="100" r="15.996"/></g>'}</svg>`)}">
${SITE ? `<link rel="apple-touch-icon" href="${SITE}/icon-180.png">
<meta property="og:image" content="${SITE}/og.png">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta name="twitter:image" content="${SITE}/og.png">` : ""}
<title>${esc(title)}</title>
<meta name="description" content="${esc(ogDesc)}">
<meta property="og:type" content="website">
<meta property="og:locale" content="fa_IR">
<meta property="og:site_name" content="${esc(r.fundName)}">
<meta property="og:title" content="${esc(title)}">
<meta property="og:description" content="${esc(ogDesc)}">
<meta name="twitter:card" content="${SITE ? "summary_large_image" : "summary"}">
<script>if(!(window.matchMedia&&matchMedia("(prefers-reduced-motion: reduce)").matches))document.documentElement.classList.add("anim")<\/script>
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
    <div class="logo">${LOGO}</div>
    ${norm(r.fundName) === norm(WORDMARK_NAME)
      ? `<h1 class="wm">${WORDMARK}<span class="sr">${esc(r.fundName)}</span></h1>`
      : `<h1>${esc(r.fundName)}</h1>`}
  </div>
  <div class="pill">تاریخ گزارش: <b>${upd}</b></div>
</header>

<main class="grid-main">
  <div class="grp">نمای کلی</div>
  <section class="card hero span12" data-anim>
    <div class="label">دارایی کل صندوق</div>
    <div class="big">${moneyParts(r.capital)}</div>
    <div class="readout" id="growthRead">${growthBadge}</div>
    <div class="growth" id="growth"></div>
    <p class="hint">رشد دارایی کل از ${r.growth.start}. روی نمودار بزنید یا نشانگر را ببرید تا دارایی صندوق در اول هر ماه را ببینید.</p>
  </section>

  <div class="stats span12${r.wait ? " has-wait" : ""}">${stats.join("")}</div>

  <div class="grp">وام‌ها</div>
  <section class="card span12 lip" data-anim>
    <div class="lip-a">
      <h2>وام‌های در جریان</h2>
      <div class="big2">${intN(a.n)} <small>وام</small></div>
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
      <div class="sub"><span class="sub-d">هر دایره یک وام در جریان است</span><span class="sub-m">سهم وام‌های در جریان در هر وضعیت</span></div>
      <div class="sbar" aria-hidden="true">${["ok", "ahead", "late1", "late2"].filter((k) => s[k] > 0).map((k, i) => `<i class="s-${k}" style="flex:${s[k]};--i:${i}"></i>`).join("")}</div>
      <div class="dots" aria-hidden="true">${r.dots.map((d, i) => `<span class="dot s-${d}" style="--i:${i}"></span>`).join("")}</div>
      <ul class="key">
        <li><span class="dot s-ok"></span><b>${fmtInt(s.ok)}</b> وام طبق برنامه</li>
        <li><span class="dot s-ahead"></span><b>${fmtInt(s.ahead)}</b> وام جلوتر از برنامه</li>
        <li><span class="dot s-late1"></span><b>${fmtInt(s.late1)}</b> وام یک قسط معوق</li>
        <li><span class="dot s-late2"></span><b>${fmtInt(s.late2)}</b> وام دو قسط یا بیشتر معوق</li>
      </ul>
      <p class="note">اقساط ماهانه فرض شده‌اند و اولین سررسید هر وام یک ماه پس از دریافت آن است.</p>
    </div>
  </section>

  <section class="card span12 chartcard" data-kind="loans" id="ch-loans" data-anim>
    <div class="cc-main">
      <div class="chead"><div><h2>وام پرداخت‌شده</h2><div class="sub">مبلغ وام‌هایی که صندوق در هر ماه به اعضا داده است</div></div>${ranges()}</div>
      <div class="mstrip"><div class="mv">${intN(mo.loanN)} <small>وام</small></div><div><div><span class="mh">این ماه</span> · ${mo.loanN ? `به ارزش <b>${fmtMoney(mo.loanAmt, true)}</b>` : "هنوز وامی پرداخت نشده"}</div>${mo.avg12 ? `<div>میانگین ماهانه‌ی ${fmtInt(CONFIG.waitMonths)} ماه اخیر: <b>${fmtMoney(mo.avg12, true)}</b></div>` : ""}</div></div>
      <div class="legend"></div>
      <div class="chart"></div>
      <div class="slegend"></div>
      <div class="readout" aria-live="polite"></div>
    </div>
    <aside class="side">
      <div class="side-h">این ماه <span>${esc(monthRange)}</span></div>
      <div class="v">${intN(mo.loanN)} <small>وام</small></div>
      <div class="k">${mo.loanN ? `به ارزش <b>${fmtMoney(mo.loanAmt, true)}</b>` : "هنوز وامی پرداخت نشده"}</div>
      ${mo.list.length ? `<div class="chips" aria-label="وام‌های این ماه">${mo.list.slice(0, 4).map((l) => `<span><b>${fmtMoney(l.amount)}</b>${fmtInt(l.count)} قسط</span>`).join("")}${mo.list.length > 4 ? `<span class="more">و ${fmtInt(mo.list.length - 4)} وام دیگر</span>` : ""}</div>` : ""}
      ${mo.avg12 ? `<div class="cmp">میانگین ماهانه‌ی ${fmtInt(CONFIG.waitMonths)} ماه اخیر: <b>${fmtMoney(mo.avg12, true)}</b></div>` : ""}
    </aside>
  </section>

  <div class="grp">ورودی صندوق</div>
  <section class="card span12 chartcard" data-kind="coll" id="ch-coll" data-anim>
    <div class="cc-main">
      <div class="chead"><div><h2>مجموع وصولی</h2><div class="sub">هر چه در هر ماه وارد صندوق شده، به تفکیک نوع واریز</div></div>${ranges()}</div>
      <div class="mstrip">${mo.due ? ringSVG(collectPct / 100, `${fmtInt(Math.round(collectPct))}٪`, "", 64, 7) : ""}<div><div><span class="mh">این ماه</span> · وصول اقساط</div>${mo.due ? `<div class="mv">${intN(mo.paid)} <small>قسط از ${fmtInt(mo.due)} قسط</small></div>` : `<div>هنوز قسطی سررسید نشده</div>`}<div>وصولی این ماه تا امروز: <b>${fmtMoney(curColl, true)}</b></div></div></div>
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
      <div class="v">${intN(mo.paid)} <small>قسط از ${fmtInt(mo.due)} قسط</small></div>
      <div class="k">قسطی که تا امروز سررسید شده، پرداخت شده</div>`
        : `<div class="v">—</div><div class="k">هنوز قسطی در این ماه سررسید نشده</div>`}
      <div class="cmp">وصولی این ماه تا امروز: <b>${fmtMoney(curColl, true)}</b></div>
    </aside>
  </section>
</main>

<nav class="contact" aria-label="ارتباط با صندوق">
  <a href="${CONTACT.bale.url}" target="_blank" rel="noopener noreferrer">${ICONS.bale}<span>کانال بله: <b>${esc(CONTACT.bale.handle)}</b></span></a>
  <a href="mailto:${CONTACT.email}">${ICONS.mail}<span>ایمیل: <b>${esc(CONTACT.email)}</b></span></a>
</nav>

<footer>این داشبورد از خروجی نرم‌افزار صندوق در ${upd} ساخته شده و فقط ارقام کلی را نشان می‌دهد؛ اطلاعات هیچ عضوی در آن نیست.</footer>
</div>
<script>
window.__D = ${JSON.stringify(D)};
(${dashRuntime.toString()})();
<\/script>
</body>
</html>`;
}
// ============================================================================
// قالب ایمیل خبرنامه‌ی صندوق برای اعضا. خروجی HTML ساده و جدول‌محور است تا در
// Gmail و بقیه‌ی ایمیل‌خوان‌ها درست دیده شود. هیچ نام یا اطلاعات فردی در آن نمی‌آید.
// ============================================================================
const escE = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const EMAIL = {
  navy: "#0A1A4F", blue: "#0049E8", ink: "#1A2333", muted: "#5B6B85",
  line: "#E3E8F1", soft: "#F4F6FB", card: "#FFFFFF",
  font: "Tahoma, 'Segoe UI', Arial, sans-serif",
};

function renderEmail(r, o) {
  o = o || {};
  const E = EMAIL;
  const site = String(o.siteUrl || "").trim();
  const esc2 = (s) => escE(String(s == null ? "" : s));
  const rtl = "direction:rtl;text-align:right;"; // برای Gmail: هنگام چسباندن، dir روی <html> دور ریخته می‌شود
  const money = (n) => fmtMoney(n, true);
  const upd = fmtDate(r.asOf);
  const mo = r.month;
  const cm = r.months[r.months.length - 1] || {};
  const curColl = (cm.repay || 0) + (cm.fee || 0) + (cm.deposit || 0) + (cm.other || 0);
  // نام ماه از خود داده می‌آید؛ ایمیل اولِ مهر، خلاصه‌ی شهریور را نشان می‌دهد و باید همان را بنویسد
  const moName = MONTHS[J.d2j(mo.start).jm - 1];

  // کارت‌ها با inline-block چیده می‌شوند، نه با ستون جدول: در عرض کم خودشان زیر هم می‌آیند
  // بدون اینکه به media query نیاز باشد (Gmail هنگام چسباندن، بخش <style> را دور می‌ریزد).
  const cell = (label, value, sub) => `<div style="display:inline-block;vertical-align:top;width:50%;min-width:272px;max-width:100%;padding:8px;box-sizing:border-box;font-size:15px;${rtl}">`
      + `<table role="presentation" dir="rtl" width="100%" cellpadding="0" cellspacing="0" style="background:${E.soft};border-radius:12px;${rtl}">`
      + `<tr><td dir="rtl" style="padding:14px 16px;font-family:${E.font};${rtl}">`
      + `<div style="font-size:13px;color:${E.muted};line-height:1.7;">${esc2(label)}</div>`
      + `<div style="font-size:21px;font-weight:bold;color:${E.ink};line-height:1.6;">${esc2(value)}</div>`
      + (sub ? `<div style="font-size:12px;color:${E.muted};line-height:1.6;">${esc2(sub)}</div>` : "")
      + `</td></tr></table></div>`;

  const btn = (text, href, primary) => `
    <table role="presentation" dir="rtl" cellpadding="0" cellspacing="0" align="center" style="margin:6px auto;direction:rtl;">
      <tr><td align="center" bgcolor="${primary ? E.blue : "#FFFFFF"}" style="border-radius:10px;${primary ? "" : `border:1px solid ${E.line};`}">
        <a href="${esc2(href)}" style="display:inline-block;padding:${primary ? "13px 34px" : "11px 26px"};font-family:${E.font};font-size:${primary ? "16px" : "14px"};font-weight:bold;color:${primary ? "#FFFFFF" : E.blue};text-decoration:none;border-radius:10px;">${esc2(text)}</a>
      </td></tr>
    </table>`;

  const rows = [];
  rows.push(`<tr><td dir="rtl" style="padding:22px 24px 6px;font-family:${E.font};font-size:15px;line-height:2;color:${E.ink};${rtl}">
      ${o.note ? esc2(o.note).replace(/\n/g, "<br>") : `سلام،<br>گزارش تازه‌ی صندوق تا ${esc2(upd)} آماده است.`}
    </td></tr>`);
  rows.push(`<tr><td style="padding:6px 16px;">
      <div dir="rtl" style="font-size:0;direction:rtl;text-align:center;">${
        cell("دارایی کل صندوق", money(r.capital), `${fmtInt(r.memberCount)} عضو فعال`)
      }${cell("وام‌های در جریان", `${fmtInt(r.active.n)} وام`, `مانده ${money(r.active.remaining)}`)
      }${cell(`وام‌های ${moName}`, mo.loanN ? `${fmtInt(mo.loanN)} وام` : "—", mo.loanN ? `به ارزش ${money(mo.loanAmt)}` : `در ${moName} وامی پرداخت نشده`)
      }${cell(`وصولی ${moName}`, curColl ? money(curColl) : "—", mo.due ? `${fmtInt(mo.paid)} از ${fmtInt(mo.due)} قسط سررسیدشده پرداخت شده` : `در ${moName} قسطی سررسید نشده`)}</div>
    </td></tr>`);
  if (site) rows.push(`<tr><td style="padding:14px 24px 4px;">${btn("مشاهده داشبورد صندوق", site, true)}</td></tr>
    <tr><td dir="rtl" style="padding:0 24px 8px;font-family:${E.font};font-size:12px;color:${E.muted};line-height:1.8;direction:rtl;text-align:center;">
      داشبورد با هر گزارش تازه‌ی صندوق خودکار به‌روز می‌شود.</td></tr>`);

  return `<!doctype html>
<html lang="fa" dir="rtl">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc2(r.fundName)}</title></head>
<body style="margin:0;padding:0;background:${E.soft};">
<div dir="rtl" style="${rtl}">
<div style="display:none;max-height:0;overflow:hidden;opacity:0;">دارایی کل ${money(r.capital)} · ${fmtInt(r.active.n)} وام در جریان · به‌روز تا ${esc2(upd)}</div>
<table role="presentation" dir="rtl" width="100%" cellpadding="0" cellspacing="0" style="background:${E.soft};padding:18px 10px;${rtl}">
  <tr><td align="center">
    <table role="presentation" dir="rtl" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:${E.card};border-radius:16px;overflow:hidden;${rtl}">
      <tr><td dir="rtl" style="background:${E.navy};padding:26px 24px 22px;font-family:${E.font};direction:rtl;text-align:center;" align="center">
        ${site
          ? `<img src="${esc2(site)}/email-logo.png" width="260" height="54" alt="${esc2(r.fundName)}" style="display:block;margin:0 auto 8px;border:0;outline:none;text-decoration:none;width:260px;max-width:72%;height:auto;">`
          : `<div style="font-size:19px;font-weight:bold;color:#FFFFFF;line-height:1.6;">${esc2(r.fundName)}</div>`}
        <div style="font-size:13px;color:#B9C6E8;line-height:1.7;">گزارش صندوق · به‌روز تا ${esc2(upd)}</div>
      </td></tr>
      ${rows.join("\n")}
      <tr><td style="padding:12px 16px 4px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${E.soft};border-radius:12px;">
          <tr><td dir="rtl" align="center" style="padding:14px 16px;font-family:${E.font};font-size:13px;color:${E.muted};line-height:2.1;direction:rtl;text-align:center;">
            کانال بله: <span dir="ltr"><a href="${esc2(CONTACT.bale.url)}" style="color:${E.navy};font-weight:bold;text-decoration:none;">${esc2(CONTACT.bale.handle)}</a></span><br>
            ایمیل: <span dir="ltr"><a href="mailto:${esc2(CONTACT.email)}" style="color:${E.navy};font-weight:bold;text-decoration:none;">${esc2(CONTACT.email)}</a></span>
          </td></tr>
        </table>
      </td></tr>
      <tr><td dir="rtl" style="padding:16px 24px 22px;border-top:1px solid ${E.line};font-family:${E.font};font-size:12px;color:${E.muted};line-height:1.9;${rtl}" align="right">
        این ایمیل فقط ارقام کلی صندوق را دارد و هیچ اطلاعاتی درباره‌ی اعضا در آن نیست.<br>
        اگر نمی‌خواهید این ایمیل را دریافت کنید، به مدیر صندوق خبر بدهید.
      </td></tr>
    </table>
  </td></tr>
</table>
</div>
</body>
</html>`;
}

// نسخه‌ی متنی ساده (برای ایمیل‌خوان‌هایی که HTML نشان نمی‌دهند یا کپی سریع در پیام‌رسان)
function renderEmailText(r, o) {
  o = o || {};
  const mo = r.month, upd = fmtDate(r.asOf);
  const cm = r.months[r.months.length - 1] || {};
  const curColl = (cm.repay || 0) + (cm.fee || 0) + (cm.deposit || 0) + (cm.other || 0);
  const moName = MONTHS[J.d2j(mo.start).jm - 1];
  return [
    `${r.fundName} — گزارش تا ${upd}`,
    o.note || "",
    `دارایی کل: ${fmtMoney(r.capital, true)} (${fmtInt(r.memberCount)} عضو فعال)`,
    `وام‌های در جریان: ${fmtInt(r.active.n)} وام، مانده ${fmtMoney(r.active.remaining, true)}`,
    mo.loanN ? `وام‌های ${moName}: ${fmtInt(mo.loanN)} وام به ارزش ${fmtMoney(mo.loanAmt, true)}` : `در ${moName} وامی پرداخت نشده است.`,
    curColl ? `وصولی ${moName}: ${fmtMoney(curColl, true)}` + (mo.due ? `؛ ${fmtInt(mo.paid)} از ${fmtInt(mo.due)} قسط سررسیدشده پرداخت شده` : "") : "",
    o.siteUrl ? `داشبورد: ${o.siteUrl}` : "",
    `کانال بله: ${CONTACT.bale.handle} (${CONTACT.bale.url})`,
    `ایمیل: ${CONTACT.email}`,
  ].filter(Boolean).join("\n");
}

module.exports = { CONFIG, THEME, CONTACT, parseWorkbook, parseRequests, detectKind, compute, renderReport, renderEmail, renderEmailText, loadNameFixes, fmtDate, fmtNum, fmtInt, fmtMoney, ReportError };
