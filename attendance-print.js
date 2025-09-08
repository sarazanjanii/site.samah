/* =========================================================================
   attendance-print.js
   ======================================================================= */

/* ---------- Notify (اختیاری) ---------- */
function _notify(msg, type = 'info') {
  if (typeof window.notifySafe === 'function') return notifySafe(msg, type);
  const map = { error: 'error', warning: 'warn', info: 'log', success: 'log' };
  console[map[type] || 'log']('[notify]', msg);
}

/* ---------- Helpers ---------- */
function toFaDigits(val) { return String(val ?? '').replace(/\d/g, d => '۰۱۲۳۴۵۶۷۸۹'[Number(d)]); }

function arToFa(text) {
  if (text == null) return '';
  let s = String(text);
  // نُرمال‌سازی ی، ک، ه و ... + حذف حرکات
  const map = {
    // ی و ک
    '\u064A': '\u06CC', // ي → ی
    '\u06D2': '\u06CC', // ے → ی
    '\u0649': '\u06CC', // ى → ی
    '\u0643': '\u06A9', // ك → ک
    // ه و ة
    '\u0629': '\u0647', // ة → ه
    // همزه‌دارها را نگه می‌داریم، فقط اشکال ترکیبی را ساده می‌کنیم
    '\u0623': '\u0627', // أ → ا
    '\u0625': '\u0627', // إ → ا
    // فاصله کشیده
    '\u0640': '',       // ـ (کَشیده) → حذف
  };
  s = s.replace(/[\u064A\u06D2\u0649\u0643\u0629\u0623\u0625\u0622\u0640]/g, ch => map[ch] || ch);
  s = s.replace(/[\u064B-\u065F\u0670\u06D6-\u06ED]/g, '');      // حذف حرکات
  s = s.replace(/[0-9\u0660-\u0669]/g, d => toFaDigits(d));      // اعداد به فارسی
  return s.replace(/\s+/g, ' ').trim();
}

function formatJalaliFa(dateInput) {
  if (!dateInput) return '-';
  try {
    const d = new Date(dateInput);
    const fmt = new Intl.DateTimeFormat('fa-IR-u-ca-persian', { year: 'numeric', month: '2-digit', day: '2-digit' });
    return fmt.format(d).replace(/\//g, '/');
  } catch { return '-'; }
}
function sortSessions(sessions) {
  return [...(sessions || [])].sort((a, b) => {
    const ad = new Date(a.date), bd = new Date(b.date);
    if (ad - bd) return ad - bd;
    return String(a.startTime || '').localeCompare(String(b.startTime || ''), 'fa');
  });
}
function topicShort(s) {
  const txt = arToFa(s);
  if (!txt) return '-';
  const words = txt.split(' ');
  return words.slice(0, 4).join(' ') + (words.length > 4 ? '...' : '');
}
function hourOnly(t) {
  const hh = String(t || '').split(':')[0] || '';
  return toFaDigits(hh);
}

/* ---------- Token ---------- */
const _sleep = ms => new Promise(r => setTimeout(r, ms));
async function _waitForLocalforage(maxWaitMs = 1500, stepMs = 100) {
  const start = Date.now();
  while (Date.now() - start < maxWaitMs) {
    if (window.localforage && typeof window.localforage.getItem === 'function') return window.localforage;
    await _sleep(stepMs);
  }
  return null;
}
async function _getTokenFromLocalforage() {
  const lf = await _waitForLocalforage(); if (!lf) return null;
  try {
    const t = await lf.getItem('sameh_token');
    if (t) return t;
  } catch { }
  return null;
}
function _getTokenFromLocalStorage() {
  try {
    const t = localStorage.getItem('sameh_token');
    if (t) return t;
  } catch { }
  return null;
}
async function getSamehToken() {
  let token = await _getTokenFromLocalforage(); if (token) return token;
  token = _getTokenFromLocalStorage(); if (token) return token;
  throw new Error('توکن سامح یافت نشد. لطفاً sameh_token را در localforage/localStorage تنظیم کنید.');
}

/* ---------- API ---------- */
async function fetchAttendanceInfo(classId) {
  const token = await getSamehToken();
  const api = `https://sameh.behdasht.gov.ir/api/v2/edu/attendanceInfo?id=${encodeURIComponent(classId)}`;
  const r = await fetch(api, {
    headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json' },
    credentials: 'omit', mode: 'cors', cache: 'no-store'
  });
  if (!r.ok) {
    const tt = await r.text().catch(() => '');
    throw new Error(`API ${r.status}: ${tt || r.statusText}`);
  }
  const j = await r.json();
  if (!j || !j.data) throw new Error('ساختار پاسخ API نامعتبر است.');
  return j.data;
}

/* ---------- Font scales (قابل تنظیم) ---------- */
function getFontScale(S) {
  const cfg = (typeof window !== 'undefined' && window.attendanceFontScales) || null;
  const defaults = {
    1: { base: 9, h1: 10, t1: 9, sub: 9 },
    2: { base: 10.8, h1: 13, t1: 12, sub: 8.5 },
    4: { base: 9.8, h1: 12, t1: 11, sub: 8 },
    6: { base: 8.8, h1: 11, t1: 10, sub: 7.2 },
    10: { base: 8, h1: 10.2, t1: 9.2, sub: 6.5 },
    12: { base: 7.5, h1: 9.6, t1: 8.6, sub: 6.1 },
  };
  const ref = cfg || defaults;
  const keys = Object.keys(ref).map(Number).sort((a, b) => a - b);
  let pick = keys[keys.length - 1];
  for (const k of keys) { if (S <= k) { pick = k; break; } }
  return ref[pick];
}

/* ---------- Rows-per-page rules (قابل تغییر) ---------- */
/*
window.attendanceRowsPerPageRules = { lte:4, rowsWhenLte:20, gte:5, rowsWhenGte:10 };
یا:
window.attendanceRowsPerPageFn = s => (s <= 4 ? 20 : 10);
*/
function getRowsPerPage(sessionCount) {
  if (typeof window !== 'undefined' && typeof window.attendanceRowsPerPageFn === 'function') {
    const v = Number(window.attendanceRowsPerPageFn(sessionCount));
    if (Number.isFinite(v) && v > 0) return Math.floor(v);
  }
  const rules = (typeof window !== 'undefined' && window.attendanceRowsPerPageRules) || {
    lte: 4, rowsWhenLte: 20, gte: 5, rowsWhenGte: 10
  };
  if (sessionCount <= Number(rules.lte ?? 4)) return Math.floor(rules.rowsWhenLte ?? 20);
  if (sessionCount >= Number(rules.gte ?? 5)) return Math.floor(rules.rowsWhenGte ?? 10);
  return Math.floor(((Number(rules.rowsWhenLte ?? 20) + Number(rules.rowsWhenGte ?? 10)) / 2));
}

/* ---------- HEAD + CSS ---------- */
function buildHeadHTML(sessionCount, meta) {
  const orient = (sessionCount <= 4) ? 'portrait' : 'landscape';
  const scale = getFontScale(sessionCount);
  const basePt = scale.base, h1Pt = scale.h1, t1Pt = scale.t1, subPt = scale.sub;
  const baseHref = document.baseURI || (location.origin + location.pathname.replace(/[^/]+$/, ''));

  return `
  <meta charset="utf-8">
  <base href="${baseHref}">
 <title>لیست حضور و غیاب  - دوره ${toFaDigits(meta?.id || '')}</title>
  <style>
    @font-face {
      font-family: 'Shabnam-Bold';
      src: url('css/fonts/Shabnam-Bold-FD.eot');
      src: url('css/fonts/Shabnam-Bold-FD.eot?#iefix') format('embedded-opentype'),
           url('css/fonts/Shabnam-Bold-FD.woff2') format('woff2'),
           url('css/fonts/Shabnam-Bold-FD.woff') format('woff'),
           url('css/fonts/Shabnam-Bold-FD.ttf') format('truetype');
      font-weight: bold; font-style: normal;
    }
    @media print {
      @page { size: A4 ${orient}; margin: 10mm; }
      table.tbl{ break-inside: avoid; page-break-inside: avoid; }
      table.tbl + table.tbl{ break-before: page; page-break-before: always; }
      .page-break{ display:none !important; }
    }
    *{box-sizing:border-box}
    html,body{
      margin:0; padding:0; direction: rtl;
      font-family: Shabnam-Bold, sans-serif;
      color:#000; background:#fff; font-size: ${basePt}pt;
    }
    .wrapper{padding: 0 4px;}

    .tbl{
      width:100%; border-collapse: collapse; table-layout:fixed;
      border:1.6px solid #000; margin: 0;
      --rowh: 28px;   /* از JS ست می‌شود */
      --extra: 0px;   /* از JS ست می‌شود */
    }

    /* وسط‌چین کامل همه سلول‌ها */
    .th, .td{
      font-weight:700;font-size:${basePt}pt;border:1px solid #000; padding:4px 4px; text-align:center;
      vertical-align:middle; line-height:1.8; white-space: normal;
      overflow: hidden; text-overflow: clip; overflow-wrap: break-word;
    }

    thead .th{font-weight:700; white-space: normal; font-size:${basePt}pt;}
    tbody .td{font-size:${basePt}pt;}
    .head-row .th{background:#d9d9d9;}
    .th-title p{margin:2px 0;line-height: 2;}
    .t1{font-size:${h1Pt}pt}
    .t2, .t3, .t4{font-size:${t1Pt}pt}
    .time{font-weight:500;font-size:${basePt}pt}

    /* --- قوانین نمایش محتوا در tbody --- */
    /* پیش‌فرض: یک‌خط با ellipsis برای کنترل ارتفاع */
    tbody .td{ font-weight:700;font-size:${basePt}pt;white-space: nowrap; }

    /* نام و نام خانوادگی و دوره: چندخطی، بدون ellipsis (نمایش کامل) */
    .td-name, .td-course{
      white-space: normal !important;
      text-overflow: clip !important;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .td-nid, .td-mobile{
      direction: ltr;
      unicode-bidi: plaintext;
      text-align: center;
      text-overflow: clip !important;
      /* اعداد منظم‌تر (عرض برابر) */
      font-variant-numeric: tabular-nums;
      /* اگر فونت فعلی اعداد را یکنواخت نکند، یک مونو‌اسپیس هم اضافه کن */
      font-family: Shabnam-Bold, ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace, sans-serif;
      /* برای حالت‌های خیلی تنگ، اجازه‌ی شکستن اجباری (روی مرز هر کاراکتر) */
      overflow-wrap: anywhere;
      word-break: break-all;
      white-space: normal; /* از حالت nowrap خارج می‌شویم تا بتواند بشکند */
      line-height: 1.8;
    }

    /* عرض ستون‌ها (اصلاح‌شده) */
    .td-idx{width:14px; min-width:14px}          /* ردیف باریک‌تر */
    .td-name{text-align:center; padding-right:4px; width:200px}
    .td-nid{width:115px}
    .td-mobile{width:140px}
    .td-course{width:150px}                      /* دوره کمی عریض‌تر */
    .td-sess{width:54px}
    .td-sign{width:90px}                         /* امضا بزرگ‌تر شد */
    .td-score{width:80px}

    tbody tr:nth-child(even) .td{background:#f5f5f5}
    
    .td-nid, .td-mobile, thead .th-nid, thead .th-mobile {
      white-space: nowrap !important;
      overflow: hidden !important;
      text-overflow: clip !important;
      direction: ltr;
      unicode-bidi: plaintext;
      text-align: center;
    }

    .topic { display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; text-overflow: ellipsis; text-align: center; white-space: normal; line-height: 1.8; }

    /* ارتفاع ردیف‌ها */
    .tbl tbody tr{ height: var(--rowh); }
    .tbl tbody tr.row-stretch{ height: calc(var(--rowh) + var(--extra)); }
    .row-teacher{ height: calc(2 * var(--rowh)); }
    .row-teacher.row-stretch{ height: calc(2 * var(--rowh) + var(--extra)); }

    /* امضاهای پایانی (دو برابرِ قبل) */
    .signatures{
      width:100%; border-collapse:collapse; table-layout:fixed;
      margin-top:8px; border:1.2px solid #000;
      break-inside: avoid; page-break-inside: avoid;
    }
    .signatures tr{ height: 160px; } /* 2× 80px */
    .signatures td{
      border:0.8px solid #000; font-size:${t1Pt}pt;
      padding:15px 8px; text-align:center; vertical-align:text-top;
      white-space: nowrap; overflow:hidden; text-overflow:ellipsis;
    }
  </style>`;
}

/* ---------- THEAD ---------- */
function buildTheadHTML(meta, sessions) {
  const S = Math.max(1, Math.min(12, (sessions || []).length || 1));
  const classTitle = arToFa(meta.className || '-');
  const startFa = formatJalaliFa(meta.startDate);
  const endFa = formatJalaliFa(meta.endDate);
  const examFa = formatJalaliFa(meta.examDate);

  const totalCols = 5 + S + 2;
  const titleColspan = totalCols - 2;

  const sessionsText = (S === 1)
    ? `تعداد جلسه : ${toFaDigits(S)}`
    : `تعداد جلسات : 1 تا ${toFaDigits(S)}`;

  const datesRow = sessions.slice(0, S).map(s => `<th class="th">${formatJalaliFa(s.date)}</th>`).join('');
  const headsRow = sessions.slice(0, S).map(s => {
    const inst = arToFa((s.instructor && s.instructor.firstAndLastName) ? s.instructor.firstAndLastName : '-');
    const timeStr = `${hourOnly(s.endTime)} - ${hourOnly(s.startTime)}`;
    return `<th class="th">
              <p>${inst}</p>
              <p class="topic">${topicShort(s.topic)}</p>
              <p class="time">${timeStr}</p>
            </th>`;
  }).join('');

  return `
    <thead>
      <tr>
        <th class="th th-logo" colspan="2">
          <img src="../" alt="sameh-logo" style="width: 6rem; margin: auto; display:block;">
        </th>
        <th class="th th-title" colspan="${titleColspan}">
          <p class="t1">لیست حضور و غیاب شرکت‌کنندگان ${classTitle}</p>
          <p class="t2">( شناسه دوره ${toFaDigits(meta.id || '')} )</p>
          <p class="t3">تاریخ برگزاری دوره : ${startFa} الی ${endFa} | ${sessionsText}</p>
          <p class="t4">تاریخ آزمون : ${examFa}</p>        </th>
      </tr>
      <tr>
        <th class="th" colspan="5">تاریخ‌های برگزاری دوره</th>
        ${datesRow}
        <th class="th" rowspan="2">امضا</th>
        <th class="th" rowspan="2">نمره آزمون</th>
      </tr>
      <tr class="head-row">
        <th class="th th-idx">ردیف</th>
        <th class="th th-name">نام و نام خانوادگی</th>
        <th class="th th-nid">کدملی</th>
        <th class="th th-mobile">موبایل</th>
        <th class="th">دوره ثبت نام شده</th>
        ${headsRow}
      </tr>
    </thead>`;
}

/* ---------- ردیف‌ها ---------- */
function makeRowHTML(a, rowNumberFa, sessionCount) {
  const fullName = arToFa([a.firstName || '', a.lastName || ''].filter(Boolean).join(' '));
  const course = arToFa(a.registrationCourse || '-');
  return `
    <tr>
      <td class="td td-idx">${rowNumberFa}</td>
      <td class="td td-name">${fullName}</td>
      <td class="td td-nid">${toFaDigits(a.nationalId || '')}</td>
      <td class="td td-mobile">${toFaDigits(a.cellphone || '')}</td>
      <td class="td td-course">${course}</td>
      ${Array.from({ length: sessionCount }).map(() => `<td class="td td-sess"></td>`).join('')}
      <td class="td td-sign"></td>
      <td class="td td-score"></td>
    </tr>`;
}
function makeEmptyRowHTML(sessionCount) {
  return `
    <tr>
      <td class="td td-idx"></td>
      <td class="td td-name"></td>
      <td class="td td-nid"></td>
      <td class="td td-mobile"></td>
      <td class="td td-course"></td>
      ${Array.from({ length: sessionCount }).map(() => `<td class="td td-sess"></td>`).join('')}
      <td class="td td-sign"></td>
      <td class="td td-score"></td>
    </tr>`;
}
function makeTeacherRowHTML(sessionCount) {
  return `
    <tr class="row-teacher">
      <td class="td" colspan="5">${arToFa('امضای مدرس')}</td>
      ${Array.from({ length: sessionCount }).map(() => `<td class="td td-sess"></td>`).join('')}
      <td class="td"></td>
      <td class="td"></td>
    </tr>`;
}

/* ---------- امضاهای پایانی ---------- */
function buildSignaturesHTML() {
  return `
    <table class="signatures" id="finalSignatures">
      <tr>
        <td>${arToFa('امضای مدیر / مسئول فنی دفتر خدمات سلامت')}</td>
        <td>${arToFa('امضای ناظر دانشگاه')}</td>
      </tr>
    </table>`;
}

/* ---------- اسکلت سند چاپ ---------- */
function buildPrintableHTMLSkeleton(sessionCount, meta) {
  const headHTML = buildHeadHTML(sessionCount, meta);
  return `<!doctype html>
 <html lang="fa" dir="rtl">
 <head>${headHTML}</head>
 <body>
   <div class="wrapper" id="wrapper"></div>
 </body>
 </html>`;
}

/* ---------- اندازه‌گیری ارتفاع‌ها ---------- */
function measureHeights(doc, theadHTML, signaturesHTML, pagePx) {
  const probe = doc.createElement('div');
  probe.style.position = 'absolute'; probe.style.visibility = 'hidden'; probe.style.left = '-9999px'; probe.style.top = '0';
  probe.innerHTML = `
    <table class="tbl">${theadHTML}<tbody><tr><td class="td">x</td></tr></tbody></table>
    ${signaturesHTML}
  `;
  doc.body.appendChild(probe);

  const probeTable = probe.querySelector('table.tbl');
  const theadH = probeTable.tHead.getBoundingClientRect().height;

  const sigEl = probe.querySelector('#finalSignatures');
  const sigH = sigEl.getBoundingClientRect().height + 8; // margin-top تقریبی

  const minRowH = 8;

  probe.remove();
  return { theadH, sigH, minRowH, pagePx };
}

/* ---------- فیتِ دقیق جدول روی صفحه ---------- */
function fitTable(table, { pagePx, units, theadH, reserveBottomPx, minRowH, maxIter = 10 }) {
  const tbody = table.tBodies[0];
  if (!tbody) return;

  // حدس اولیه
  let base = Math.floor((pagePx - theadH - reserveBottomPx - 1) / units);
  if (!Number.isFinite(base) || base < minRowH) base = minRowH;
  let extra = 0;

  for (let i = 0; i < maxIter; i++) {
    table.style.setProperty('--rowh', `${base}px`);
    table.style.setProperty('--extra', `0px`);
    const totalH = table.getBoundingClientRect().height + reserveBottomPx;
    const diff = Math.round(pagePx - totalH); // + => جا داریم | - => سربار

    if (diff === 0) { extra = 0; break; }
    if (diff > 0) { extra = diff; break; }

    const overshoot = -diff;
    const delta = Math.max(1, Math.ceil(overshoot / units));
    base = Math.max(minRowH, base - delta);
  }

  // یک ردیف stretch برای جذب extra
  Array.from(tbody.querySelectorAll('tr.row-stretch')).forEach(tr => tr.classList.remove('row-stretch'));
  const lastRow = tbody.lastElementChild;
  if (lastRow) {
    table.style.setProperty('--extra', `${extra}px`);
    if (extra > 0) lastRow.classList.add('row-stretch');
  }
}

/* ---------- فیتِ فونتِ سلول‌ها برای نمایش کامل محتوا (بدون "...") ---------- */
function fitCellsFontToRow(table) {
  const tbody = table.tBodies[0];
  if (!tbody) return;
  // ارتفاع هدف هر ردیف
  const rowH = parseFloat(getComputedStyle(table).getPropertyValue('--rowh')) || 28;
  const extra = parseFloat(getComputedStyle(table).getPropertyValue('--extra')) || 0;

  // سلکتور سلول‌هایی که باید کامل نمایش داده شوند
  const targets = tbody.querySelectorAll('.td-name, .td-course, .td-nid, .td-mobile');
  targets.forEach(td => {
    // اگر ردیف stretch است، ارتفاع کمی بیشتر است
    const tr = td.closest('tr');
    const targetH = tr && tr.classList.contains('row-stretch') ? (rowH + extra) : rowH;

    // کاهش تدریجی فونت تا جا شود (حداقل 75%)
    let fs = parseFloat(getComputedStyle(td).fontSize) || 12;
    const baseFs = fs;
    for (let i = 0; i < 10; i++) {
      // اندازه‌گیری با scrollHeight
      if (td.scrollHeight <= targetH - 6 /* paddings و خطا */) break;
      fs = Math.max(baseFs * 0.75, fs - Math.max(0.5, baseFs * 0.05));
      td.style.fontSize = fs + 'px';
    }

    // برای نام و دوره: حذف ellipsis قطعی
    if (td.classList.contains('td-name') || td.classList.contains('td-course')) {
      td.style.textOverflow = 'clip';
      td.style.whiteSpace = 'normal';
    }
    // برای کد ملی و موبایل: اطمینان از بدون ellipsis
    if (td.classList.contains('td-nid') || td.classList.contains('td-mobile')) {
      td.style.textOverflow = 'clip';
      td.style.whiteSpace = 'nowrap';
    }
  });
}

/* ---------- اجرای اصلی چاپ ---------- */
async function printAttendanceListEmbedded() {
  try {
    const classIdEl = document.querySelector('#classSelect');
    const classId = classIdEl ? classIdEl.value : null;
    if (!classId || classId === '0') {
      _notify('لطفاً ابتدا یک کلاس را انتخاب کنید.', 'warning');
      return;
    }

    const data = await fetchAttendanceInfo(classId);
    const meta = data.guildClass || {};
    const sessionsSorted = sortSessions(data.sessions || []);
    const applicants = data.applicants || [];
    const S = Math.max(1, Math.min(12, sessionsSorted.length || 1));

    const ROWS_PER_PAGE = getRowsPerPage(S);

    // پنجره چاپ
    const w = window.open('', '_blank', 'width=1200,height=800');
    if (!w) { _notify('پنجره چاپ مسدود شده است. مسدودکننده پاپ‌آپ را غیرفعال کنید.', 'error'); return; }

    // اسکلت HTML
    w.document.open('text/html');
    w.document.write(buildPrintableHTMLSkeleton(S, meta));
    w.document.close();

    const d = w.document;
    const wrapper = d.getElementById('wrapper');

    // پارامترهای A4
    const mm = 3.7795275591; // px per mm
    const orient = (S <= 4) ? 'portrait' : 'landscape';
    const pageInnerHeightMM = (orient === 'portrait') ? (297 - 20) : (210 - 20); // A4 - 2×10mm margin
    const PAGE_PX = pageInnerHeightMM * mm;

    // HTMLهای ثابت
    const theadHTML = buildTheadHTML(meta, sessionsSorted);
    const signaturesHTML = buildSignaturesHTML();

    // اندازه‌گیری header و signatures
    const metrics = measureHeights(d, theadHTML, signaturesHTML, PAGE_PX);
    const { theadH, sigH, minRowH } = metrics;

    // سازندهٔ جدول
    function newTable() {
      const tbl = d.createElement('table');
      tbl.className = 'tbl';
      tbl.innerHTML = theadHTML + '<tbody></tbody>';
      return tbl;
    }

    // تقسیم داده‌ها به صفحات
    const pages = [];
    for (let i = 0; i < applicants.length; i += ROWS_PER_PAGE) {
      pages.push(applicants.slice(i, i + ROWS_PER_PAGE));
    }
    if (pages.length === 0) pages.push([]);

    const lastPageIndex = pages.length - 1;
    let globalRowIndex = 1;

    pages.forEach((chunk, pageIdx) => {
      const table = newTable();
      wrapper.appendChild(table);
      const tbody = table.querySelector('tbody');
      const isLast = (pageIdx === lastPageIndex);

      // ردیف‌های داده
      for (const a of chunk) {
        const tmp = d.createElement('tbody');
        tmp.innerHTML = makeRowHTML(a, toFaDigits(globalRowIndex++), S);
        tbody.appendChild(tmp.firstElementChild);
      }

      // صفحه‌های غیرآخر: تا عدد ثابت پر کن
      if (!isLast) {
        const need = ROWS_PER_PAGE - chunk.length;
        for (let k = 0; k < need; k++) {
          const tmp = d.createElement('tbody'); tmp.innerHTML = makeEmptyRowHTML(S);
          tbody.appendChild(tmp.firstElementChild);
        }
      } else {
        // صفحهٔ آخر: ردیف امضای مدرس (۲ واحد) داخل همین جدول
        const tmpT = d.createElement('tbody'); tmpT.innerHTML = makeTeacherRowHTML(S);
        tbody.appendChild(tmpT.firstElementChild);
      }

      // تعداد «واحد-ردیف» این صفحه
      const units = isLast ? (chunk.length + 2) : ROWS_PER_PAGE;

      // رزرو پایین صفحه: فقط صفحهٔ آخر باید جا برای signatures نگه دارد
      const reserveBottomPx = isLast ? (sigH + 8) : 0;

      // فیت دقیق جدول روی A4
      fitTable(table, {
        pagePx: PAGE_PX,
        units,
        theadH,
        reserveBottomPx,
        minRowH
      });

      // فیت فونتِ سلول‌ها برای نمایش کامل محتوا (بدون «...») در ارتفاع ثابت
      fitCellsFontToRow(table);

      // بعد از جدولِ آخر، جدول امضاها را اضافه کن (همان صفحه)
      if (isLast) {
        const sigWrap = d.createElement('div');
        sigWrap.innerHTML = signaturesHTML;
        wrapper.appendChild(sigWrap.firstElementChild);
      }
    });

    // چاپ
    w.addEventListener('load', function () {
      try { w.print(); } catch (e) { }
      setTimeout(function () { try { w.close(); } catch (e) { } }, 400);
    });

  } catch (err) {
    console.error(err);
    _notify(err.message || 'امکان تهیه لیست حضور و غیاب وجود ندارد.', 'error');
  }
}
async function renderAttendancePreview(classId) {
  if (!classId) { alert("لطفاً کلاس را انتخاب کنید."); return; }

  // ایجاد یا انتخاب container
  let container = document.getElementById("attendanceContainer");
  if (!container) {
    container = document.createElement("div");
    container.id = "attendanceContainer";
    document.body.appendChild(container);
  }
  container.innerHTML = "";

  // دریافت اطلاعات از API
  let data;
  try { data = await fetchAttendanceInfo(classId); }
  catch (err) { container.textContent = "خطا در دریافت داده‌ها"; return; }

  const meta = data.guildClass || {};
  const sessions = data.sessions || [];
  const applicants = data.applicants || [];
  const sessionCount = sessions.length;

  // بررسی حالت landscape
  const totalColumns = 7 + sessionCount; // 7 ستون اصلی + تعداد جلسات
  const isLandscape = sessionCount > 4;

  // افزودن CSS داخل JS
  const style = document.createElement("style");
  style.textContent = `
        .attendance-wrapper {
            border:2px solid #000;
            border-radius:12px;
            padding:20px;
            max-width:${isLandscape ? "1600px" : "1200px"};
            margin:20px auto;
            background:#fff;
            font-family:Tahoma,sans-serif;
            color:#000;
            overflow-x:auto;
        }
        .attendance-header {
            display:flex;
            justify-content:space-between;
            align-items:flex-start;
            border-bottom:2px solid #000;
            padding-bottom:15px;
            margin-bottom:20px;
        }
        .attendance-header-text { width:80%; text-align:center; font-size:18px; font-weight:bold; line-height:1.8; }
        .attendance-header-text span { font-size:14px; font-weight:normal; }
        .attendance-logo { width:20%; display:flex; justify-content:center; align-items:center; border-left:2px solid #000; padding-left:10px; }
        .attendance-logo img { height:60px; object-fit:contain; }
        .attendance-table { width:100%; border-collapse:collapse; table-layout:fixed; font-size:13px; }
        .attendance-table th, .attendance-table td { border:2px solid #000; padding:6px; text-align:center; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        .attendance-table th { background:#e0e0e0; font-weight:bold; }
        .attendance-footer { margin-top:40px; display:flex; justify-content:space-between; font-size:14px; font-weight:bold; }
        @media print { body * { visibility:hidden; } .attendance-wrapper, .attendance-wrapper * { visibility:visible; } .attendance-wrapper { position:absolute; top:0; left:0; width:100%; } }
    `;
  container.appendChild(style);

  // wrapper
  const wrapper = document.createElement("div");
  wrapper.className = "attendance-wrapper";

  // هدر
  const header = document.createElement("div");
  header.className = "attendance-header";
  const headerText = document.createElement("div");
  headerText.className = "attendance-header-text";
  headerText.innerHTML = `
        لیست حضور و غیاب شرکت‌کنندگان ${meta.className || '-'}<br>
        <span>تاریخ برگزاری: ${meta.startDate} الی ${meta.endDate}</span>
    `;
  const logoBox = document.createElement("div");
  logoBox.className = "attendance-logo";
  const logoImg = document.createElement("img");
  logoImg.src = "img/samo-logo.png";
  logoImg.alt = "لوگو سامو";
  logoBox.appendChild(logoImg);
  header.appendChild(headerText);
  header.appendChild(logoBox);
  wrapper.appendChild(header);

  // جدول
  const table = document.createElement("table");
  table.className = "attendance-table";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  const headers = ["ردیف", "نام و نام خانوادگی", "کد ملی", "موبایل", "دوره ثبت شده"];
  sessions.forEach((s, i) => headers.push("جلسه " + toFaDigits(i + 1)));
  headers.push("امضا", "نمره آزمون");

  headers.forEach(h => { const th = document.createElement("th"); th.textContent = h; headRow.appendChild(th); });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  applicants.forEach((a, idx) => {
    const tr = document.createElement("tr");
    const values = [
      toFaDigits(idx + 1),
      [a.firstName, a.lastName].filter(Boolean).join(' '),
      toFaDigits(a.nationalId || ''),
      toFaDigits(a.cellphone || ''),
      a.registrationCourse || '-'
    ];
    sessions.forEach(_ => values.push(""));
    values.push("", ""); // امضا و نمره
    values.forEach(v => { const td = document.createElement("td"); td.textContent = v; tr.appendChild(td); });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrapper.appendChild(table);

  // فوتر
  const footer = document.createElement("div");
  footer.className = "attendance-footer";
  footer.innerHTML = `
        <div>امضای رئیس جلسه</div>
        <div>امضای ناظر دانشگاه</div>
        <div>امضای مدیر / مسئول فنی دفتر خدمات سلامت</div>
    `;
  wrapper.appendChild(footer);

  container.appendChild(wrapper);
}

// بایند روی دکمه نمایش
document.getElementById("loadAttendanceBtn").addEventListener("click", function () {
  const classId = document.getElementById("classSelect").value;
  renderAttendancePreview(classId);
});
// ---------- شبیه‌سازی API ----------
async function fetchAttendanceInfo(classId) {
  await new Promise(r => setTimeout(r, 200));
  return {
    guildClass: { className: `کلاس ${classId}`, startDate: '1402/06/01', endDate: '1402/06/10' },
    sessions: ['جلسه ۱', 'جلسه ۲', 'جلسه ۳', 'جلسه ۴', 'جلسه ۵'], // اینجا تعداد جلسه واقعی از API میاد
    applicants: [
      { firstName: 'فاطمه', lastName: 'محمدی', nationalId: '045045XXXX', cellphone: '093645XXXXX', registrationCourse: 'دوره عمومی' },
      { firstName: 'علی', lastName: 'رضایی', nationalId: '045046XXXX', cellphone: '093646XXXXX', registrationCourse: 'دوره تخصصی' }
      // تعداد ردیف‌ها هم از API گرفته می‌شود
    ]
  };
}

// ---------- تبدیل اعداد به فارسی ----------
function toFaDigits(input) {
  const fa = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹'];
  return input.toString().replace(/\d/g, d => fa[d]);
}

// ---------- پیش‌نمایش حضور و غیاب ----------
async function renderAttendancePreview(classId) {
  if (!classId) { alert("لطفاً کلاس را انتخاب کنید."); return; }

  let container = document.getElementById("attendanceContainer");
  if (!container) {
    container = document.createElement("div");
    container.id = "attendanceContainer";
    document.body.appendChild(container);
  }
  container.innerHTML = "";

  // افزودن استایل داخل JS
  const style = document.createElement("style");
  style.textContent = `
        .container { border:2px solid black; margin:30px auto; padding:20px; max-width:1200px; background:#fff; font-family:Tahoma,sans-serif; }
        .header { display:flex; justify-content:space-between; align-items:center; border-bottom:2px solid black; padding-bottom:20px; margin-bottom:30px; }
        .header-text { width:78%; text-align:center; font-size:18px; font-weight:bold; line-height:2.2; }
        .header-text span { display:block; font-size:15px; font-weight:normal; margin-top:8px; }
        .logo-box { width:22%; display:flex; justify-content:center; align-items:center; border-left:2px solid black; padding-left:10px; }
        .logo-box img { height:70px; width:auto; object-fit:contain; }
        table { width:100%; border-collapse:collapse; table-layout:fixed; font-size:13px; }
        thead th { background:#e0e0e0; border:2px solid black; padding:10px; font-weight:bold; text-align:center; white-space:nowrap; }
        tbody td { border:2px solid black; padding:8px; text-align:center; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; background:#fff; }
        tbody tr:nth-child(even) td { background:#f9f9f9; }
        .footer { margin-top:50px; display:flex; justify-content:space-between; font-size:14px; font-weight:bold; }
        @media print { body * { visibility:hidden; } .container, .container * { visibility:visible; } .container { position:absolute; top:0; left:0; width:100%; } }
    `;
  container.appendChild(style);

  // دریافت داده از API
  let data;
  try { data = await fetchAttendanceInfo(classId); }
  catch (e) { container.textContent = "خطا در دریافت داده‌ها"; return; }

  const meta = data.guildClass || {};
  const sessions = data.sessions || [];
  const applicants = data.applicants || [];

  const isLandscape = sessions.length > 4; // اگر جلسه >4 افقی کنیم
  const wrapper = document.createElement("div");
  wrapper.className = "container";
  if (isLandscape) wrapper.style.maxWidth = "1600px";

  // هدر
  const header = document.createElement("div"); header.className = "header";
  const headerText = document.createElement("div"); headerText.className = "header-text";
  headerText.innerHTML = `لیست حضور و غیاب شرکت‌کنندگان ${meta.className || '-'}<br>
                            <span>تاریخ برگزاری: ${meta.startDate} الی ${meta.endDate}</span>`;
  const logoBox = document.createElement("div"); logoBox.className = "logo-box";
  const logoImg = document.createElement("img"); logoImg.src = "img/samo-logo.png"; logoImg.alt = "لوگو سامو";
  logoBox.appendChild(logoImg);
  header.appendChild(headerText); header.appendChild(logoBox); wrapper.appendChild(header);

  // جدول
  const table = document.createElement("table");
  const thead = document.createElement("thead"); const headRow = document.createElement("tr");
  const headers = ["ردیف", "نام و نام خانوادگی", "کد ملی", "موبایل", "دوره ثبت شده"];
  sessions.forEach((s, i) => headers.push("جلسه " + toFaDigits(i + 1)));
  headers.push("امضا", "نمره آزمون");
  headers.forEach(h => { const th = document.createElement("th"); th.textContent = h; headRow.appendChild(th); });
  thead.appendChild(headRow); table.appendChild(thead);

  const tbody = document.createElement("tbody");
  applicants.forEach((a, idx) => {
    const tr = document.createElement("tr");
    const values = [toFaDigits(idx + 1), [a.firstName, a.lastName].filter(Boolean).join(' '), toFaDigits(a.nationalId || ''), toFaDigits(a.cellphone || ''), a.registrationCourse || '-'];
    sessions.forEach(_ => values.push("")); values.push("", ""); // امضا و نمره
    values.forEach(v => { const td = document.createElement("td"); td.textContent = v; tr.appendChild(td); });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody); wrapper.appendChild(table);

  // فوتر
  const footer = document.createElement("div"); footer.className = "footer";
  footer.innerHTML = `<div>امضای رئیس جلسه</div><div>امضای ناظر دانشگاه</div><div>امضای مدیر / مسئول فنی دفتر خدمات سلامت</div>`;
  wrapper.appendChild(footer);

  container.appendChild(wrapper);
}

// بایند روی دکمه
document.getElementById("loadAttendanceBtn").addEventListener("click", function () {
  const classId = document.getElementById("classSelect").value;
  renderAttendancePreview(classId);
});
