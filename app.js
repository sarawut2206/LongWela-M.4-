/* ==========================================================================
   บันทึกการมาสาย ม.4  —  ทำงานในเบราว์เซอร์ล้วน ไม่ต้องมีเซิร์ฟเวอร์
   ข้อมูลเก็บใน localStorage และส่งออกเป็นไฟล์ Word (.docx) ตามฟอร์มโรงเรียน
   ========================================================================== */

const KEY = 'malaisai_m4_v2';
const SESS = { '0750': '07.50', '0830': '08.30' };
const TH_MONTH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

/* ลิงก์ Apps Script (Google Sheet) — เปลี่ยนได้ในแท็บ "ข้อมูล / สำรอง" */
const GS_URL_DEFAULT =
  'https://script.google.com/macros/s/AKfycbxykvR2ZPolzozxvtT1uDxsbsqSHMsNTPU9WiGPW5rgJmsoUgIBWUhvMmiqo-Tj3Bojag/exec';

/* ---------- state ---------- */
let S = load();

function baselineCarry() {
  const B = window.BASELINE || {};
  return { '0750': Object.assign({}, B['0750']), '0830': Object.assign({}, B['0830']) };
}
function blank() {
  return {
    v: 2,
    sess: '0750',
    cur: mondayISO(new Date()),
    room: null,
    weeks: {},          // isoMonday -> { no, label, marks:{ '0750':{ 'ห้อง-เลขที่': จำนวนครั้ง }, '0830':{} } }
    carry: baselineCarry(),   // เริ่มต้น = ยอดสะสมถึงสัปดาห์ที่ 12 จากไฟล์ Word เดิม
    roster: null,       // null = ใช้รายชื่อจาก students.js
    cloud: { url: GS_URL_DEFAULT, auto: true, rev: 0, at: '' },
    pending: { late: {}, carry: {} },   // รายการที่แก้แล้วยังไม่ได้ส่งขึ้น Google Sheet
    meta: {
      signers: [
        { name: 'นายศราวุธ  พิมศร', pos: 'รองหัวหน้าระดับชั้นมัธยมศึกษาปีที่ 4' },
        { name: 'นางมานิดา ยอดเมือง', pos: 'รองหัวหน้าระดับชั้นมัธยมศึกษาปีที่ 4' }
      ],
      who: 0,
      level: '4', thr: 4, note: true
    }
  };
}
function signer() { return S.meta.signers[S.meta.who] || S.meta.signers[0]; }
function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return blank();
    const s = Object.assign(blank(), JSON.parse(raw));
    s.meta = Object.assign(blank().meta, s.meta || {});
    if (!Array.isArray(s.meta.signers) || s.meta.signers.length < 2) s.meta.signers = blank().meta.signers;
    if (s.meta.who !== 0 && s.meta.who !== 1) s.meta.who = 0;
    // ถ้าไฟล์ students.js ถูกอัปเดตเป็นรายชื่อชุดใหม่ ให้ทิ้งรายชื่อที่เคยแก้ไว้กับชุดเก่า
    if (s.roster && s.rosterVer !== window.ROSTER_VERSION) { s.roster = null; s.rosterVer = null; }
    s.carry = Object.assign({ '0750': {}, '0830': {} }, s.carry || {});
    s.cloud = Object.assign({ url: GS_URL_DEFAULT, auto: true, rev: 0, at: '' }, s.cloud || {});
    s.pending = Object.assign({ late: {}, carry: {} }, s.pending || {});
    // ข้อมูลรูปแบบเก่า (ติ๊กรายวัน 5 ช่อง) -> แปลงเป็นจำนวนครั้ง
    Object.values(s.weeks || {}).forEach(w => ['0750', '0830'].forEach(se => {
      const m = (w.marks || {})[se]; if (!m) return;
      Object.keys(m).forEach(k => { if (Array.isArray(m[k])) m[k] = m[k].filter(Boolean).length || undefined; });
      Object.keys(m).forEach(k => { if (!m[k]) delete m[k]; });
    }));
    return s;
  } catch (e) { return blank(); }
}
function save() { localStorage.setItem(KEY, JSON.stringify(S)); autoSync(); }

/* ---------- helpers: วันที่ ---------- */
function mondayISO(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const diff = (x.getDay() + 6) % 7;          // จันทร์ = 0
  x.setDate(x.getDate() - diff);
  return iso(x);
}
function iso(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function parseISO(s) { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function rangeLabel(startISO) {
  const a = parseISO(startISO), b = addDays(a, 4);
  const be = b.getFullYear() + 543;
  if (a.getMonth() === b.getMonth())
    return `${a.getDate()} – ${b.getDate()} ${TH_MONTH[b.getMonth()]} ${be}`;
  return `${a.getDate()} ${TH_MONTH[a.getMonth()]} – ${b.getDate()} ${TH_MONTH[b.getMonth()]} ${be}`;
}

/* ---------- helpers: ข้อมูล ---------- */
function roster() { return S.roster || window.ROSTER_DEFAULT || []; }
function rooms() { return [...new Set(roster().map(s => s.r))].sort((a, b) => a - b); }
function keyOf(s) { return s.r + '-' + s.n; }
function week(isoStart, create) {
  let w = S.weeks[isoStart];
  if (!w && create) {
    const prev = weekList().filter(k => k < isoStart).pop();
    let no = '';
    if (prev) {
      const gap = Math.round((parseISO(isoStart) - parseISO(prev)) / 604800000);
      const pn = parseInt(S.weeks[prev].no, 10);
      if (!isNaN(pn)) no = String(pn + gap);
    }
    w = S.weeks[isoStart] = { no, label: rangeLabel(isoStart), marks: { '0750': {}, '0830': {} } };
  }
  return w;
}
function weekList() { return Object.keys(S.weeks).sort(); }
function marks(isoStart, sess) {
  const w = week(isoStart, true);
  if (!w.marks[sess]) w.marks[sess] = {};
  return w.marks[sess];
}
/* จำนวนครั้งที่สายในสัปดาห์นั้น (เก็บเป็นตัวเลขล้วน ไม่ผูกกับวัน) */
function weekCount(isoStart, sess, k) {
  const m = (S.weeks[isoStart]?.marks?.[sess] || {})[k];
  if (Array.isArray(m)) return m.filter(Boolean).length;   // ข้อมูลรูปแบบเก่า
  return Number(m || 0);
}
/* จำนวนครั้งสะสม = ยอดยกมา + ทุกสัปดาห์ที่วันที่ <= สัปดาห์ที่เลือก */
function total(sess, k, uptoISO) {
  let t = Number(S.carry[sess]?.[k] || 0);
  for (const wk of weekList()) {
    if (uptoISO && wk > uptoISO) continue;
    t += weekCount(wk, sess, k);
  }
  return t;
}
/* ยอดสะสม "ก่อน" สัปดาห์ที่กำลังบันทึก = ยอดยกมา + สัปดาห์ที่เก่ากว่า */
function prevTotal(sess, k) {
  let t = Number(S.carry[sess]?.[k] || 0);
  for (const wk of weekList()) if (wk < S.cur) t += weekCount(wk, sess, k);
  return t;
}
/* ตั้งจำนวนครั้งของสัปดาห์นี้ */
function setCount(k, n) {
  const m = marks(S.cur, S.sess);
  n = Math.max(0, Math.min(99, Math.round(Number(n) || 0)));
  if (n) m[k] = n; else delete m[k];
  S.pending.late[S.sess + '|' + S.cur + '|' + k] = true;
  save();
}
function addCount(k, d) { setCount(k, weekCount(S.cur, S.sess, k) + d); }
/* ตั้ง "ยอดก่อนสัปดาห์นี้" โดยปรับที่ยอดยกมาให้ผลรวมออกมาตามที่พิมพ์ */
function setPrev(k, n) {
  n = Math.max(0, Math.min(999, Math.round(Number(n) || 0)));
  const fromWeeks = prevTotal(S.sess, k) - Number(S.carry[S.sess]?.[k] || 0);
  const carry = Math.max(0, n - fromWeeks);
  if (!S.carry[S.sess]) S.carry[S.sess] = {};
  if (carry) S.carry[S.sess][k] = carry; else delete S.carry[S.sess][k];
  S.pending.carry[S.sess + '|' + k] = true;
  save();
}

/* ---------- ข้อความแจ้งผลสั้น ๆ มุมล่าง ---------- */
let toastTimer = null;
function toast(msg, bad) {
  const el = $('#toast'); if (!el) return;
  el.innerHTML = msg;
  el.className = 'on' + (bad ? ' bad' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = ''; }, bad ? 5000 : 2600);
}

/* ==========================================================================
   TABS
   ========================================================================== */
$$('nav button').forEach(b => b.onclick = () => {
  $$('nav button').forEach(x => x.classList.toggle('on', x === b));
  ['rec', 'sum', 'dash', 'carry', 'roster', 'data'].forEach(t => $('#tab-' + t).hidden = (t !== b.dataset.tab));
  if (b.dataset.tab === 'sum') renderSummary();
  if (b.dataset.tab === 'dash') renderDash();
  if (b.dataset.tab === 'carry') renderCarry();
  if (b.dataset.tab === 'roster') renderRoster();
  if (b.dataset.tab === 'data') renderData();
});

/* session switchers (ทั้ง 3 หน้าใช้ค่าเดียวกัน) */
function setSess(s) {
  S.sess = s; save();
  $$('#segSess button,#segSess2 button,#segSess3 button').forEach(b => b.classList.toggle('on', b.dataset.s === s));
  $('#hdSess').textContent = 'สาย ' + SESS[s] + ' น.';
  renderRec(); if (!$('#tab-sum').hidden) renderSummary(); if (!$('#tab-carry').hidden) renderCarry();
}
$$('#segSess button,#segSess2 button,#segSess3 button').forEach(b => b.onclick = () => setSess(b.dataset.s));

/* ==========================================================================
   หน้า 1 : ติ๊กมาสาย
   ========================================================================== */
$('#wkStart').onchange = () => {
  S.cur = mondayISO(parseISO($('#wkStart').value || iso(new Date())));
  week(S.cur, true); save(); syncWeekBar(); renderRec();
};
$('#wkNo').oninput = () => { week(S.cur, true).no = $('#wkNo').value; save(); syncWeekBar(true); };
$('#wkLabel').oninput = () => { week(S.cur, true).label = $('#wkLabel').value; save(); };
$('#q').oninput = renderRec;

function syncWeekBar(skipInputs) {
  const w = week(S.cur, true);
  if (!skipInputs) {
    $('#wkStart').value = S.cur;
    $('#wkNo').value = w.no;
    $('#wkLabel').value = w.label;
  }
  $('#hdWeek').textContent = 'สัปดาห์ที่ ' + (w.no || '–') + ' • ' + w.label;
}

function renderRooms() {
  const box = $('#rooms'); box.innerHTML = '';
  const m = marks(S.cur, S.sess);
  const rs = rooms();
  if (S.room === null || !rs.includes(S.room)) S.room = rs[0];
  rs.forEach(r => {
    const cnt = roster().filter(s => s.r === r && m[keyOf(s)]).length;
    const b = document.createElement('button');
    b.className = (r === S.room ? 'on' : '');
    b.innerHTML = 'ห้อง ' + r + (cnt ? `<span class="n">${cnt}</span>` : '');
    b.onclick = () => { S.room = r; save(); renderRec(); };
    box.appendChild(b);
  });
}

function renderRec() {
  syncWeekBar();
  renderRooms();
  const q = $('#q').value.trim();
  const list = q
    ? roster().filter(s => s.name.includes(q) || String(s.n) === q)
    : roster().filter(s => s.r === S.room);
  const thr = Number(S.meta.thr) || 4;

  let h = '<thead><tr>' + (q ? '<th>ห้อง</th>' : '') +
    '<th>เลขที่</th><th style="text-align:left">ชื่อ-สกุล</th>' +
    '<th title="ยอดสะสมก่อนสัปดาห์นี้ — พิมพ์แก้ได้">ยอดเดิม</th>' +
    '<th>สายสัปดาห์นี้ (ครั้ง)</th>' +
    '<th>รวมทั้งหมด</th></tr></thead><tbody>';

  list.forEach(s => {
    const k = keyOf(s);
    const base = prevTotal(S.sess, k);
    const wc = weekCount(S.cur, S.sess, k);
    const tt = base + wc;
    const cls = tt >= thr ? 'over' : (wc ? 'hit' : '');
    h += `<tr class="${cls}">` + (q ? `<td>${s.r}</td>` : '') +
      `<td>${s.n}</td><td class="name">${s.name}</td>` +
      `<td><input class="base" type="number" min="0" inputmode="numeric" value="${base || ''}" data-k="${k}" placeholder="0"></td>` +
      `<td class="cnt">` +
      `<button class="stp" data-k="${k}" data-d="-1" ${wc ? '' : 'disabled'}>−</button>` +
      `<input class="num" type="number" min="0" inputmode="numeric" value="${wc || ''}" data-k="${k}" placeholder="0">` +
      `<button class="stp add" data-k="${k}" data-d="1">+</button></td>` +
      `<td class="tot ${tt >= thr ? 'hi' : ''}">${tt || ''}</td></tr>`;
  });
  h += '</tbody>';
  $('#tblRec').innerHTML = h;

  $('#tblRec').querySelectorAll('.stp').forEach(el => el.onclick = () => {
    addCount(el.dataset.k, Number(el.dataset.d)); renderRec();
  });
  $('#tblRec').querySelectorAll('.num').forEach(el => el.onchange = () => {
    setCount(el.dataset.k, el.value); renderRec();
  });
  $('#tblRec').querySelectorAll('.base').forEach(el => el.onchange = () => {
    setPrev(el.dataset.k, el.value); renderRec();
  });

  const m = marks(S.cur, S.sess);
  const nRoom = roster().filter(s => s.r === S.room && weekCount(S.cur, S.sess, keyOf(s))).length;
  const sumRoom = roster().filter(s => s.r === S.room)
    .reduce((a, s) => a + weekCount(S.cur, S.sess, keyOf(s)), 0);
  const nAll = Object.keys(m).length;
  $('#recInfo').innerHTML = q
    ? `พบ ${list.length} คน`
    : `ห้อง ${S.room} • ${list.length} คน • สัปดาห์นี้สาย <b>${nRoom}</b> คน รวม <b>${sumRoom}</b> ครั้ง` +
      ` (ทุกห้อง ${nAll} คน) — ช่อง <b>ยอดเดิม</b> พิมพ์แก้ได้ ถ้าเด็กมีสถิติเดิมในสมุดแต่ยังไม่มีในระบบ`;
}

/* ==========================================================================
   หน้า 2 : สรุป & พิมพ์
   ========================================================================== */
$('#thr').oninput = () => { S.meta.thr = Number($('#thr').value) || 4; save(); renderSummary(); renderRec(); };
$('#chkNote').onchange = () => { S.meta.note = $('#chkNote').checked; save(); renderSummary(); };
$('#btnPrint').onclick = () => printReport();
$('#btnPdf').onclick = () => printReport(true);
/* พิมพ์ / บันทึกเป็น PDF — ตั้งชื่อไฟล์ให้ตรงกับชื่อเอกสารก่อนสั่งพิมพ์ */
function printReport(isPdf) {
  renderSummary();
  const w = week(S.cur, true);
  const old = document.title;
  document.title = `สรุปสาย ${SESS[S.sess]} น. สัปดาห์ที่ ${w.no || ''} ${(w.label || '').replace(/\s*–\s*/g, '-')}`;
  if (isPdf) toast('ในหน้าต่างที่เปิดขึ้น ให้เลือกเครื่องพิมพ์เป็น <b>Save as PDF / บันทึกเป็น PDF</b>');
  setTimeout(() => {
    window.print();
    setTimeout(() => { document.title = old; }, 800);
  }, isPdf ? 400 : 0);
}
$('#btnDocx').onclick = () => downloadDocx(S.sess);
$('#btnDocxBoth').onclick = () => { downloadDocx('0750'); setTimeout(() => downloadDocx('0830'), 600); };

function qualified(sess) {
  const thr = Number(S.meta.thr) || 4;
  return roster()
    .map(s => ({ s, t: total(sess, keyOf(s), S.cur) }))
    .filter(x => x.t >= thr)
    .sort((a, b) => a.s.r - b.s.r || a.s.n - b.s.n);
}

function fillWhoSel() {
  ['#whoSel', '#whoSel2'].forEach(sel => {
    const el = $(sel); if (!el) return;
    el.innerHTML = '';
    S.meta.signers.forEach((g, i) => el.add(new Option(g.name, i)));
    el.value = S.meta.who;
    el.onchange = () => {
      S.meta.who = Number(el.value) || 0; save();
      fillWhoSel(); renderSummary();
    };
  });
}

function renderSummary() {
  $('#thr').value = S.meta.thr;
  $('#chkNote').checked = !!S.meta.note;
  fillWhoSel();
  const w = week(S.cur, true);
  const rowsQ = qualified(S.sess);
  $('#sumInfo').textContent =
    `นับสะสมถึงสัปดาห์ที่ ${w.no || '–'} (${w.label}) — มีนักเรียนถึงเกณฑ์ ${rowsQ.length} คน` +
    ` • สัปดาห์ที่บันทึกไว้ทั้งหมด ${weekList().length} สัปดาห์`;

  const t = SESS[S.sess];
  let h = `<div class="ptitle">สรุปการมาสายของนักเรียนหลังเวลา ${t} น. ระดับชั้นมัธยมศึกษาปีที่ ${S.meta.level}<br>` +
    `สัปดาห์ที่ ${w.no || ''}  ระหว่างวันที่ ${w.label}</div>`;
  if (S.meta.note) h += `<div class="pnote">หมายเหตุ : เฉพาะนักเรียนมาสายตั้งแต่ ${S.meta.thr} ครั้งขึ้นไป</div>`;
  h += '<table><tr><th>ที่</th><th>ห้อง</th><th>เลขที่</th><th>ชื่อ-สกุล</th><th>สาย (จำนวนครั้ง)</th><th>หมายเหตุ</th></tr>';
  rowsQ.forEach((x, i) => {
    h += `<tr><td>${i + 1}</td><td>${x.s.r}</td><td>${x.s.n}</td>` +
      `<td style="text-align:left">${x.s.name}</td><td>${x.t}</td><td></td></tr>`;
  });
  if (!rowsQ.length) h += '<tr><td colspan="6" style="color:#888">— ไม่มีนักเรียนถึงเกณฑ์ —</td></tr>';
  h += '</table>';
  h += `<div style="margin-top:18px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงเรียนมาเพื่อโปรดทราบ</div>`;
  h += `<div class="sign">( ${signer().name})<br>${signer().pos}</div>`;
  $('#preview').innerHTML = h;
}

/* ==========================================================================
   หน้า 3 : ยอดยกมา
   ========================================================================== */
$('#carryRoom').onchange = renderCarry;
$('#carryQ').oninput = renderCarry;
$('#btnBaseline').onclick = () => {
  const B = window.BASELINE;
  if (!B) return alert('ไม่พบไฟล์ baseline.js');
  if (!confirm(B.label + '\n\nจะเขียนทับยอดยกมาเดิมของทั้ง 2 ช่วงเวลา ดำเนินการต่อ ?')) return;
  S.carry = baselineCarry(); save(); renderCarry(); renderRec();
  alert('โหลดแล้ว — ' + B.label);
};
$('#btnCarryClear').onclick = () => {
  if (!confirm('ล้างยอดยกมาทั้งหมดของ สาย ' + SESS[S.sess] + ' น. ?')) return;
  S.carry[S.sess] = {}; save(); renderCarry(); renderRec();
};

function renderCarry() {
  const sel = $('#carryRoom');
  if (!sel.options.length) rooms().forEach(r => sel.add(new Option('ห้อง ' + r, r)));
  const q = $('#carryQ').value.trim();
  const list = q ? roster().filter(s => s.name.includes(q))
    : roster().filter(s => String(s.r) === String(sel.value || rooms()[0]));
  let h = '<thead><tr><th>ห้อง</th><th>เลขที่</th><th style="text-align:left">ชื่อ-สกุล</th><th>ยอดยกมา</th><th>ติ๊กในโปรแกรม</th><th>รวม</th></tr></thead><tbody>';
  list.forEach(s => {
    const k = keyOf(s);
    const c = Number(S.carry[S.sess]?.[k] || 0);
    const inApp = total(S.sess, k, S.cur) - c;
    h += `<tr><td>${s.r}</td><td>${s.n}</td><td class="name">${s.name}</td>` +
      `<td><input type="number" min="0" style="width:80px" value="${c || ''}" data-k="${k}"></td>` +
      `<td>${inApp || ''}</td><td class="tot">${(c + inApp) || ''}</td></tr>`;
  });
  h += '</tbody>';
  $('#tblCarry').innerHTML = h;
  $('#tblCarry').querySelectorAll('input').forEach(el => el.onchange = () => {
    const v = Number(el.value) || 0;
    if (!S.carry[S.sess]) S.carry[S.sess] = {};
    if (v) S.carry[S.sess][el.dataset.k] = v; else delete S.carry[S.sess][el.dataset.k];
    save(); renderCarry();
  });
}

/* ==========================================================================
   หน้า 4 : รายชื่อนักเรียน
   ========================================================================== */
$('#btnAdd').onclick = () => {
  const r = Number($('#adRoom').value), n = Number($('#adNo').value), name = $('#adName').value.trim();
  if (!r || !n || !name) return alert('กรอก ห้อง / เลขที่ / ชื่อ-สกุล ให้ครบ');
  mergeRoster([{ r, n, name }]);
  $('#adName').value = ''; $('#adNo').value = n + 1;
  renderRoster(); renderRec();
};
$('#btnMerge').onclick = () => {
  const list = parsePaste($('#paste').value);
  if (!list.length) return alert('ไม่พบข้อมูลที่อ่านได้');
  mergeRoster(list); $('#paste').value = '';
  alert('เพิ่ม/อัปเดต ' + list.length + ' รายการแล้ว'); renderRoster(); renderRec();
};
$('#btnReplace').onclick = () => {
  const list = parsePaste($('#paste').value);
  if (!list.length) return alert('ไม่พบข้อมูลที่อ่านได้');
  if (!confirm('แทนที่รายชื่อทั้งหมดด้วย ' + list.length + ' รายชื่อ ?')) return;
  S.roster = list.sort((a, b) => a.r - b.r || a.n - b.n);
  S.rosterVer = window.ROSTER_VERSION; save();
  $('#paste').value = ''; renderRoster(); renderRec();
};
$('#btnResetRoster').onclick = () => {
  if (!confirm('คืนค่ารายชื่อเริ่มต้นจากไฟล์ students.js ?')) return;
  S.roster = null; S.rosterVer = null; save(); renderRoster(); renderRec();
};
function parsePaste(txt) {
  const out = [];
  txt.split(/\r?\n/).forEach(line => {
    if (!line.trim()) return;
    const p = line.split(/\t|,|\s{2,}/).map(x => x.trim()).filter(x => x !== '');
    if (p.length < 3) return;
    const r = parseInt(p[0], 10), n = parseInt(p[1], 10);
    if (!r || !n) return;
    out.push({ r, n, name: p.slice(2).join(' ').replace(/\s+/g, ' ').trim() });
  });
  return out;
}
function mergeRoster(list) {
  const cur = roster().map(s => ({ ...s }));
  list.forEach(x => {
    const i = cur.findIndex(s => s.r === x.r && s.n === x.n);
    if (i >= 0) cur[i].name = x.name; else cur.push(x);
  });
  S.roster = cur.sort((a, b) => a.r - b.r || a.n - b.n);
  S.rosterVer = window.ROSTER_VERSION; save();
}
function renderRoster() {
  const rs = rooms();
  $('#rosterInfo').innerHTML = `ตอนนี้มี <b>${roster().length}</b> คน / <b>${rs.length}</b> ห้อง (ห้อง ${rs.join(', ')})` +
    (S.roster ? ' — <b>แก้ไขเพิ่มเองแล้ว</b> (เก็บในเครื่องนี้)'
      : ' — รายชื่อชุด <b>' + (window.ROSTER_VERSION || '-') + '</b> จากไฟล์ students.js');
  let h = '<thead><tr><th>ห้อง</th><th>เลขที่</th><th style="text-align:left">ชื่อ-สกุล</th><th></th></tr></thead><tbody>';
  roster().forEach(s => {
    h += `<tr><td>${s.r}</td><td>${s.n}</td><td class="name">${s.name}</td>` +
      `<td><button class="btn" data-r="${s.r}" data-n="${s.n}" style="padding:2px 8px">ลบ</button></td></tr>`;
  });
  h += '</tbody>';
  $('#tblRoster').innerHTML = h;
  $('#tblRoster').querySelectorAll('button').forEach(b => b.onclick = () => {
    const r = Number(b.dataset.r), n = Number(b.dataset.n);
    S.roster = roster().filter(s => !(s.r === r && s.n === n));
    S.rosterVer = window.ROSTER_VERSION; save(); renderRoster(); renderRec();
  });
}

/* ==========================================================================
   หน้า 5 : ข้อมูล / สำรอง
   ========================================================================== */
$('#level').oninput = () => { S.meta.level = $('#level').value; save(); };
[['#sg0name', 0, 'name'], ['#sg0pos', 0, 'pos'], ['#sg1name', 1, 'name'], ['#sg1pos', 1, 'pos']]
  .forEach(([sel, i, f]) => $(sel).oninput = () => {
    S.meta.signers[i][f] = $(sel).value; save(); fillWhoSel();
  });
$('#btnExport').onclick = () => {
  downloadBlob(new Blob([JSON.stringify(S, null, 1)], { type: 'application/json' }),
    'สำรองข้อมูลมาสาย-' + iso(new Date()) + '.json');
};
$('#fileImp').onchange = (e) => {
  const f = e.target.files[0]; if (!f) return;
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const d = JSON.parse(rd.result);
      if (!d.weeks) throw 0;
      if (!confirm('นำเข้าข้อมูลนี้ทับข้อมูลเดิมทั้งหมด ?')) return;
      S = Object.assign(blank(), d); save(); boot();
      alert('นำเข้าเรียบร้อย');
    } catch (err) { alert('ไฟล์ไม่ถูกต้อง'); }
  };
  rd.readAsText(f); e.target.value = '';
};
/* ==========================================================================
   📊 นำเข้าจากไฟล์ Excel (.xlsx/.xls/.csv) — ใช้ SheetJS โหลดจาก CDN เฉพาะตอนกดนำเข้า
   รองรับ 1) ไฟล์ Google Sheet ที่ดาวน์โหลดมา (ชีต students/late/carry/meta)
          2) ไฟล์รายชื่อ เช่น ม.4.xls (ห้อง/เลขที่/ชื่อ)
   ========================================================================== */
function loadSheetJS() {
  if (window.XLSX) return Promise.resolve();
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = () => res(); s.onerror = () => rej(new Error('โหลดตัวอ่าน Excel ไม่ได้ (ต้องต่ออินเทอร์เน็ตตอนนำเข้า)'));
    document.head.appendChild(s);
  });
}
const norm = (x) => String(x ?? '').replace(/\s+/g, ' ').trim();
function normSess(v) {
  let s = norm(v).replace(/[.:]/g, '');
  if (!s) return '';
  if (/^\d+$/.test(s)) s = s.padStart(4, '0');
  return s === '0750' || s === '0830' ? s : (s.includes('830') ? '0830' : (s.includes('750') ? '0750' : ''));
}
function normISO(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getTime() + 12 * 3600 * 1000).toISOString().slice(0, 10); // กันเวลาเพี้ยน ±11 ชม.
  if (typeof v === 'number' && window.XLSX) { const d = XLSX.SSF.parse_date_code(v); if (d) return iso(new Date(d.y, d.m - 1, d.d)); }
  const s = norm(v);
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return iso(new Date(+m[1], m[2] - 1, +m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) { let y = +m[3]; if (y > 2400) y -= 543; return iso(new Date(y, m[2] - 1, +m[1])); }
  return '';
}
/* หาคอลัมน์จากชื่อหัวตาราง (รองรับหลายชื่อ) */
function colIndex(head, names) {
  const H = head.map(h => norm(h).toLowerCase());
  for (const n of names) { const i = H.indexOf(n.toLowerCase()); if (i >= 0) return i; }
  for (const n of names) { const i = H.findIndex(h => h.includes(n.toLowerCase())); if (i >= 0) return i; }
  return -1;
}
function parseWorkbook(wb) {
  const out = { weeks: {}, carry: { '0750': {}, '0830': {} }, roster: [], meta: null, found: [] };
  wb.SheetNames.forEach(name => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, cellDates: false, defval: '' });
    if (!rows.length) return;
    let hi = rows.findIndex(r => r.some(c => norm(c) !== ''));
    if (hi < 0) return;
    const head = rows[hi], body = rows.slice(hi + 1).filter(r => r.some(c => norm(c) !== ''));
    const lname = name.toLowerCase();

    // ---- late : ช่วงเวลา | สัปดาห์ที่ | วันจันทร์ | ช่วงวันที่ | ห้อง | เลขที่ | ชื่อ-สกุล | จำนวนครั้ง
    const cS = colIndex(head, ['ช่วงเวลา']), cD = colIndex(head, ['วันจันทร์']);
    if (lname === 'late' || (cS >= 0 && cD >= 0)) {
      const cNo = colIndex(head, ['สัปดาห์ที่']), cL = colIndex(head, ['ช่วงวันที่']), cR = colIndex(head, ['ห้อง']),
        cN = colIndex(head, ['เลขที่']), cC = colIndex(head, ['จำนวนครั้ง', 'รวมสัปดาห์นี้', 'รวม']);
      let n = 0;
      body.forEach(r => {
        const sess = normSess(r[cS]), isoW = normISO(r[cD]);
        const cnt = Number(r[cC]) || 0, rm = Number(r[cR]), no = Number(r[cN]);
        if (!sess || !isoW || !rm || !no) return;
        const w = out.weeks[isoW] || (out.weeks[isoW] = { no: norm(r[cNo]), label: norm(r[cL]) || rangeLabel(isoW), marks: { '0750': {}, '0830': {} } });
        if (!w.no && norm(r[cNo])) w.no = norm(r[cNo]);
        if (cnt > 0) w.marks[sess][rm + '-' + no] = cnt;
        n++;
      });
      out.found.push(`มาสาย ${n} แถว (${Object.keys(out.weeks).length} สัปดาห์)`);
      return;
    }
    // ---- carry : ช่วงเวลา | ห้อง | เลขที่ | ชื่อ-สกุล | ยอดยกมา
    const cY = colIndex(head, ['ยอดยกมา']);
    if (lname === 'carry' || (cS >= 0 && cY >= 0)) {
      const cR = colIndex(head, ['ห้อง']), cN = colIndex(head, ['เลขที่']);
      let n = 0;
      body.forEach(r => {
        const sess = normSess(r[cS]), cnt = Number(r[cY]) || 0, rm = Number(r[cR]), no = Number(r[cN]);
        if (!sess || !rm || !no || !cnt) return;
        out.carry[sess][rm + '-' + no] = cnt; n++;
      });
      out.found.push(`ยอดยกมา ${n} คน`);
      return;
    }
    // ---- meta : key | value
    if (lname === 'meta') {
      const m = {}; body.forEach(r => { if (norm(r[0])) m[norm(r[0])] = r[1]; });
      out.meta = m; out.found.push('ค่าตั้งค่า'); return;
    }
    // ---- students / รายชื่อ : ห้อง(room) | เลขที่ | ชื่อ-สกุล | เลขประจำตัว
    const cR = colIndex(head, ['ห้อง', 'room']), cN = colIndex(head, ['เลขที่', 'no']),
      cNm = colIndex(head, ['ชื่อ-สกุล', 'ชื่อ นามสกุล', 'ชื่อ', 'name']), cId = colIndex(head, ['เลขประจำตัว', 'รหัสประจำตัว', 'รหัส', 'id']);
    if (cR >= 0 && cN >= 0 && cNm >= 0) {
      let n = 0;
      body.forEach(r => {
        const rm = Number(r[cR]), no = Number(r[cN]), nm = norm(r[cNm]);
        if (!rm || !no || !nm) return;
        out.roster.push({ r: rm, n: no, id: cId >= 0 ? norm(r[cId]).replace(/\.0$/, '') : '', name: nm }); n++;
      });
      if (n) out.found.push(`รายชื่อ ${n} คน`);
    }
  });
  return out;
}
async function importExcelBuffer(buf, fname) {
  await loadSheetJS();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false });
  const p = parseWorkbook(wb);
  const hasLate = Object.keys(p.weeks).length > 0, hasCarry = Object.keys(p.carry['0750']).length + Object.keys(p.carry['0830']).length > 0;
  if (!hasLate && !hasCarry && !p.roster.length) throw new Error('ไม่พบข้อมูลที่รู้จักในไฟล์ (ต้องมีชีต late / carry / students หรือคอลัมน์ ห้อง-เลขที่-ชื่อ)');
  const msgs = [];
  if (hasLate || hasCarry) msgs.push('• ข้อมูลมาสาย + ยอดยกมา จะ "แทน" ของเดิมในเครื่องทั้งหมด');
  if (p.roster.length) msgs.push('• รายชื่อนักเรียน ' + p.roster.length + ' คน จะ "แทน" รายชื่อเดิม');
  if (!confirm(`ไฟล์: ${fname}\nพบ: ${p.found.join(', ')}\n\n${msgs.join('\n')}\n\nดำเนินการต่อ ?`)) return null;

  if (hasLate || hasCarry) {
    S.weeks = hasLate ? p.weeks : S.weeks;
    S.carry = hasCarry ? p.carry : S.carry;
    S.pending = { late: {}, carry: {} };            // ข้อมูลมาจากชีตแล้ว ไม่ต้องส่งกลับ
    const ws = weekList(); if (ws.length) S.cur = ws[ws.length - 1];
  }
  if (p.roster.length) {
    p.roster.sort((a, b) => a.r - b.r || a.n - b.n);
    S.roster = p.roster; S.rosterVer = window.ROSTER_VERSION;
  }
  if (p.meta) {
    try { const sg = JSON.parse(p.meta['signers'] || '[]'); if (sg.length >= 2) S.meta.signers = sg; } catch (e) { }
    if (p.meta['thr']) S.meta.thr = Number(p.meta['thr']);
    if (p.meta['level']) S.meta.level = String(p.meta['level']);
  }
  localStorage.setItem(KEY, JSON.stringify(S));
  ready = false; boot(); renderData(); ready = true;
  return p;
}
$('#fileXlsx').onchange = (e) => {
  const f = e.target.files[0]; if (!f) return;
  $('#xlsxInfo').textContent = 'กำลังอ่าน ' + f.name + '…';
  const rd = new FileReader();
  rd.onload = async () => {
    try {
      const p = await importExcelBuffer(rd.result, f.name);
      if (!p) { $('#xlsxInfo').textContent = 'ยกเลิก'; return; }
      $('#xlsxInfo').innerHTML = '<b style="color:#1b7f3b">นำเข้าแล้ว:</b> ' + p.found.join(', ');
      toast('นำเข้าจาก Excel เรียบร้อย — ' + p.found.join(', '));
    } catch (err) {
      $('#xlsxInfo').innerHTML = '<b style="color:#c62828">นำเข้าไม่สำเร็จ:</b> ' + err.message;
      alert('นำเข้าไม่สำเร็จ\n' + err.message);
    }
  };
  rd.readAsArrayBuffer(f); e.target.value = '';
};

$('#btnWipe').onclick = () => {
  if (!confirm('ล้างข้อมูลทั้งหมด (รวมทุกสัปดาห์) ?')) return;
  if (!confirm('ยืนยันอีกครั้ง — ข้อมูลจะหายถาวร')) return;
  localStorage.removeItem(KEY); S = blank(); boot();
};
function renderData() {
  $('#sg0name').value = S.meta.signers[0].name; $('#sg0pos').value = S.meta.signers[0].pos;
  $('#sg1name').value = S.meta.signers[1].name; $('#sg1pos').value = S.meta.signers[1].pos;
  $('#level').value = S.meta.level;
  const tot = allSummary();
  const w = week(S.cur, true);
  $('#allInfo').innerHTML = `สัปดาห์ที่ ${w.no || '–'} (${w.label}) — สาย <b>${tot.ppl}</b> คน รวม <b>${tot.times}</b> ครั้ง` +
    (S.cloud.at ? ` • ซิงก์ล่าสุด ${S.cloud.at}` : ' • ยังไม่เคยซิงก์');
  $('#gsUrl').value = S.cloud.url || '';
  $('#chkAuto').checked = !!S.cloud.auto;
  if (S.cloud.at) setCloudInfo('ซิงก์ล่าสุด ' + S.cloud.at + ' • rev ' + S.cloud.rev, '☁ พร้อม');
  fillWhoSel();
  let h = '<thead><tr><th>สัปดาห์ที่</th><th>ช่วงวันที่</th><th>สาย 07.50</th><th>สาย 08.30</th><th></th></tr></thead><tbody>';
  weekList().forEach(k => {
    const w = S.weeks[k];
    const c1 = Object.keys(w.marks['0750'] || {}).length, c2 = Object.keys(w.marks['0830'] || {}).length;
    h += `<tr${k === S.cur ? ' class="hit"' : ''}><td>${w.no || '–'}</td><td>${w.label}</td><td>${c1} คน</td><td>${c2} คน</td>` +
      `<td><button class="btn go" data-k="${k}" style="padding:2px 8px">เปิด</button> ` +
      `<button class="btn warn del" data-k="${k}" style="padding:2px 8px">ลบ</button></td></tr>`;
  });
  if (!weekList().length) h += '<tr><td colspan="5" style="color:#888">ยังไม่มีข้อมูล</td></tr>';
  h += '</tbody>';
  $('#tblWeeks').innerHTML = h;
  $('#tblWeeks').querySelectorAll('.go').forEach(b => b.onclick = () => {
    S.cur = b.dataset.k; save(); $$('nav button')[0].click(); renderRec();
  });
  $('#tblWeeks').querySelectorAll('.del').forEach(b => b.onclick = () => {
    if (!confirm('ลบข้อมูลสัปดาห์นี้ ?')) return;
    delete S.weeks[b.dataset.k];
    if (S.cur === b.dataset.k) S.cur = weekList().pop() || mondayISO(new Date());
    save(); renderData(); renderRec();
  });
}

/* ==========================================================================
   ☁ ซิงก์กับ Google Sheet (ผ่าน Apps Script /exec)
   - อ่าน  : GET  ?action=load   (ถ้าเบราว์เซอร์บล็อก จะสลับไปใช้ JSONP อัตโนมัติ)
   - เขียน : POST {action:'save', data, baseRev}
   ========================================================================== */
let syncTimer = null, syncing = false, ready = false;

function cloudUrl() { return (S.cloud && S.cloud.url || '').trim(); }

function setCloudInfo(txt, cls) {
  const el = $('#gsInfo'); if (el) el.innerHTML = txt;
  const b = $('#hdCloud'); if (b) b.textContent = cls || txt.replace(/<[^>]*>/g, '').slice(0, 40);
}

/* อ่านข้อมูลด้วย JSONP — ใช้ได้เสมอแม้เบราว์เซอร์จะบล็อก CORS */
function jsonp(url, timeout = 20000) {
  return new Promise((resolve, reject) => {
    const cb = 'gscb_' + Math.floor(performance.now() * 1000) + '_' + Math.floor(Math.random() * 1e6);
    const sc = document.createElement('script');
    const done = (fn, arg) => { delete window[cb]; sc.remove(); clearTimeout(tm); fn(arg); };
    const tm = setTimeout(() => done(reject, new Error('หมดเวลารอ')), timeout);
    window[cb] = (data) => done(resolve, data);
    sc.onerror = () => done(reject, new Error('เรียกใช้งานลิงก์ไม่สำเร็จ'));
    sc.src = url + (url.includes('?') ? '&' : '?') + 'callback=' + cb;
    document.body.appendChild(sc);
  });
}

async function cloudGet(action) {
  const u = cloudUrl();
  if (!u) throw new Error('ยังไม่ได้ใส่ลิงก์ Apps Script');
  const url = u + (u.includes('?') ? '&' : '?') + 'action=' + action + '&t=' + Math.floor(performance.now());
  try {
    const r = await fetch(url, { redirect: 'follow' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return await r.json();
  } catch (e) {
    return await jsonp(url);      // สำรอง
  }
}

async function cloudPost(payload) {
  const u = cloudUrl();
  if (!u) throw new Error('ยังไม่ได้ใส่ลิงก์ Apps Script');
  try {
    const r = await fetch(u, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // ไม่ให้เกิด preflight
      body: JSON.stringify(payload)
    });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return await r.json();
  } catch (e) {
    // สำรอง: ส่งแบบอ่านคำตอบไม่ได้ แล้วไปเช็คผลด้วย ping แทน
    await fetch(u, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });
    const p = await cloudGet('ping').catch(() => null);
    if (p && p.ok) return { ok: true, rev: p.rev, blind: true };
    throw e;
  }
}

/* ---------- สร้างแถวที่จะส่งขึ้นชีตจากรายการที่ค้าง (pending) ---------- */
function nameOfKey(k) {
  const s = roster().find(x => keyOf(x) === k);
  return s ? s.name : '';
}
function buildUpsert(roomFilter) {
  const rows = [], carry = [];
  const only = (roomFilter !== null && roomFilter !== undefined) ? String(roomFilter) + '-' : null;
  Object.keys(S.pending.late).forEach(pk => {
    const [sess, isoW, k] = pk.split('|');
    if (only && !k.startsWith(only)) return;
    const w = S.weeks[isoW] || {};
    const [r, n] = k.split('-').map(Number);
    rows.push({ sess, iso: isoW, no: w.no || '', label: w.label || rangeLabel(isoW), r, n,
      name: nameOfKey(k), count: weekCount(isoW, sess, k), _pk: pk });
  });
  Object.keys(S.pending.carry).forEach(pk => {
    const [sess, k] = pk.split('|');
    if (only && !k.startsWith(only)) return;
    const [r, n] = k.split('-').map(Number);
    carry.push({ sess, r, n, name: nameOfKey(k), count: Number(S.carry[sess]?.[k] || 0), _pk: pk });
  });
  return { rows, carry };
}
function pendingCount() { return Object.keys(S.pending.late).length + Object.keys(S.pending.carry).length; }

let needResync = false;

/* ⬆ ส่งรายการที่ค้างขึ้น Google Sheet (อัปเดตทีละแถว ไม่ลบของเดิมในชีต)
   roomFilter = เลขห้อง -> ส่งเฉพาะห้องนั้น, null -> ส่งทุกห้อง */
async function cloudPush(roomFilter, silent) {
  if (typeof roomFilter === 'boolean') roomFilter = null;
  if (syncing) { needResync = true; return false; }
  const { rows, carry } = buildUpsert(roomFilter);
  if (!rows.length && !carry.length) {
    setCloudInfo('ไม่มีรายการค้างส่ง • ซิงก์ล่าสุด ' + (S.cloud.at || '-'), '☁ ตรงกัน');
    return true;
  }
  syncing = true;
  $('#hdSync').disabled = true;
  setCloudInfo(`กำลังส่ง ${rows.length + carry.length} รายการขึ้น Google Sheet…`, '☁ กำลังส่ง…');
  try {
    const payload = {
      action: 'upsert',
      rows: rows.map(({ _pk, ...x }) => x),
      carry: carry.map(({ _pk, ...x }) => x),
      meta: S.meta,
      roster: S.cloud.rev ? [] : roster()          // ครั้งแรกสุดส่งรายชื่อไปตั้งต้นชีต students
    };
    const res = await cloudPost(payload);
    if (!res || !res.ok) throw new Error((res && res.error) || 'บันทึกไม่สำเร็จ');
    rows.forEach(x => delete S.pending.late[x._pk]);
    carry.forEach(x => delete S.pending.carry[x._pk]);
    S.cloud.rev = res.rev || S.cloud.rev || 1;
    S.cloud.at = new Date().toLocaleString('th-TH', { hour12: false });
    localStorage.setItem(KEY, JSON.stringify(S));
    const left = pendingCount();
    setCloudInfo(`ขึ้น Google Sheet แล้ว (${rows.length + carry.length} รายการ) • ${S.cloud.at}` +
      (left ? ` • ยังค้างอีก ${left} รายการ (ห้องอื่น)` : ''), left ? `☁ ค้าง ${left}` : '☁ บันทึกแล้ว');
    return true;
  } catch (e) {
    setCloudInfo('<b style="color:#c62828">ส่งขึ้นชีตไม่สำเร็จ:</b> ' + e.message +
      ` — ข้อมูลยังอยู่ในเครื่อง (ค้างส่ง ${pendingCount()} รายการ) จะลองส่งใหม่เมื่อบันทึกครั้งถัดไป`, '☁ ค้าง ' + pendingCount());
    if (!silent) alert('ส่งขึ้น Google Sheet ไม่สำเร็จ\n' + e.message + '\n\nข้อมูลยังอยู่ในเครื่องครบ');
    return false;
  } finally {
    syncing = false; $('#hdSync').disabled = false;
    if (needResync) { needResync = false; setTimeout(() => cloudPush(null, true), 300); }
  }
}

async function cloudPull() {
  try {
    setCloudInfo('กำลังดึงข้อมูล…', '☁ กำลังดึง…');
    const res = await cloudGet('load');
    if (!res || !res.ok) throw new Error((res && res.error) || 'ดึงข้อมูลไม่สำเร็จ');
    const d = res.data || {};
    const nWeeks = Object.keys(d.weeks || {}).length;
    const pend = pendingCount();
    if (!confirm(`ข้อมูลบน Google Sheet: ${nWeeks} สัปดาห์, ยอดยกมา ${Object.keys((d.carry || {})['0750'] || {}).length} คน, รายชื่อ ${(d.roster || []).length} คน\n\n` +
      'จะนำมาแทนข้อมูลในเครื่องนี้ทั้งหมด (ใช้ตอนเปิดจากเครื่องใหม่ หรือหลังแก้ข้อมูลในชีต) ดำเนินการต่อ ?' +
      (pend ? `\n\n⚠ มี ${pend} รายการในเครื่องที่ยังไม่ได้ส่งขึ้นชีต จะหายไป — ถ้าไม่แน่ใจให้กด "บันทึกทุกห้อง" ก่อน` : '')))
      { setCloudInfo('ยกเลิก', '☁ พร้อม'); return; }
    if (d.weeks) S.weeks = d.weeks;
    if (d.carry) S.carry = Object.assign({ '0750': {}, '0830': {} }, d.carry);
    if (d.roster && d.roster.length) { S.roster = d.roster; S.rosterVer = window.ROSTER_VERSION; }
    if (d.meta) {
      if (d.meta.signers && d.meta.signers.length) S.meta.signers = d.meta.signers;
      if (d.meta.who !== null && d.meta.who !== undefined) S.meta.who = Number(d.meta.who);
      if (d.meta.level) S.meta.level = d.meta.level;
      if (d.meta.thr) S.meta.thr = Number(d.meta.thr);
      if (d.meta.note !== null && d.meta.note !== undefined) S.meta.note = !!d.meta.note;
    }
    S.pending = { late: {}, carry: {} };
    S.cloud.rev = d.rev || 1; S.cloud.at = new Date().toLocaleString('th-TH', { hour12: false });
    const ws = weekList(); if (ws.length && !S.weeks[S.cur]) S.cur = ws[ws.length - 1];
    localStorage.setItem(KEY, JSON.stringify(S));
    ready = false; boot(); renderData(); ready = true;
    setCloudInfo(`ดึงข้อมูลจากชีตแล้ว • ${nWeeks} สัปดาห์`, '☁ ตรงกัน');
    toast('ดึงข้อมูลจาก Google Sheet มาแทนในเครื่องแล้ว');
  } catch (e) {
    setCloudInfo('<b style="color:#c62828">ดึงข้อมูลไม่สำเร็จ:</b> ' + e.message, '☁ ผิดพลาด');
    alert('ดึงจาก Google Sheet ไม่สำเร็จ\n' + e.message);
  }
}

/* ซิงก์อัตโนมัติ: แก้ตัวเลขปุ๊บ ส่งขึ้นชีตภายใน 1.5 วินาที */
function autoSync() {
  if (!ready) return;
  const n = pendingCount();
  if (!n) return;
  if (!S.cloud || !S.cloud.auto || !cloudUrl()) { setCloudInfo(`ค้างส่ง ${n} รายการ (ปิดซิงก์อัตโนมัติ)`, '☁ ค้าง ' + n); return; }
  clearTimeout(syncTimer);
  setCloudInfo(`มีการแก้ไข ${n} รายการ — กำลังส่งขึ้น Google Sheet…`, '☁ รอส่ง ' + n);
  syncTimer = setTimeout(() => cloudPush(null, true), 1500);
}

/* สรุปยอดของห้องหนึ่งในสัปดาห์ที่กำลังบันทึก */
function roomSummary(r) {
  const list = roster().filter(s => s.r === r);
  let ppl = 0, times = 0;
  list.forEach(s => { const c = weekCount(S.cur, S.sess, keyOf(s)); if (c) { ppl++; times += c; } });
  return { ppl, times };
}
function allSummary() {
  let ppl = 0, times = 0;
  ['0750', '0830'].forEach(se => {
    const m = (S.weeks[S.cur] || {}).marks ? S.weeks[S.cur].marks[se] || {} : {};
    Object.keys(m).forEach(k => { ppl++; times += weekCount(S.cur, se, k); });
  });
  return { ppl, times };
}

/* 💾 ปุ่มบันทึก — เซฟลงเครื่อง แล้วส่งรายการค้างขึ้น Google Sheet ทันที */
async function saveAndPush(roomFilter, head) {
  clearTimeout(syncTimer);
  localStorage.setItem(KEY, JSON.stringify(S));
  if (!cloudUrl()) { toast(head + ' (เก็บในเครื่อง — ยังไม่ได้ตั้งลิงก์ Google Sheet)', true); return; }
  const n = buildUpsert(roomFilter); const cnt = n.rows.length + n.carry.length;
  if (!cnt) { toast(head + ' • ข้อมูลตรงกับ Google Sheet อยู่แล้ว ✓'); return; }
  toast(head + ` • กำลังส่ง ${cnt} รายการขึ้น Google Sheet…`);
  const ok = await cloudPush(roomFilter, true);
  toast(ok ? head + ` • ขึ้น Google Sheet แล้ว ${cnt} รายการ ✓`
           : head + ' • <b>ส่งขึ้น Google Sheet ไม่สำเร็จ</b> (ข้อมูลยังอยู่ในเครื่อง จะส่งใหม่อัตโนมัติ)', !ok);
  if (!$('#tab-data').hidden) renderData();
  if ($('#tab-dash') && !$('#tab-dash').hidden) renderDash();
}
$('#btnSaveRoom').onclick = () => {
  const s = roomSummary(S.room);
  saveAndPush(S.room, `บันทึกห้อง ${S.room} — สาย <b>${s.ppl}</b> คน รวม <b>${s.times}</b> ครั้ง`);
};
$('#btnSaveAllRec').onclick = $('#btnSaveAll').onclick = () => {
  const tot = allSummary();
  saveAndPush(null, `บันทึกทุกห้อง — สัปดาห์นี้ สาย <b>${tot.ppl}</b> คน รวม <b>${tot.times}</b> ครั้ง`);
};

$('#hdSync').onclick = () => saveAndPush(null, 'ซิงก์');
/* ส่งข้อมูลทั้งหมดในเครื่องขึ้นชีตอีกรอบ (ใช้ซ่อมกรณีชีตขาดข้อมูล) — ไม่ลบอะไรในชีต */
$('#btnPush').onclick = () => {
  if (!confirm('จะส่งข้อมูลมาสายทุกสัปดาห์ + ยอดยกมาทั้งหมดในเครื่องนี้ขึ้น Google Sheet อีกครั้ง\n(แถวที่มีอยู่จะถูกอัปเดต แถวที่ไม่มีจะถูกเพิ่ม ไม่ลบอะไร) ดำเนินการต่อ ?')) return;
  Object.keys(S.weeks).forEach(isoW => ['0750', '0830'].forEach(se =>
    Object.keys((S.weeks[isoW].marks || {})[se] || {}).forEach(k => S.pending.late[se + '|' + isoW + '|' + k] = true)));
  ['0750', '0830'].forEach(se => Object.keys(S.carry[se] || {}).forEach(k => S.pending.carry[se + '|' + k] = true));
  saveAndPush(null, 'ส่งข้อมูลทั้งหมด');
};
$('#btnPull').onclick = () => cloudPull();
$('#btnPushRoster').onclick = async () => {
  if (!confirm('ส่งรายชื่อ ' + roster().length + ' คน ขึ้นชีต students (ทับของเดิม) ?')) return;
  try {
    const res = await cloudPost({ action: 'saveRoster', roster: roster() });
    if (!res || !res.ok) throw new Error((res && res.error) || 'ส่งไม่สำเร็จ');
    setCloudInfo('ส่งรายชื่อขึ้นชีตแล้ว ' + (res.students || roster().length) + ' คน', '☁ พร้อม');
    alert('ส่งรายชื่อขึ้นชีตแล้ว');
  } catch (e) { alert('ส่งรายชื่อไม่สำเร็จ\n' + e.message); }
};
$('#btnPing').onclick = async () => {
  try {
    setCloudInfo('กำลังทดสอบ…', '☁ ทดสอบ…');
    const r = await cloudGet('ping');
    if (!r || !r.ok) throw new Error((r && r.error) || 'ไม่ตอบกลับ');
    setCloudInfo('เชื่อมต่อได้ • rev บนคลาวด์ = ' + r.rev, '☁ พร้อม');
    alert('เชื่อมต่อ Google Sheet ได้ปกติ\nrev = ' + r.rev);
  } catch (e) {
    setCloudInfo('<b style="color:#c62828">เชื่อมต่อไม่ได้:</b> ' + e.message, '☁ ผิดพลาด');
    alert('เชื่อมต่อไม่ได้\n' + e.message +
      '\n\nตรวจสอบ: Deploy → Who has access = Anyone, และลิงก์ต้องลงท้าย /exec');
  }
};
$('#gsUrl').oninput = () => { S.cloud.url = $('#gsUrl').value.trim(); localStorage.setItem(KEY, JSON.stringify(S)); };
$('#chkAuto').onchange = () => { S.cloud.auto = $('#chkAuto').checked; localStorage.setItem(KEY, JSON.stringify(S)); };

/* ==========================================================================
   ③ แดชบอร์ด — สรุปรายห้อง + แผนภูมิแท่ง (SVG วาดเอง ไม่ใช้ไลบรารี)
   ========================================================================== */
let dashSess = '0750', dashScope = 'week';
$$('#segSessDash button').forEach(b => b.onclick = () => {
  dashSess = b.dataset.s;
  $$('#segSessDash button').forEach(x => x.classList.toggle('on', x === b));
  renderDash();
});
$$('#segScope button').forEach(b => b.onclick = () => {
  dashScope = b.dataset.scope;
  $$('#segScope button').forEach(x => x.classList.toggle('on', x === b));
  renderDash();
});
$('#btnDashPrint').onclick = () => {
  const old = document.title;
  const w = week(S.cur, true);
  document.title = `แดชบอร์ดมาสาย ม.4 ${dashScope === 'week' ? 'สัปดาห์ที่ ' + (w.no || '') : 'สะสมทั้งเทอม'}`;
  setTimeout(() => { window.print(); setTimeout(() => { document.title = old; }, 800); }, 100);
};

/* ยอดของนักเรียน 1 คน ตามช่วงเวลา/ขอบเขตที่เลือกในแดชบอร์ด */
function dashVal(k, sess) {
  const one = (se) => dashScope === 'week' ? weekCount(S.cur, se, k) : total(se, k, S.cur);
  return sess === 'both' ? one('0750') + one('0830') : one(sess);
}
function dashStats() {
  const thr = Number(S.meta.thr) || 4;
  const out = rooms().map(r => {
    const list = roster().filter(s => s.r === r);
    let times = 0, ppl = 0, over = 0;
    list.forEach(s => {
      const v = dashVal(keyOf(s), dashSess);
      if (v) { times += v; ppl++; }
      const cum = dashSess === 'both'
        ? total('0750', keyOf(s), S.cur) + total('0830', keyOf(s), S.cur)
        : total(dashSess, keyOf(s), S.cur);
      if (cum >= thr) over++;
    });
    return { r, times, ppl, over, size: list.length };
  });
  return out;
}

function renderDash() {
  const w = week(S.cur, true);
  const st = dashStats();
  const sessTxt = dashSess === 'both' ? 'รวมสาย 07.50 น. และ 08.30 น.' : 'สาย ' + SESS[dashSess] + ' น.';
  $('#dashTitle').textContent = (dashScope === 'week'
    ? `สัปดาห์ที่ ${w.no || '–'} (${w.label})` : `สะสมทั้งเทอม ถึงสัปดาห์ที่ ${w.no || '–'}`) + ' • ' + sessTxt;

  const totalTimes = st.reduce((a, x) => a + x.times, 0);
  const totalPpl = st.reduce((a, x) => a + x.ppl, 0);
  const totalOver = st.reduce((a, x) => a + x.over, 0);
  const max = Math.max(0, ...st.map(x => x.times));
  const tops = st.filter(x => x.times === max && max > 0);
  const zeros = st.filter(x => x.times === 0);

  /* ---- การ์ดตัวเลข ---- */
  $('#dashCards').innerHTML =
    `<div class="stat blue"><div class="l">มาสายรวมทั้งระดับ</div><div class="v">${totalTimes}</div><div class="s">ครั้ง • ${totalPpl} คน</div></div>` +
    `<div class="stat red"><div class="l">ห้องที่สายมากที่สุด</div><div class="v">${tops.length ? tops.map(x => 'ห้อง ' + x.r).join(', ') : '–'}</div><div class="s">${max ? max + ' ครั้ง' : 'ไม่มีข้อมูล'}</div></div>` +
    `<div class="stat green"><div class="l">ห้องที่ไม่มีใครสายเลย</div><div class="v">${zeros.length}</div><div class="s">${zeros.length ? 'ห้อง ' + zeros.map(x => x.r).join(', ') : 'ทุกห้องมีคนสาย'}</div></div>` +
    `<div class="stat"><div class="l">ถึงเกณฑ์ ${S.meta.thr || 4} ครั้งขึ้นไป (สะสม)</div><div class="v">${totalOver}</div><div class="s">คน</div></div>`;

  /* ---- แผนภูมิแท่ง (SVG) ---- */
  const n = st.length, W = 1000, H = 360, padL = 46, padR = 16, padT = 34, padB = 44;
  const cw = (W - padL - padR) / n, bw = Math.min(cw * 0.62, 60);
  const ymax = Math.max(4, Math.ceil(max * 1.15));
  const y = v => padT + (H - padT - padB) * (1 - v / ymax);
  let svg = `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="แผนภูมิแท่งจำนวนครั้งมาสายรายห้อง">`;
  const step = ymax <= 10 ? 2 : ymax <= 30 ? 5 : ymax <= 60 ? 10 : 20;
  for (let v = 0; v <= ymax; v += step) {
    svg += `<line x1="${padL}" x2="${W - padR}" y1="${y(v)}" y2="${y(v)}" stroke="#e5e9f2" stroke-width="1"/>` +
      `<text x="${padL - 8}" y="${y(v) + 5}" text-anchor="end" font-size="14" fill="#6b7280">${v}</text>`;
  }
  st.forEach((x, i) => {
    const cx = padL + cw * i + cw / 2;
    const h = x.times ? Math.max(3, y(0) - y(x.times)) : 0;
    const col = x.times === 0 ? '#22a55b' : (x.times === max ? '#d32f2f' : '#2b59c3');
    svg += `<rect x="${cx - bw / 2}" y="${y(0) - h}" width="${bw}" height="${h}" rx="6" fill="${col}"/>`;
    if (x.times === 0)
      svg += `<text x="${cx}" y="${y(0) - 10}" text-anchor="middle" font-size="22" fill="#22a55b">✓</text>`;
    else
      svg += `<text x="${cx}" y="${y(x.times) - 8}" text-anchor="middle" font-size="17" font-weight="700" fill="${col}">${x.times}</text>` +
        `<text x="${cx}" y="${y(x.times) - 8 + 0}" dy="-18" text-anchor="middle" font-size="12" fill="#6b7280">(${x.ppl} คน)</text>`;
    svg += `<text x="${cx}" y="${H - padB + 22}" text-anchor="middle" font-size="16" font-weight="${x.times === max && max ? 700 : 400}" fill="#22303f">ห้อง ${x.r}</text>`;
  });
  svg += `<line x1="${padL}" x2="${W - padR}" y1="${y(0)}" y2="${y(0)}" stroke="#9aa3b2" stroke-width="1.5"/></svg>`;
  $('#dashChart').innerHTML = svg;

  /* ---- อันดับ ---- */
  const ranked = [...st].filter(x => x.times > 0).sort((a, b) => b.times - a.times || b.ppl - a.ppl);
  const medal = ['🥇', '🥈', '🥉'];
  let rank = '<h3 style="margin:0 0 8px">ห้องที่สายมากที่สุด</h3>';
  rank += ranked.length
    ? '<ul class="rank">' + ranked.slice(0, 5).map((x, i) =>
      `<li class="${i === 0 ? 'top' : ''}"><span class="medal">${medal[i] || (i + 1) + '.'}</span>ห้อง ${x.r}<span class="hint" style="margin:0">${x.ppl} คน</span><span class="n">${x.times} ครั้ง</span></li>`).join('') + '</ul>'
    : '<p class="hint">ยังไม่มีข้อมูลมาสายในช่วงที่เลือก</p>';
  rank += '<h3 style="margin:14px 0 8px">🌟 ห้องที่ไม่มีใครสายเลย</h3>';
  rank += zeros.length
    ? '<div>' + zeros.map(x => `<span class="pill">ห้อง ${x.r}</span>`).join('') + '</div>'
    : '<p class="hint">ทุกห้องมีคนสายอย่างน้อย 1 คน</p>';
  $('#dashRank').innerHTML = rank;

  /* ---- ตาราง ---- */
  let h = '<thead><tr><th>ห้อง</th><th>นักเรียน</th><th>คนที่สาย</th><th>จำนวนครั้ง</th><th>ถึงเกณฑ์ (สะสม)</th><th style="text-align:left">สถานะ</th></tr></thead><tbody>';
  st.forEach(x => {
    const cls = x.times === 0 ? '' : (x.times === max ? 'over' : '');
    const status = x.times === 0 ? '<span style="color:#1b7f3b;font-weight:700">✓ ไม่มีคนสาย</span>'
      : (x.times === max ? '<span style="color:#c62828;font-weight:700">สายมากที่สุด</span>' : '');
    h += `<tr class="${cls}"><td>${x.r}</td><td>${x.size}</td><td>${x.ppl || ''}</td><td class="tot ${x.times === max && max ? 'hi' : ''}">${x.times || ''}</td><td>${x.over || ''}</td><td style="text-align:left">${status}</td></tr>`;
  });
  h += `<tr style="font-weight:700;background:#f4f6fb"><td>รวม</td><td>${st.reduce((a, x) => a + x.size, 0)}</td><td>${totalPpl}</td><td>${totalTimes}</td><td>${totalOver}</td><td></td></tr></tbody>`;
  $('#dashTable').innerHTML = h;
}

/* ==========================================================================
   สร้างไฟล์ Word (.docx) — เขียน OOXML + ZIP เองทั้งหมด ไม่ต้องพึ่งไลบรารีนอก
   ========================================================================== */
const CRC_T = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(u8) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
const enc = new TextEncoder();
function cat(arrs) {
  let n = 0; arrs.forEach(a => n += a.length);
  const out = new Uint8Array(n); let o = 0; arrs.forEach(a => { out.set(a, o); o += a.length; });
  return out;
}
function zipStore(files) {
  const d = new Date();
  const time = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF;
  const date = (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;
  const locals = [], centrals = []; let off = 0;
  files.forEach(f => {
    const nb = enc.encode(f.name), data = enc.encode(f.xml), crc = crc32(data);
    const lh = new Uint8Array(30 + nb.length), lv = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, 0, true); lv.setUint16(10, time, true); lv.setUint16(12, date, true);
    lv.setUint32(14, crc, true); lv.setUint32(18, data.length, true); lv.setUint32(22, data.length, true);
    lv.setUint16(26, nb.length, true); lh.set(nb, 30);
    locals.push(lh, data);
    const ch = new Uint8Array(46 + nb.length), cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true); cv.setUint16(10, 0, true); cv.setUint16(12, time, true);
    cv.setUint16(14, date, true); cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true); cv.setUint16(28, nb.length, true); cv.setUint32(42, off, true);
    ch.set(nb, 46); centrals.push(ch);
    off += lh.length + data.length;
  });
  const cd = cat(centrals);
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true);
  ev.setUint32(12, cd.length, true); ev.setUint32(16, off, true);
  return cat([...locals, cd, end]);
}

const X = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function run(text, o = {}) {
  const rpr = '<w:rPr>' + (o.b ? '<w:b/><w:bCs/>' : '') +
    (o.sz ? `<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>` : '') +
    '<w:cs/><w:lang w:bidi="th-TH"/></w:rPr>';
  return `<w:r>${rpr}<w:t xml:space="preserve">${X(text)}</w:t></w:r>`;
}
function para(text, o = {}) {
  const ppr = '<w:pPr><w:spacing w:after="' + (o.after ?? 0) + '"' +
    (o.line === false ? '' : ' w:line="240" w:lineRule="auto"') + '/>' +
    (o.align ? `<w:jc w:val="${o.align}"/>` : '') + '</w:pPr>';
  let body = '';
  for (let i = 0; i < (o.tabs || 0); i++) body += '<w:r><w:tab/></w:r>';
  if (text) body += run(text, o);
  return `<w:p>${ppr}${body}</w:p>`;
}
const COLW = [604, 759, 701, 3003, 1737, 2547];
function cell(w, text, left) {
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>` +
    `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>${left ? '' : '<w:jc w:val="center"/>'}</w:pPr>` +
    (text === '' ? '' : run(text)) + '</w:p></w:tc>';
}
function trow(cells, header) {
  return '<w:tr><w:trPr><w:trHeight w:val="581"/>' + (header ? '<w:tblHeader/>' : '') + '</w:trPr>' +
    cells.map((t, i) => cell(COLW[i], t, i === 3 && !header)).join('') + '</w:tr>';
}

function docxXml(sess) {
  const w = week(S.cur, true);
  const rowsQ = qualified(sess);
  const t = SESS[sess];
  let body = '';
  body += para(`สรุปการมาสายของนักเรียนหลังเวลา ${t} น. ระดับชั้นมัธยมศึกษาปีที่ ${S.meta.level}`,
    { b: true, sz: 36, align: 'center', line: false });
  body += para(`สัปดาห์ที่ ${w.no || ''}  ระหว่างวันที่ ${w.label}`,
    { b: true, sz: 36, align: 'center', line: false });
  body += para('', { after: 120 });
  if (S.meta.note) body += para(`หมายเหตุ : เฉพาะนักเรียนมาสายตั้งแต่ ${S.meta.thr} ครั้งขึ้นไป`, { after: 60 });

  const borders = '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map(b => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
  let tbl = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9351" w:type="dxa"/>` +
    `<w:jc w:val="center"/>${borders}<w:tblLook w:val="04A0"/></w:tblPr>` +
    '<w:tblGrid>' + COLW.map(c => `<w:gridCol w:w="${c}"/>`).join('') + '</w:tblGrid>';
  tbl += trow(['ที่', 'ห้อง', 'เลขที่', 'ชื่อ-สกุล', 'สาย (จำนวนครั้ง)', 'หมายเหตุ'], true);
  rowsQ.forEach((x, i) => tbl += trow([String(i + 1), String(x.s.r), String(x.s.n), x.s.name, String(x.t), '']));
  tbl += '</w:tbl>';
  body += tbl;

  body += para('');
  body += para('จึงเรียนมาเพื่อโปรดทราบ', { tabs: 1 });
  body += para('');
  body += para(`           ( ${signer().name})`, { tabs: 8 });
  body += para(`           ${signer().pos}`, { tabs: 7 });
  body += '<w:sectPr><w:pgSz w:w="11909" w:h="16834"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="431" w:footer="431" w:gutter="0"/>' +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="435"/></w:sectPr>';

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body>' + body + '</w:body></w:document>';
}

function buildDocx(sess) {
  const files = [
    {
      name: '[Content_Types].xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '</Types>'
    },
    {
      name: '_rels/.rels', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>'
    },
    {
      name: 'word/_rels/document.xml.rels', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'
    },
    {
      name: 'word/styles.xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:cs="TH SarabunPSK"/>' +
        '<w:sz w:val="32"/><w:szCs w:val="32"/><w:lang w:val="en-US" w:bidi="th-TH"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>' +
        '<w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>' +
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
        '</w:tblBorders></w:tblPr></w:style></w:styles>'
    },
    { name: 'word/document.xml', xml: docxXml(sess) }
  ];
  return zipStore(files);
}

function downloadDocx(sess) {
  const w = week(S.cur, true);
  const n = qualified(sess).length;
  if (!n && !confirm('สัปดาห์นี้ยังไม่มีนักเรียนถึงเกณฑ์ ' + S.meta.thr + ' ครั้ง — ต้องการสร้างไฟล์เปล่าหรือไม่ ?')) return;
  const name = `สรุปสาย ${SESS[sess]} น. สัปดาห์ที่ ${w.no || ''} ${w.label.replace(/\s*–\s*/g, '-')}.docx`;
  downloadBlob(new Blob([buildDocx(sess)],
    { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), name);
}
function downloadBlob(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/* ==========================================================================
   start
   ========================================================================== */
function boot() {
  week(S.cur, true); save();
  $('#carryRoom').innerHTML = '';
  if (window.BASELINE) $('#btnBaseline').textContent = '↺ โหลด' + window.BASELINE.label;
  setSess(S.sess);
  syncWeekBar();
  renderRec();
}
boot();
ready = true;
