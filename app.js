/* ==========================================================================
   บันทึกการมาสาย ม.4  —  ทำงานในเบราว์เซอร์ล้วน ไม่ต้องมีเซิร์ฟเวอร์
   ข้อมูลเก็บใน localStorage และส่งออกเป็นไฟล์ Word (.docx) ตามฟอร์มโรงเรียน
   ========================================================================== */

const KEY = 'malaisai_m4_v1';
const SESS = { '0750': '07.50', '0830': '08.30' };
const DAYNAME = ['จ', 'อ', 'พ', 'พฤ', 'ศ'];
const TH_MONTH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

/* ---------- state ---------- */
let S = load();

function blank() {
  return {
    v: 1,
    sess: '0750',
    cur: mondayISO(new Date()),
    room: null,
    weeks: {},          // isoMonday -> { no, label, marks:{ '0750':{key:[5 bools]}, '0830':{} } }
    carry: { '0750': {}, '0830': {} },
    roster: null,       // null = ใช้รายชื่อจาก students.js
    meta: {
      signer: 'นางมานิดา ยอดเมือง',
      position: 'รองหัวหน้าระดับชั้นมัธยมศึกษาปีที่ 4',
      level: '4', thr: 4, note: true
    }
  };
}
function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return blank();
    const s = Object.assign(blank(), JSON.parse(raw));
    s.meta = Object.assign(blank().meta, s.meta || {});
    s.carry = Object.assign({ '0750': {}, '0830': {} }, s.carry || {});
    return s;
  } catch (e) { return blank(); }
}
function save() { localStorage.setItem(KEY, JSON.stringify(S)); }

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
function weekCount(isoStart, sess, k) {
  const m = (S.weeks[isoStart]?.marks?.[sess] || {})[k];
  return m ? m.filter(Boolean).length : 0;
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
function isMarked(k, di) { return !!(marks(S.cur, S.sess)[k] || [])[di]; }
function toggle(k, di) {
  const m = marks(S.cur, S.sess);
  if (!m[k]) m[k] = [false, false, false, false, false];
  m[k][di] = !m[k][di];
  if (!m[k].some(Boolean)) delete m[k];
  save();
}

/* ==========================================================================
   TABS
   ========================================================================== */
$$('nav button').forEach(b => b.onclick = () => {
  $$('nav button').forEach(x => x.classList.toggle('on', x === b));
  ['rec', 'sum', 'carry', 'roster', 'data'].forEach(t => $('#tab-' + t).hidden = (t !== b.dataset.tab));
  if (b.dataset.tab === 'sum') renderSummary();
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
  const mon = parseISO(S.cur);
  const thr = Number(S.meta.thr) || 4;

  let h = '<thead><tr>' + (q ? '<th>ห้อง</th>' : '') + '<th>เลขที่</th><th style="text-align:left">ชื่อ-สกุล</th>';
  for (let i = 0; i < 5; i++) h += `<th>${DAYNAME[i]}<br><span style="font-weight:400">${addDays(mon, i).getDate()}</span></th>`;
  h += '<th>สัปดาห์นี้</th><th>รวมสะสม</th></tr></thead><tbody>';

  list.forEach(s => {
    const k = keyOf(s);
    const wc = weekCount(S.cur, S.sess, k);
    const tt = total(S.sess, k, S.cur);
    const cls = tt >= thr ? 'over' : (wc ? 'hit' : '');
    h += `<tr class="${cls}">` + (q ? `<td>${s.r}</td>` : '') + `<td>${s.n}</td><td class="name">${s.name}</td>`;
    for (let i = 0; i < 5; i++)
      h += `<td class="tk"><span class="tick ${isMarked(k, i) ? 'on' : ''}" data-k="${k}" data-d="${i}">✓</span></td>`;
    h += `<td>${wc || ''}</td><td class="tot ${tt >= thr ? 'hi' : ''}">${tt || ''}</td></tr>`;
  });
  h += '</tbody>';
  $('#tblRec').innerHTML = h;

  $('#tblRec').querySelectorAll('.tick').forEach(el => el.onclick = () => {
    toggle(el.dataset.k, Number(el.dataset.d));
    renderRec();
  });

  const m = marks(S.cur, S.sess);
  const nRoom = roster().filter(s => s.r === S.room && m[keyOf(s)]).length;
  const nAll = Object.keys(m).length;
  $('#recInfo').textContent = q
    ? `พบ ${list.length} คน`
    : `ห้อง ${S.room} • ${list.length} คน • สัปดาห์นี้มีคนสาย ${nRoom} คน (รวมทุกห้อง ${nAll} คน)`;
}

/* ==========================================================================
   หน้า 2 : สรุป & พิมพ์
   ========================================================================== */
$('#thr').oninput = () => { S.meta.thr = Number($('#thr').value) || 4; save(); renderSummary(); renderRec(); };
$('#chkNote').onchange = () => { S.meta.note = $('#chkNote').checked; save(); renderSummary(); };
$('#btnPrint').onclick = () => window.print();
$('#btnDocx').onclick = () => downloadDocx(S.sess);
$('#btnDocxBoth').onclick = () => { downloadDocx('0750'); setTimeout(() => downloadDocx('0830'), 600); };

function qualified(sess) {
  const thr = Number(S.meta.thr) || 4;
  return roster()
    .map(s => ({ s, t: total(sess, keyOf(s), S.cur) }))
    .filter(x => x.t >= thr)
    .sort((a, b) => a.s.r - b.s.r || a.s.n - b.s.n);
}

function renderSummary() {
  $('#thr').value = S.meta.thr;
  $('#chkNote').checked = !!S.meta.note;
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
  h += `<div class="sign">( ${S.meta.signer})<br>${S.meta.position}</div>`;
  $('#preview').innerHTML = h;
}

/* ==========================================================================
   หน้า 3 : ยอดยกมา
   ========================================================================== */
$('#carryRoom').onchange = renderCarry;
$('#carryQ').oninput = renderCarry;
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
  S.roster = list.sort((a, b) => a.r - b.r || a.n - b.n); save();
  $('#paste').value = ''; renderRoster(); renderRec();
};
$('#btnResetRoster').onclick = () => {
  if (!confirm('คืนค่ารายชื่อเริ่มต้นจากไฟล์ students.js ?')) return;
  S.roster = null; save(); renderRoster(); renderRec();
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
  S.roster = cur.sort((a, b) => a.r - b.r || a.n - b.n); save();
}
function renderRoster() {
  const rs = rooms();
  $('#rosterInfo').innerHTML = `ตอนนี้มี <b>${roster().length}</b> คน / <b>${rs.length}</b> ห้อง (ห้อง ${rs.join(', ')})` +
    (S.roster ? ' — <b>แก้ไขแล้ว</b> (เก็บในเครื่องนี้)' : ' — รายชื่อเริ่มต้นจากไฟล์ students.js');
  let h = '<thead><tr><th>ห้อง</th><th>เลขที่</th><th style="text-align:left">ชื่อ-สกุล</th><th></th></tr></thead><tbody>';
  roster().forEach(s => {
    h += `<tr><td>${s.r}</td><td>${s.n}</td><td class="name">${s.name}</td>` +
      `<td><button class="btn" data-r="${s.r}" data-n="${s.n}" style="padding:2px 8px">ลบ</button></td></tr>`;
  });
  h += '</tbody>';
  $('#tblRoster').innerHTML = h;
  $('#tblRoster').querySelectorAll('button').forEach(b => b.onclick = () => {
    const r = Number(b.dataset.r), n = Number(b.dataset.n);
    S.roster = roster().filter(s => !(s.r === r && s.n === n)); save(); renderRoster(); renderRec();
  });
}

/* ==========================================================================
   หน้า 5 : ข้อมูล / สำรอง
   ========================================================================== */
['signer', 'position', 'level'].forEach(f => $('#' + f).oninput = () => {
  S.meta[f] = $('#' + f).value; save();
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
$('#btnWipe').onclick = () => {
  if (!confirm('ล้างข้อมูลทั้งหมด (รวมทุกสัปดาห์) ?')) return;
  if (!confirm('ยืนยันอีกครั้ง — ข้อมูลจะหายถาวร')) return;
  localStorage.removeItem(KEY); S = blank(); boot();
};
function renderData() {
  $('#signer').value = S.meta.signer; $('#position').value = S.meta.position; $('#level').value = S.meta.level;
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
  body += para(`           ( ${S.meta.signer})`, { tabs: 8 });
  body += para(`           ${S.meta.position}`, { tabs: 7 });
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
  setSess(S.sess);
  syncWeekBar();
  renderRec();
}
boot();
