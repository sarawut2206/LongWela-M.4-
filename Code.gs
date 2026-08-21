/**
 * ===========================================================================
 *  บันทึกการมาสาย ม.4  —  Google Apps Script (ฐานข้อมูลบน Google Sheet)
 * ---------------------------------------------------------------------------
 *  วิธีติดตั้ง
 *   1) เปิด Google Sheet ที่จะใช้เก็บข้อมูล → เมนู ส่วนขยาย (Extensions) → Apps Script
 *   2) ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้ลงไป → บันทึก
 *   3) กด Deploy → Manage deployments → (ดินสอ) Edit → Version: New version
 *      Execute as: Me      Who has access: Anyone
 *      → Deploy   (ลิงก์ /exec จะเป็นลิงก์เดิม ไม่ต้องแก้ในเว็บ)
 *   4) ครั้งแรกจะขึ้นให้อนุญาตสิทธิ์ (Authorize) ให้กดอนุญาตจนจบ
 *
 *  ชีตที่ระบบสร้างให้อัตโนมัติ
 *   - students : รายชื่อนักเรียน (ห้อง / เลขที่ / เลขประจำตัว / ชื่อ-สกุล)
 *   - late     : บันทึกมาสายรายสัปดาห์ (1 แถว = นักเรียน 1 คน ต่อ 1 สัปดาห์ ต่อ 1 ช่วงเวลา)
 *   - carry    : ยอดยกมา (ยอดสะสมก่อนเริ่มใช้โปรแกรม)
 *   - meta     : ค่าตั้งค่า เช่น ผู้ลงนาม เกณฑ์ครั้ง และเลขรุ่นข้อมูล (rev)
 * ===========================================================================
 */

var SH_STUDENTS = 'students';
var SH_LATE = 'late';
var SH_CARRY = 'carry';
var SH_META = 'meta';

var HEAD_STUDENTS = ['ห้อง', 'เลขที่', 'เลขประจำตัว', 'ชื่อ-สกุล'];
var HEAD_LATE = ['ช่วงเวลา', 'สัปดาห์ที่', 'วันจันทร์', 'ช่วงวันที่', 'ห้อง', 'เลขที่', 'ชื่อ-สกุล',
  'จำนวนครั้ง'];
var HEAD_CARRY = ['ช่วงเวลา', 'ห้อง', 'เลขที่', 'ชื่อ-สกุล', 'ยอดยกมา'];
var HEAD_META = ['key', 'value'];

/* ---------------------------------------------------------------- helpers */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, head) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function body_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (err) { return {}; }
  }
  return {};
}

function key_(r, n) { return r + '-' + n; }

/* ------------------------------------------------------------------ entry */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'load';
  var out;
  try {
    if (action === 'ping') out = { ok: true, ts: new Date().toISOString(), rev: getRev_() };
    else if (action === 'load') out = { ok: true, data: loadAll_() };
    else out = { ok: false, error: 'ไม่รู้จักคำสั่ง: ' + action };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  // รองรับ JSONP (เผื่อเบราว์เซอร์บล็อกการอ่านข้ามโดเมนแบบปกติ)
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(out);
}

function doPost(e) {
  var b = body_(e);
  var action = b.action || (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'save') return json_(saveAll_(b.data, b.baseRev, b.force));
    if (action === 'saveRoster') return json_(saveRoster_(b.roster));
    if (action === 'load') return json_({ ok: true, data: loadAll_() });
    return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ------------------------------------------------------------------- meta */
function metaMap_() {
  var sh = sheet_(SH_META, HEAD_META);
  var v = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < v.length; i++) if (v[i][0] !== '') out[v[i][0]] = v[i][1];
  return out;
}

function getRev_() {
  var r = metaMap_()['rev'];
  return Number(r || 0);
}

function writeMeta_(map) {
  var sh = sheet_(SH_META, HEAD_META);
  var rows = [];
  Object.keys(map).forEach(function (k) { rows.push([k, map[k]]); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_META.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* ------------------------------------------------------------------- load */
function loadAll_() {
  var meta = metaMap_();

  // ---- รายชื่อนักเรียน
  var shS = sheet_(SH_STUDENTS, HEAD_STUDENTS);
  var vs = shS.getDataRange().getValues();
  var roster = [];
  for (var i = 1; i < vs.length; i++) {
    var r = vs[i][0], n = vs[i][1];
    if (r === '' || n === '') continue;
    roster.push({ r: Number(r), n: Number(n), id: String(vs[i][2] || ''), name: String(vs[i][3] || '').trim() });
  }
  roster.sort(function (a, b) { return a.r - b.r || a.n - b.n; });

  // ---- ยอดยกมา
  var shC = sheet_(SH_CARRY, HEAD_CARRY);
  var vc = shC.getDataRange().getValues();
  var carry = { '0750': {}, '0830': {} };
  for (var i = 1; i < vc.length; i++) {
    var sess = String(vc[i][0] || '').trim();
    if (!carry[sess]) continue;
    var cnt = Number(vc[i][4] || 0);
    if (!cnt) continue;
    carry[sess][key_(Number(vc[i][1]), Number(vc[i][2]))] = cnt;
  }

  // ---- บันทึกมาสายรายสัปดาห์
  var shL = sheet_(SH_LATE, HEAD_LATE);
  var vl = shL.getDataRange().getValues();
  var weeks = {};
  for (var i = 1; i < vl.length; i++) {
    var sess = String(vl[i][0] || '').trim();
    var iso = vl[i][2];
    if (!sess || !iso) continue;
    if (iso instanceof Date) iso = Utilities.formatDate(iso, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    iso = String(iso).slice(0, 10);
    if (!weeks[iso]) weeks[iso] = { no: String(vl[i][1] || ''), label: String(vl[i][3] || ''), marks: { '0750': {}, '0830': {} } };
    if (!weeks[iso].marks[sess]) weeks[iso].marks[sess] = {};
    var cnt = Number(vl[i][7] || 0);
    if (!cnt) continue;
    weeks[iso].marks[sess][key_(Number(vl[i][4]), Number(vl[i][5]))] = cnt;
  }

  var signers = [];
  try { signers = JSON.parse(meta['signers'] || '[]'); } catch (err) { signers = []; }

  return {
    rev: Number(meta['rev'] || 0),
    updatedAt: meta['updatedAt'] || '',
    roster: roster,
    carry: carry,
    weeks: weeks,
    meta: {
      signers: signers.length ? signers : null,
      who: meta['who'] === '' || meta['who'] === undefined ? null : Number(meta['who']),
      level: meta['level'] || null,
      thr: meta['thr'] ? Number(meta['thr']) : null,
      note: meta['note'] === '' || meta['note'] === undefined ? null : (String(meta['note']) === 'true')
    }
  };
}

/* ------------------------------------------------------------------- save */
function saveAll_(data, baseRev, force) {
  if (!data) return { ok: false, error: 'ไม่มีข้อมูลที่ส่งมา' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'ระบบกำลังบันทึกรายการอื่นอยู่ ลองใหม่อีกครั้ง' };

  try {
    var rev = getRev_();
    if (!force && baseRev !== undefined && baseRev !== null && Number(baseRev) !== rev) {
      return { ok: false, conflict: true, rev: rev, error: 'ข้อมูลบนคลาวด์ถูกแก้ไขจากที่อื่นแล้ว (rev ' + rev + ')' };
    }

    var roster = data.roster || [];
    // ชื่อ-สกุลสำหรับเขียนลงชีต: เอาจากชีต students เดิม + ทับด้วยที่ส่งมาใหม่
    var nameOf = {};
    var shPrev = sheet_(SH_STUDENTS, HEAD_STUDENTS).getDataRange().getValues();
    for (var i = 1; i < shPrev.length; i++) {
      if (shPrev[i][0] === '' || shPrev[i][1] === '') continue;
      nameOf[key_(Number(shPrev[i][0]), Number(shPrev[i][1]))] = String(shPrev[i][3] || '');
    }
    roster.forEach(function (s) { nameOf[key_(s.r, s.n)] = s.name; });
    var names = data.names || {};
    Object.keys(names).forEach(function (k) { nameOf[k] = names[k]; });

    // ---- late
    var rows = [];
    var weeks = data.weeks || {};
    Object.keys(weeks).sort().forEach(function (iso) {
      var w = weeks[iso];
      ['0750', '0830'].forEach(function (sess) {
        var m = (w.marks || {})[sess] || {};
        Object.keys(m).forEach(function (k) {
          var cnt = m[k];
          // รองรับข้อมูลรูปแบบเก่าที่เป็นการติ๊กรายวัน
          if (Object.prototype.toString.call(cnt) === '[object Array]') {
            cnt = cnt.filter(function (x) { return x; }).length;
          }
          cnt = Number(cnt || 0);
          if (!cnt) return;
          var p = k.split('-');
          rows.push([sess, w.no || '', iso, w.label || '', Number(p[0]), Number(p[1]), nameOf[k] || '', cnt]);
        });
      });
    });
    rows.sort(function (a, b) {
      return String(a[2]).localeCompare(String(b[2])) || String(a[0]).localeCompare(String(b[0])) || a[4] - b[4] || a[5] - b[5];
    });
    writeRows_(SH_LATE, HEAD_LATE, rows, [3]);   // คอลัมน์ 3 = วันจันทร์ (เก็บเป็นข้อความ)

    // ---- carry
    var crows = [];
    ['0750', '0830'].forEach(function (sess) {
      var c = (data.carry || {})[sess] || {};
      Object.keys(c).forEach(function (k) {
        var cnt = Number(c[k] || 0);
        if (!cnt) return;
        var p = k.split('-');
        crows.push([sess, Number(p[0]), Number(p[1]), nameOf[k] || '', cnt]);
      });
    });
    crows.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])) || a[1] - b[1] || a[2] - b[2]; });
    writeRows_(SH_CARRY, HEAD_CARRY, crows);

    // ---- students (เขียนเฉพาะตอนที่ชีตยังว่าง หรือสั่งมาให้เขียน)
    var shS = sheet_(SH_STUDENTS, HEAD_STUDENTS);
    if (roster.length && (shS.getLastRow() < 2 || data.writeRoster)) {
      writeRows_(SH_STUDENTS, HEAD_STUDENTS, roster.map(function (s) {
        return [s.r, s.n, s.id || '', s.name];
      }));
    }

    // ---- meta
    var m = data.meta || {};
    rev = rev + 1;
    writeMeta_({
      rev: rev,
      updatedAt: Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      signers: JSON.stringify(m.signers || []),
      who: m.who === undefined ? 0 : m.who,
      level: m.level || '4',
      thr: m.thr || 4,
      note: m.note ? 'true' : 'false'
    });

    return { ok: true, rev: rev, rows: rows.length, carryRows: crows.length, students: roster.length };
  } finally {
    lock.releaseLock();
  }
}

function saveRoster_(roster) {
  if (!roster || !roster.length) return { ok: false, error: 'ไม่มีรายชื่อที่ส่งมา' };
  writeRows_(SH_STUDENTS, HEAD_STUDENTS, roster.map(function (s) { return [s.r, s.n, s.id || '', s.name]; }));
  return { ok: true, students: roster.length };
}

function writeRows_(name, head, rows, textCols) {
  var sh = sheet_(name, head);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, head.length).clearContent();
  if (!rows.length) return;
  // บังคับให้คอลัมน์วันที่เป็น "ข้อความ" กัน Google Sheet แปลงรูปแบบวันที่เอง
  (textCols || []).forEach(function (c) {
    sh.getRange(2, c, rows.length, 1).setNumberFormat('@');
  });
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}

/* -------------------------------------------------- เมนูช่วยงานในตัวชีตเอง */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('มาสาย ม.4')
    .addItem('สร้าง/ตรวจชีตให้ครบ', 'setupSheets')
    .addItem('ดูสถานะข้อมูล', 'showStatus')
    .addToUi();
}

function setupSheets() {
  sheet_(SH_STUDENTS, HEAD_STUDENTS);
  sheet_(SH_LATE, HEAD_LATE);
  sheet_(SH_CARRY, HEAD_CARRY);
  sheet_(SH_META, HEAD_META);
  SpreadsheetApp.getUi().alert('สร้างชีตครบแล้ว: students / late / carry / meta');
}

function showStatus() {
  var d = loadAll_();
  SpreadsheetApp.getUi().alert(
    'รายชื่อนักเรียน: ' + d.roster.length + ' คน\n' +
    'สัปดาห์ที่บันทึก: ' + Object.keys(d.weeks).length + ' สัปดาห์\n' +
    'ยอดยกมา 07.50: ' + Object.keys(d.carry['0750']).length + ' คน\n' +
    'ยอดยกมา 08.30: ' + Object.keys(d.carry['0830']).length + ' คน\n' +
    'rev: ' + d.rev + '   อัปเดตล่าสุด: ' + d.updatedAt
  );
}
