/*****************************************************************
 * ระบบตารางงานร้านยา — ตัวเชื่อมหลังบ้าน (Google Apps Script)
 * ผูกกับ Google Sheet ที่ import มาจาก pharmacy_google_sheets.xlsx
 *
 * วิธีใช้: Extensions > Apps Script > วางโค้ดนี้ > Deploy > New deployment
 *          > Web app > Execute as: Me > Who has access: Anyone
 * แล้วเอา URL ที่ได้ไปใส่ในไฟล์ index.html (ตัวแปร API_URL)
 *****************************************************************/

/* ===================================================================
   ⚙️ ไฟล์ที่ระบบจะอ่าน/เขียน
   ใส่ ID ของไฟล์ Google Sheet เดือนปัจจุบัน (ข้อความยาวๆ ใน URL ระหว่าง /d/ กับ /edit)
   ขึ้นเดือนใหม่ = แก้ ID ตรงนี้ช่องเดียว ไม่ต้อง deploy ใหม่
   ถ้าเว้นว่าง "" = ใช้ไฟล์ที่วางสคริปต์นี้ไว้ (แบบเดิม)
   =================================================================== */
var SHEET_ID = "1psoXL4gM_HT_FMvKfYUhcasfactM5CT2C6ZZsFj8hrc";

// ไฟล์เดือนปัจจุบัน: ถ้าตั้งค่าผ่านเว็บไว้ (Script Properties) จะใช้ค่านั้นก่อน
// ไม่งั้นใช้ SHEET_ID ด้านบนเป็นค่าเริ่มต้น → เปลี่ยนเดือนได้จากเว็บ ไม่ต้องแก้โค้ด
function activeId_() {
  var p = PropertiesService.getScriptProperties().getProperty('ACTIVE_SHEET_ID');
  return (p && p.trim()) ? p.trim() : SHEET_ID;
}

// ไฟล์ที่กำลังทำงานในคำขอนี้ (ถ้าเว็บส่ง file มา = เดือนที่เลือก, ไม่งั้น = ค่าเริ่มต้น)
var _CUR_FILE = null;
function ss_() {
  var id = _CUR_FILE || activeId_();
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}
function extractId_(s) {
  s = String(s || '').trim();
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);   // เผื่อวางมาทั้งลิงก์
  return m ? m[1] : s;
}

/* ---------- ทะเบียนเดือน (เก็บใน Script Properties = ใช้ร่วมทุกเดือน) ---------- */
function getMonths_() {
  var raw = PropertiesService.getScriptProperties().getProperty('MONTHS');
  var arr = [];
  if (raw) { try { arr = JSON.parse(raw); } catch (e) { arr = []; } }
  if (!arr.length) {                 // ยังไม่เคยตั้ง → เริ่มด้วยไฟล์ปัจจุบัน
    var def = activeId_();
    if (def) {
      var label = 'เดือนปัจจุบัน';
      try { label = SpreadsheetApp.openById(def).getName(); } catch (e) {}
      arr = [{ id: def, label: label }];
    }
  }
  return arr;
}
function saveMonths_(arr) {
  PropertiesService.getScriptProperties().setProperty('MONTHS', JSON.stringify(arr));
}
function resolveFile_(p) {
  var f = (p && p.file || '').toString().trim();
  if (f) {
    var months = getMonths_();
    for (var i = 0; i < months.length; i++) if (months[i].id === f) return f;
  }
  return activeId_();
}

var SETTINGS_SHEET = 'Settings';
var CONFIG_SHEET   = 'Config';
var REQ_SHEET      = 'Requests';

// แถวข้อมูลพนักงานใน Settings = แถว 5 ถึง 21
var EMP_FIRST_ROW = 5;
var EMP_LAST_ROW  = 21;
// แถวรายการ shift ใน Settings คอลัมน์ L = แถว 5 ถึง 32
var SHIFT_FIRST_ROW = 5;
var SHIFT_LAST_ROW  = 32;
// ในชีตตารางสาขา: หัวตารางอยู่แถว 4, ข้อมูลวันเริ่มแถว 5
var SCHED_HEADER_ROW = 4;
var SCHED_FIRST_ROW  = 5;

/* ---------- helper: แปลงเลขสาขา <-> ชื่อ <-> ชื่อชีต ---------- */
function branchIdToName(v) {
  v = String(v).replace('.0', '').trim();
  if (v === '5' || v === 'บางขุนศรี') return 'บางขุนศรี';
  if (v === 'สาขา 1' || v === 'สาขา 2' || v === 'สาขา 3' || v === 'สาขา 4') return v;
  return 'สาขา ' + v;
}
function branchNameToSheet(name) {
  return 'ตาราง_' + name;   // 'สาขา 1' -> 'ตาราง_สาขา 1' ; 'บางขุนศรี' -> 'ตาราง_บางขุนศรี'
}

/* ---------- อ่าน Config ---------- */
function getConfig_() {
  var sh = ss_().getSheetByName(CONFIG_SHEET);
  var vals = sh.getDataRange().getValues();
  var o = {};
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0]) o[String(vals[i][0]).trim()] = String(vals[i][1]).trim();
  }
  return o;
}

/* ---------- อ่านพนักงานทั้งหมด ---------- */
function getEmployees_() {
  var sh = ss_().getSheetByName(SETTINGS_SHEET);
  var rng = sh.getRange(EMP_FIRST_ROW, 1, EMP_LAST_ROW - EMP_FIRST_ROW + 1, 11).getValues();
  var list = [];
  for (var i = 0; i < rng.length; i++) {
    var name = rng[i][0];
    if (!name) continue;
    list.push({
      name: String(name).trim(),
      branch: branchIdToName(rng[i][1]),
      otrate: rng[i][8],
      pin: String(rng[i][10]).trim()
    });
  }
  return list;
}

/* ---------- รายการ shift ---------- */
function getShifts_() {
  var sh = ss_().getSheetByName(SETTINGS_SHEET);
  var rng = sh.getRange(SHIFT_FIRST_ROW, 12, SHIFT_LAST_ROW - SHIFT_FIRST_ROW + 1, 1).getValues(); // col L
  var out = [];
  for (var i = 0; i < rng.length; i++) {
    if (rng[i][0]) out.push(String(rng[i][0]).trim());
  }
  if (out.indexOf('หยุด') === -1) out.unshift('หยุด');
  return out;
}

/* ---------- map: ชื่อ shift -> OT (ชม.) สำหรับลงสี ---------- */
function getShiftMap_() {
  var sh = ss_().getSheetByName(SETTINGS_SHEET);
  var rng = sh.getRange(SHIFT_FIRST_ROW, 12, SHIFT_LAST_ROW - SHIFT_FIRST_ROW + 1, 3).getValues(); // L:N
  var m = {};
  for (var i = 0; i < rng.length; i++) {
    if (rng[i][0] !== '' && rng[i][0] !== null) {
      m[String(rng[i][0]).trim()] = Number(rng[i][2]) || 0; // N = OT
    }
  }
  return m;
}

/* ---------- อ่านตารางของสาขา (พร้อมสีจริงจากชีต) ---------- */
function readSchedule_(branchName) {
  var sheet = ss_().getSheetByName(branchNameToSheet(branchName));
  if (!sheet) return null;
  var rng = sheet.getDataRange();
  var data = rng.getValues();
  var bgs  = rng.getBackgrounds();   // ← สีพื้นของทุกเซลล์ ช่วงเดียวกับ data
  var header = data[SCHED_HEADER_ROW - 1];   // 0-based
  var headerBg = bgs[SCHED_HEADER_ROW - 1];
  // คอลัมน์พนักงาน = ตั้งแต่ col index 2 เป็นต้นไปที่มีหัวตาราง
  var cols = [];
  for (var c = 2; c < header.length; c++) {
    if (header[c]) cols.push({ idx: c, name: String(header[c]).trim() });
  }
  // สีหัวตาราง: คอลัมน์ 'วัน' ใช้สีเซลล์แรกของแถวหัว, ที่เหลือตามเซลล์หัวของแต่ละคน
  var headerColors = { 'วัน': headerBg[0] };
  for (var h = 0; h < cols.length; h++) {
    headerColors[cols[h].name] = headerBg[cols[h].idx];
  }
  var days = [];
  for (var r = SCHED_FIRST_ROW - 1; r < data.length; r++) {
    var day = data[r][0];
    if (day === '' || day === null) continue;
    var cells = {}, colors = {};
    for (var k = 0; k < cols.length; k++) {
      var v = data[r][cols[k].idx];
      cells[cols[k].name]  = (v === null) ? '' : String(v);
      colors[cols[k].name] = bgs[r][cols[k].idx];   // สีจริงของช่องนั้นในชีต
    }
    days.push({ day: Number(day), dow: String(data[r][1] || ''), cells: cells, colors: colors });
  }
  return {
    columns: cols.map(function (x) { return x.name; }),
    days: days,
    headerColors: headerColors
  };
}

/* ---------- คำขอที่ค้างอยู่ ---------- */
function readRequests_(filterBranch, onlyPending) {
  var sh = ss_().getSheetByName(REQ_SHEET);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var rec = {
      reqId: String(row[0]), ts: row[1], branch: String(row[2]), day: Number(row[3]),
      employee: String(row[4]), oldShift: String(row[5]), newShift: String(row[6]),
      note: String(row[7]), status: String(row[8]), decidedAt: row[9], decidedBy: String(row[10]),
      rowNum: i + 1
    };
    if (onlyPending && rec.status !== 'pending') continue;
    if (filterBranch && rec.branch !== filterBranch) continue;
    out.push(rec);
  }
  return out;
}

/* ============================ ACTIONS ============================ */

function actBootstrap_() {
  var emps = getEmployees_();
  var branchSet = {};
  emps.forEach(function (e) { branchSet[e.branch] = true; });
  var branches = ['สาขา 1', 'สาขา 2', 'สาขา 3', 'สาขา 4', 'บางขุนศรี'].filter(function (b) { return branchSet[b]; });
  return {
    ok: true,
    branches: branches,
    employees: emps.map(function (e) { return { name: e.name, branch: e.branch }; }),
    shifts: getShifts_(),
    shiftOT: getShiftMap_(),
    months: getMonths_(),
    defaultFile: activeId_()
  };
}

function actLogin_(p) {
  var name = (p.name || '').trim();
  var pin = (p.pin || '').trim();
  var cfg = getConfig_();
  if (pin && pin === String(cfg.AdminPIN).trim() && (name === '' || name === cfg.AdminName)) {
    return { ok: true, role: 'admin', name: cfg.AdminName || 'แอดมิน' };
  }
  var emps = getEmployees_();
  for (var i = 0; i < emps.length; i++) {
    if (emps[i].name === name && emps[i].pin === pin) {
      return { ok: true, role: 'staff', name: emps[i].name, branch: emps[i].branch };
    }
  }
  return { ok: false, error: 'ชื่อหรือ PIN ไม่ถูกต้อง' };
}

function verifyEmp_(name, pin) {
  var emps = getEmployees_();
  for (var i = 0; i < emps.length; i++) {
    if (emps[i].name === name && emps[i].pin === pin) return emps[i];
  }
  return null;
}
function isAdmin_(pin) {
  var cfg = getConfig_();
  return pin && pin === String(cfg.AdminPIN).trim();
}

function actSchedule_(p) {
  var branch = branchIdToName(p.branch || '');
  var sched = readSchedule_(branch);
  if (!sched) return { ok: false, error: 'ไม่พบสาขา ' + branch };
  return {
    ok: true,
    branch: branch,
    columns: sched.columns,
    days: sched.days,
    headerColors: sched.headerColors,
    shifts: getShifts_(),
    shiftOT: getShiftMap_(),
    pending: readRequests_(branch, true)
  };
}

function actRequest_(p) {
  var name = (p.name || '').trim();
  var pin = (p.pin || '').trim();
  var emp = verifyEmp_(name, pin);
  var admin = isAdmin_(pin);
  if (!emp && !admin) return { ok: false, error: 'PIN ไม่ถูกต้อง' };

  var branch = branchIdToName(p.branch || (emp ? emp.branch : ''));
  var day = Number(p.day);
  var employee = (p.employee || name).trim();
  var newShift = (p.newShift || '').trim();
  var note = (p.note || '').trim();

  // พนักงานทั่วไปแก้ได้เฉพาะของตัวเอง
  if (!admin && employee !== name) return { ok: false, error: 'แก้ได้เฉพาะตารางของตัวเอง' };
  if (!newShift && !note) return { ok: false, error: 'กรุณาเลือกเวลาใหม่ หรือใส่หมายเหตุ' };

  // หา shift เดิม
  var sched = readSchedule_(branch);
  var oldShift = '';
  if (sched) {
    for (var i = 0; i < sched.days.length; i++) {
      if (sched.days[i].day === day) { oldShift = sched.days[i].cells[employee] || ''; break; }
    }
  }

  var sh = ss_().getSheetByName(REQ_SHEET);
  var reqId = 'R' + new Date().getTime();
  sh.appendRow([reqId, new Date(), branch, day, employee, oldShift, newShift, note, 'pending', '', '']);
  return { ok: true, reqId: reqId };
}

function actPending_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  return { ok: true, pending: readRequests_(null, true) };
}

function actDecide_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var reqId = (p.reqId || '').trim();
  var action = (p.decision || '').trim();   // approve / reject  (ใช้ชื่อ decision กันชนกับ action ที่เป็น router)
  var decNote = (p.note || '').trim();
  var cfg = getConfig_();
  var reqs = readRequests_(null, false);
  var rec = null;
  for (var i = 0; i < reqs.length; i++) { if (reqs[i].reqId === reqId) { rec = reqs[i]; break; } }
  if (!rec) return { ok: false, error: 'ไม่พบคำขอ' };

  var sh = ss_().getSheetByName(REQ_SHEET);

  if (action === 'approve') {
    // เขียนค่าใหม่ลงชีตตาราง (เฉพาะกรณีมี newShift)
    if (rec.newShift) {
      var sheet = ss_().getSheetByName(branchNameToSheet(rec.branch));
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        var header = data[SCHED_HEADER_ROW - 1];
        var colIdx = -1;
        for (var c = 2; c < header.length; c++) {
          if (String(header[c]).trim() === rec.employee) { colIdx = c; break; }
        }
        var rowIdx = -1;
        for (var r = SCHED_FIRST_ROW - 1; r < data.length; r++) {
          if (Number(data[r][0]) === rec.day) { rowIdx = r; break; }
        }
        if (colIdx >= 0 && rowIdx >= 0) {
          sheet.getRange(rowIdx + 1, colIdx + 1).setValue(rec.newShift);
        }
      }
    }
    sh.getRange(rec.rowNum, 9, 1, 3).setValues([['approved', new Date(), cfg.AdminName || 'admin']]);
    if (decNote) sh.getRange(rec.rowNum, 8).setValue((rec.note ? rec.note + ' | ' : '') + 'แอดมิน: ' + decNote);
    return { ok: true };
  } else if (action === 'reject') {
    sh.getRange(rec.rowNum, 9, 1, 3).setValues([['rejected', new Date(), cfg.AdminName || 'admin']]);
    if (decNote) sh.getRange(rec.rowNum, 8).setValue((rec.note ? rec.note + ' | ' : '') + 'แอดมิน: ' + decNote);
    return { ok: true };
  }
  return { ok: false, error: 'action ไม่ถูกต้อง' };
}

function actMyRequests_(p) {
  var name = (p.name || '').trim();
  var emp = verifyEmp_(name, (p.pin || '').trim());
  if (!emp && !isAdmin_((p.pin || '').trim())) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  var all = readRequests_(null, false);
  var mine = all.filter(function (r) { return r.employee === name; });
  return { ok: true, requests: mine };
}

/* ---------- จัดการ PIN (เฉพาะแอดมิน) ---------- */
function actPins_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var cfg = getConfig_();
  return {
    ok: true,
    adminPin: String(cfg.AdminPIN || ''),
    activeFileId: activeId_(),
    activeFileName: (function () { try { return ss_().getName(); } catch (e) { return ''; } })(),
    employees: getEmployees_().map(function (e) { return { name: e.name, branch: e.branch, pin: e.pin }; })
  };
}

function actSetActiveFile_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var id = extractId_(p.fileId || '');
  if (!id) return { ok: false, error: 'กรุณาใส่ ID หรือลิงก์ไฟล์' };
  var name = '';
  try {
    var test = SpreadsheetApp.openById(id);
    name = test.getName();
    if (!test.getSheetByName(SETTINGS_SHEET)) return { ok: false, error: 'ไฟล์นี้ไม่มีแท็บ Settings (โครงสร้างไม่ตรงเทมเพลต)' };
  } catch (e) {
    return { ok: false, error: 'เปิดไฟล์ไม่ได้ — ตรวจ ID หรือสิทธิ์การเข้าถึง' };
  }
  PropertiesService.getScriptProperties().setProperty('ACTIVE_SHEET_ID', id);
  return { ok: true, activeFileId: id, activeFileName: name };
}

/* ---------- จัดการเดือน (เฉพาะแอดมิน) ---------- */
function actMonths_(p) {
  return { ok: true, months: getMonths_(), defaultFile: activeId_() };
}
function actAddMonth_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var id = extractId_(p.fileId || '');
  if (!id) return { ok: false, error: 'กรุณาใส่ลิงก์หรือ ID ไฟล์' };
  var name = '';
  try {
    var test = SpreadsheetApp.openById(id);
    name = test.getName();
    if (!test.getSheetByName(SETTINGS_SHEET)) return { ok: false, error: 'ไฟล์นี้ไม่มีแท็บ Settings (โครงสร้างไม่ตรง)' };
  } catch (e) { return { ok: false, error: 'เปิดไฟล์ไม่ได้ — ตรวจ ID หรือสิทธิ์' }; }
  var label = (p.label || '').trim() || name;
  var months = getMonths_();
  for (var i = 0; i < months.length; i++) {
    if (months[i].id === id) { months[i].label = label; saveMonths_(months); return { ok: true, months: months }; }
  }
  months.push({ id: id, label: label });
  saveMonths_(months);
  return { ok: true, months: months };
}
function actRemoveMonth_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var id = (p.fileId || '').trim();
  var months = getMonths_().filter(function (m) { return m.id !== id; });
  saveMonths_(months);
  return { ok: true, months: months };
}

function actSetPin_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var target = (p.target || '').trim();
  var newpin = (p.newpin || '').trim();
  if (!/^[0-9]{4,6}$/.test(newpin)) return { ok: false, error: 'PIN ต้องเป็นตัวเลข 4–6 หลัก' };
  var sh = ss_().getSheetByName(SETTINGS_SHEET);
  for (var r = EMP_FIRST_ROW; r <= EMP_LAST_ROW; r++) {
    if (String(sh.getRange(r, 1).getValue()).trim() === target) {
      sh.getRange(r, 11).setValue(newpin);  // คอลัมน์ K
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบพนักงาน ' + target };
}

function actSetAdminPin_(p) {
  if (!isAdmin_((p.pin || '').trim())) return { ok: false, error: 'ต้องใช้ PIN แอดมิน' };
  var newpin = (p.newpin || '').trim();
  if (!/^[0-9]{4,6}$/.test(newpin)) return { ok: false, error: 'PIN ต้องเป็นตัวเลข 4–6 หลัก' };
  var sh = ss_().getSheetByName(CONFIG_SHEET);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'AdminPIN') { sh.getRange(i + 1, 2).setValue(newpin); return { ok: true }; }
  }
  return { ok: false, error: 'ไม่พบ AdminPIN ใน Config' };
}

/* ============================ ROUTER ============================ */
function route_(p) {
  var a = (p.action || '').trim();
  switch (a) {
    case 'bootstrap':   return actBootstrap_();
    case 'login':       return actLogin_(p);
    case 'schedule':    return actSchedule_(p);
    case 'request':     return actRequest_(p);
    case 'pending':     return actPending_(p);
    case 'decide':      return actDecide_(p);
    case 'myrequests':  return actMyRequests_(p);
    case 'pins':        return actPins_(p);
    case 'setpin':      return actSetPin_(p);
    case 'setadminpin': return actSetAdminPin_(p);
    case 'setactivefile': return actSetActiveFile_(p);
    case 'months':      return actMonths_(p);
    case 'addmonth':    return actAddMonth_(p);
    case 'removemonth': return actRemoveMonth_(p);
    default:            return { ok: false, error: 'ไม่รู้จัก action: ' + a };
  }
}

function respond_(p) {
  var result;
  try { _CUR_FILE = resolveFile_(p); result = route_(p); }
  catch (err) { result = { ok: false, error: String(err) }; }
  var json = JSON.stringify(result);
  var cb = p.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e)  { return respond_(e.parameter); }
function doPost(e) {
  var p = e.parameter || {};
  if (e.postData && e.postData.contents) {
    try { var b = JSON.parse(e.postData.contents); for (var k in b) p[k] = b[k]; } catch (x) {}
  }
  return respond_(p);
}

/* ทดสอบจากในตัว editor ได้ */
function _test() { Logger.log(JSON.stringify(actBootstrap_())); }
