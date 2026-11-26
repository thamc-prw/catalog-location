/**
 * THAMC Inventory API — Google Apps Script (Web App)
 * ผูกกับ Google Sheet ที่มีแท็บ FRZ294 (หลัก), Sheet3 (รหัสกลุ่ม)
 * แท็บ Extra และ Log จะถูกสร้างอัตโนมัติเมื่อรัน setup() หนึ่งครั้ง
 *
 * คอลัมน์แท็บ FRZ294:
 * A Item code | B TUH Code | C Item name | D Bin location | E Vender | F Expiry date
 * G Uom | H Item cost | I Freeze qty | J รูป | K เคลื่อนไหว | L รหัสกลุ่ม | M ประเภท
 */
const CFG = {
  SHEET_ITEMS: 'FRZ294',
  SHEET_GROUPS: 'Sheet3',
  SHEET_EXTRA: 'Extra',
  SHEET_LOG: 'Log',
  SHEET_LOC: 'Locations',
  SPREADSHEET_ID: '',   // ใส่ ID ของ Google Sheet ถ้าสร้างสคริปต์แบบ standalone (ไม่ได้เปิดจาก Extensions > Apps Script ในไฟล์ Sheet)
  PHOTO_FOLDER_ID: '',   // ใส่ ID โฟลเดอร์ Google Drive สำหรับเก็บรูป (ปล่อยว่าง = สร้างโฟลเดอร์ "THAMC Item Photos")
  ADMIN_PIN: '1234',     // รหัสผ่านแอดมิน (เปลี่ยนก่อนใช้งานจริง) — เฉพาะแอดมินเท่านั้นที่เขียนข้อมูลได้
  TOKEN: ''              // ใส่รหัสผ่านร่วม (ถ้าต้องการ) แล้วส่งมาใน ?token=... / body.token
};
const EXTRA_HEAD = ['code', 'ref', 'barcodes', 'photo1', 'photo2', 'photo3', 'isNew', 'updated'];
const SEP = '|';

/** รันครั้งเดียวจากตัวแก้ไข Apps Script เพื่อสร้างแท็บ Extra / Log */
function setup() {
  const ss = ss_();
  [[CFG.SHEET_EXTRA, EXTRA_HEAD], [CFG.SHEET_LOG, ['time', 'user', 'action', 'code', 'detail']], [CFG.SHEET_LOC, ['cab', 'shelves', 'slots']]].forEach(([n, h]) => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) { sh.appendRow(h); sh.setFrozenRows(1); }
  });
  folder_();
}

function doGet(e) {
  try {
    auth_(e.parameter.token);
    const a = e.parameter.action || 'items';
    if (a === 'items') return json_({ ok: true, items: readItems_(), groups: readGroups_(), cabs: readCabs_() });
    if (a === 'item') return json_({ ok: true, item: readItems_().find(i => i.code === e.parameter.code) || null });
    if (a === 'login') return json_({ ok: e.parameter.pin === CFG.ADMIN_PIN });
    if (a === 'ping') return json_({ ok: true, time: new Date() });
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    auth_(b.token);
    if (b.adminPin !== CFG.ADMIN_PIN) return json_({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
    if (b.action === 'login') return json_({ ok: true });
    if (b.action === 'update') return json_(updateItem_(b));
    if (b.action === 'cab') return json_(saveCab_(b));
    if (b.action === 'delCab') return json_(delCab_(b));
    if (b.action === 'setLocs') return json_(setLocs_(b));
    if (b.action === 'delItem') return json_(delItem_(b));
    if (b.action === 'import') return json_(importRows_(b));
    if (b.action === 'delPhoto') return json_(delPhoto_(b));
    if (b.action === 'migrate') return json_(migratePhotos_(b));
    if (b.action === 'photo') return json_(savePhoto_(b));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}

/* ---------- read ---------- */
function readItems_() {
  const ss = ss_();
  const mainSh = items_();
  const main = mainSh.getDataRange().getValues();
  const groups = readGroups_();
  const extra = {};
  const ex = ss.getSheetByName(CFG.SHEET_EXTRA);
  if (ex && ex.getLastRow() > 1) ex.getDataRange().getValues().slice(1).forEach(r => { extra[String(r[0])] = r; });
  return main.slice(1).filter(r => r[0] !== '').map(r => {
    const code = String(r[0]), x = extra[code] || [];
    const gcode = String(r[11] || code.slice(0, 6));
    const mainId = driveId_(r[9]);
    const ids = [...new Set([mainId, driveId_(x[3]), driveId_(x[4]), driveId_(x[5])].filter(Boolean))].slice(0, 3);
    return {
      code, tuh: String(r[1] || ''), name: String(r[2] || ''), loc: String(r[3] || ''), vendor: String(r[4] || ''),
      exp: fmtDate_(r[5]), uom: String(r[6] || ''), cost: Number(r[7]) || 0, qty: Number(r[8]) || 0,
      moving: String(r[10]).trim() === 'เคลื่อนไหว', gcode, cat: String(r[12] || groups[gcode] || ''),
      ref: split_(x[1]), bc: split_(x[2]), photoIds: ids, isNew: x[6] === true || x[6] === 'TRUE' || x[6] === 1
    };
  });
}
function readGroups_() {
  const sh = ss_().getSheetByName(CFG.SHEET_GROUPS), g = {};
  if (sh) sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0] !== '') g[String(r[0])] = String(r[1]); });
  return g;
}

function readCabs_() {
  const sh = ss_().getSheetByName(CFG.SHEET_LOC);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '').map(r => ({ cab: String(r[0]), shelves: Number(r[1]) || 0, slots: Number(r[2]) || 0 }));
}

/* ---------- write ---------- */
function saveCab_(b) {
  const s = ss_();
  let sh = s.getSheetByName(CFG.SHEET_LOC);
  if (!sh) { setup(); sh = s.getSheetByName(CFG.SHEET_LOC); }
  const row = findRow_(sh, b.cab);
  const vals = [String(b.cab), Number(b.shelves) || 0, Number(b.slots) || 0];
  if (row) sh.getRange(row, 1, 1, 3).setValues([vals]); else sh.appendRow(vals);
  log_(b.user, 'cab', b.cab, b.shelves + 'x' + b.slots);
  return { ok: true };
}

/** ตั้ง/ล้าง Location หลายรายการพร้อมกัน: changes = [{code, loc}] (loc '' = นำออกจากช่อง) */
function setLocs_(b) {
  const sh = items_(), last = sh.getLastRow(); if (last < 2) return { ok: true, n: 0 };
  const idx = {}; sh.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => { idx[String(r[0])] = i; });
  const rng = sh.getRange(2, 4, last - 1, 1), vals = rng.getValues(); let n = 0;
  (b.changes || []).forEach(c => { const i = idx[String(c.code)]; if (i !== undefined) { vals[i][0] = c.loc || ''; n++; } });
  rng.setValues(vals); log_(b.user, 'setLocs', '', n + ' items');
  return { ok: true, n };
}

/** ลบสินค้าออกจากระบบถาวร (แถวใน FRZ294 และแถวใน Extra) */
function delItem_(b) {
  const sh = items_(), row = findRow_(sh, b.code);
  if (!row) return { ok: false, error: 'ไม่พบรหัส ' + b.code };
  sh.deleteRow(row);
  const ex = ss_().getSheetByName(CFG.SHEET_EXTRA), er = ex ? findRow_(ex, b.code) : 0;
  if (er) ex.deleteRow(er);
  log_(b.user, 'delItem', b.code, '');
  return { ok: true };
}

function delCab_(b) {
  const prefix = String(b.cab) + '-', sh = items_(), last = sh.getLastRow();
  if (last > 1 && sh.getRange(2, 4, last - 1, 1).getValues().some(r => String(r[0]).indexOf(prefix) === 0)) return { ok: false, error: 'ตู้ ' + b.cab + ' ยังมีสินค้า' };
  const loc = ss_().getSheetByName(CFG.SHEET_LOC), row = loc ? findRow_(loc, b.cab) : 0;
  if (row) loc.deleteRow(row);
  log_(b.user, 'delCab', b.cab, '');
  return { ok: true };
}

/** นำเข้า/อัปเดตหลายแถวจากไฟล์ Excel — upsert ตามรหัสสินค้า ไม่ลบแถวเดิม ไม่แตะแท็บ Extra */
function importRows_(b) {
  const sh = items_(), rows = b.rows || [];
  const last = sh.getLastRow();
  const data = sh.getRange(1, 1, last, 14).getValues();
  const idx = {};
  for (let i = 1; i < data.length; i++) idx[String(data[i][0])] = i;
  let added = 0, updated = 0;
  rows.forEach(o => {
    const code = String(o.code || '').trim(); if (!code) return;
    const vals = [/^\d+$/.test(code) ? Number(code) : code, o.tuh || '', o.name || '', o.loc || '', o.vendor || '', o.exp || '', o.uom || '', Number(o.cost) || 0, Number(o.qty) || 0, o.img || '', o.moving ? 'เคลื่อนไหว' : '', o.gcode || code.slice(0, 6), o.cat || ''];
    if (idx[code] !== undefined) {
      const r = data[idx[code]], keepImg = r[9];
      for (let c = 0; c < 13; c++) r[c] = vals[c];
      if (!o.img) r[9] = keepImg; else if (String(o.img) !== String(keepImg)) r[13] = '';
      updated++;
    } else { const nr = vals.slice(); nr.push(''); data.push(nr); idx[code] = data.length - 1; added++; }
  });
  sh.getRange(1, 1, data.length, 14).setValues(data);
  log_(b.user, 'import', '', 'added ' + added + ', updated ' + updated);
  return { ok: true, added, updated };
}
function updateItem_(b) {
  const ss = ss_(), sh = items_();
  const row = findRow_(sh, b.code);
  if (!row) return { ok: false, error: 'ไม่พบรหัส ' + b.code };
  const detail = [];
  if (b.loc !== undefined) {
    const old = sh.getRange(row, 4).getValue();
    if (String(old) !== String(b.loc)) { sh.getRange(row, 4).setValue(b.loc); detail.push('loc ' + old + ' → ' + b.loc); }
  }
  if (b.moving !== undefined) { sh.getRange(row, 11).setValue(b.moving ? 'เคลื่อนไหว' : ''); detail.push('moving=' + b.moving); }
  const ex = extraRow_(b.code);
  if (b.ref !== undefined) { ex.sh.getRange(ex.row, 2).setValue(b.ref.join(SEP)); detail.push('ref'); }
  if (b.bc !== undefined) { ex.sh.getRange(ex.row, 3).setValue(b.bc.join(SEP)); detail.push('barcode'); }
  if (b.isNew !== undefined) ex.sh.getRange(ex.row, 7).setValue(!!b.isNew);
  if (b.photoCount !== undefined) {
    const vals = ex.sh.getRange(ex.row, 4, 1, 3).getValues()[0];
    for (let k = b.photoCount; k < 3; k++) vals[k] = '';
    ex.sh.getRange(ex.row, 4, 1, 3).setValues([vals]);
    if (b.photoCount === 0) sh.getRange(row, 10).setValue('');
    detail.push('photos=' + b.photoCount);
  }
  ex.sh.getRange(ex.row, 8).setValue(new Date());
  log_(b.user, 'update', b.code, detail.join('; '));
  return { ok: true };
}

/** เพิ่มรูป (ต่อท้าย ไม่ทับรูปเดิม) — สูงสุด 3 รูปต่อสินค้า รูปแรกจะถูกใช้เป็นรูปหลักในคอลัมน์ J */
function savePhoto_(b) {
  const sh = items_(), row = findRow_(sh, b.code);
  if (!row) return { ok: false, error: 'ไม่พบรหัส ' + b.code };
  const ex = extraRow_(b.code), cur = ex.sh.getRange(ex.row, 4, 1, 3).getValues()[0].map(v => String(v || ''));
  const jId = driveId_(sh.getRange(row, 10).getValue());
  const have = new Set([jId, ...cur.map(driveId_)].filter(Boolean));
  if (have.size >= 3) return { ok: false, error: 'สินค้านี้มีรูปครบ 3 รูปแล้ว ลบรูปเดิมก่อน' };
  const slot = cur.findIndex(v => !v);
  if (slot < 0) return { ok: false, error: 'ช่องเก็บรูปเต็ม ลบรูปเดิมก่อน' };
  const bytes = Utilities.base64Decode(String(b.data).replace(/^data:[^,]+,/, ''));
  const blob = Utilities.newBlob(bytes, b.mime || 'image/jpeg', b.code + '_' + Date.now() + '.jpg');
  const file = folder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  ex.sh.getRange(ex.row, 4 + slot).setValue(file.getUrl());
  if (!jId) sh.getRange(row, 10).setValue(file.getUrl());
  log_(b.user, 'photo', b.code, 'add ' + file.getId());
  return { ok: true, id: file.getId(), slot };
}

/** ลบรูปเดียวตามรหัสไฟล์ Drive (ถอดลิงก์ออกจากชีต ไม่ลบไฟล์ใน Drive) */
function delPhoto_(b) {
  const id = String(b.id || ''); if (!id) return { ok: false, error: 'ไม่พบรหัสรูป' };
  const sh = items_(), row = findRow_(sh, b.code);
  if (!row) return { ok: false, error: 'ไม่พบรหัส ' + b.code };
  if (driveId_(sh.getRange(row, 10).getValue()) === id) sh.getRange(row, 10).setValue('');
  const ex = extraRow_(b.code), rng = ex.sh.getRange(ex.row, 4, 1, 3), v = rng.getValues()[0].map(x => driveId_(x) === id ? '' : x);
  rng.setValues([v]);
  if (!String(sh.getRange(row, 10).getValue() || '')) { const first = v.find(Boolean); if (first) sh.getRange(row, 10).setValue(first); }
  log_(b.user, 'photo', b.code, 'remove ' + id);
  return { ok: true };
}

/**
 * ย้ายรูปจากลิงก์ Drive เดิม (คอลัมน์ J) มาทำสำเนาเก็บในโฟลเดอร์ Drive ของระบบ แล้วแทนที่ลิงก์
 * คอลัมน์ N = สถานะ: ok / error: ... (ว่าง = ยังไม่ย้าย) · เรียกซ้ำจนกว่า pending = 0
 * b.limit = จำนวนไฟล์ต่อรอบ (0 = ดูสถานะเฉยๆ) · b.retry = true เพื่อลองรายการ error ใหม่
 */
function migratePhotos_(b) {
  const t0 = Date.now(), sh = items_(), last = sh.getLastRow();
  if (last < 2) return { ok: true, total: 0, okN: 0, errN: 0, pending: 0, progressed: 0 };
  if (!String(sh.getRange(1, 14).getValue())) sh.getRange(1, 14).setValue('รูปย้ายแล้ว');
  const rng = sh.getRange(2, 10, last - 1, 5).getValues(); // J..N
  const limit = Number(b.limit) || 0, folder = limit ? folder_() : null;
  let total = 0, okN = 0, errN = 0, pending = 0, progressed = 0;
  for (let i = 0; i < rng.length; i++) {
    const link = String(rng[i][0] || ''), mark = String(rng[i][4] || '');
    if (!link) continue;
    total++;
    if (mark === 'ok') { okN++; continue; }
    if (mark && !b.retry) { errN++; continue; }
    if (!limit || progressed >= limit || Date.now() - t0 > 240000) { pending++; continue; }
    const row = i + 2, id = driveId_(link);
    progressed++;
    if (!id) { sh.getRange(row, 14).setValue('error: ไม่พบรหัสไฟล์ในลิงก์'); errN++; continue; }
    try {
      const src = DriveApp.getFileById(id);
      const copy = src.makeCopy(src.getName(), folder);
      copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sh.getRange(row, 10).setValue(copy.getUrl());
      sh.getRange(row, 14).setValue('ok');
      okN++;
    } catch (e) { sh.getRange(row, 14).setValue('error: ' + String(e).slice(0, 90)); errN++; }
  }
  if (progressed) log_(b.user, 'migrate', '', 'ok ' + okN + ', error ' + errN + ', pending ' + pending);
  return { ok: true, total, okN, errN, pending, progressed };
}

/* ---------- helpers ---------- */
function findRow_(sh, code) {
  const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) if (String(col[i][0]) === String(code)) return i + 1;
  return 0;
}
function extraRow_(code) {
  const ss = ss_();
  let sh = ss.getSheetByName(CFG.SHEET_EXTRA);
  if (!sh) { setup(); sh = ss.getSheetByName(CFG.SHEET_EXTRA); }
  let row = findRow_(sh, code);
  if (!row) { sh.appendRow([String(code), '', '', '', '', '', false, '']); row = sh.getLastRow(); sh.getRange(row, 1).setNumberFormat('@').setValue(String(code)); }
  return { sh, row };
}
function folder_() {
  if (CFG.PHOTO_FOLDER_ID) return DriveApp.getFolderById(CFG.PHOTO_FOLDER_ID);
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('PHOTO_FOLDER');
  if (id) return DriveApp.getFolderById(id);
  const f = DriveApp.createFolder('THAMC Item Photos');
  p.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}
function log_(user, action, code, detail) {
  const sh = ss_().getSheetByName(CFG.SHEET_LOG);
  if (sh) sh.appendRow([new Date(), user || '', action, code, detail]);
}
function auth_(t) { if (CFG.TOKEN && t !== CFG.TOKEN) throw new Error('unauthorized'); }
function split_(v) { return String(v || '').split(SEP).map(s => s.trim()).filter(Boolean); }
function driveId_(v) { const m = String(v || '').match(/[-\w]{25,}/); return m ? m[0] : ''; }
function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Bangkok', 'dd MMM yyyy') : String(v || ''); }
function items_() {
  const s = ss_();
  const sh = s.getSheetByName(CFG.SHEET_ITEMS) || s.getSheets().find(x => ![CFG.SHEET_EXTRA, CFG.SHEET_LOG, CFG.SHEET_GROUPS].includes(x.getName()));
  if (!sh) throw new Error('ไม่พบแท็บข้อมูลสินค้า (ตั้งชื่อแท็บ ' + CFG.SHEET_ITEMS + ' หรือวางข้อมูลในแท็บแรก)');
  return sh;
}
function ss_() {
  const s = CFG.SPREADSHEET_ID ? SpreadsheetApp.openById(CFG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
  if (!s) throw new Error('ไม่พบ Google Sheet — ใส่ CFG.SPREADSHEET_ID');
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
