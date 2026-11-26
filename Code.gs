// Deposit Digital backend (v1.8.0) — วางใน Apps Script ของ Google Sheet (Extensions > Apps Script)
//
// ตั้งค่า
// 1) สร้าง Token: ในแอปกด "🎲 สร้าง Token สุ่ม" (แนะนำ) แล้ววางที่ Project Settings > Script properties > TOKEN
//    หรือเปิดสคริปต์นี้แล้วรันฟังก์ชัน setupToken() หนึ่งครั้ง ดู Token ใน Execution log แล้วคัดลอกไปใส่ในแอป
// 2) Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
// 3) นำ URL (ลงท้าย /exec) และ TOKEN ไปใส่ในแอป
//
// หมายเหตุความปลอดภัย: Token ต้องยาวอย่างน้อย 24 ตัวอักษร และข้อมูลที่รับต้องเป็นไฟล์เข้ารหัสด้วยรหัสผ่านสำรอง (bp=true)
const FOLDER = 'DepositDigital Backups';
const CHUNK = 40000;          // ขนาดต่อเซลล์ (จำกัดของ Sheets = 50,000 ตัวอักษร)
const MIN_TOKEN = 24;
const MAX_BYTES = 8000000;    // ขนาดสูงสุดของไฟล์สำรอง (ตัวอักษร)
const LOCK_MS = 30000;

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out({ ok: true, app: 'Deposit Digital' }); }

// รันครั้งเดียวในตัวแก้ไข เพื่อสร้าง TOKEN สุ่ม 244 บิตและบันทึกลง Script properties
function setupToken() {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', t);
  Logger.log('TOKEN = ' + t + '\nคัดลอกไปใส่ในแอป แล้วลบ log นี้ได้');
}

// เทียบข้อความแบบใช้เวลาเท่ากัน ลดช่องให้เดาจากเวลาตอบ
function safeEq(a, b) {
  a = String(a == null ? '' : a); b = String(b == null ? '' : b);
  let d = a.length ^ b.length;
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n; i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return d === 0;
}

function folder() {
  const it = DriveApp.getFoldersByName(FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
}
function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function checkFile(f) {
  if (!f || typeof f !== 'object') return 'ไฟล์สำรองไม่ถูกต้อง';
  if (f.bp !== true || !f.iv || !f.data || !f.salt) return 'ต้องเป็นไฟล์ที่เข้ารหัสด้วยรหัสผ่านสำรองเท่านั้น';
  return null;
}

function doPost(e) {
  let lock = null;
  try {
    const t = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!t || t.length < MIN_TOKEN) return out({ ok: false, error: 'TOKEN ฝั่งสคริปต์สั้นเกินไป (ต้อง ≥ ' + MIN_TOKEN + ' ตัวอักษร)' });
    if (!e || !e.postData || !e.postData.contents || e.postData.contents.length > MAX_BYTES + 1000) return out({ ok: false, error: 'คำขอไม่ถูกต้องหรือใหญ่เกินไป' });
    const r = JSON.parse(e.postData.contents);
    if (!safeEq(r.token, t)) return out({ ok: false, error: 'token ไม่ถูกต้อง' });

    // กันเขียน/อ่านซ้อนกัน (เช่น แอปส่งอัตโนมัติพร้อมกับกดส่งเอง)
    lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_MS)) return out({ ok: false, error: 'ระบบกำลังบันทึกอยู่ ลองใหม่อีกครั้ง' });

    if (r.action === 'sheet_save') {
      const bad = checkFile(r.file); if (bad) return out({ ok: false, error: bad });
      const str = JSON.stringify(r.file);
      if (str.length > MAX_BYTES) return out({ ok: false, error: 'ไฟล์ใหญ่เกินไป' });
      const rows = [];
      for (let i = 0; i < str.length; i += CHUNK) rows.push([str.slice(i, i + CHUNK)]);
      const sh = sheet('Latest');
      // เขียนข้อมูลใหม่ก่อน แล้วค่อยล้างแถวส่วนเกิน — ถ้าพังกลางทางข้อมูลเก่ายังอยู่
      const rng = sh.getRange(1, 1, rows.length, 1);
      rng.setNumberFormat('@');   // บังคับเป็นข้อความ กันเซลล์ที่ขึ้นต้นด้วย = ถูกตีเป็นสูตร
      rng.setValues(rows);
      const last = sh.getLastRow();
      if (last > rows.length) sh.getRange(rows.length + 1, 1, last - rows.length, 1).clearContent();
      sheet('Log').appendRow([new Date(), 'sheet_save', str.length]);
      return out({ ok: true });
    }
    if (r.action === 'sheet_get') {
      const sh = sheet('Latest'), n = sh.getLastRow();
      if (n < 1) return out({ ok: false, error: 'ยังไม่มีข้อมูลในชีต' });
      const str = sh.getRange(1, 1, n, 1).getValues().map(x => x[0]).join('');
      return out({ ok: true, file: JSON.parse(str) });
    }
    if (r.action === 'drive_save') {
      const bad = checkFile(r.file); if (bad) return out({ ok: false, error: bad });
      const str = JSON.stringify(r.file);
      if (str.length > MAX_BYTES) return out({ ok: false, error: 'ไฟล์ใหญ่เกินไป' });
      const name = 'DepositDigital_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '.json';
      const f = folder().createFile(name, str, MimeType.PLAIN_TEXT);
      return out({ ok: true, id: f.getId(), name: name });
    }
    if (r.action === 'list') {
      const files = [], it = folder().getFiles();
      while (it.hasNext()) { const f = it.next(); files.push({ id: f.getId(), name: f.getName(), t: f.getDateCreated().getTime() }); }
      files.sort((a, b) => b.t - a.t);
      return out({ ok: true, files: files.slice(0, 20) });
    }
    if (r.action === 'drive_get') {
      const f = DriveApp.getFileById(String(r.id));
      const ps = f.getParents();
      if (!ps.hasNext() || ps.next().getId() !== folder().getId()) return out({ ok: false, error: 'ไฟล์อยู่นอกโฟลเดอร์สำรอง' });
      return out({ ok: true, file: JSON.parse(f.getBlob().getDataAsString()) });
    }
    return out({ ok: false, error: 'unknown action' });
  } catch (x) {
    return out({ ok: false, error: String(x) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (_) {} }
  }
}
