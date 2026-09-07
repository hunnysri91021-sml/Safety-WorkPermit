/**
 * SHE Patrol Digital System — Google Apps Script backend
 * Siam Motors Logistics Co., Ltd. (SML)
 *
 * ทำหน้าที่แทน SharePoint + Power Automate เดิม เพื่อให้พนักงานที่ไม่มี license
 * Microsoft 365 ใช้งานได้เหมือนกันทุกคน — ไม่มีค่าใช้จ่ายเพิ่ม (Google Sheets/Apps Script
 * ฟรีสำหรับผู้เข้าถึงผ่าน Web App ที่ deploy แบบ "Anyone")
 *
 * โครงสร้างและเทคนิคยึดตาม WorkPermit-AppsScript.gs ของโปรเจกต์พี่น้อง (Safety-WorkPermit)
 * ที่ใช้งานจริงอยู่แล้ว — โดยเฉพาะ:
 *   - front-end ต้องส่ง Content-Type: text/plain (ไม่ใช่ application/json) กัน browser
 *     ยิง CORS preflight (OPTIONS) ซึ่ง Apps Script Web App ไม่รองรับ — ฝั่งนี้ parse
 *     เนื้อหาเป็น JSON ตามปกติจาก e.postData.contents
 *   - LockService กันข้อมูลชนกันตอนหลายคนบันทึกพร้อมกัน
 *   - DriveApp เก็บรูป ตั้งค่า ANYONE_WITH_LINK/VIEW ให้เปิดดูได้ตรงจาก <img src> เลย
 *     (ง่ายกว่าฝั่ง SharePoint เดิมที่ต้องมี action getPhoto แยกเพราะ SharePoint บังคับ auth)
 *
 * วิธี deploy:
 *   1. เปิด Google Sheet ใหม่ (ว่างเปล่า) → Extensions → Apps Script
 *   2. วางโค้ดไฟล์นี้ทั้งหมดแทนโค้ดเริ่มต้น → บันทึก
 *   3. รันฟังก์ชัน setupSheet() หนึ่งครั้ง (เลือกจาก dropdown ด้านบน แล้วกด Run) —
 *      จะสร้างชีต Findings / Users / Settings ให้อัตโนมัติ อนุญาตสิทธิ์ตามที่ถาม
 *   4. กรอกอีเมลต่างๆ ในชีต "Settings" (Frontend_Base_URL, Safety_Officer_Email, ฯลฯ)
 *   5. เพิ่มแถวแรกในชีต "Users" ด้วยตัวเอง (Admin คนแรก) — ดูหมายเหตุใน README
 *   6. Deploy → New deployment → เลือกประเภท "Web app"
 *        Execute as: Me
 *        Who has access: Anyone
 *      กด Deploy แล้วคัดลอก URL ที่ได้ (ลงท้ายด้วย /exec)
 *   7. นำ URL ไปใส่ใน APP_CONFIG.API_URL ของ index.html — และตรวจว่า APP_CONFIG.API_TOKEN
 *      ในไฟล์เดียวกันตรงกับค่า API_TOKEN ด้านล่างนี้เป๊ะ (กัน API ถูกเรียกตรง ๆ นอกหน้าเว็บ)
 *   8. (ทำครั้งเดียว) ตั้ง time-driven trigger ให้ checkSlaEscalation() รันทุกวัน:
 *      Apps Script editor → รูปนาฬิกา (Triggers) → Add Trigger →
 *      Function: checkSlaEscalation, Event source: Time-driven, Day timer, 8-9am
 *   9. (แนะนำ ถ้าบัญชี Google ที่ใช้เป็นบัญชีส่วนตัว ไม่ใช่ M365 tenant ของบริษัท) กรอก
 *      MS365_Backup_Email ในแท็บ Settings เป็นอีเมล Outlook ของบริษัท แล้วตั้ง time-driven
 *      trigger อีกตัว: Function: backupToMS365Email, Event source: Time-driven, Week timer
 *      — จะได้ไฟล์สำรอง .xlsx ส่งเข้าอีเมล M365 ทุกสัปดาห์ ตั้ง Power Automate flow ตาม
 *      power-automate/Flow6-Backup-To-OneDrive.md เพิ่มถ้าอยากให้เซฟเข้า OneDrive อัตโนมัติด้วย
 *   10. ล็อกอินตอนนี้ต้องใส่รหัสผ่านด้วย (เก็บเป็น SHA-256 hash ในคอลัมน์ Password ของชีต Users
 *      ไม่ใช่ plaintext) — Admin คนแรกที่เพิ่มในข้อ 5 ยังไม่มีรหัสผ่าน เข้าแอปไม่ได้จนกว่าจะตั้งรหัส
 *      ผ่านให้ตัวเองก่อน ใช้เมนู "SHE Patrol" → "ตั้งรหัสผ่านเริ่มต้นให้ผู้ใช้งาน (bootstrap)" ใน
 *      Google Sheet เพื่อตั้งรหัสผ่านแรกโดยไม่ต้องล็อกอินเข้าแอปก่อน — หลังจากนั้น Admin จัดการ
 *      รหัสผ่านของคนอื่นต่อได้ผ่านเมนู "จัดการผู้ใช้งาน" ในแอปตามปกติ
 */

const FINDINGS_SHEET_NAME = "Findings";
const USERS_SHEET_NAME = "Users";
const SETTINGS_SHEET_NAME = "Settings";
const LOG_SHEET_NAME = "ActivityLog";
const PHOTOS_FOLDER_NAME = "SHE Patrol Photos";

// กันเรียก API ตรง ๆ ข้าม index.html — ต้องตรงกับ APP_CONFIG.API_TOKEN ในไฟล์ index.html
// (ไม่ใช่การป้องกันแบบสมบูรณ์ เพราะไฟล์ index.html เปิดสาธารณะอยู่แล้วใครอ่านโค้ดก็เจอค่านี้ได้
//  แต่กันการเรียกสุ่ม/บอทเก็บข้อมูลอัตโนมัติที่ไม่ได้อ่านโค้ดหน้าเว็บก่อน)
const API_TOKEN = "1cc4f7797f5b8ddfe61452bd40d8845967f77ffeb1881bb6";

const FINDINGS_HEADERS = [
  "Id", "TypeOfAudit", "PatrolDate", "Shop", "Place", "Description", "Grade", "Category",
  "PhotoBeforeUrl", "DueDate", "RootCause", "ActionResponsible", "Countermeasure",
  "PhotoAfterUrl", "Status", "VerifiedBy", "Rules_Confirmed_DateTime",
];
const USERS_HEADERS = ["Id", "Name", "Email", "Role", "Shop", "Active", "Password", "LastActiveAt"];
const ONLINE_THRESHOLD_MS = 2 * 60 * 1000; // ถือว่า "กำลังใช้งานอยู่" ถ้า heartbeat ล่าสุดไม่เกิน 2 นาที
const LOG_HEADERS = ["Timestamp", "UserId", "UserName", "Action", "TargetType", "TargetId", "Details"];

const STATUS_OPEN = ["เปิดใหม่", "รอดำเนินการ", "ดำเนินการแล้ว", "รอตรวจสอบ"];

const DEFAULT_SETTINGS = [
  ["Key", "Value", "คำอธิบาย"],
  ["Frontend_Base_URL", "", "URL ของเว็บที่โฮสต์ index.html เช่น https://ชื่อบัญชี.github.io/SHE-Patrol/ — ใช้สร้าง deep link ในอีเมล"],
  ["TypeOfAudit_Options", "Planned,Unplanned,Follow-up,Special", "รายการตัวเลือก 'ประเภทการตรวจ' ในฟอร์มบันทึกรายการตรวจใหม่ — คั่นแต่ละตัวเลือกด้วยจุลภาค (,) แก้ไขแล้วมีผลทันที"],
  ["Shop_Options", "PDI,Acc,Yard,Washing,Touch up,Store,อื่นๆ", "รายการตัวเลือก 'หน่วยงาน/Shop' ในฟอร์ม/ตัวกรอง/จัดการผู้ใช้งาน — คั่นด้วยจุลภาค (,) — หมายเหตุ: ชื่อ Shop ที่เพิ่มใหม่จะยังไม่มีอีเมลแจ้งเตือนเฉพาะจนกว่าจะเพิ่มคีย์ Shop_Email_<ชื่อ Shop> เอง ไม่งั้นจะ fallback ไปที่ Shop_Email_อื่นๆ"],
  ["Category_Options", "F,S,EES,5S", "รายการตัวเลือก 'ประเภท (STD)' ในฟอร์มบันทึกรายการตรวจใหม่ — คั่นด้วยจุลภาค (,) แก้ไขแล้วมีผลทันที"],
  ["SLA_Days_A", "7", "จำนวนวันจนถึงกำหนดปิดงาน (Due Date) สำหรับ Grade A — คำนวณอัตโนมัติตอนสร้างรายการตรวจใหม่ (วันที่ตรวจ + จำนวนวันนี้) แก้ไขแล้วมีผลกับรายการที่สร้างใหม่เท่านั้น ไม่ย้อนแก้รายการเก่า"],
  ["SLA_Days_B", "14", "จำนวนวันจนถึงกำหนดปิดงานสำหรับ Grade B"],
  ["SLA_Days_C", "30", "จำนวนวันจนถึงกำหนดปิดงานสำหรับ Grade C"],
  ["SLA_Days_Others", "30", "จำนวนวันจนถึงกำหนดปิดงานสำหรับ Grade Others"],
  ["Safety_Officer_Email", "", "รับแจ้งเตือนตอนมีรายการ 'รอตรวจสอบ'"],
  ["MGR_Email", "", "อีเมล MGR สำหรับ escalate Grade A/B ที่เลย/ใกล้กำหนด"],
  ["AGMGM_Email", "", "อีเมล AGM/GM สำหรับ escalate Grade A ที่เลย/ใกล้กำหนด"],
  ["Shop_Email_PDI", "", "อีเมลหัวหน้างาน PDI"],
  ["Shop_Email_Acc", "", "อีเมลหัวหน้างาน Acc"],
  ["Shop_Email_Yard", "", "อีเมลหัวหน้างาน Yard"],
  ["Shop_Email_Washing", "", "อีเมลหัวหน้างาน Washing"],
  ["Shop_Email_Touch up", "", "อีเมลหัวหน้างาน Touch up"],
  ["Shop_Email_Store", "", "อีเมลหัวหน้างาน Store"],
  ["Shop_Email_อื่นๆ", "", "อีเมลหัวหน้างาน อื่นๆ (ใช้เป็น fallback ด้วยถ้า Shop อื่นไม่ได้กรอกอีเมล)"],
  ["SLA_Escalation_Days_Before", "2", "แจ้งเตือนล่วงหน้ากี่วันก่อนถึงกำหนด (checkSlaEscalation รันทุกวัน)"],
  ["MS365_Backup_Email", "", "อีเมล Microsoft 365 (Outlook) ที่จะรับไฟล์สำรองข้อมูล .xlsx รายสัปดาห์ — เว้นว่างถ้ายังไม่ต้องการสำรอง (ดู backupToMS365Email + power-automate/Flow6-Backup-To-OneDrive.md)"],
];

// ---------------------------------------------------------------
// ติดตั้ง/อัปเดตโครงสร้างชีตทั้งหมด (รันซ้ำได้อย่างปลอดภัย)
// ---------------------------------------------------------------
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let findingsSheet = ss.getSheetByName(FINDINGS_SHEET_NAME);
  if (!findingsSheet) findingsSheet = ss.insertSheet(FINDINGS_SHEET_NAME);
  // Migration: ชีตเก่าไม่มีคอลัมน์ TypeOfAudit (คอลัมน์ B เดิมคือ PatrolDate) — แทรกคอลัมน์ว่างที่
  // ตำแหน่ง B ก่อน เพื่อไม่ให้ข้อมูลเดิมเลื่อนทับผิดคอลัมน์ ทำครั้งเดียว (เช็ค header เดิมก่อนแทรก)
  const existingFindingsHeaderB = findingsSheet.getRange(1, 2).getValue();
  if (existingFindingsHeaderB && existingFindingsHeaderB !== "TypeOfAudit") {
    findingsSheet.insertColumnBefore(2);
  }
  findingsSheet.getRange(1, 1, 1, FINDINGS_HEADERS.length).setValues([FINDINGS_HEADERS]).setFontWeight("bold");
  findingsSheet.setFrozenRows(1);
  // เก็บวันที่/เวลาเป็นข้อความล้วน กัน Sheets auto-parse เป็น Date แล้วรูปแบบเพี้ยน
  ["PatrolDate", "DueDate", "Rules_Confirmed_DateTime"].forEach(h => {
    const col = FINDINGS_HEADERS.indexOf(h) + 1;
    findingsSheet.getRange(1, col, 2000, 1).setNumberFormat("@");
  });

  let usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) usersSheet = ss.insertSheet(USERS_SHEET_NAME);
  usersSheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]).setFontWeight("bold");
  usersSheet.setFrozenRows(1);

  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  // เก็บคอลัมน์ Value เป็นข้อความล้วน กัน Sheets auto-parse ค่าที่ดูเหมือนตัวเลข (เช่น
  // SLA_Escalation_Days_Before = "2") ให้กลายเป็นตัวเลขจริง ซึ่งทำให้หน้าตั้งค่าใน index.html พังได้
  settingsSheet.getRange(1, 2, 2000, 1).setNumberFormat("@");
  settingsSheet.getRange(1, 1, 1, 3).setValues([DEFAULT_SETTINGS[0]]);
  // เติมเฉพาะ "คีย์ที่ยังไม่มี" ต่อท้ายชีตเสมอ — เทียบตามคีย์ ไม่ใช่ตำแหน่ง/จำนวนแถว เพราะถ้าเทียบ
  // แค่จำนวนแถว การแทรกคีย์ใหม่ตรงกลาง DEFAULT_SETTINGS (ไม่ใช่ต่อท้าย) จะทำให้ชีตที่เคยรัน setupSheet
  // ไปแล้วได้ค่าเพี้ยน (index ไม่ตรงของเดิม)
  const existingRows = settingsSheet.getLastRow();
  const existingKeys = existingRows >= 2 ? settingsSheet.getRange(2, 1, existingRows - 1, 1).getValues().flat() : [];
  const missingRows = DEFAULT_SETTINGS.slice(1).filter(row => existingKeys.indexOf(row[0]) === -1);
  if (missingRows.length) {
    settingsSheet.getRange(Math.max(existingRows, 1) + 1, 1, missingRows.length, 3).setValues(missingRows);
  }
  // แปลงค่าที่มีอยู่แล้วให้เป็นข้อความ เผื่อเคยถูก Sheets auto-parse เป็นตัวเลข/boolean ไปก่อนหน้านี้
  if (existingRows >= 2) {
    const valRange = settingsSheet.getRange(2, 2, existingRows - 1, 1);
    valRange.setValues(valRange.getValues().map(row => [row[0] === "" ? "" : String(row[0])]));
  }
  settingsSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  settingsSheet.setFrozenRows(1);
  settingsSheet.autoResizeColumns(1, 3);

  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) logSheet = ss.insertSheet(LOG_SHEET_NAME);
  // เก็บ Timestamp เป็นข้อความล้วน (ISO string) กัน Sheets auto-parse เป็น Date แล้วรูปแบบเพี้ยน
  // เหมือนคอลัมน์วันที่ใน Findings — เรียงตามข้อความ ISO ได้ผลเหมือนเรียงตามเวลาจริงอยู่แล้ว
  logSheet.getRange(1, 1, 2000, 1).setNumberFormat("@");
  logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight("bold");
  logSheet.setFrozenRows(1);
  logSheet.autoResizeColumns(1, LOG_HEADERS.length);

  SpreadsheetApp.getUi().alert("ตั้งค่าชีตเรียบร้อย — ไปกรอกอีเมลในแท็บ Settings แล้วเพิ่ม Admin คนแรกในแท็บ Users ต่อได้เลย");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SHE Patrol")
    .addItem("ตั้งค่าชีต (Setup)", "setupSheet")
    .addItem("ตั้งรหัสผ่านเริ่มต้นให้ผู้ใช้งาน (bootstrap)", "setInitialAdminPassword")
    .addItem("แก้ลิงก์รูปที่โหลดไม่ขึ้น (รูปแตก)", "fixBrokenPhotoUrls")
    .addItem("เช็ค SLA เลยกำหนด (ทดสอบ)", "checkSlaEscalation")
    .addItem("สำรองข้อมูลเข้า MS365 ตอนนี้ (ทดสอบ)", "backupToMS365Email")
    .addToUi();
}

// ---------------------------------------------------------------
// รูปที่อัปโหลดก่อนหน้านี้ (ก่อนเปลี่ยนมาใช้ endpoint "thumbnail") ใช้ลิงก์แบบ
// "uc?export=view" ซึ่ง Google มักเด้งหน้า "ไม่สามารถสแกนไวรัสได้" แทนรูปจริง ทำให้
// รูปที่บันทึกไปแล้วเห็นเป็นไอคอนแตก — รันเมนูนี้ครั้งเดียวเพื่อแปลงลิงก์เก่าทุกอันในชีต
// Findings ให้เป็นรูปแบบใหม่ที่โหลดได้จริง (รันซ้ำได้ปลอดภัย ไม่กระทบรูปที่แก้ไปแล้ว)
function fixBrokenPhotoUrls() {
  const sheet = getSheet_(FINDINGS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("ยังไม่มีข้อมูลในชีต Findings");
    return;
  }
  const oldPattern = /^https:\/\/drive\.google\.com\/uc\?export=view&id=([^&]+)$/;
  ["PhotoBeforeUrl", "PhotoAfterUrl"].forEach(header => {
    const col = FINDINGS_HEADERS.indexOf(header) + 1;
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const values = range.getValues();
    let fixedCount = 0;
    const fixed = values.map(row => {
      const url = row[0];
      const match = typeof url === "string" && url.match(oldPattern);
      if (!match) return row;
      fixedCount++;
      return [`https://drive.google.com/thumbnail?id=${match[1]}&sz=w1000`];
    });
    if (fixedCount > 0) range.setValues(fixed);
  });
  SpreadsheetApp.getUi().alert("แก้ลิงก์รูปเก่าที่โหลดไม่ขึ้นเรียบร้อย — รีเฟรชหน้าเว็บแล้วลองเปิดดูรายการที่รูปเคยแตกอีกครั้ง");
}

// ---------------------------------------------------------------
// ใช้ตอน bootstrap ครั้งแรก (หรือกู้บัญชี Admin ที่ล็อกอินไม่ได้) — ตั้งรหัสผ่านให้ผู้ใช้งาน
// คนใดคนหนึ่งได้โดยไม่ต้องล็อกอินเข้าแอปก่อน เพราะรันจากเมนูใน Google Sheet โดยตรง
// ---------------------------------------------------------------
function setInitialAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const emailResp = ui.prompt("ตั้งรหัสผ่านผู้ใช้งาน", "อีเมลของผู้ใช้งานที่จะตั้งรหัสผ่าน (ต้องมีอยู่ในแท็บ Users แล้ว):", ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const email = emailResp.getResponseText().trim();

  const pwResp = ui.prompt("ตั้งรหัสผ่านผู้ใช้งาน", "รหัสผ่านใหม่สำหรับ " + email + ":", ui.ButtonSet.OK_CANCEL);
  if (pwResp.getSelectedButton() !== ui.Button.OK) return;
  const password = pwResp.getResponseText();
  if (!password) { ui.alert("กรุณาใส่รหัสผ่าน — ยกเลิกการตั้งค่า"); return; }

  const rows = readAllRows_(USERS_SHEET_NAME, USERS_HEADERS);
  const row = rows.find(r => String(r.Email).toLowerCase() === email.toLowerCase());
  if (!row) { ui.alert("ไม่พบอีเมลนี้ในแท็บ Users — เพิ่มผู้ใช้งานคนนี้ก่อน"); return; }

  const sheet = getSheet_(USERS_SHEET_NAME);
  const rowNum = findRowNumberById_(USERS_SHEET_NAME, USERS_HEADERS, row.Id);
  const pwCol = USERS_HEADERS.indexOf("Password") + 1;
  sheet.getRange(rowNum, pwCol).setValue(hashPassword_(password));
  ui.alert("ตั้งรหัสผ่านให้ " + row.Name + " เรียบร้อยแล้ว — ล็อกอินเข้าแอปด้วยรหัสผ่านนี้ได้เลย");
}

function getSetting(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) return "";
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return "";
}

function shopLeadEmail_(shop) {
  return getSetting("Shop_Email_" + shop) || getSetting("Shop_Email_อื่นๆ");
}

// ---------------------------------------------------------------
// Settings — อ่าน/แก้ไขทั้งหมดผ่านแท็บ "ตั้งค่า" ในแอป (Admin เท่านั้น)
// แทนการเปิด Google Sheet ไปแก้ทีละแถวในแท็บ Settings เอง
// ---------------------------------------------------------------
function listSettings_() {
  const sheet = getSheet_(SETTINGS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(row => row[0] !== "")
    .map(row => ({ Key: row[0], Value: row[1], Description: row[2] }));
}

// เช็คว่า Web App ยังตอบสนองอยู่จริง (ใช้แสดงสถานะออนไลน์/ออฟไลน์ใน index.html)
// ไม่แตะชีตเลย จึงเบาและไม่ติด LockService
function ping_() {
  return { status: "ok", serverTime: new Date().toISOString() };
}

function updateSettings_(values, actingUser) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const changedParts = [];
  try {
    const sheet = getSheet_(SETTINGS_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    Object.keys(values || {}).forEach(k => {
      const idx = keys.indexOf(k);
      if (idx === -1) return;
      const before = sheet.getRange(idx + 2, 2).getValue();
      if (String(before) === String(values[k])) return;
      sheet.getRange(idx + 2, 2).setValue(values[k]);
      changedParts.push(`${k}: "${before}" → "${values[k]}"`);
    });
  } finally {
    lock.releaseLock();
  }
  if (changedParts.length) logActivity_(actingUser, "แก้ไขตั้งค่า", "Settings", "", changedParts.join(" | "));
  return listSettings_();
}

// ---------------------------------------------------------------
// Router — front-end (index.html) POST { action, ...payload } ทุกครั้ง
// (Content-Type: text/plain ฝั่ง front-end กัน CORS preflight — ดูหมายเหตุหัวไฟล์)
// ---------------------------------------------------------------
function doGet(e) {
  return ContentService.createTextOutput("SHE Patrol API is running. ใช้ POST พร้อม {action, ...} เท่านั้น")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: "รูปแบบข้อมูลไม่ถูกต้อง" });
  }

  if (data.token !== API_TOKEN) {
    return jsonResponse({ ok: false, error: "ไม่ได้รับอนุญาต" });
  }

  // ทุก action ยกเว้น 3 ตัวนี้ต้องมี session token ที่ยังไม่หมดอายุ (ออกให้ตอน login_() สำเร็จ) —
  // ป้องกันไม่ให้ใครก็ตามที่รู้แค่ API_TOKEN (ซึ่งฝังอยู่ใน index.html แบบสาธารณะ) เรียกข้อมูลจริง
  // ได้โดยไม่เคย login เลย — session token เซ็นด้วย secret ฝั่งเซิร์ฟเวอร์ที่ไม่เคยส่งให้ client
  const PUBLIC_ACTIONS_ = ["login", "loginOptions", "ping"];
  let sessionUser = null;
  if (PUBLIC_ACTIONS_.indexOf(data.action) === -1) {
    sessionUser = verifySessionToken_(data.sessionToken);
    if (!sessionUser) return jsonResponse({ ok: false, error: "เซสชันหมดอายุ กรุณาล็อกอินใหม่", sessionExpired: true });
  }

  try {
    switch (data.action) {
      case "listFindings": return jsonResponse({ ok: true, data: listFindings_() });
      case "getFinding": return jsonResponse({ ok: true, data: getFinding_(data.id) });
      case "createFinding":
        requireRole_(sessionUser, ["admin", "auditor"]);
        return jsonResponse({ ok: true, data: createFinding_(data.fields || {}, data.photos, sessionUser) });
      case "updateFinding":
        requireCanEditFinding_(sessionUser, data.id, data.fields || {});
        return jsonResponse({ ok: true, data: updateFinding_(data.id, data.fields || {}, data.photos, sessionUser) });
      case "listUsers": return jsonResponse({ ok: true, data: listUsers_() });
      case "createUser":
        requireRole_(sessionUser, ["admin"]);
        return jsonResponse({ ok: true, data: createUser_(data.fields || {}, sessionUser) });
      case "updateUser":
        requireRole_(sessionUser, ["admin"]);
        return jsonResponse({ ok: true, data: updateUser_(data.id, data.fields || {}, sessionUser) });
      case "login": return jsonResponse({ ok: true, data: login_(data.id, data.password) });
      case "loginOptions": return jsonResponse({ ok: true, data: loginOptions_() });
      case "uploadPhoto": return jsonResponse({ ok: true, data: uploadPhoto_(data.findingId, data.stage, data.fileName, data.contentBase64) });
      case "listSettings": return jsonResponse({ ok: true, data: listSettings_() });
      case "updateSettings":
        requireRole_(sessionUser, ["admin"]);
        return jsonResponse({ ok: true, data: updateSettings_(data.values || {}, sessionUser) });
      case "listActivityLog":
        requireRole_(sessionUser, ["admin"]);
        return jsonResponse({ ok: true, data: listActivityLog_() });
      case "heartbeat":
        heartbeat_(sessionUser);
        return jsonResponse({ ok: true, data: listOnlineUsers_() });
      case "listOnlineUsers": return jsonResponse({ ok: true, data: listOnlineUsers_() });
      case "ping": return jsonResponse({ ok: true, data: ping_() });
      default: return jsonResponse({ ok: false, error: "ไม่รู้จัก action: " + data.action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// Generic sheet-as-table helpers (ใช้ร่วมกันทั้ง Findings และ Users)
// ---------------------------------------------------------------
function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`ไม่พบชีต '${name}' — รัน setupSheet() ก่อน`);
  return sheet;
}

function readAllRows_(sheetName, headers) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .map(row => rowToObject_(headers, row))
    .filter(r => r.Id !== "" && r.Id !== null && r.Id !== undefined);
}

function rowToObject_(headers, rowValues) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = rowValues[i]; });
  return obj;
}

function findRowNumberById_(sheetName, headers, id) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const idCol = headers.indexOf("Id") + 1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat().map(Number);
  const idx = ids.indexOf(Number(id));
  return idx === -1 ? -1 : idx + 2;
}

function nextId_(sheetName, headers) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const idCol = headers.indexOf("Id") + 1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat().map(Number).filter(n => !isNaN(n));
  return ids.length ? Math.max.apply(null, ids) + 1 : 1;
}

// ---------------------------------------------------------------
// Activity Log — บันทึกทุกครั้งที่มี login/สร้าง/แก้ไข/ลบ เพื่อตรวจสอบย้อนหลังได้ว่า
// ใครทำอะไรกับข้อมูลไหน เมื่อไหร่ — เขียนนอก LockService lock ของ action หลักเสมอ (กัน
// เพิ่ม log ช้าไปถ่วงเวลา request หลัก) และห่อ try/catch ไว้เสมอ (log พังต้องไม่ทำให้
// action หลักที่ผู้ใช้กำลังรออยู่พังตาม เช่นกรณีชีตยังไม่ได้รัน setupSheet() เวอร์ชันใหม่)
// ---------------------------------------------------------------
function logActivity_(user, action, targetType, targetId, details) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    if (!sheet) return;
    sheet.appendRow([
      new Date().toISOString(),
      user ? user.Id : "",
      user ? user.Name : "",
      action,
      targetType || "",
      targetId !== undefined && targetId !== null ? targetId : "",
      details || "",
    ]);
  } catch (err) {
    // เงียบไว้ — ดู log ไม่ได้ยังดีกว่า action หลักใช้งานไม่ได้
  }
}

// เทียบค่าเดิม/ใหม่เฉพาะ key ที่ถูกส่งมาแก้ (ไม่ใช่ทั้งแถว) แล้วคืนข้อความสรุปแบบ
// "field: เดิม → ใหม่" คั่นด้วย " | " — ใช้กับทั้ง Finding/User/Settings เพราะโครงสร้างเหมือนกัน
function diffText_(before, after, keys, redactKeys) {
  const redact = redactKeys || [];
  const parts = [];
  keys.forEach(k => {
    const b = before[k] === undefined || before[k] === null ? "" : before[k];
    const a = after[k] === undefined || after[k] === null ? "" : after[k];
    if (String(b) === String(a)) return;
    if (redact.indexOf(k) !== -1) { parts.push(`${k}: (เปลี่ยนแล้ว)`); return; }
    parts.push(`${k}: "${b}" → "${a}"`);
  });
  return parts.join(" | ");
}

// เลือก action label ให้อ่านง่ายตามฟิลด์ที่แก้ — ปิดงาน/เปลี่ยนสถานะ/แก้ข้อมูลรายการ/แก้ความคืบหน้า
function findingActionLabel_(fields) {
  const detailFields = ["TypeOfAudit", "PatrolDate", "Shop", "Place", "Description", "Grade", "Category"];
  if (fields.Status === "ปิดงาน") return "ปิดงาน";
  if (fields.Status) return "เปลี่ยนสถานะ";
  if (Object.keys(fields).some(k => detailFields.includes(k))) return "แก้ไขข้อมูลรายการ";
  return "แก้ไขความคืบหน้า";
}

function listActivityLog_() {
  // ไม่ใช้ readAllRows_() เพราะฟังก์ชันนั้น filter แถวว่างด้วยคอลัมน์ "Id" ซึ่งชีต Log
  // ไม่มี (คอลัมน์แรกคือ Timestamp) — อ่านตรงแล้ว filter ด้วย Timestamp แทน
  const sheet = getSheet_(LOG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues()
    .map(row => rowToObject_(LOG_HEADERS, row))
    .filter(r => r.Timestamp !== "" && r.Timestamp !== null && r.Timestamp !== undefined);
  // ใหม่สุดก่อน — เรียงตาม Timestamp (ISO string เรียงเป็นลำดับเวลาได้ตรงอยู่แล้ว)
  rows.sort((a, b) => String(b.Timestamp).localeCompare(String(a.Timestamp)));
  // จำกัดจำนวนแถวที่ส่งกลับ กันชีต log โตมากแล้ว payload ใหญ่/ช้าเกินไป
  return rows.slice(0, 1000);
}

// ---------------------------------------------------------------
// Findings CRUD
// ---------------------------------------------------------------
function listFindings_() {
  return readAllRows_(FINDINGS_SHEET_NAME, FINDINGS_HEADERS).sort((a, b) => b.Id - a.Id);
}

function getFinding_(id) {
  const row = readAllRows_(FINDINGS_SHEET_NAME, FINDINGS_HEADERS).find(r => r.Id === Number(id));
  if (!row) throw new Error("ไม่พบ finding รหัส " + id);
  return row;
}

function createFinding_(fields, photos, actingUser) {
  // อัปโหลดรูปก่อนเข้า lock (ถ้ามี) — รวมสร้างรายการ + แนบรูป (ได้หลายรูป) เป็นคำขอเดียว
  // แทนที่จะสร้างก่อนแล้วค่อยยิงอีกรอบเพื่อแนบรูป (ลด round-trip ไป Apps Script จาก 3 ครั้งเหลือ 1)
  if (photos && photos.length) {
    fields = Object.assign({}, fields, { PhotoBeforeUrl: uploadPhotoFiles_("ก่อนแก้ไข", photos) });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let record;
  try {
    const id = nextId_(FINDINGS_SHEET_NAME, FINDINGS_HEADERS);
    record = Object.assign({
      Id: id, PhotoBeforeUrl: "", PhotoAfterUrl: "", RootCause: "", ActionResponsible: "",
      Countermeasure: "", Status: "เปิดใหม่", VerifiedBy: "", Rules_Confirmed_DateTime: "",
    }, fields, { Id: id });
    const row = FINDINGS_HEADERS.map(h => (record[h] !== undefined && record[h] !== null) ? record[h] : "");
    getSheet_(FINDINGS_SHEET_NAME).appendRow(row);
  } finally {
    lock.releaseLock();
  }
  // ส่งอีเมล/บันทึก log นอก lock — กันคนอื่นที่กำลังจะบันทึกพร้อมกันต้องรอ
  notifyNewFinding_(record);
  logActivity_(actingUser, "สร้างรายการใหม่", "Finding", record.Id, `${record.Shop} · ${record.Place || "-"} · Grade ${record.Grade}`);
  return record;
}

function updateFinding_(id, fields, photos, actingUser) {
  // เช่นเดียวกับ createFinding_ — อัปโหลดรูป (ได้หลายรูป) ก่อนเข้า lock แล้วรวมเป็นคำขอเดียว
  // รูปใหม่จะ "ต่อท้าย" รูปเดิมของรายการ (คั่นด้วยจุลภาคในคอลัมน์เดียวกัน) ไม่ใช่แทนที่ —
  // เลยต้องอ่านค่าปัจจุบันก่อน (นอก lock เพื่อไม่ให้ยึด lock ไว้นานระหว่างอัปโหลดรูปที่ช้ากว่า)
  const beforeList = photos && photos.before ? photos.before : [];
  const afterList = photos && photos.after ? photos.after : [];
  const photoNote = [];
  if (beforeList.length || afterList.length) {
    fields = Object.assign({}, fields);
    const current = getFinding_(id);
    if (beforeList.length) {
      const newUrls = uploadPhotoFiles_("ก่อนแก้ไข", beforeList);
      fields.PhotoBeforeUrl = [current.PhotoBeforeUrl, newUrls].filter(Boolean).join(",");
      photoNote.push(`แนบรูปก่อนแก้ไขเพิ่ม ${beforeList.length} รูป`);
    }
    if (afterList.length) {
      const newUrls = uploadPhotoFiles_("หลังแก้ไข", afterList);
      fields.PhotoAfterUrl = [current.PhotoAfterUrl, newUrls].filter(Boolean).join(",");
      photoNote.push(`แนบรูปหลังแก้ไขเพิ่ม ${afterList.length} รูป`);
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let before, after;
  try {
    const rowNum = findRowNumberById_(FINDINGS_SHEET_NAME, FINDINGS_HEADERS, id);
    if (rowNum === -1) throw new Error("ไม่พบ finding รหัส " + id);
    const sheet = getSheet_(FINDINGS_SHEET_NAME);
    before = rowToObject_(FINDINGS_HEADERS, sheet.getRange(rowNum, 1, 1, FINDINGS_HEADERS.length).getValues()[0]);
    after = Object.assign({}, before, fields, { Id: before.Id });
    const row = FINDINGS_HEADERS.map(h => (after[h] !== undefined && after[h] !== null) ? after[h] : "");
    sheet.getRange(rowNum, 1, 1, FINDINGS_HEADERS.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  // ส่งอีเมล/บันทึก log นอก lock — กันคนอื่นที่กำลังจะบันทึกพร้อมกันต้องรอ
  notifyFindingStatusChange_(before, after);
  const changedKeys = FINDINGS_HEADERS.filter(h => h !== "PhotoBeforeUrl" && h !== "PhotoAfterUrl" && Object.prototype.hasOwnProperty.call(fields, h));
  const details = [diffText_(before, after, changedKeys), ...photoNote].filter(Boolean).join(" | ");
  logActivity_(actingUser, findingActionLabel_(fields), "Finding", id, details);
  return after;
}

// ---------------------------------------------------------------
// Users CRUD — Password เก็บเป็น SHA-256 hash เท่านั้น ไม่ใช่ plaintext และ
// ไม่ส่งกลับไปให้ client เด็ดขาด (sanitizeUser_ ตัดออกก่อน return ทุกครั้ง)
// ---------------------------------------------------------------
function hashPassword_(plain) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(plain), Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, "0")).join("");
}

function sanitizeUser_(u) {
  const copy = Object.assign({}, u);
  delete copy.Password;
  return copy;
}

function listUsers_() {
  return readAllRows_(USERS_SHEET_NAME, USERS_HEADERS).map(sanitizeUser_);
}

// เฉพาะไว้แสดงในหน้า login (เลือกชื่อตัวเอง) — เรียกได้ก่อน login เสมอ ตัด Email/Shop
// ออกเพราะไม่จำเป็นสำหรับหน้านี้ และเป็น PII ที่ไม่ควรให้ดึงได้ก่อนยืนยันตัวตน
function loginOptions_() {
  return readAllRows_(USERS_SHEET_NAME, USERS_HEADERS)
    .filter(u => u.Active)
    .map(u => ({ Id: u.Id, Name: u.Name, Role: u.Role, Active: true }));
}

// ---------------------------------------------------------------
// "กำลังใช้งานอยู่" — front-end เรียก heartbeat ทุก ~45 วิระหว่างเปิดแอปค้างไว้
// (ดู startPresenceHeartbeat() ใน index.html) เขียนเวลาล่าสุดลงคอลัมน์ LastActiveAt
// ของผู้ใช้คนนั้น แล้ว listOnlineUsers กรองคนที่ heartbeat ไม่เกิน ONLINE_THRESHOLD_MS
// ที่แล้วมาโชว์ — ไม่มี mechanism แจ้งตอน "ออฟไลน์ทันที" (ปิดแท็บ/ปิดเบราว์เซอร์) เพราะ
// Apps Script ไม่มี WebSocket ให้ใช้ ต้องพึ่งเวลาหมดอายุแทน จึงอาจเห็นคนที่เพิ่งปิดแท็บ
// ไปค้างอยู่ในรายการได้สูงสุด ~2 นาที
// ---------------------------------------------------------------
function heartbeat_(user) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rowNum = findRowNumberById_(USERS_SHEET_NAME, USERS_HEADERS, user.Id);
    if (rowNum === -1) return;
    const col = USERS_HEADERS.indexOf("LastActiveAt") + 1;
    getSheet_(USERS_SHEET_NAME).getRange(rowNum, col).setValue(new Date().toISOString());
  } finally {
    lock.releaseLock();
  }
}

function listOnlineUsers_() {
  const cutoff = Date.now() - ONLINE_THRESHOLD_MS;
  return readAllRows_(USERS_SHEET_NAME, USERS_HEADERS)
    .filter(u => u.Active && u.LastActiveAt && new Date(u.LastActiveAt).getTime() >= cutoff)
    .map(u => ({ Id: u.Id, Name: u.Name, Role: u.Role }))
    .sort((a, b) => a.Name.localeCompare(b.Name));
}

// ---------------------------------------------------------------
// Session token — เซ็นด้วย secret ที่สุ่มครั้งแรกที่ใช้งานและเก็บใน Script Properties
// (ไม่เคยอยู่ในโค้ด/ไม่เคยส่งให้ client) ต่างจาก API_TOKEN ที่ฝังอยู่ใน index.html แบบเปิดเผย —
// ตัวนี้ออกให้เฉพาะตอน login สำเร็จเท่านั้น และ doPost() บังคับเช็คทุก action ที่ไม่ใช่
// login/loginOptions/ping ก่อนทำงาน เป็นการป้องกันระดับ API จริง ไม่ใช่แค่กันบอท
// ---------------------------------------------------------------
const SESSION_TOKEN_TTL_MS = 24 * 60 * 60 * 1000; // หมดอายุใน 24 ชม. ต้อง login ใหม่

function getSessionSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("SESSION_SECRET");
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty("SESSION_SECRET", secret);
  }
  return secret;
}

function signSessionPayload_(payload) {
  const bytes = Utilities.computeHmacSha256Signature(payload, getSessionSecret_());
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, "0")).join("");
}

function issueSessionToken_(userId) {
  const payload = userId + "|" + (Date.now() + SESSION_TOKEN_TTL_MS);
  return payload + "|" + signSessionPayload_(payload);
}

// คืน user (sanitize แล้ว) ถ้า token ถูกต้อง/ยังไม่หมดอายุ/บัญชียังใช้งานอยู่ — ไม่งั้นคืน null
function verifySessionToken_(token) {
  if (!token || typeof token !== "string") return null;
  const parts = token.split("|");
  if (parts.length !== 3) return null;
  const payload = parts[0] + "|" + parts[1];
  if (signSessionPayload_(payload) !== parts[2]) return null;
  const exp = Number(parts[1]);
  if (!exp || Date.now() > exp) return null;
  const userId = Number(parts[0]);
  const row = readAllRows_(USERS_SHEET_NAME, USERS_HEADERS).find(r => r.Id === userId);
  if (!row || !row.Active) return null;
  return sanitizeUser_(row);
}

const ROLE_CODE_MAP_ = { "SHE-Auditor": "auditor", "SHE-Dept-Responsible": "dept", "SHE-Safety-Admin": "admin", "SHE-Executive-Viewer": "exec", "SHE-Safety-Officer": "safety" };
function userRoleCodes_(user) {
  return (user.Role || "").split(",").map(s => ROLE_CODE_MAP_[s.trim()]).filter(Boolean);
}
function requireRole_(user, allowedCodes) {
  const codes = userRoleCodes_(user);
  if (!allowedCodes.some(c => codes.includes(c))) throw new Error("ไม่มีสิทธิ์ทำรายการนี้");
}
// mirror ของ canEditProgress()/canEditDetails()/canClose() ฝั่ง client — ข้อมูลเดิมของรายการ
// (TypeOfAudit/PatrolDate/Shop/Place/Description/Grade/Category) แก้ได้ทั้ง Admin/Auditor/
// Safety Officer (SHE-Safety-Officer — ไม่ใช่ Safety Admin เต็มรูปแบบ), ฟิลด์ความคืบหน้าที่เหลือ
// แก้ได้ทั้งสามบทบาทนั้นบวกหน่วยงานที่ตรง Shop ของตัวเอง, และปิดงาน (Status: "ปิดงาน")
// ทำได้เฉพาะ Admin กับ Safety Officer เท่านั้น (ไม่ใช่ Auditor/หน่วยงาน)
function requireCanEditFinding_(user, findingId, fields) {
  const codes = userRoleCodes_(user);
  const detailFields = ["TypeOfAudit", "PatrolDate", "Shop", "Place", "Description", "Grade", "Category"];
  if (Object.keys(fields).some(k => detailFields.includes(k)) && !codes.includes("admin") && !codes.includes("auditor") && !codes.includes("safety")) {
    throw new Error("แก้ไขข้อมูลเดิมของรายการได้เฉพาะ Admin, ผู้ตรวจ (Auditor) หรือ จนท.ความปลอดภัย");
  }
  if (fields.Status === "ปิดงาน" && !codes.includes("admin") && !codes.includes("safety")) {
    throw new Error("ปิดงานได้เฉพาะ Admin หรือ จนท.ความปลอดภัย");
  }
  if (codes.includes("admin") || codes.includes("auditor") || codes.includes("safety")) return;
  if (codes.includes("dept")) {
    const finding = getFinding_(findingId);
    if (user.Shop && user.Shop === finding.Shop) return;
  }
  throw new Error("ไม่มีสิทธิ์แก้ไขรายการนี้");
}

function login_(id, password) {
  const row = readAllRows_(USERS_SHEET_NAME, USERS_HEADERS).find(r => r.Id === Number(id));
  if (!row) throw new Error("ไม่พบผู้ใช้งาน");
  if (!row.Active) throw new Error("บัญชีนี้ถูกปิดใช้งาน ติดต่อ Safety Admin");
  if (!row.Password) throw new Error("บัญชีนี้ยังไม่ได้ตั้งรหัสผ่าน ติดต่อ Safety Admin ให้ตั้งรหัสผ่านก่อน");
  if (hashPassword_(password) !== row.Password) throw new Error("รหัสผ่านไม่ถูกต้อง");
  const user = sanitizeUser_(row);
  user.SessionToken = issueSessionToken_(user.Id);
  logActivity_(user, "เข้าสู่ระบบ", "User", user.Id, "");
  return user;
}

function createUser_(fields, actingUser) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let record;
  try {
    const id = nextId_(USERS_SHEET_NAME, USERS_HEADERS);
    record = Object.assign({ Id: id, Active: true, Shop: "", Password: "" }, fields, { Id: id });
    if (record.Password) record.Password = hashPassword_(record.Password);
    const row = USERS_HEADERS.map(h => (record[h] !== undefined && record[h] !== null) ? record[h] : "");
    getSheet_(USERS_SHEET_NAME).appendRow(row);
  } finally {
    lock.releaseLock();
  }
  logActivity_(actingUser, "เพิ่มผู้ใช้งาน", "User", record.Id, `${record.Name || "-"} · ${record.Role || "-"}`);
  return sanitizeUser_(record);
}

function updateUser_(id, fields, actingUser) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let before, after, patch;
  try {
    const rowNum = findRowNumberById_(USERS_SHEET_NAME, USERS_HEADERS, id);
    if (rowNum === -1) throw new Error("ไม่พบผู้ใช้งานรหัส " + id);
    const sheet = getSheet_(USERS_SHEET_NAME);
    before = rowToObject_(USERS_HEADERS, sheet.getRange(rowNum, 1, 1, USERS_HEADERS.length).getValues()[0]);
    patch = Object.assign({}, fields);
    if (patch.Password) patch.Password = hashPassword_(patch.Password);
    else delete patch.Password; // เว้นว่าง = ไม่เปลี่ยนรหัสผ่านเดิม
    after = Object.assign({}, before, patch, { Id: before.Id });
    const row = USERS_HEADERS.map(h => (after[h] !== undefined && after[h] !== null) ? after[h] : "");
    sheet.getRange(rowNum, 1, 1, USERS_HEADERS.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  const changedKeys = USERS_HEADERS.filter(h => Object.prototype.hasOwnProperty.call(patch, h));
  const details = diffText_(before, after, changedKeys, ["Password"]);
  logActivity_(actingUser, "แก้ไขผู้ใช้งาน", "User", id, details);
  return sanitizeUser_(after);
}

// ---------------------------------------------------------------
// รูปภาพ — เก็บใน Google Drive ตั้งค่าดูได้ตรงจากลิงก์เลย (ไม่ต้องมี auth)
// ---------------------------------------------------------------
function uploadPhoto_(findingId, stage, fileName, contentBase64) {
  return { fileUrl: uploadPhotoFile_(stage, fileName, contentBase64) };
}

// อัปโหลดหลายไฟล์ (findings[].{fileName, contentBase64}) แล้วคืน URL คั่นด้วยจุลภาคเดียว —
// ใช้เก็บในคอลัมน์ PhotoBeforeUrl/PhotoAfterUrl เดียวกัน (ไม่เพิ่มคอลัมน์ใหม่ เหมือน
// Shop_Options/Category_Options ที่เก็บหลายค่าคั่นจุลภาคในคอลัมน์เดียวกันอยู่แล้ว)
function uploadPhotoFiles_(stage, files) {
  return (files || [])
    .filter(f => f && f.contentBase64)
    .map(f => uploadPhotoFile_(stage, f.fileName, f.contentBase64))
    .join(",");
}

// ใช้ร่วมกันทั้ง action "uploadPhoto" เดี่ยวๆ (แนบรูปเพิ่มทีหลัง) และตอน
// createFinding_/updateFinding_ แนบรูปมาพร้อมคำขอเดียวกันเลย
function uploadPhotoFile_(stage, fileName, contentBase64) {
  const bytes = Utilities.base64Decode(contentBase64);
  const safeName = `${stage === "ก่อนแก้ไข" ? "before" : "after"}_${new Date().getTime()}.jpg`;
  const blob = Utilities.newBlob(bytes, "image/jpeg", safeName);
  const folder = getOrCreateFolder_(PHOTOS_FOLDER_NAME);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // ใช้ endpoint "thumbnail" ไม่ใช่ "uc?export=view" — แบบหลังนี้ Google มักเด้งหน้า
  // "ไม่สามารถสแกนไวรัสได้" แทนรูปจริงเวลาฝัง <img src> ตรงๆ ทำให้รูปพังเป็นไอคอนแตก
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1000`;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ---------------------------------------------------------------
// การแจ้งเตือนทางอีเมล (แทน Power Automate Flow 1/2/4 เดิม)
// ---------------------------------------------------------------
function findingLink_(id) {
  const base = getSetting("Frontend_Base_URL");
  return base ? `${base.replace(/\/?$/, "/")}index.html?id=${id}` : "";
}

function notifyNewFinding_(record) {
  const to = shopLeadEmail_(record.Shop);
  if (!to) return; // ยังไม่ได้กรอกอีเมลใน Settings — ข้ามการแจ้งเตือน
  const link = findingLink_(record.Id);
  const subject = `[SHE Patrol] พบ Finding ใหม่ Grade ${record.Grade} - ${record.Shop}`;
  const body = [
    "พบ Finding ใหม่จากการตรวจ SHE Patrol", "",
    `Shop: ${record.Shop}`,
    `จุดที่พบ: ${record.Place || "-"}`,
    `รายละเอียด: ${record.Description || "-"}`,
    `Grade: ${record.Grade} | Category: ${record.Category || "-"}`,
    `กำหนดเสร็จ: ${record.DueDate}`, "",
    link ? `ดูรายละเอียดและอัปเดตมาตรการแก้ไข: ${link}` : "",
  ].join("\n");
  MailApp.sendEmail({ to: to, subject: subject, body: body });
}

function notifyFindingStatusChange_(before, after) {
  if (before.Status === after.Status) return; // ไม่เปลี่ยนสถานะ ไม่ต้องแจ้ง
  const link = findingLink_(after.Id);

  if (after.Status === "รอตรวจสอบ") {
    const to = getSetting("Safety_Officer_Email");
    if (!to) return;
    MailApp.sendEmail({
      to: to,
      subject: `[SHE Patrol] รอตรวจสอบ - ${after.Shop} / ${after.Place || "-"}`,
      body: [
        "หน่วยงานได้อัปเดตมาตรการแก้ไขและแนบรูปหลังแล้ว รอตรวจสอบและปิดงาน", "",
        `Shop: ${after.Shop} | Grade: ${after.Grade}`,
        `มาตรการแก้ไข: ${after.Countermeasure || "-"}`,
        `ผู้รับผิดชอบ: ${after.ActionResponsible || "-"}`, "",
        link ? `เปิดหน้าตรวจสอบและปิดงาน: ${link}` : "",
      ].join("\n"),
    });
  }

  if (after.Status === "ปิดงาน") {
    const to = shopLeadEmail_(after.Shop);
    if (!to) return;
    MailApp.sendEmail({
      to: to,
      subject: `[SHE Patrol] ปิดงานแล้ว - ${after.Shop} / ${after.Place || "-"}`,
      body: [
        "Finding นี้ถูกปิดงานเรียบร้อยแล้ว", "",
        `Grade: ${after.Grade} | Shop: ${after.Shop}`,
        `ผู้ตรวจยืนยัน: ${after.VerifiedBy || "-"}`,
        `เวลายืนยัน: ${after.Rules_Confirmed_DateTime || "-"}`, "",
        link ? `ดูรายละเอียด: ${link}` : "",
      ].join("\n"),
    });
  }
}

// ---------------------------------------------------------------
// เช็ค SLA เลยกำหนด/ใกล้กำหนดทุกวัน — ตั้ง time-driven trigger เรียกฟังก์ชันนี้
// (แทน Power Automate Flow 3 เดิม)
// ---------------------------------------------------------------
function checkSlaEscalation() {
  const daysBefore = Number(getSetting("SLA_Escalation_Days_Before")) || 2;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rows = listFindings_().filter(r => STATUS_OPEN.indexOf(r.Status) !== -1);
  rows.forEach(r => {
    const due = new Date(r.DueDate);
    due.setHours(0, 0, 0, 0);
    const daysLeft = Math.round((due - today) / 86400000);
    if (daysLeft > daysBefore) return; // ยังไม่ใกล้กำหนด ข้าม

    const to = [];
    if (r.Grade === "A") {
      if (getSetting("MGR_Email")) to.push(getSetting("MGR_Email"));
      if (getSetting("AGMGM_Email")) to.push(getSetting("AGMGM_Email"));
    } else if (r.Grade === "B") {
      if (getSetting("MGR_Email")) to.push(getSetting("MGR_Email"));
    } else {
      const lead = shopLeadEmail_(r.Shop);
      if (lead) to.push(lead);
    }
    if (!to.length) return;

    const link = findingLink_(r.Id);
    const overdue = daysLeft < 0;
    MailApp.sendEmail({
      to: to.join(","),
      subject: `[SHE Patrol] ${overdue ? "เลยกำหนด" : "ใกล้กำหนด"} Grade ${r.Grade} - ${r.Shop}`,
      body: [
        `Finding Grade ${r.Grade} ที่ ${r.Shop} / ${r.Place || "-"}`,
        overdue ? `เลยกำหนดมาแล้ว ${Math.abs(daysLeft)} วัน` : `เหลืออีก ${daysLeft} วัน`,
        `สถานะปัจจุบัน: ${r.Status}`, "",
        link ? `เปิดหน้า Finding: ${link}` : "",
      ].join("\n"),
    });
  });
}

// ---------------------------------------------------------------
// สำรองข้อมูลเข้า Microsoft 365 — เพราะสเปรดชีตนี้ผูกกับบัญชี Google ส่วนตัว
// ไม่ใช่ M365 tenant ของบริษัท จึงส่งสำเนา .xlsx ทั้งไฟล์เข้าอีเมล M365 เป็นระยะ
// (ตั้ง time-driven trigger ให้ backupToMS365Email รันรายสัปดาห์ — ดู README)
//
// ฝั่ง Microsoft 365 ใช้ Power Automate flow มาตรฐาน (ไม่ใช่ premium — ต่างจาก Flow 5 เดิม)
// คอยดักอีเมลนี้แล้วเซฟไฟล์แนบเข้า OneDrive/SharePoint ให้อัตโนมัติ ดู
// power-automate/Flow6-Backup-To-OneDrive.md
// ---------------------------------------------------------------
function backupToMS365Email() {
  const to = getSetting("MS365_Backup_Email");
  if (!to) return; // ยังไม่ได้ตั้งค่า — ข้าม

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: "Bearer " + token } });
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const blob = response.getBlob().setName(`SHE_Patrol_Backup_${dateStr}.xlsx`);

  MailApp.sendEmail({
    to: to,
    subject: `[SHE Patrol] สำรองข้อมูลรายสัปดาห์ ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy")}`,
    body: [
      "ไฟล์แนบคือสำเนาข้อมูล SHE Patrol ทั้งหมด (Findings + Users) ณ เวลาที่ส่งอีเมลนี้",
      "ส่งมาเพื่อสำรองข้อมูลไว้ใน Microsoft 365 เนื่องจากข้อมูลหลักอยู่ใน Google Sheets ที่ผูกกับ",
      "บัญชี Google ส่วนบุคคล ไม่ใช่ M365 tenant ของบริษัท",
      "",
      "ถ้าตั้งค่า Power Automate flow ตาม power-automate/Flow6-Backup-To-OneDrive.md ไว้แล้ว",
      "ไฟล์นี้จะถูกบันทึกเข้า OneDrive/SharePoint ให้อัตโนมัติ ไม่ต้องทำอะไรเพิ่ม",
    ].join("\n"),
    attachments: [blob],
  });
}
