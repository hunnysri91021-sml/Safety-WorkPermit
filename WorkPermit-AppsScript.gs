/**
 * WorkPermit-AppsScript.gs
 * ระบบ Work Permit ครบวงจร: รับคำขอ, ออกเลขที่ Permit, แจ้งเตือน Safety Team,
 * อนุมัติ/ไม่อนุมัติ, แบบทดสอบก่อนออกบัตร (เฉพาะงานเสี่ยงสูง), ออกบัตรผู้รับเหมา
 * เป็น PDF พร้อม QR Code, บันทึกเข้า-ออกพื้นที่โดย รปภ. (ยืนยันตัวด้วย PIN),
 * รับรูปประเมินพื้นที่จากผู้รับเหมา, และสรุปข้อมูลสำหรับ Dashboard
 * บริษัท สยามกลการโลจิสติกส์ จำกัด (SML)
 *
 * ====================================================================
 * วิธีติดตั้ง
 * ====================================================================
 *  1. สร้าง Google Sheet ใหม่ (เช่น ตั้งชื่อ "SML Work Permit")
 *  2. เปิด Extensions > Apps Script แล้ววางไฟล์นี้ทับโค้ดเริ่มต้น (Code.gs)
 *  3. รันฟังก์ชัน setupSheet() หนึ่งครั้งจากตัวแก้ไข Apps Script
 *     (ครั้งแรกจะขอ authorize สิทธิ์เข้าถึง Sheet + Drive + ส่งอีเมล)
 *     จะสร้างแท็บ WorkPermit, Settings, Guards, Quiz, ScanLog, SiteReports
 *     ให้อัตโนมัติ พร้อมข้อมูลตัวอย่าง
 *  4. Deploy > New deployment > เลือกประเภท "Web app"
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     กด Deploy แล้วคัดลอก Web app URL ที่ได้
 *  5. นำ URL ไปใส่ค่าคงที่ GOOGLE_SCRIPT_URL ในไฟล์ request-permit.html,
 *     view.html, quiz.html, guard-login.html, site-report.html, dashboard.html
 *     (index.html เป็นหน้าแรก/เมนูเลือกใช้งาน ไม่ต้องใส่ GOOGLE_SCRIPT_URL)
 *  6. เปิดแท็บ "Settings" ในชีต แล้วกรอก:
 *       - Safety_Team_Emails: อีเมล Safety Team (คั่นด้วยจุลภาคถ้ามีหลายคน)
 *       - Site_Base_URL: URL ของเว็บที่โฮสต์ไฟล์ .html ทั้งหมด เช่น
 *         https://hunnysri91021-sml.github.io/Safety-WorkPermit/
 *       - Quiz_Required_Keywords: คำที่ใช้จับคู่กับ "ลักษณะงาน/ประเภทงาน" ของ
 *         คำขอ ถ้าตรงคำใดคำหนึ่งจะบังคับให้ทำแบบทดสอบก่อนออกบัตร (ค่าเริ่มต้น
 *         อ้างอิงจากตัวเลือกลักษณะงานเสี่ยงสูงในฟอร์ม)
 *       - Quiz_Pass_Percent: เกณฑ์ผ่านแบบทดสอบ (ค่าเริ่มต้น 80)
 *       - Dashboard_Password: รหัสผ่านเข้าดู dashboard.html (ตั้งเอง ไม่มีค่า
 *         เริ่มต้น — ถ้าไม่ตั้ง dashboard จะเข้าไม่ได้)
 *     เปลี่ยนได้ตลอดเวลาโดยไม่ต้องแก้โค้ด
 *  7. เปิดแท็บ "Guards" เพิ่มรายชื่อ รปภ. จริง (รวมถึง รปภ. บริษัท outsource)
 *     พร้อมตั้ง PIN ให้แต่ละคน (ลบ/แก้แถวตัวอย่างที่ Active=No ออกได้)
 *  8. เปิดแท็บ "Quiz" แก้ไข/เพิ่มคำถามแบบทดสอบตามต้องการ (มีตัวอย่างให้ 5 ข้อ)
 *  9. ตั้งค่าแชร์ของ Sheet เป็น "Restricted" แล้วเพิ่มเฉพาะอีเมลของ
 *     Safety Team ที่ต้องดูข้อมูล/อนุมัติคำขอ (ให้สิทธิ์ Editor) ห้ามเปิดเป็น
 *     "Anyone with the link" เด็ดขาด เพื่อรักษาข้อมูลส่วนบุคคลตาม PDPA
 *
 * ====================================================================
 * วิธีใช้งาน
 * ====================================================================
 *  Safety Team: เปิด Sheet แท็บ "WorkPermit" คลิกแถวคำขอที่จะพิจารณา ไปที่เมนู
 *  "คำขอ Work Permit" เลือก "✅ อนุมัติแถวที่เลือก" หรือ "❌ ไม่อนุมัติแถวที่เลือก"
 *    - งานที่ไม่ตรงกับ Quiz_Required_Keywords → ออกบัตร PDF + อีเมลทันที
 *    - งานที่ตรงกับ Quiz_Required_Keywords → อีเมลลิงก์แบบทดสอบไปก่อน บัตรจะ
 *      ออกให้อัตโนมัติก็ต่อเมื่อผู้รับเหมาทำแบบทดสอบผ่าน (quiz.html)
 *
 *  รปภ.: เปิด guard-login.html บนอุปกรณ์ที่ป้อม กรอก PIN ครั้งเดียว (จำไว้ใน
 *  เครื่องนั้น) จากนั้นใช้กล้องมือถือ/แท็บเล็ตสแกน QR บนบัตรผู้รับเหมาได้เลย
 *  (เปิดเป็น view.html ตามปกติ) หน้านั้นจะมีปุ่ม "บันทึกเข้า/ออกพื้นที่" โผล่ขึ้น
 *  มาให้กดเพราะเครื่องนั้น login เป็น รปภ. อยู่แล้ว
 *
 *  ผู้รับเหมา: หลังแลกบัตรเข้าพื้นที่แล้ว เปิด site-report.html?id=<เลขที่ Permit>
 *  เพื่อถ่ายรูปส่งประเมินพื้นที่ก่อนเริ่มงานได้
 *
 * ====================================================================
 * หมายเหตุด้านความปลอดภัยของข้อมูล (PDPA)
 * ====================================================================
 *  - Web app "Anyone" เรียก doPost()/doGet() ได้เฉพาะ action ที่กำหนดไว้เท่านั้น
 *  - การดูสถานะผ่าน view.html ส่งกลับเฉพาะข้อมูลจำเป็นต่อการตรวจสอบหน้างาน
 *  - Dashboard และ Guard scan ต้องใช้รหัสผ่าน/PIN ก่อนเข้าถึงข้อมูลเสมอ
 *  - ไฟล์ PDF/รูปภาพที่สร้างจะแชร์แบบ "Anyone with the link" เฉพาะไฟล์นั้น
 *    ไฟล์เดียว (ไม่ใช่ทั้งโฟลเดอร์/ทั้งชีต)
 *  - รันซ้ำ setupSheet() ได้อย่างปลอดภัย จะไม่ลบข้อมูล/ค่าตั้งค่าที่มีอยู่
 *  - บัญชี Gmail ส่วนตัวส่งอีเมลผ่าน MailApp ได้ฟรีประมาณวันละ 100 ฉบับ
 */

const SHEET_NAME = "WorkPermit";
const SETTINGS_SHEET_NAME = "Settings";
const GUARDS_SHEET_NAME = "Guards";
const QUIZ_SHEET_NAME = "Quiz";
const SCANLOG_SHEET_NAME = "ScanLog";
const SITEREPORTS_SHEET_NAME = "SiteReports";
const PDF_FOLDER_NAME = "Work Permit PDFs";
const SITEREPORT_FOLDER_NAME = "Site Report Photos";

const HEADERS = [
  "Permit_ID", "Submit_DateTime", "Status", "Requester_Name", "Company", "Contact_Email",
  "Contact_Phone", "Work_Location", "Work_Date", "Time_From", "Time_To",
  "Work_Nature", "Work_Type", "Worker_Count", "Worker_Names", "PPE_Used",
  "Safety_Equipment", "Tools_List", "Rules_Confirmed_DateTime", "Consent_PDPA",
  "Approver_Name", "Approved_DateTime", "Reject_Reason", "Card_Exchanged", "Remarks", "PDF_Link",
  "Quiz_Required", "Quiz_Passed", "Quiz_Score", "Quiz_Attempts"
];

const GUARD_HEADERS = ["Guard_PIN", "Guard_Name", "Guard_Company", "Active"];
const DEFAULT_GUARDS = [
  GUARD_HEADERS,
  ["0000", "(ตัวอย่าง) แก้ไข/ลบแถวนี้แล้วเพิ่มรายชื่อจริง", "บริษัท รปภ. ตัวอย่าง", "No"]
];

const QUIZ_HEADERS = ["Question_ID", "Question", "Choice_A", "Choice_B", "Choice_C", "Choice_D", "Correct"];
const DEFAULT_QUIZ = [
  QUIZ_HEADERS,
  ["Q1", "ก่อนเข้าปฏิบัติงานต้องทำอะไรก่อน?", "แลกบัตรที่ป้อม รปภ.", "เดินเข้าไปได้เลย", "โทรแจ้งเพื่อน", "ไม่ต้องทำอะไร", "A"],
  ["Q2", "หากทำงานบนที่สูงเกิน 2 เมตร ต้องสวมใส่อุปกรณ์ใด?", "ถุงมือผ้า", "เข็มขัดนิรภัย", "รองเท้าแตะ", "หมวกไหมพรม", "B"],
  ["Q3", "หากเกิดอัคคีภัยขณะทำงาน ต้องทำอย่างไร?", "วิ่งหนีคนเดียวโดยไม่แจ้งใคร", "แจ้งหัวหน้างาน/รปภ. ทันที", "เก็บของก่อนแจ้ง", "ไม่ต้องแจ้งใคร", "B"],
  ["Q4", "ห้ามนำสิ่งใดเข้าพื้นที่บริษัทฯ โดยเด็ดขาด?", "กล้องถ่ายรูป", "อาวุธและสิ่งเสพติด", "โทรศัพท์มือถือ", "กระเป๋าเอกสาร", "B"],
  ["Q5", "ความเร็วจำกัดภายในพื้นที่บริษัทฯ คือเท่าใด?", "30 กม./ชม.", "60 กม./ชม.", "80 กม./ชม.", "ไม่จำกัด", "A"]
];

const SCANLOG_HEADERS = ["Timestamp", "Permit_ID", "Direction", "Guard_PIN", "Guard_Name", "Guard_Company"];
const SITEREPORT_HEADERS = ["Timestamp", "Permit_ID", "Photo_Link", "Note"];

const DEFAULT_SETTINGS = [
  ["Key", "Value", "คำอธิบาย"],
  ["Safety_Team_Emails", "", "อีเมล Safety Team ที่รับแจ้งเตือนคำขอใหม่ (คั่นด้วยจุลภาคถ้ามีหลายคน)"],
  ["Company_Name", "บริษัท สยามกลการโลจิสติกส์ จำกัด", "ชื่อบริษัท ใช้ในอีเมล/เอกสาร PDF"],
  ["Site_Base_URL", "", "URL ของเว็บที่โฮสต์ไฟล์ .html ทั้งหมด เช่น https://ชื่อบัญชี.github.io/Safety-WorkPermit/"],
  ["Quiz_Required_Keywords", "อับอากาศ,ที่สูงเกิน 2 เมตร,สารเคมี,ประกายไฟ", "คำที่ใช้จับคู่กับลักษณะงาน/ประเภทงาน — ถ้าตรงจะบังคับทำแบบทดสอบก่อนออกบัตร (คั่นด้วยจุลภาค)"],
  ["Quiz_Pass_Percent", "80", "เกณฑ์ % ที่ต้องทำถูกจึงจะผ่านแบบทดสอบ"],
  ["Dashboard_Password", "", "รหัสผ่านเข้าดู dashboard.html — ต้องตั้งเองก่อนใช้งาน"]
];

function colIndex(name) {
  return HEADERS.indexOf(name) + 1;
}

// ---------------------------------------------------------------
// ติดตั้ง/อัปเดตโครงสร้างชีตทั้งหมด (รันซ้ำได้อย่างปลอดภัย)
// ---------------------------------------------------------------
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  setupKeyValueSheet_(ss, SETTINGS_SHEET_NAME, DEFAULT_SETTINGS, 3);
  setupTableSheet_(ss, GUARDS_SHEET_NAME, DEFAULT_GUARDS);
  setupTableSheet_(ss, QUIZ_SHEET_NAME, DEFAULT_QUIZ);
  setupHeaderOnlySheet_(ss, SCANLOG_SHEET_NAME, SCANLOG_HEADERS);
  setupHeaderOnlySheet_(ss, SITEREPORTS_SHEET_NAME, SITEREPORT_HEADERS);
}

function setupKeyValueSheet_(ss, name, defaultRows, numCols) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const existingRows = sheet.getLastRow();
  if (existingRows < defaultRows.length) {
    sheet.getRange(existingRows + 1, 1, defaultRows.length - existingRows, numCols)
      .setValues(defaultRows.slice(existingRows));
  }
  sheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, numCols);
}

function setupTableSheet_(ss, name, defaultRows) {
  let sheet = ss.getSheetByName(name);
  const isNew = !sheet;
  if (!sheet) sheet = ss.insertSheet(name);
  if (isNew) {
    sheet.getRange(1, 1, defaultRows.length, defaultRows[0].length).setValues(defaultRows);
  } else {
    sheet.getRange(1, 1, 1, defaultRows[0].length).setValues([defaultRows[0]]);
  }
  sheet.getRange(1, 1, 1, defaultRows[0].length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function setupHeaderOnlySheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("คำขอ Work Permit")
    .addItem("✅ อนุมัติแถวที่เลือก", "approveSelectedRow")
    .addItem("❌ ไม่อนุมัติแถวที่เลือก", "rejectSelectedRow")
    .addToUi();
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

function getViewUrl_(permitId) {
  const siteBaseUrl = getSetting("Site_Base_URL");
  if (!siteBaseUrl) return "";
  return siteBaseUrl.replace(/\/?$/, "/") + "view.html?id=" + encodeURIComponent(permitId);
}

// ---------------------------------------------------------------
// เราท์เตอร์หลัก — GET ใช้สำหรับอ่านข้อมูล (สถานะ/แบบทดสอบ/dashboard)
// POST ใช้สำหรับเขียนข้อมูล (ส่งคำขอ/ทำแบบทดสอบ/scan/ส่งรูป)
// ---------------------------------------------------------------
function doGet(e) {
  const action = e.parameter.action || "status";
  try {
    switch (action) {
      case "status":
        return handleGetStatus_(e);
      case "quiz":
        return handleGetQuiz_(e);
      case "dashboard":
        return handleGetDashboard_(e);
      default:
        return jsonResponse({ result: "error", message: "ไม่รู้จัก action นี้" });
    }
  } catch (err) {
    return jsonResponse({ result: "error", message: err.message });
  }
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ result: "error", message: "รูปแบบข้อมูลไม่ถูกต้อง" });
  }

  const action = data.action || "submitPermit";
  try {
    switch (action) {
      case "submitPermit":
        return handleSubmitPermit_(data);
      case "submitQuiz":
        return handleSubmitQuiz_(data);
      case "guardLogin":
        return handleGuardLogin_(data);
      case "guardScan":
        return handleGuardScan_(data);
      case "submitSiteReport":
        return handleSubmitSiteReport_(data);
      default:
        return jsonResponse({ result: "error", message: "ไม่รู้จัก action นี้" });
    }
  } catch (err) {
    return jsonResponse({ result: "error", message: err.message });
  }
}

// ---------------------------------------------------------------
// 1) รับคำขอใหม่จากฟอร์ม (request-permit.html)
// ---------------------------------------------------------------
function handleSubmitPermit_(data) {
  if (!data.Requester_Name || !data.Company || !data.Contact_Email || !data.Work_Location || !data.Work_Date) {
    return jsonResponse({ result: "error", message: "ข้อมูลที่จำเป็นไม่ครบถ้วน" });
  }
  if (!data.Consent_PDPA) {
    return jsonResponse({ result: "error", message: "ต้องยินยอมให้เก็บข้อมูลส่วนบุคคลก่อนส่งคำขอ" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // ล็อกกันเลขที่ Permit ซ้ำกันถ้ามีคนส่งฟอร์มพร้อมกันหลายคน
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let permitId;
  try {
    permitId = generatePermitId_(sheet);

    const record = {
      Permit_ID: permitId,
      Submit_DateTime: new Date(), // ใช้เวลาเซิร์ฟเวอร์ ไม่พึ่งนาฬิกาเครื่องผู้ส่ง
      Status: "รออนุมัติ",
      Requester_Name: data.Requester_Name,
      Company: data.Company,
      Contact_Email: data.Contact_Email,
      Contact_Phone: data.Contact_Phone || "",
      Work_Location: data.Work_Location,
      Work_Date: data.Work_Date,
      Time_From: data.Time_From || "",
      Time_To: data.Time_To || "",
      Work_Nature: Array.isArray(data.Work_Nature) ? data.Work_Nature.join(", ") : "",
      Work_Type: Array.isArray(data.Work_Type) ? data.Work_Type.join(", ") : "",
      Worker_Count: data.Worker_Count || "",
      Worker_Names: data.Worker_Names || "",
      PPE_Used: Array.isArray(data.PPE_Used) ? data.PPE_Used.join(", ") : "",
      Safety_Equipment: Array.isArray(data.Safety_Equipment) ? data.Safety_Equipment.join(", ") : "",
      Tools_List: data.Tools_List || "",
      Rules_Confirmed_DateTime: data.Rules_Confirmed_DateTime || "",
      Consent_PDPA: "ยินยอม"
    };

    sheet.appendRow(HEADERS.map(h => (record[h] !== undefined ? record[h] : "")));
  } finally {
    lock.releaseLock();
  }

  notifySafetyTeamOfNewRequest_(data, permitId);

  return jsonResponse({ result: "ok", permitId: permitId });
}

function generatePermitId_(sheet) {
  const prefix = "WP-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMM") + "-";
  const lastRow = sheet.getLastRow();
  let ids = [];
  if (lastRow > 1) {
    ids = sheet.getRange(2, colIndex("Permit_ID"), lastRow - 1, 1).getValues().flat();
  }
  const countThisMonth = ids.filter(id => String(id).indexOf(prefix) === 0).length;
  const seq = String(countThisMonth + 1).padStart(3, "0");
  return prefix + seq;
}

// ---------------------------------------------------------------
// 2) ค้นหาสถานะ Permit ตามเลขที่ (ใช้โดย view.html ตอนสแกน QR)
// ---------------------------------------------------------------
function handleGetStatus_(e) {
  const id = e.parameter.id;
  if (!id) return jsonResponse({ result: "error", message: "ไม่พบเลขที่ Permit ใน URL" });

  const record = findPermitById_(id);
  if (!record) return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });

  // ส่งกลับเฉพาะข้อมูลที่จำเป็นต่อการตรวจสอบหน้างาน — ไม่ส่งอีเมล/เบอร์โทร/
  // รายชื่อผู้ปฏิบัติงานแบบเต็ม/เหตุผลไม่อนุมัติ เพื่อรักษาข้อมูลส่วนบุคคลตาม PDPA
  return jsonResponse({
    result: "ok",
    Permit_ID: record.Permit_ID,
    Status: record.Status,
    Company: record.Company,
    Requester_Name: record.Requester_Name,
    Work_Location: record.Work_Location,
    Work_Date: formatDateOnly_(record.Work_Date),
    Time_From: record.Time_From,
    Time_To: record.Time_To,
    Worker_Count: record.Worker_Count,
    Approved_DateTime: formatDateTime_(record.Approved_DateTime),
    Quiz_Required: record.Quiz_Required === "Yes",
    Quiz_Passed: record.Quiz_Passed === "Yes"
  });
}

function findPermitById_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const idCol = colIndex("Permit_ID") - 1;
  const rowValues = values.find(r => r[idCol] === id);
  return rowValues ? rowToRecord_(rowValues) : null;
}

function findPermitRowNumberById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, colIndex("Permit_ID"), lastRow - 1, 1).getValues().flat();
  const idx = ids.indexOf(id);
  return idx === -1 ? -1 : idx + 2;
}

function rowToRecord_(rowValues) {
  const record = {};
  HEADERS.forEach((h, i) => { record[h] = rowValues[i]; });
  return record;
}

function formatDateOnly_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(value || "");
}

function formatDateTime_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  }
  return "";
}

// ---------------------------------------------------------------
// แจ้งเตือนอีเมล
// ---------------------------------------------------------------
function notifySafetyTeamOfNewRequest_(data, permitId) {
  const emails = getSetting("Safety_Team_Emails");
  if (!emails) return; // ยังไม่ได้ตั้งค่าอีเมลในแท็บ Settings — ข้ามการแจ้งเตือน

  const companyName = getSetting("Company_Name") || "SML";
  const subject = `[Work Permit ใหม่ ${permitId}] ${data.Company} — ${data.Work_Location}`;
  const body = [
    `มีคำขอ Work Permit ใหม่รอพิจารณาจาก ${companyName}`,
    "",
    `เลขที่ Permit: ${permitId}`,
    `ผู้แจ้ง: ${data.Requester_Name}`,
    `บริษัท: ${data.Company}`,
    `สถานที่ปฏิบัติงาน: ${data.Work_Location}`,
    `วันที่: ${data.Work_Date}  เวลา: ${data.Time_From || "-"}–${data.Time_To || "-"}`,
    `จำนวนผู้ปฏิบัติงาน: ${data.Worker_Count || "-"}`,
    "",
    "กรุณาเปิด Google Sheet เพื่อพิจารณาอนุมัติ/ไม่อนุมัติผ่านเมนู \"คำขอ Work Permit\""
  ].join("\n");

  MailApp.sendEmail({ to: emails, subject: subject, body: body });
}

// ---------------------------------------------------------------
// เมนูอนุมัติ/ไม่อนุมัติในชีต
// ---------------------------------------------------------------
function approveSelectedRow() {
  processDecision_(true);
}

function rejectSelectedRow() {
  processDecision_(false);
}

function isQuizRequired_(record) {
  const keywords = String(getSetting("Quiz_Required_Keywords") || "")
    .split(",").map(k => k.trim()).filter(Boolean);
  if (!keywords.length) return false;
  const haystack = `${record.Work_Nature || ""} ${record.Work_Type || ""}`;
  return keywords.some(k => haystack.indexOf(k) !== -1);
}

function processDecision_(isApproved) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) {
    ui.alert(`กรุณาไปที่แท็บ "${SHEET_NAME}" แล้วเลือกแถวคำขอก่อน`);
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row === 1) {
    ui.alert("แถวที่เลือกเป็นหัวตาราง กรุณาเลือกแถวคำขอจริง");
    return;
  }

  const rowValues = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const record = rowToRecord_(rowValues);
  if (!record.Requester_Name || !record.Contact_Email) {
    ui.alert("แถวนี้ไม่มีข้อมูลคำขอ (อาจเป็นแถวว่าง)");
    return;
  }

  let rejectReason = "";
  if (!isApproved) {
    const resp = ui.prompt(
      "เหตุผลที่ไม่อนุมัติ",
      "กรุณาระบุเหตุผล (จะถูกส่งไปในอีเมลแจ้งผู้รับเหมา):",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    rejectReason = resp.getResponseText().trim();
  }

  const approverEmail = Session.getActiveUser().getEmail() || "Safety Team";
  const now = new Date();

  sheet.getRange(row, colIndex("Status")).setValue(isApproved ? "อนุมัติ" : "ไม่อนุมัติ");
  sheet.getRange(row, colIndex("Approver_Name")).setValue(approverEmail);
  sheet.getRange(row, colIndex("Approved_DateTime")).setValue(now);
  if (!isApproved) {
    sheet.getRange(row, colIndex("Reject_Reason")).setValue(rejectReason);
  }

  record.Status = isApproved ? "อนุมัติ" : "ไม่อนุมัติ";
  record.Approved_DateTime = now;

  if (!isApproved) {
    emailRejected_(record, rejectReason);
    ui.alert(`บันทึกผล "ไม่อนุมัติ" (${record.Permit_ID}) และส่งอีเมลแจ้งผู้รับเหมาแล้ว`);
    return;
  }

  const quizRequired = isQuizRequired_(record);
  sheet.getRange(row, colIndex("Quiz_Required")).setValue(quizRequired ? "Yes" : "No");
  record.Quiz_Required = quizRequired ? "Yes" : "No";

  if (quizRequired) {
    emailApprovedNeedsQuiz_(record);
    ui.alert(`อนุมัติแล้ว (${record.Permit_ID}) — งานนี้ต้องทำแบบทดสอบก่อนออกบัตร ระบบส่งลิงก์แบบทดสอบให้ผู้รับเหมาทางอีเมลแล้ว`);
    return;
  }

  try {
    issuePermitCard_(sheet, row, record);
    ui.alert(`บันทึกผล "อนุมัติ" (${record.Permit_ID}) และส่งบัตร PDF ให้ผู้รับเหมาทางอีเมลแล้ว`);
  } catch (pdfErr) {
    ui.alert("อนุมัติสำเร็จ แต่ออกบัตร PDF ไม่สำเร็จ: " + pdfErr.message);
  }
}

// ---------------------------------------------------------------
// ออกบัตร PDF พร้อม QR Code แล้วอีเมลให้ผู้รับเหมา (ใช้ทั้งกรณีอนุมัติตรง
// และกรณีผ่านแบบทดสอบ)
// ---------------------------------------------------------------
function issuePermitCard_(sheet, row, record) {
  const viewUrl = getViewUrl_(record.Permit_ID);
  const pdfFile = generatePermitPdf_(record, viewUrl);
  sheet.getRange(row, colIndex("PDF_Link")).setValue(pdfFile.getUrl());
  emailApprovedWithCard_(record, pdfFile.getBlob());
}

function emailApprovedWithCard_(record, pdfBlob) {
  const companyName = getSetting("Company_Name") || "SML";
  const subject = `Work Permit ${record.Permit_ID} ได้รับการอนุมัติแล้ว — ${companyName}`;
  const body = [
    `เรียน คุณ${record.Requester_Name}`,
    "",
    `เลขที่ Permit: ${record.Permit_ID}`,
    "คำขอ Work Permit ของท่านได้รับการอนุมัติแล้ว กรุณาแสดงไฟล์ PDF ที่แนบมา (มี QR Code สำหรับตรวจสอบ) พร้อมบัตรประชาชนที่ป้อม รปภ. ก่อนเข้าปฏิบัติงานตามวันเวลาที่แจ้งไว้",
    "",
    "หลังเข้าพื้นที่แล้ว สามารถส่งรูปประเมินพื้นที่ปฏิบัติงานผ่านเว็บ site-report.html โดยใช้เลขที่ Permit นี้"
  ].join("\n");

  MailApp.sendEmail({ to: record.Contact_Email, subject: subject, body: body, attachments: [pdfBlob] });
}

function emailApprovedNeedsQuiz_(record) {
  const companyName = getSetting("Company_Name") || "SML";
  const siteBaseUrl = getSetting("Site_Base_URL");
  const quizUrl = siteBaseUrl
    ? siteBaseUrl.replace(/\/?$/, "/") + "quiz.html?id=" + encodeURIComponent(record.Permit_ID)
    : "(ยังไม่ได้ตั้งค่า Site_Base_URL ในแท็บ Settings)";

  const subject = `Work Permit ${record.Permit_ID} ได้รับการอนุมัติเบื้องต้น — กรุณาทำแบบทดสอบ — ${companyName}`;
  const body = [
    `เรียน คุณ${record.Requester_Name}`,
    "",
    `เลขที่ Permit: ${record.Permit_ID}`,
    "คำขอ Work Permit ของท่านผ่านการอนุมัติเบื้องต้นแล้ว เนื่องจากลักษณะงานนี้จัดเป็นงานเสี่ยงสูง กรุณาทำแบบทดสอบความปลอดภัยให้ผ่านก่อน จึงจะออกบัตรผู้รับเหมาให้ได้",
    "",
    `ลิงก์แบบทดสอบ: ${quizUrl}`,
    "",
    "เมื่อทำแบบทดสอบผ่านแล้ว ระบบจะออกบัตร PDF พร้อม QR Code ส่งให้ทางอีเมลนี้ทันที"
  ].join("\n");

  MailApp.sendEmail({ to: record.Contact_Email, subject: subject, body: body });
}

function emailRejected_(record, rejectReason) {
  const companyName = getSetting("Company_Name") || "SML";
  const subject = `Work Permit ${record.Permit_ID} ไม่ได้รับการอนุมัติ — ${companyName}`;
  const body = [
    `เรียน คุณ${record.Requester_Name}`,
    "",
    `เลขที่ Permit: ${record.Permit_ID}`,
    "คำขอ Work Permit ของท่านไม่ได้รับการอนุมัติ",
    `เหตุผล: ${rejectReason || "-"}`,
    "กรุณาติดต่อ Safety Team ของบริษัทฯ หากต้องการสอบถามเพิ่มเติม"
  ].join("\n");

  MailApp.sendEmail({ to: record.Contact_Email, subject: subject, body: body });
}

// ---------------------------------------------------------------
// ออกเอกสาร PDF (บัตรผู้รับเหมา) พร้อม QR Code
// ---------------------------------------------------------------
function generatePermitPdf_(record, viewUrl) {
  const companyName = getSetting("Company_Name") || "บริษัท สยามกลการโลจิสติกส์ จำกัด";
  const workDateText = formatDateOnly_(record.Work_Date);
  const approvedText = formatDateTime_(record.Approved_DateTime);

  let qrImgTag = "";
  if (viewUrl) {
    const qrImageUrl = "https://api.qrserver.com/v1/create-qr-code/?size=260x260&data=" + encodeURIComponent(viewUrl);
    const qrBlob = UrlFetchApp.fetch(qrImageUrl).getBlob();
    const qrDataUri = "data:image/png;base64," + Utilities.base64Encode(qrBlob.getBytes());
    qrImgTag = `<div class="qr-box"><img src="${qrDataUri}"><div class="cap">สแกนเพื่อตรวจสอบสถานะ / บันทึกเข้า-ออกพื้นที่</div></div>`;
  }

  const html = `<html><head><meta charset="UTF-8"><style>
    body{font-family:'Sarabun','Tahoma',sans-serif;color:#1C2B34;padding:24px;}
    .header{background:#0F3D5C;color:#fff;padding:18px 22px;border-radius:10px;margin-bottom:18px;}
    .header .co{font-size:12px;color:#BFD8EC;}
    .header h1{margin:4px 0 0;font-size:20px;}
    .badge{display:inline-block;padding:6px 14px;border-radius:999px;background:#E8F5EE;color:#2F8F5B;font-weight:bold;font-size:13px;margin-top:8px;}
    table{width:100%;border-collapse:collapse;margin-top:16px;}
    td{padding:8px 6px;font-size:13px;border-bottom:1px solid #E2E0D6;vertical-align:top;}
    td.k{color:#6B7580;width:170px;}
    td.v{font-weight:600;}
    .qr-box{text-align:center;margin-top:24px;}
    .qr-box img{width:150px;height:150px;}
    .qr-box .cap{font-size:11px;color:#6B7580;margin-top:6px;}
    .foot{margin-top:24px;font-size:11px;color:#6B7580;text-align:center;}
  </style></head>
  <body>
    <div class="header">
      <div class="co">${companyName} · SAFETY FIRST</div>
      <h1>บัตรผู้รับเหมา / ใบอนุญาตปฏิบัติงาน (Work Permit)</h1>
      <div class="badge">เลขที่ ${record.Permit_ID} — ${record.Status}</div>
    </div>
    <table>
      <tr><td class="k">ผู้แจ้ง</td><td class="v">${record.Requester_Name}</td></tr>
      <tr><td class="k">บริษัท/หน่วยงาน</td><td class="v">${record.Company}</td></tr>
      <tr><td class="k">สถานที่ปฏิบัติงาน</td><td class="v">${record.Work_Location}</td></tr>
      <tr><td class="k">วันเวลาปฏิบัติงาน</td><td class="v">${workDateText}  ${record.Time_From || '-'}–${record.Time_To || '-'}</td></tr>
      <tr><td class="k">จำนวนผู้ปฏิบัติงาน</td><td class="v">${record.Worker_Count || '-'}</td></tr>
      <tr><td class="k">ลักษณะงาน</td><td class="v">${record.Work_Nature || '-'}</td></tr>
      <tr><td class="k">ประเภทงาน</td><td class="v">${record.Work_Type || '-'}</td></tr>
      <tr><td class="k">อนุมัติเมื่อ</td><td class="v">${approvedText || '-'}</td></tr>
    </table>
    ${qrImgTag}
    <div class="foot">เอกสารนี้ออกโดยระบบอัตโนมัติ — กรุณาแสดงเอกสารนี้พร้อมบัตรประชาชนที่ป้อม รปภ.</div>
  </body></html>`;

  const pdfBlob = Utilities.newBlob(html, "text/html", record.Permit_ID + ".html")
    .getAs("application/pdf")
    .setName(`WorkPermit_${record.Permit_ID}.pdf`);

  const folder = getOrCreateFolder_(PDF_FOLDER_NAME);
  const file = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ---------------------------------------------------------------
// 3) แบบทดสอบก่อนออกบัตร (quiz.html)
// ---------------------------------------------------------------
function handleGetQuiz_(e) {
  const id = e.parameter.id;
  if (!id) return jsonResponse({ result: "error", message: "ไม่พบเลขที่ Permit ใน URL" });

  const record = findPermitById_(id);
  if (!record) return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });
  if (record.Status !== "อนุมัติ") {
    return jsonResponse({ result: "error", message: "Permit นี้ยังไม่ได้รับการอนุมัติ" });
  }
  if (record.Quiz_Required !== "Yes") {
    return jsonResponse({ result: "error", message: "งานนี้ไม่จำเป็นต้องทำแบบทดสอบ" });
  }
  if (record.Quiz_Passed === "Yes") {
    return jsonResponse({ result: "ok", permitId: id, alreadyPassed: true, questions: [] });
  }

  const quizSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QUIZ_SHEET_NAME);
  const rows = quizSheet.getDataRange().getValues();
  const questions = rows.slice(1).filter(r => r[0]).map(r => ({
    id: r[0],
    question: r[1],
    choices: { A: r[2], B: r[3], C: r[4], D: r[5] }
  }));

  return jsonResponse({ result: "ok", permitId: id, alreadyPassed: false, questions: questions });
}

function handleSubmitQuiz_(data) {
  const id = data.permitId;
  if (!id || !data.answers) return jsonResponse({ result: "error", message: "ข้อมูลไม่ครบถ้วน" });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const row = findPermitRowNumberById_(sheet, id);
  if (row === -1) return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });

  const record = rowToRecord_(sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0]);
  if (record.Status !== "อนุมัติ" || record.Quiz_Required !== "Yes") {
    return jsonResponse({ result: "error", message: "Permit นี้ไม่อยู่ในสถานะที่ทำแบบทดสอบได้" });
  }
  if (record.Quiz_Passed === "Yes") {
    return jsonResponse({ result: "ok", passed: true, alreadyPassed: true });
  }

  const quizSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QUIZ_SHEET_NAME);
  const quizRows = quizSheet.getDataRange().getValues().slice(1).filter(r => r[0]);

  let correctCount = 0;
  quizRows.forEach(r => {
    const qId = r[0];
    const correctAnswer = String(r[6]).trim().toUpperCase();
    const givenAnswer = String(data.answers[qId] || "").trim().toUpperCase();
    if (givenAnswer && givenAnswer === correctAnswer) correctCount++;
  });

  const total = quizRows.length;
  const scorePercent = total > 0 ? Math.round((correctCount / total) * 100) : 0;
  const passPercent = Number(getSetting("Quiz_Pass_Percent")) || 80;
  const passed = scorePercent >= passPercent;

  const attempts = (Number(record.Quiz_Attempts) || 0) + 1;
  sheet.getRange(row, colIndex("Quiz_Attempts")).setValue(attempts);
  sheet.getRange(row, colIndex("Quiz_Score")).setValue(scorePercent);
  sheet.getRange(row, colIndex("Quiz_Passed")).setValue(passed ? "Yes" : "No");

  if (passed) {
    record.Quiz_Passed = "Yes";
    try {
      issuePermitCard_(sheet, row, record);
    } catch (pdfErr) {
      return jsonResponse({ result: "error", message: "ทำแบบทดสอบผ่านแล้ว แต่ออกบัตรไม่สำเร็จ กรุณาติดต่อ Safety Team: " + pdfErr.message });
    }
  }

  return jsonResponse({ result: "ok", passed: passed, score: scorePercent, correctCount: correctCount, total: total, passPercent: passPercent });
}

// ---------------------------------------------------------------
// 4) รปภ. login ด้วย PIN + สแกนบันทึกเข้า-ออก
// ---------------------------------------------------------------
function handleGuardLogin_(data) {
  const pin = String(data.pin || "").trim();
  if (!pin) return jsonResponse({ result: "error", message: "กรุณากรอก PIN" });

  const guard = findActiveGuardByPin_(pin);
  if (!guard) return jsonResponse({ result: "error", message: "PIN ไม่ถูกต้อง หรือยังไม่ได้เปิดใช้งาน" });

  return jsonResponse({ result: "ok", guardPin: pin, guardName: guard.name, guardCompany: guard.company });
}

function findActiveGuardByPin_(pin) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUARDS_SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [rowPin, name, company, active] = rows[i];
    if (String(rowPin).trim() === pin && String(active).trim().toLowerCase() === "yes") {
      return { name: name, company: company };
    }
  }
  return null;
}

function handleGuardScan_(data) {
  const permitId = data.permitId;
  const pin = String(data.guardPin || "").trim();
  const direction = data.direction === "ออก" ? "ออก" : "เข้า";

  if (!permitId || !pin) return jsonResponse({ result: "error", message: "ข้อมูลไม่ครบถ้วน" });

  const guard = findActiveGuardByPin_(pin);
  if (!guard) return jsonResponse({ result: "error", message: "PIN รปภ. ไม่ถูกต้อง หรือหมดสิทธิ์ใช้งาน กรุณา login ใหม่" });

  const record = findPermitById_(permitId);
  if (!record) return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });
  if (record.Status !== "อนุมัติ") {
    return jsonResponse({ result: "error", message: "Permit นี้ยังไม่ได้รับการอนุมัติ ห้ามเข้าพื้นที่" });
  }

  const scanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCANLOG_SHEET_NAME);
  scanSheet.appendRow([new Date(), permitId, direction, pin, guard.name, guard.company]);

  return jsonResponse({ result: "ok", direction: direction, guardName: guard.name, permitId: permitId });
}

// ---------------------------------------------------------------
// 5) รับรูปประเมินพื้นที่จากผู้รับเหมา (site-report.html)
// ---------------------------------------------------------------
function handleSubmitSiteReport_(data) {
  const permitId = data.permitId;
  if (!permitId || !data.photoBase64) {
    return jsonResponse({ result: "error", message: "ข้อมูลไม่ครบถ้วน กรุณาแนบรูปภาพ" });
  }

  const record = findPermitById_(permitId);
  if (!record) return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });

  const commaIdx = data.photoBase64.indexOf(",");
  const base64Data = commaIdx !== -1 ? data.photoBase64.slice(commaIdx + 1) : data.photoBase64;
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, "image/jpeg", `${permitId}_${new Date().getTime()}.jpg`);

  const folder = getOrCreateFolder_(SITEREPORT_FOLDER_NAME);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SITEREPORTS_SHEET_NAME);
  sheet.appendRow([new Date(), permitId, file.getUrl(), data.note || ""]);

  return jsonResponse({ result: "ok" });
}

// ---------------------------------------------------------------
// 6) Dashboard สรุปข้อมูล realtime (dashboard.html) — ต้องใส่รหัสผ่าน
// ---------------------------------------------------------------
function handleGetDashboard_(e) {
  const password = getSetting("Dashboard_Password");
  if (!password) {
    return jsonResponse({ result: "error", message: "ยังไม่ได้ตั้งรหัสผ่าน Dashboard ในแท็บ Settings" });
  }
  if (e.parameter.pw !== password) {
    return jsonResponse({ result: "error", message: "รหัสผ่านไม่ถูกต้อง" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues() : [];
  const records = rows.map(rowToRecord_);

  const statusCol = { "รออนุมัติ": 0, "อนุมัติ": 0, "ไม่อนุมัติ": 0 };
  records.forEach(r => { if (statusCol[r.Status] !== undefined) statusCol[r.Status]++; });

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  const scanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCANLOG_SHEET_NAME);
  const scanLastRow = scanSheet.getLastRow();
  const scanRows = scanLastRow > 1 ? scanSheet.getRange(2, 1, scanLastRow - 1, SCANLOG_HEADERS.length).getValues() : [];

  let todayIn = 0, todayOut = 0;
  const onSitePermits = {};
  scanRows.forEach(r => {
    const [timestamp, permitId, direction] = r;
    if (!(timestamp instanceof Date)) return;
    if (Utilities.formatDate(timestamp, tz, "yyyy-MM-dd") !== todayStr) return;
    if (direction === "เข้า") { todayIn++; onSitePermits[permitId] = true; }
    if (direction === "ออก") { todayOut++; delete onSitePermits[permitId]; }
  });

  const recentPermits = records
    .slice(-10)
    .reverse()
    .map(r => ({
      Permit_ID: r.Permit_ID,
      Company: r.Company,
      Status: r.Status,
      Work_Location: r.Work_Location,
      Work_Date: formatDateOnly_(r.Work_Date)
    }));

  return jsonResponse({
    result: "ok",
    totalPermits: records.length,
    pendingCount: statusCol["รออนุมัติ"],
    approvedCount: statusCol["อนุมัติ"],
    rejectedCount: statusCol["ไม่อนุมัติ"],
    todayEntries: todayIn,
    todayExits: todayOut,
    onSiteNow: Object.keys(onSitePermits).length,
    recentPermits: recentPermits
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
