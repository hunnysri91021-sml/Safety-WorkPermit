/**
 * WorkPermit-AppsScript.gs
 * รับข้อมูลจาก work_permit_form.html แล้วบันทึกลง Google Sheet
 * แจ้งเตือน Safety Team เมื่อมีคำขอใหม่ และแจ้งผู้รับเหมาเมื่ออนุมัติ/ไม่อนุมัติ
 * บริษัท สยามกลการโลจิสติกส์ จำกัด (SML)
 *
 * วิธีติดตั้ง:
 *  1. สร้าง Google Sheet ใหม่ (เช่น ตั้งชื่อ "SML Work Permit")
 *  2. เปิด Extensions > Apps Script แล้ววางไฟล์นี้ทับโค้ดเริ่มต้น (Code.gs)
 *  3. รันฟังก์ชัน setupSheet() หนึ่งครั้งจากตัวแก้ไข Apps Script
 *     (ครั้งแรกจะขอ authorize สิทธิ์เข้าถึง Sheet + ส่งอีเมลของบัญชี Google ที่สร้างไว้)
 *     ขั้นตอนนี้จะสร้างแท็บ "WorkPermit" (ข้อมูลคำขอ) และแท็บ "Settings"
 *     (หน้าตั้งค่า) ให้อัตโนมัติ
 *  4. Deploy > New deployment > เลือกประเภท "Web app"
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     กด Deploy แล้วคัดลอก Web app URL ที่ได้
 *  5. นำ URL ไปใส่ค่าคงที่ GOOGLE_SCRIPT_URL ในไฟล์ index.html
 *  6. เปิดแท็บ "Settings" ในชีต แล้วกรอกอีเมล Safety Team ที่ต้องการให้รับ
 *     แจ้งเตือน (คั่นด้วยจุลภาคถ้ามีหลายคน) — เปลี่ยนได้ตลอดเวลาโดยไม่ต้อง
 *     แก้โค้ด แค่แก้ค่าในเซลล์
 *  7. ตั้งค่าแชร์ของ Sheet เป็น "Restricted" แล้วเพิ่มเฉพาะอีเมลของ
 *     Safety Team ที่ต้องดูข้อมูล/อนุมัติคำขอ (ให้สิทธิ์ Editor) ห้ามเปิดเป็น
 *     "Anyone with the link" เด็ดขาด เพื่อรักษาข้อมูลส่วนบุคคลตาม PDPA
 *
 * วิธีใช้งาน (Safety Team):
 *  - เปิด Sheet แท็บ "WorkPermit" คลิกแถวคำขอที่จะพิจารณา (คลิกเซลล์ใดก็ได้ในแถวนั้น)
 *  - ไปที่เมนู "คำขอ Work Permit" ด้านบน เลือก "✅ อนุมัติแถวที่เลือก" หรือ
 *    "❌ ไม่อนุมัติแถวที่เลือก" — ระบบจะอัปเดตสถานะและอีเมลแจ้งผู้รับเหมาให้อัตโนมัติ
 *
 * หมายเหตุ:
 *  - Web app ที่ "Who has access: Anyone" อนุญาตให้ใครก็ได้เรียก doPost()
 *    เพื่อ "เพิ่มแถวใหม่" เท่านั้น สคริปต์นี้ไม่มีฟังก์ชันอ่าน/ลบข้อมูลผ่าน URL นี้
 *    ผู้ที่ไม่มีสิทธิ์เข้า Sheet จึงเห็นข้อมูลที่ส่งไปแล้วไม่ได้
 *  - รันซ้ำ setupSheet() ได้อย่างปลอดภัย จะไม่ลบข้อมูล/ค่าตั้งค่าที่มีอยู่
 *  - บัญชี Gmail ส่วนตัวส่งอีเมลผ่าน MailApp ได้ฟรีประมาณวันละ 100 ฉบับ
 *    เพียงพอสำหรับปริมาณคำขอ Work Permit ทั่วไป
 */

const SHEET_NAME = "WorkPermit";
const SETTINGS_SHEET_NAME = "Settings";

const HEADERS = [
  "Submit_DateTime", "Status", "Requester_Name", "Company", "Contact_Email",
  "Contact_Phone", "Work_Location", "Work_Date", "Time_From", "Time_To",
  "Work_Nature", "Work_Type", "Worker_Count", "Worker_Names", "PPE_Used",
  "Safety_Equipment", "Tools_List", "Rules_Confirmed_DateTime", "Consent_PDPA",
  "Approver_Name", "Approved_DateTime", "Reject_Reason", "Card_Exchanged", "Remarks"
];

const DEFAULT_SETTINGS = [
  ["Key", "Value", "คำอธิบาย"],
  ["Safety_Team_Emails", "", "อีเมล Safety Team ที่รับแจ้งเตือนคำขอใหม่ (คั่นด้วยจุลภาคถ้ามีหลายคน)"],
  ["Company_Name", "บริษัท สยามกลการโลจิสติกส์ จำกัด", "ชื่อบริษัท ใช้ในอีเมลแจ้งเตือน"]
];

function colIndex(name) {
  return HEADERS.indexOf(name) + 1;
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }
  // เติมเฉพาะแถวที่ยังว่าง ไม่ทับค่าที่ admin ตั้งไว้แล้วถ้ารันซ้ำ
  const existingRows = settingsSheet.getLastRow();
  if (existingRows < DEFAULT_SETTINGS.length) {
    settingsSheet.getRange(existingRows + 1, 1, DEFAULT_SETTINGS.length - existingRows, 3)
      .setValues(DEFAULT_SETTINGS.slice(existingRows));
  }
  settingsSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  settingsSheet.setFrozenRows(1);
  settingsSheet.autoResizeColumns(1, 3);
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

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (!data.Requester_Name || !data.Company || !data.Contact_Email || !data.Work_Location || !data.Work_Date) {
      return jsonResponse({ result: "error", message: "ข้อมูลที่จำเป็นไม่ครบถ้วน" });
    }
    if (!data.Consent_PDPA) {
      return jsonResponse({ result: "error", message: "ต้องยินยอมให้เก็บข้อมูลส่วนบุคคลก่อนส่งคำขอ" });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    const row = [
      new Date(),                                              // Submit_DateTime — ใช้เวลาเซิร์ฟเวอร์ ไม่พึ่งนาฬิกาเครื่องผู้ส่ง
      "รออนุมัติ",                                              // Status
      data.Requester_Name,
      data.Company,
      data.Contact_Email,
      data.Contact_Phone || "",
      data.Work_Location,
      data.Work_Date,
      data.Time_From || "",
      data.Time_To || "",
      Array.isArray(data.Work_Nature) ? data.Work_Nature.join(", ") : "",
      Array.isArray(data.Work_Type) ? data.Work_Type.join(", ") : "",
      data.Worker_Count || "",
      data.Worker_Names || "",
      Array.isArray(data.PPE_Used) ? data.PPE_Used.join(", ") : "",
      Array.isArray(data.Safety_Equipment) ? data.Safety_Equipment.join(", ") : "",
      data.Tools_List || "",
      data.Rules_Confirmed_DateTime || "",
      "ยินยอม",
      "", "", "", "", ""                                       // Approver_Name, Approved_DateTime, Reject_Reason, Card_Exchanged, Remarks — Safety Team กรอกภายหลัง
    ];

    sheet.appendRow(row);
    notifySafetyTeamOfNewRequest(data);

    return jsonResponse({ result: "ok" });
  } catch (err) {
    return jsonResponse({ result: "error", message: err.message });
  }
}

function notifySafetyTeamOfNewRequest(data) {
  const emails = getSetting("Safety_Team_Emails");
  if (!emails) return; // ยังไม่ได้ตั้งค่าอีเมลในแท็บ Settings — ข้ามการแจ้งเตือน

  const companyName = getSetting("Company_Name") || "SML";
  const subject = `[Work Permit ใหม่] ${data.Company} — ${data.Work_Location}`;
  const body = [
    `มีคำขอ Work Permit ใหม่รอพิจารณาจาก ${companyName}`,
    "",
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

function approveSelectedRow() {
  processDecision_(true);
}

function rejectSelectedRow() {
  processDecision_(false);
}

function processDecision_(isApproved) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) {
    SpreadsheetApp.getUi().alert(`กรุณาไปที่แท็บ "${SHEET_NAME}" แล้วเลือกแถวคำขอก่อน`);
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert("แถวที่เลือกเป็นหัวตาราง กรุณาเลือกแถวคำขอจริง");
    return;
  }

  const requesterName = sheet.getRange(row, colIndex("Requester_Name")).getValue();
  const contactEmail = sheet.getRange(row, colIndex("Contact_Email")).getValue();
  if (!requesterName || !contactEmail) {
    SpreadsheetApp.getUi().alert("แถวนี้ไม่มีข้อมูลคำขอ (อาจเป็นแถวว่าง)");
    return;
  }

  let rejectReason = "";
  if (!isApproved) {
    const resp = SpreadsheetApp.getUi().prompt(
      "เหตุผลที่ไม่อนุมัติ",
      "กรุณาระบุเหตุผล (จะถูกส่งไปในอีเมลแจ้งผู้รับเหมา):",
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
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

  notifyContractorOfDecision_(requesterName, contactEmail, isApproved, rejectReason);
  SpreadsheetApp.getUi().alert(`บันทึกผล "${isApproved ? "อนุมัติ" : "ไม่อนุมัติ"}" และส่งอีเมลแจ้งผู้รับเหมาแล้ว`);
}

function notifyContractorOfDecision_(requesterName, contactEmail, isApproved, rejectReason) {
  const companyName = getSetting("Company_Name") || "SML";
  const subject = isApproved
    ? `Work Permit ของท่านได้รับการอนุมัติแล้ว — ${companyName}`
    : `Work Permit ของท่านไม่ได้รับการอนุมัติ — ${companyName}`;

  const bodyLines = [`เรียน คุณ${requesterName}`, ""];
  if (isApproved) {
    bodyLines.push("คำขอ Work Permit ของท่านได้รับการอนุมัติแล้ว กรุณาแลกบัตรที่ป้อม รปภ. ก่อนเข้าปฏิบัติงานตามวันเวลาที่แจ้งไว้");
  } else {
    bodyLines.push("คำขอ Work Permit ของท่านไม่ได้รับการอนุมัติ");
    bodyLines.push(`เหตุผล: ${rejectReason || "-"}`);
    bodyLines.push("กรุณาติดต่อ Safety Team ของบริษัทฯ หากต้องการสอบถามเพิ่มเติม");
  }

  MailApp.sendEmail({ to: contactEmail, subject: subject, body: bodyLines.join("\n") });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
