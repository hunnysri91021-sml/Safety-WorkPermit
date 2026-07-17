/**
 * WorkPermit-AppsScript.gs
 * รับข้อมูลจาก index.html แล้วบันทึกลง Google Sheet, ออกเลขที่ Permit,
 * แจ้งเตือน Safety Team เมื่อมีคำขอใหม่, ให้ Safety Team อนุมัติ/ไม่อนุมัติ
 * ผ่านเมนูในชีต, แจ้งผลไปยังผู้รับเหมาพร้อมไฟล์ PDF ที่มี QR Code สำหรับ
 * ตรวจสอบสถานะที่หน้างาน
 * บริษัท สยามกลการโลจิสติกส์ จำกัด (SML)
 *
 * วิธีติดตั้ง:
 *  1. สร้าง Google Sheet ใหม่ (เช่น ตั้งชื่อ "SML Work Permit")
 *  2. เปิด Extensions > Apps Script แล้ววางไฟล์นี้ทับโค้ดเริ่มต้น (Code.gs)
 *  3. รันฟังก์ชัน setupSheet() หนึ่งครั้งจากตัวแก้ไข Apps Script
 *     (ครั้งแรกจะขอ authorize สิทธิ์เข้าถึง Sheet + Drive + ส่งอีเมล)
 *     จะสร้างแท็บ "WorkPermit" (ข้อมูลคำขอ) และแท็บ "Settings" ให้อัตโนมัติ
 *  4. Deploy > New deployment > เลือกประเภท "Web app"
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     กด Deploy แล้วคัดลอก Web app URL ที่ได้
 *  5. นำ URL ไปใส่ค่าคงที่ GOOGLE_SCRIPT_URL ในไฟล์ index.html และ view.html
 *  6. เปิดแท็บ "Settings" ในชีต แล้วกรอก:
 *       - Safety_Team_Emails: อีเมล Safety Team (คั่นด้วยจุลภาคถ้ามีหลายคน)
 *       - Site_Base_URL: URL ของเว็บที่โฮสต์ index.html/view.html เช่น
 *         https://hunnysri91021-sml.github.io/Safety-WorkPermit/
 *     เปลี่ยนได้ตลอดเวลาโดยไม่ต้องแก้โค้ด
 *  7. ตั้งค่าแชร์ของ Sheet เป็น "Restricted" แล้วเพิ่มเฉพาะอีเมลของ
 *     Safety Team ที่ต้องดูข้อมูล/อนุมัติคำขอ (ให้สิทธิ์ Editor) ห้ามเปิดเป็น
 *     "Anyone with the link" เด็ดขาด เพื่อรักษาข้อมูลส่วนบุคคลตาม PDPA
 *
 * วิธีใช้งาน (Safety Team):
 *  - เปิด Sheet แท็บ "WorkPermit" คลิกแถวคำขอที่จะพิจารณา (คลิกเซลล์ใดก็ได้ในแถวนั้น)
 *  - ไปที่เมนู "คำขอ Work Permit" ด้านบน เลือก "✅ อนุมัติแถวที่เลือก" หรือ
 *    "❌ ไม่อนุมัติแถวที่เลือก" — ระบบจะอัปเดตสถานะ, ออก PDF พร้อม QR Code
 *    (เฉพาะกรณีอนุมัติ), และอีเมลแจ้งผู้รับเหมาให้อัตโนมัติ
 *
 * หมายเหตุด้านความปลอดภัยของข้อมูล (PDPA):
 *  - Web app "Anyone" เรียก doPost() ได้แค่ "เพิ่มแถวใหม่" เท่านั้น
 *  - doGet() (ใช้โดย view.html ตอนสแกน QR) ส่งคืนเฉพาะข้อมูลที่จำเป็นต่อการ
 *    ตรวจสอบหน้างาน (ชื่อผู้แจ้ง บริษัท สถานที่ วันเวลา สถานะ) — ไม่ส่งอีเมล
 *    เบอร์โทร รายชื่อผู้ปฏิบัติงานแบบเต็ม หรือข้อมูลอ่อนไหวอื่นกลับไป
 *  - ไฟล์ PDF ที่สร้างจะแชร์แบบ "Anyone with the link" เฉพาะไฟล์นั้นไฟล์เดียว
 *    (ไม่ใช่ทั้งโฟลเดอร์/ทั้งชีต) เพื่อให้ผู้รับเหมาที่ไม่มีบัญชีองค์กรเปิดได้
 *    เนื่องจากชื่อไฟล์/ลิงก์คาดเดายาก และส่งให้เฉพาะเจ้าของคำขอทางอีเมลเท่านั้น
 *  - รันซ้ำ setupSheet() ได้อย่างปลอดภัย จะไม่ลบข้อมูล/ค่าตั้งค่าที่มีอยู่
 *  - บัญชี Gmail ส่วนตัวส่งอีเมลผ่าน MailApp ได้ฟรีประมาณวันละ 100 ฉบับ
 */

const SHEET_NAME = "WorkPermit";
const SETTINGS_SHEET_NAME = "Settings";
const PDF_FOLDER_NAME = "Work Permit PDFs";

const HEADERS = [
  "Permit_ID", "Submit_DateTime", "Status", "Requester_Name", "Company", "Contact_Email",
  "Contact_Phone", "Work_Location", "Work_Date", "Time_From", "Time_To",
  "Work_Nature", "Work_Type", "Worker_Count", "Worker_Names", "PPE_Used",
  "Safety_Equipment", "Tools_List", "Rules_Confirmed_DateTime", "Consent_PDPA",
  "Approver_Name", "Approved_DateTime", "Reject_Reason", "Card_Exchanged", "Remarks", "PDF_Link"
];

const DEFAULT_SETTINGS = [
  ["Key", "Value", "คำอธิบาย"],
  ["Safety_Team_Emails", "", "อีเมล Safety Team ที่รับแจ้งเตือนคำขอใหม่ (คั่นด้วยจุลภาคถ้ามีหลายคน)"],
  ["Company_Name", "บริษัท สยามกลการโลจิสติกส์ จำกัด", "ชื่อบริษัท ใช้ในอีเมล/เอกสาร PDF"],
  ["Site_Base_URL", "", "URL ของเว็บที่โฮสต์ index.html/view.html เช่น https://ชื่อบัญชี.github.io/Safety-WorkPermit/"]
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

// ---------------------------------------------------------------
// รับคำขอใหม่จากฟอร์ม (index.html)
// ---------------------------------------------------------------
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

    notifySafetyTeamOfNewRequest(data, permitId);

    return jsonResponse({ result: "ok", permitId: permitId });
  } catch (err) {
    return jsonResponse({ result: "error", message: err.message });
  }
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
// ค้นหาสถานะ Permit ตามเลขที่ (ใช้โดย view.html ตอนสแกน QR)
// ---------------------------------------------------------------
function doGet(e) {
  const id = e.parameter.id;
  if (!id) {
    return jsonResponse({ result: "error", message: "ไม่พบเลขที่ Permit ใน URL" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const idCol = colIndex("Permit_ID") - 1;
  const rowValues = values.find(r => r[idCol] === id);
  if (!rowValues) {
    return jsonResponse({ result: "error", message: "ไม่พบข้อมูล Permit นี้" });
  }

  const record = rowToRecord_(rowValues);

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
    Approved_DateTime: formatDateTime_(record.Approved_DateTime)
  });
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
function notifySafetyTeamOfNewRequest(data, permitId) {
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

  let pdfBlob = null;
  if (isApproved) {
    try {
      const siteBaseUrl = getSetting("Site_Base_URL");
      const viewUrl = siteBaseUrl
        ? siteBaseUrl.replace(/\/?$/, "/") + "view.html?id=" + encodeURIComponent(record.Permit_ID)
        : "";
      const pdfFile = generatePermitPdf_(record, viewUrl);
      sheet.getRange(row, colIndex("PDF_Link")).setValue(pdfFile.getUrl());
      pdfBlob = pdfFile.getBlob();
    } catch (pdfErr) {
      // ไม่ให้การออก PDF ที่ล้มเหลวบล็อกการอนุมัติ — แค่แจ้งเตือนใน UI
      ui.alert("อนุมัติสำเร็จ แต่สร้าง PDF ไม่สำเร็จ: " + pdfErr.message);
    }
  }

  notifyContractorOfDecision_(record, rejectReason, pdfBlob);
  ui.alert(`บันทึกผล "${isApproved ? "อนุมัติ" : "ไม่อนุมัติ"}" (${record.Permit_ID}) และส่งอีเมลแจ้งผู้รับเหมาแล้ว`);
}

function notifyContractorOfDecision_(record, rejectReason, pdfBlob) {
  const companyName = getSetting("Company_Name") || "SML";
  const isApproved = record.Status === "อนุมัติ";
  const subject = isApproved
    ? `Work Permit ${record.Permit_ID} ได้รับการอนุมัติแล้ว — ${companyName}`
    : `Work Permit ${record.Permit_ID} ไม่ได้รับการอนุมัติ — ${companyName}`;

  const bodyLines = [`เรียน คุณ${record.Requester_Name}`, "", `เลขที่ Permit: ${record.Permit_ID}`];
  if (isApproved) {
    bodyLines.push("คำขอ Work Permit ของท่านได้รับการอนุมัติแล้ว กรุณาแสดงไฟล์ PDF ที่แนบมา (มี QR Code สำหรับตรวจสอบ) พร้อมบัตรประชาชนที่ป้อม รปภ. ก่อนเข้าปฏิบัติงานตามวันเวลาที่แจ้งไว้");
  } else {
    bodyLines.push("คำขอ Work Permit ของท่านไม่ได้รับการอนุมัติ");
    bodyLines.push(`เหตุผล: ${rejectReason || "-"}`);
    bodyLines.push("กรุณาติดต่อ Safety Team ของบริษัทฯ หากต้องการสอบถามเพิ่มเติม");
  }

  const options = pdfBlob ? { attachments: [pdfBlob] } : {};
  MailApp.sendEmail({ to: record.Contact_Email, subject: subject, body: bodyLines.join("\n"), ...options });
}

// ---------------------------------------------------------------
// ออกเอกสาร PDF พร้อม QR Code (เรียกเฉพาะตอนอนุมัติ)
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
    qrImgTag = `<div class="qr-box"><img src="${qrDataUri}"><div class="cap">สแกนเพื่อตรวจสอบสถานะ Permit นี้</div></div>`;
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
      <h1>ใบอนุญาตปฏิบัติงาน (Work Permit)</h1>
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

  const folder = getOrCreatePdfFolder_();
  const file = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

function getOrCreatePdfFolder_() {
  const folders = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PDF_FOLDER_NAME);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
