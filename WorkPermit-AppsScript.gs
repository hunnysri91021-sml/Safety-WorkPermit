/**
 * WorkPermit-AppsScript.gs
 * รับข้อมูลจาก work_permit_form.html แล้วบันทึกลง Google Sheet
 * บริษัท สยามกลการโลจิสติกส์ จำกัด (SML)
 *
 * วิธีติดตั้ง:
 *  1. สร้าง Google Sheet ใหม่ (เช่น ตั้งชื่อ "SML Work Permit")
 *  2. เปิด Extensions > Apps Script แล้ววางไฟล์นี้ทับโค้ดเริ่มต้น (Code.gs)
 *  3. รันฟังก์ชัน setupSheet() หนึ่งครั้งจากตัวแก้ไข Apps Script
 *     (ครั้งแรกจะขอ authorize สิทธิ์เข้าถึง Sheet ของบัญชี Google ที่สร้างไว้)
 *  4. Deploy > New deployment > เลือกประเภท "Web app"
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     กด Deploy แล้วคัดลอก Web app URL ที่ได้
 *  5. นำ URL ไปใส่ค่าคงที่ GOOGLE_SCRIPT_URL ในไฟล์ work_permit_form.html
 *  6. ตั้งค่าแชร์ของ Sheet เป็น "Restricted" แล้วเพิ่มเฉพาะอีเมลของ
 *     Safety Team ที่ต้องดูข้อมูล ห้ามเปิดเป็น "Anyone with the link"
 *     เพื่อรักษาข้อมูลส่วนบุคคลตาม PDPA
 *
 * หมายเหตุ:
 *  - Web app ที่ "Who has access: Anyone" อนุญาตให้ใครก็ได้เรียก doPost()
 *    เพื่อ "เพิ่มแถวใหม่" เท่านั้น สคริปต์นี้ไม่มีฟังก์ชันอ่าน/แก้ไข/ลบ
 *    ข้อมูลที่มีอยู่ ผู้ที่ไม่มีสิทธิ์เข้า Sheet จึงเห็นข้อมูลที่ส่งไปแล้วไม่ได้
 *  - รันซ้ำ setupSheet() ได้อย่างปลอดภัย จะไม่ลบข้อมูลที่มีอยู่
 */

const SHEET_NAME = "WorkPermit";

const HEADERS = [
  "Submit_DateTime", "Status", "Requester_Name", "Company", "Contact_Email",
  "Contact_Phone", "Work_Location", "Work_Date", "Time_From", "Time_To",
  "Work_Nature", "Work_Type", "Worker_Count", "Worker_Names", "PPE_Used",
  "Safety_Equipment", "Tools_List", "Rules_Confirmed_DateTime", "Consent_PDPA",
  "Approver_Name", "Approved_DateTime", "Reject_Reason", "Card_Exchanged", "Remarks"
];

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);
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

    return jsonResponse({ result: "ok" });
  } catch (err) {
    return jsonResponse({ result: "error", message: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
