<#
====================================================================
 Create-WorkPermitList.ps1
 สร้าง SharePoint List "WorkPermit" พร้อมคอลัมน์ทั้งหมด 24 คอลัมน์
 อ้างอิง Spec: SML_WorkPermit_Spec.docx (ส่วนที่ 2)
 บริษัท สยามกลการโลจิสติกส์ จำกัด (SML)
====================================================================

การใช้งาน:
  1. ติดตั้ง Module (รันครั้งเดียว):
     Install-Module PnP.PowerShell -Scope CurrentUser

  2. รันสคริปต์ (แทน URL ด้วย Site จริงของบริษัท):
     .\Create-WorkPermitList.ps1 -SiteUrl "https://siammotorsgroup.sharepoint.com/sites/SafetySite"

  3. หน้าต่างเบราว์เซอร์จะเปิดให้ Login ด้วยบัญชี M365 ที่มีสิทธิ์ Site Owner/Member

หมายเหตุ:
  - สคริปต์นี้ตรวจสอบก่อนว่ามี List / คอลัมน์อยู่แล้วหรือยัง ถ้ามีจะข้ามไป (รันซ้ำได้อย่างปลอดภัย)
  - ไม่ใช้ Premium Connector หรือ License พิเศษใดๆ ใช้สิทธิ์ SharePoint ปกติของผู้ใช้
====================================================================
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "URL ของ SharePoint Site เช่น https://yourtenant.sharepoint.com/sites/SafetySite")]
    [string]$SiteUrl,

    [string]$ListName = "WorkPermit"
)

$ErrorActionPreference = "Stop"

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " SML Work Permit — SharePoint List Setup" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

# ---------------------------------------------------------------
# 1) เชื่อมต่อ SharePoint
# ---------------------------------------------------------------
Write-Host "`nกำลังเชื่อมต่อไปยัง: $SiteUrl" -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive

# ---------------------------------------------------------------
# 2) สร้าง List (ถ้ายังไม่มี)
# ---------------------------------------------------------------
$existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if (-not $existingList) {
    Write-Host "`nสร้าง List ใหม่: '$ListName'" -ForegroundColor Green
    New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch | Out-Null
}
else {
    Write-Host "`nพบ List '$ListName' อยู่แล้ว — จะข้ามการสร้าง List และเพิ่มเฉพาะคอลัมน์ที่ยังไม่มี" -ForegroundColor Yellow
}

# ---------------------------------------------------------------
# 3) ฟังก์ชันช่วยสร้างคอลัมน์แบบไม่ซ้ำ (Idempotent)
# ---------------------------------------------------------------
function Add-WPField {
    param(
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type,               # Text | Note | Choice | MultiChoice | DateTime | Boolean | User
        [string[]]$Choices = @(),
        [bool]$Required = $false,
        [string]$DefaultValue = $null,
        [bool]$PlainTextNote = $true # สำหรับ Note ให้เป็น plain text ไม่ใช่ Rich Text
    )

    $existing = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  - ข้าม (มีคอลัมน์นี้แล้ว): $InternalName" -ForegroundColor DarkGray
        return
    }

    try {
        switch ($Type) {
            "Choice" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type Choice -Choices $Choices -Required:$Required -AddToDefaultView | Out-Null
            }
            "MultiChoice" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type MultiChoice -Choices $Choices -Required:$Required -AddToDefaultView | Out-Null
            }
            "Note" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type Note -Required:$Required -AddToDefaultView | Out-Null
                if ($PlainTextNote) {
                    Set-PnPField -List $ListName -Identity $InternalName -Values @{ RichText = $false } | Out-Null
                }
            }
            "DateTime" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type DateTime -Required:$Required -AddToDefaultView | Out-Null
            }
            "Boolean" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type Boolean -Required:$Required -AddToDefaultView | Out-Null
            }
            "User" {
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type User -Required:$Required -AddToDefaultView | Out-Null
            }
            default {
                # Text (Single line of text)
                Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName `
                    -Type Text -Required:$Required -AddToDefaultView | Out-Null
            }
        }

        if ($DefaultValue) {
            Set-PnPField -List $ListName -Identity $InternalName -Values @{ DefaultValue = $DefaultValue } | Out-Null
        }

        Write-Host "  + สร้างคอลัมน์: $InternalName ($Type)" -ForegroundColor Green
    }
    catch {
        Write-Host "  ! เกิดข้อผิดพลาดตอนสร้าง $InternalName : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ---------------------------------------------------------------
# 4) รายการคอลัมน์ทั้งหมด (ตรงกับ SML_WorkPermit_Spec.docx)
#    หมายเหตุ: Title เป็นคอลัมน์เริ่มต้นของ SharePoint อยู่แล้ว ไม่ต้องสร้างใหม่
#    ใช้เก็บเลขที่ Permit (WP-YYYYMM-XXX) ที่ Flow เป็นผู้กำหนดค่า
# ---------------------------------------------------------------
Write-Host "`nกำลังสร้างคอลัมน์ทั้งหมดใน List '$ListName' ..." -ForegroundColor Cyan

Add-WPField -InternalName "Status" -DisplayName "สถานะ" -Type Choice `
    -Choices @("รออนุมัติ", "อนุมัติ", "ไม่อนุมัติ", "หมดอายุ") -Required $true -DefaultValue "รออนุมัติ"

Add-WPField -InternalName "Submit_DateTime" -DisplayName "วันเวลาที่ยื่นคำขอ" -Type DateTime -Required $true

Add-WPField -InternalName "Requester_Name" -DisplayName "ชื่อ-สกุล ผู้แจ้ง" -Type Text -Required $true

Add-WPField -InternalName "Company" -DisplayName "บริษัท/หน่วยงาน" -Type Text -Required $true

Add-WPField -InternalName "Contact_Email" -DisplayName "อีเมลติดต่อ" -Type Text -Required $true

Add-WPField -InternalName "Contact_Phone" -DisplayName "เบอร์โทรติดต่อ" -Type Text -Required $true

Add-WPField -InternalName "Work_Location" -DisplayName "สถานที่ปฏิบัติงาน" -Type Text -Required $true

Add-WPField -InternalName "Work_Date" -DisplayName "วันที่เข้าปฏิบัติงาน" -Type DateTime -Required $true

Add-WPField -InternalName "Time_From" -DisplayName "เวลาเริ่ม" -Type Text -Required $true

Add-WPField -InternalName "Time_To" -DisplayName "เวลาสิ้นสุด" -Type Text -Required $true

Add-WPField -InternalName "Work_Nature" -DisplayName "ลักษณะงาน" -Type MultiChoice -Choices @(
    "งานก่อให้เกิดประกายไฟ/สะเก็ดไฟ",
    "งานที่ทำบนที่สูงเกิน 2 เมตร",
    "งานเกี่ยวกับสารเคมี",
    "งานที่ทำในสถานที่อับอากาศ",
    "งานอื่นๆ"
) -Required $false

Add-WPField -InternalName "Work_Type" -DisplayName "ประเภทงาน" -Type MultiChoice -Choices @(
    "งานก่อสร้าง/เจาะ/ขุด/ก่อ/ฉาบ/ถมดิน",
    "งานระบบไฟฟ้า",
    "งานติดตั้ง/ซ่อมแซมเครื่องจักร",
    "งานระบบประปา น้ำ ท่อน้ำ น้ำเสีย",
    "งานอื่นๆ"
) -Required $false

Add-WPField -InternalName "Worker_Count" -DisplayName "จำนวนผู้เข้าปฏิบัติงาน" -Type Choice -Choices @(
    "1-5 คน", "6-10 คน", "11-15 คน", "16 คนขึ้นไป"
) -Required $true

Add-WPField -InternalName "Worker_Names" -DisplayName "รายชื่อผู้เข้าปฏิบัติงาน" -Type Note -Required $true

Add-WPField -InternalName "PPE_Used" -DisplayName "อุปกรณ์ป้องกันภัยส่วนบุคคล (PPE)" -Type MultiChoice -Choices @(
    "หมวกนิรภัย", "รองเท้านิรภัย", "แว่นตานิรภัย", "หน้ากากเชื่อม", "ถุงมือนิรภัย", "เข็มขัดนิรภัย", "Ear Plug"
) -Required $false

Add-WPField -InternalName "Safety_Equipment" -DisplayName "อุปกรณ์ป้องกันอันตรายที่เตรียม" -Type MultiChoice -Choices @(
    "ถังดับเพลิง", "ฉากป้องกันสะเก็ดไฟ", "ผ้าใบปิดคลุมพื้นที่การทำงาน", "แนวกั้นเขตกำหนดพื้นที่"
) -Required $false

Add-WPField -InternalName "Tools_List" -DisplayName "เครื่องมือ/อุปกรณ์ไฟฟ้าที่นำเข้ามา" -Type Note -Required $false

Add-WPField -InternalName "Rules_Confirmed_DateTime" -DisplayName "วันเวลาที่ยืนยันอ่านกฎระเบียบ" -Type DateTime -Required $true

Add-WPField -InternalName "Approver_Name" -DisplayName "ผู้อนุมัติ (Safety Team)" -Type User -Required $false

Add-WPField -InternalName "Approved_DateTime" -DisplayName "วันเวลาที่อนุมัติ/ไม่อนุมัติ" -Type DateTime -Required $false

Add-WPField -InternalName "Reject_Reason" -DisplayName "เหตุผล (กรณีไม่อนุมัติ)" -Type Note -Required $false

Add-WPField -InternalName "Card_Exchanged" -DisplayName "แลกบัตรแล้ว" -Type Boolean -Required $false

Add-WPField -InternalName "Remarks" -DisplayName "หมายเหตุเพิ่มเติม" -Type Note -Required $false

# ---------------------------------------------------------------
# 5) สรุปผล
# ---------------------------------------------------------------
Write-Host "`n=====================================================" -ForegroundColor Cyan
$finalFields = Get-PnPField -List $ListName | Where-Object { -not $_.Hidden }
Write-Host " เสร็จสิ้น — List '$ListName' มีทั้งหมด $($finalFields.Count) คอลัมน์ (รวมคอลัมน์ระบบเริ่มต้น)" -ForegroundColor Green
Write-Host " ตรวจสอบผลลัพธ์ได้ที่: $SiteUrl/Lists/$ListName" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Cyan

Write-Host "`nขั้นตอนถัดไป:" -ForegroundColor Yellow
Write-Host "  1. เปิด List ตรวจสอบว่าคอลัมน์ Choice/MultiChoice แสดงตัวเลือกครบถ้วน"
Write-Host "  2. ไปที่ Power Automate สร้าง Flow ตาม SML_WorkPermit_Spec.docx ส่วนที่ 3-4"
Write-Host "  3. ตั้งค่า Permission ของ List ให้ Safety Team เป็น Contribute ขึ้นไป"
