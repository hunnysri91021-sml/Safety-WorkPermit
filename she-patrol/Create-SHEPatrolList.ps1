<#
.NOTE
    SUPERSEDED as the app's live backend: index.html now talks to
    SHE-Patrol-AppsScript.gs (Google Sheets) instead, because the Power Automate
    HTTP-trigger gateway needed to let non-M365-licensed staff use the app
    requires a paid per-user Power Automate plan, which isn't included in the
    M365 Business Basic license SML is on. This script and the provisioned
    SharePoint List/Library/Groups still exist on the tenant and remain here for
    reference / possible future use. See README.md "สถาปัตยกรรม" for the current
    architecture.

.SYNOPSIS
    Provisions the SHE Patrol Digital System on a SharePoint Online site:
    - List: SHE_Patrol_Findings (16 columns)
    - Document Library: SHE_Patrol_Photos (FindingID, PhotoStage columns)
    - List: SHE_Patrol_Users (Name, Email, Role, Shop, Active — login roster managed by Admin)
    - 4 permission groups: SHE-Auditor, SHE-Dept-Responsible, SHE-Safety-Admin, SHE-Executive-Viewer

.DESCRIPTION
    Idempotent — safe to re-run. Every list/library/field/group is created only if it
    does not already exist; existing objects are left untouched.

    Requires the PnP.PowerShell module:
        Install-Module -Name PnP.PowerShell -Scope CurrentUser

.PARAMETER SiteUrl
    Full URL of the target SharePoint site, e.g. https://siammotor.sharepoint.com/sites/Chosiya_Server

.PARAMETER ClientId
    Entra ID App (client) ID to authenticate with. Required if your tenant has disabled the
    default multi-tenant "PnP Management Shell" app — register your own with:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SHE Patrol PnP"
    then pass the AzureAppId/ClientId it prints here.

.EXAMPLE
    ./Create-SHEPatrolList.ps1 -SiteUrl "https://siammotor.sharepoint.com/sites/Chosiya_Server" -ClientId "27ebd1dc-a57c-40c3-9d8e-035dcbd23d08"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module not found. Installing for current user..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
if ($ClientId) {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# --------------------------------------------------------------------------
# 1. List: SHE_Patrol_Findings
# --------------------------------------------------------------------------

$listName = "SHE_Patrol_Findings"

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "Creating list '$listName' ..." -ForegroundColor Cyan
    $list = New-PnPList -Title $listName -Template GenericList -EnableVersioning -OnQuickLaunch
} else {
    Write-Host "List '$listName' already exists — skipping creation." -ForegroundColor DarkGray
}

# Field definitions: InternalName, DisplayName, Type, ExtraArgs (hashtable of Add-PnPField params)
$findingFields = @(
    @{ Name = "PatrolDate";               Display = "วันที่ตรวจ (Patrol Date)";              Type = "DateTime" }
    @{ Name = "Shop";                      Display = "Shop";                                    Type = "Choice"; Choices = @("PDI","Acc","Yard","Washing","Touch up","Store","อื่นๆ") }
    @{ Name = "Place";                     Display = "จุดที่พบ (Place)";                        Type = "Text" }
    @{ Name = "Description";               Display = "รายละเอียด (Description)";               Type = "Note" }
    @{ Name = "Grade";                     Display = "Grade";                                   Type = "Choice"; Choices = @("A","B","C","Others") }
    @{ Name = "Category";                  Display = "Category";                                Type = "Choice"; Choices = @("F","S","EES","5S") }
    @{ Name = "PhotoBeforeUrl";            Display = "รูปก่อนแก้ไข (Photo Before)";            Type = "URL" }
    @{ Name = "DueDate";                   Display = "กำหนดเสร็จ (Due Date)";                  Type = "DateTime" }
    @{ Name = "RootCause";                 Display = "สาเหตุ (Root Cause)";                    Type = "Note" }
    @{ Name = "ActionResponsible";         Display = "ผู้รับผิดชอบ (Action / Responsible)";    Type = "User" }
    @{ Name = "Countermeasure";            Display = "มาตรการแก้ไข (Countermeasure)";         Type = "Note" }
    @{ Name = "PhotoAfterUrl";             Display = "รูปหลังแก้ไข (Photo After)";             Type = "URL" }
    @{ Name = "Status";                    Display = "สถานะ (Status)";                         Type = "Choice"; Choices = @("เปิดใหม่","รอดำเนินการ","ดำเนินการแล้ว","รอตรวจสอบ","ปิดงาน") }
    @{ Name = "VerifiedBy";                Display = "ผู้ตรวจยืนยัน (Verified By)";             Type = "User" }
    @{ Name = "Rules_Confirmed_DateTime";  Display = "เวลาที่ยืนยันปิดงาน (Rules Confirmed)";  Type = "DateTime"; DisplayFormat = "DateTime" }
)

foreach ($f in $findingFields) {
    $existing = Get-PnPField -List $listName -Identity $f.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Field '$($f.Name)' already exists — skipping." -ForegroundColor DarkGray
        continue
    }
    Write-Host "  Adding field '$($f.Name)' ($($f.Type)) ..." -ForegroundColor Cyan
    switch ($f.Type) {
        "Choice" {
            Add-PnPField -List $listName -DisplayName $f.Display -InternalName $f.Name -Type Choice -Choices $f.Choices -AddToDefaultView | Out-Null
        }
        default {
            Add-PnPField -List $listName -DisplayName $f.Display -InternalName $f.Name -Type $f.Type -AddToDefaultView | Out-Null
        }
    }
    if ($f.Type -eq "DateTime") {
        $format = if ($f.DisplayFormat) { $f.DisplayFormat } else { "DateOnly" }
        Set-PnPField -List $listName -Identity $f.Name -Values @{ DisplayFormat = $format } | Out-Null
    }
}

# Default Status to "เปิดใหม่" for new items
$statusField = Get-PnPField -List $listName -Identity "Status" -ErrorAction SilentlyContinue
if ($statusField -and -not $statusField.DefaultValue) {
    Set-PnPField -List $listName -Identity "Status" -Values @{ DefaultValue = "เปิดใหม่" } | Out-Null
}

# --------------------------------------------------------------------------
# 2. Document Library: SHE_Patrol_Photos
# --------------------------------------------------------------------------

$libraryName = "SHE_Patrol_Photos"

$library = Get-PnPList -Identity $libraryName -ErrorAction SilentlyContinue
if (-not $library) {
    Write-Host "Creating document library '$libraryName' ..." -ForegroundColor Cyan
    $library = New-PnPList -Title $libraryName -Template DocumentLibrary -OnQuickLaunch
} else {
    Write-Host "Library '$libraryName' already exists — skipping creation." -ForegroundColor DarkGray
}

$libraryFields = @(
    @{ Name = "FindingID";   Display = "FindingID";                              Type = "Text" }
    @{ Name = "PhotoStage";  Display = "ขั้นตอนรูป (Photo Stage)"; Type = "Choice"; Choices = @("ก่อนแก้ไข","หลังแก้ไข") }
)

foreach ($f in $libraryFields) {
    $existing = Get-PnPField -List $libraryName -Identity $f.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Field '$($f.Name)' already exists — skipping." -ForegroundColor DarkGray
        continue
    }
    Write-Host "  Adding field '$($f.Name)' ($($f.Type)) ..." -ForegroundColor Cyan
    switch ($f.Type) {
        "Choice" {
            Add-PnPField -List $libraryName -DisplayName $f.Display -InternalName $f.Name -Type Choice -Choices $f.Choices -AddToDefaultView | Out-Null
        }
        default {
            Add-PnPField -List $libraryName -DisplayName $f.Display -InternalName $f.Name -Type $f.Type -AddToDefaultView | Out-Null
        }
    }
}

# --------------------------------------------------------------------------
# 2b. List: SHE_Patrol_Users (login roster — Admin assigns Role/Shop here)
# --------------------------------------------------------------------------

$usersListName = "SHE_Patrol_Users"

$usersList = Get-PnPList -Identity $usersListName -ErrorAction SilentlyContinue
if (-not $usersList) {
    Write-Host "Creating list '$usersListName' ..." -ForegroundColor Cyan
    $usersList = New-PnPList -Title $usersListName -Template GenericList -OnQuickLaunch
} else {
    Write-Host "List '$usersListName' already exists — skipping creation." -ForegroundColor DarkGray
}

$userFields = @(
    @{ Name = "Name";    Display = "ชื่อ-นามสกุล";       Type = "Text" }
    @{ Name = "Email";   Display = "อีเมล";               Type = "Text" }
    @{ Name = "Role";    Display = "บทบาท (Role)";        Type = "Choice"; Choices = @("SHE-Auditor","SHE-Dept-Responsible","SHE-Safety-Admin","SHE-Executive-Viewer") }
    @{ Name = "Shop";    Display = "Shop (เฉพาะ Dept-Responsible)"; Type = "Choice"; Choices = @("","PDI","Acc","Yard","Washing","Touch up","Store","อื่นๆ") }
    @{ Name = "Active";  Display = "ใช้งานอยู่";          Type = "Boolean" }
)

foreach ($f in $userFields) {
    $existing = Get-PnPField -List $usersListName -Identity $f.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Field '$($f.Name)' already exists — skipping." -ForegroundColor DarkGray
        continue
    }
    Write-Host "  Adding field '$($f.Name)' ($($f.Type)) ..." -ForegroundColor Cyan
    switch ($f.Type) {
        "Choice" {
            Add-PnPField -List $usersListName -DisplayName $f.Display -InternalName $f.Name -Type Choice -Choices $f.Choices -AddToDefaultView | Out-Null
        }
        default {
            Add-PnPField -List $usersListName -DisplayName $f.Display -InternalName $f.Name -Type $f.Type -AddToDefaultView | Out-Null
        }
    }
}

$activeField = Get-PnPField -List $usersListName -Identity "Active" -ErrorAction SilentlyContinue
if ($activeField -and -not $activeField.DefaultValue) {
    Set-PnPField -List $usersListName -Identity "Active" -Values @{ DefaultValue = "1" } | Out-Null
}

# --------------------------------------------------------------------------
# 3. Permission groups
# --------------------------------------------------------------------------

function Ensure-UniquePermissions {
    param([string]$ListTitle)
    $list = Get-PnPList -Identity $ListTitle
    Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments | Out-Null
    if (-not $list.HasUniqueRoleAssignments) {
        Write-Host "  Breaking role inheritance on '$ListTitle' (keeping existing access) ..." -ForegroundColor Cyan
        $list.BreakRoleInheritance($true, $false)
        Invoke-PnPQuery
    }
}

Write-Host "Preparing lists for unique permissions ..." -ForegroundColor Cyan
Ensure-UniquePermissions -ListTitle $listName
Ensure-UniquePermissions -ListTitle $libraryName
Ensure-UniquePermissions -ListTitle $usersListName

$groupDefs = @(
    @{ Name = "SHE-Auditor";            Role = "Contribute";  Description = "ผู้ตรวจ (SHE Committee / Safety Officer) — สร้าง finding ใหม่" }
    @{ Name = "SHE-Dept-Responsible";   Role = "Contribute";  Description = "หน่วยงานที่ถูกตรวจพบ — อัปเดตมาตรการแก้ไขและรูปหลัง (จำกัดเฉพาะแผนกตน ดูหมายเหตุด้านล่าง)" }
    @{ Name = "SHE-Safety-Admin";       Role = "Full Control"; Description = "เจ้าหน้าที่ความปลอดภัย — แก้ไข Grade, ตรวจสอบ, ปิดงาน, ดู dashboard รวม" }
    @{ Name = "SHE-Executive-Viewer";   Role = "Read";         Description = "ผู้บริหาร — ดู dashboard/รายงานเท่านั้น" }
)

foreach ($g in $groupDefs) {
    $group = Get-PnPGroup -Identity $g.Name -ErrorAction SilentlyContinue
    if (-not $group) {
        Write-Host "Creating group '$($g.Name)' ..." -ForegroundColor Cyan
        $group = New-PnPGroup -Title $g.Name -Description $g.Description
    } else {
        Write-Host "Group '$($g.Name)' already exists — skipping creation." -ForegroundColor DarkGray
    }

    Write-Host "  Granting '$($g.Role)' on '$listName' to '$($g.Name)' ..." -ForegroundColor Cyan
    Set-PnPListPermission -Identity $listName -Group $g.Name -AddRole $g.Role | Out-Null

    Write-Host "  Granting '$($g.Role)' on '$libraryName' to '$($g.Name)' ..." -ForegroundColor Cyan
    Set-PnPListPermission -Identity $libraryName -Group $g.Name -AddRole $g.Role | Out-Null

    # SHE_Patrol_Users: only Safety-Admin can manage the roster; everyone else needs
    # Read so the app can look up their own Role/Shop at sign-in.
    $usersRole = if ($g.Name -eq "SHE-Safety-Admin") { "Full Control" } else { "Read" }
    Write-Host "  Granting '$usersRole' on '$usersListName' to '$($g.Name)' ..." -ForegroundColor Cyan
    Set-PnPListPermission -Identity $usersListName -Group $g.Name -AddRole $usersRole | Out-Null
}

Write-Host ""
Write-Host "=== Provisioning complete ===" -ForegroundColor Green
Write-Host "Site: $SiteUrl"
Write-Host "List: $listName (16 columns)"
Write-Host "Library: $libraryName (FindingID, PhotoStage)"
Write-Host "List: $usersListName (Name, Email, Role, Shop, Active)"
Write-Host "Groups: $($groupDefs.Name -join ', ')"
Write-Host ""
Write-Host "ขั้นตอนถัดไป:" -ForegroundColor Yellow
Write-Host "  1) เพิ่ม Admin คนแรกใน '$usersListName' ด้วยตัวเอง (Name/Email/Role=SHE-Safety-Admin/Active=Yes)" -ForegroundColor Yellow
Write-Host "     ผ่านหน้า SharePoint List โดยตรง — แอปยังไม่มีใคร login ได้จนกว่าจะมีแถวนี้" -ForegroundColor Yellow
Write-Host "  2) สร้าง Power Automate Flow 5 (API Gateway) ตาม power-automate/Flow5-API-Gateway.md" -ForegroundColor Yellow
Write-Host "     แล้วนำ URL ไปใส่ใน APP_CONFIG.API_URL ของ index.html — จำเป็นเพื่อให้ผู้ใช้ที่ไม่มี M365 ใช้งานได้" -ForegroundColor Yellow
Write-Host ""
Write-Host "หมายเหตุ (SHE-Dept-Responsible scoping):" -ForegroundColor Yellow
Write-Host "  กลุ่ม SHE-Dept-Responsible ได้รับสิทธิ์ Contribute บนทั้ง List/Library ตามที่สเปกกำหนด" -ForegroundColor Yellow
Write-Host "  แต่ front-end เชื่อมผ่าน Flow 5 (Power Automate) ไม่ใช่ SharePoint ตรง จึงบังคับ 'เฉพาะแผนกตน'" -ForegroundColor Yellow
Write-Host "  ที่ระดับ UI ผ่านคอลัมน์ Shop ใน '$usersListName' แทน — กลุ่ม SharePoint เหล่านี้ยังมีประโยชน์" -ForegroundColor Yellow
Write-Host "  สำหรับคนที่เข้าดู/export List ตรงจาก SharePoint เอง" -ForegroundColor Yellow

Disconnect-PnPOnline
