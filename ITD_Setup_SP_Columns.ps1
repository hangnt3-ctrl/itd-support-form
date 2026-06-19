# ITD_Setup_SP_Columns.ps1
# Tao va cap nhat cot SharePoint List: ITD_Tickets
# Site  : https://haseduvn.sharepoint.com/sites/ITDSupport
# List  : ITD_Tickets
# Yeu cau: PnP PowerShell (Install-Module PnP.PowerShell)
# Chay  : .\ITD_Setup_SP_Columns.ps1

$SiteUrl  = "https://haseduvn.sharepoint.com/sites/ITDSupport"
$ListName = "ITD_Tickets"

# Azure AD app da dang ky cho form (cf3fc815...) -- can them quyen Sites.FullControl hoac SharePoint admin
$ClientId = "cf3fc815-e55f-45ee-a451-9dc5b0c233aa"
$TenantId = "010a6d4c-53ee-42a6-b68e-74c646a3297c"

Write-Host "`n[ITD] Ket noi SharePoint..." -ForegroundColor Cyan

# Thu cach 1: dung app da dang ky (can quyen Sites.Manage hoac SharePoint Admin)
try {
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive -ErrorAction Stop
    Write-Host "[OK] Da ket noi voi app $ClientId" -ForegroundColor Green
} catch {
    Write-Host "[!!] App $ClientId that bai: $_" -ForegroundColor Yellow
    Write-Host "     Thu cach 2: WebLogin (dung cookie trinh duyet)..." -ForegroundColor Cyan
    # Cach 2: dung session trinh duyet (khong can consent app)
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    Write-Host "[OK] Da ket noi qua WebLogin" -ForegroundColor Green
}

function Test-ColumnExists($ListName, $FieldName) {
    try { Get-PnPField -List $ListName -Identity $FieldName -ErrorAction Stop | Out-Null; return $true }
    catch { return $false }
}
function Write-OK($msg)   { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Write-New($msg)  { Write-Host "  [+]  $msg" -ForegroundColor Yellow }
function Write-Skip($msg) { Write-Host "  [--] $msg" -ForegroundColor Gray }
function Write-Err($msg)  { Write-Host "  [!!] $msg" -ForegroundColor Red }

# ============================================================
# BUOC 1 -- Tao cot co kieu dac biet (DateTime)
# ============================================================
Write-Host "`n[BUOC 1] Tao cot DateTime con thieu..." -ForegroundColor Cyan

foreach ($col in @(
    @{ Name="resolved_time"; Display="Thoi gian giai quyet"; Type="DateTime" }
)) {
    if (-not (Test-ColumnExists $ListName $col.Name)) {
        try {
            Add-PnPField -List $ListName -DisplayName $col.Display -InternalName $col.Name `
                         -Type $col.Type -AddToDefaultView $false
            Write-New "$($col.Name) ($($col.Type)) -- DA TAO"
        } catch { Write-Err "$($col.Name): $_" }
    } else {
        Write-Skip "$($col.Name) -- da ton tai"
    }
}

# ============================================================
# BUOC 2 -- TicketStatus (Choice)
# ============================================================
Write-Host "`n[BUOC 2] TicketStatus (Choice)..." -ForegroundColor Cyan

$StatusChoices = @("Cho tiep nhan","Dang xu ly","Cho xac nhan quan ly","Hoan thanh","Tu choi")

try {
    if (Test-ColumnExists $ListName "TicketStatus") {
        Set-PnPField -List $ListName -Identity "TicketStatus" -Values @{
            Choices      = $StatusChoices
            DefaultValue = "Cho tiep nhan"
        }
        Write-OK "TicketStatus -- cap nhat 5 gia tri Choice"
    } else {
        Add-PnPFieldFromXml -List $ListName -FieldXml @"
<Field Type="Choice" DisplayName="Trang thai phieu" Name="TicketStatus" Format="Dropdown" FillInChoice="FALSE">
  <Default>Cho tiep nhan</Default>
  <CHOICES>
    <CHOICE>Cho tiep nhan</CHOICE>
    <CHOICE>Dang xu ly</CHOICE>
    <CHOICE>Cho xac nhan quan ly</CHOICE>
    <CHOICE>Hoan thanh</CHOICE>
    <CHOICE>Tu choi</CHOICE>
  </CHOICES>
</Field>
"@
        Write-New "TicketStatus -- da tao moi"
    }
} catch { Write-Err "TicketStatus: $_" }

# ============================================================
# BUOC 2B -- TicketPriority (Choice)
# ============================================================
Write-Host "`n[BUOC 2B] TicketPriority (Choice)..." -ForegroundColor Cyan

try {
    if (Test-ColumnExists $ListName "TicketPriority") {
        Set-PnPField -List $ListName -Identity "TicketPriority" -Values @{
            Choices      = @("P1","P2","P3","P4")
            DefaultValue = "P3"
        }
        Write-OK "TicketPriority -- cap nhat P1/P2/P3/P4"
    } else {
        Add-PnPFieldFromXml -List $ListName -FieldXml @"
<Field Type="Choice" DisplayName="Do uu tien" Name="TicketPriority" Format="Dropdown" FillInChoice="FALSE">
  <Default>P3</Default>
  <CHOICES>
    <CHOICE>P1</CHOICE>
    <CHOICE>P2</CHOICE>
    <CHOICE>P3</CHOICE>
    <CHOICE>P4</CHOICE>
  </CHOICES>
</Field>
"@
        Write-New "TicketPriority -- da tao moi"
    }
} catch { Write-Err "TicketPriority: $_" }

# ============================================================
# BUOC 3 -- Xac minh va tao cot con thieu
# ============================================================
Write-Host "`n[BUOC 3] Kiem tra toan bo cot bat buoc..." -ForegroundColor Cyan

$RequiredColumns = @(
    # He thong
    @{ Name="ticket_id";            Display="Ma phieu";                   Type="Text"     },
    @{ Name="sla";                  Display="SLA";                        Type="Text"     },
    @{ Name="submission_time";      Display="Thoi gian gui";              Type="DateTime" },
    # Nguoi gui
    @{ Name="submitter_name";       Display="Ten nguoi gui";              Type="Text"     },
    @{ Name="submitter_email";      Display="Email nguoi gui";            Type="Text"     },
    # Phan loai
    @{ Name="TicketGroup";          Display="Nhom ho tro";                Type="Text"     },
    @{ Name="account_issue_type";   Display="Loai van de tai khoan";      Type="Text"     },
    @{ Name="permission_type";      Display="Loai phan quyen";            Type="Text"     },
    @{ Name="TicketTool";           Display="Cong cu lien quan";          Type="Text"     },
    # Noi dung
    @{ Name="TicketDescription";    Display="Mo ta yeu cau";              Type="Note"     },
    @{ Name="TicketImpact";         Display="Muc do anh huong";           Type="Text"     },
    # Quan ly (manager_form) -- ManagerName va manager_department da xoa (form khong thu thap)
    @{ Name="ManagerEmail";         Display="Email quan ly cap 1";        Type="Text"     },
    @{ Name="manager_email";        Display="Email quan ly cap 2";        Type="Text"     },
    @{ Name="manager_reason";       Display="Ly do xac nhan";             Type="Note"     },
    # Tai khoan moi
    @{ Name="new_account_role";     Display="Chuc danh tai khoan moi";    Type="Text"     },
    @{ Name="new_account_group";    Display="Nhom tai khoan moi";         Type="Text"     },
    # Nguoi dung Q14
    @{ Name="requester_name";       Display="Ho ten nguoi can tao TK";    Type="Text"     },
    @{ Name="requester_email";      Display="Email nguoi can tao TK";     Type="Text"     },
    @{ Name="requester_dept";       Display="Phong ban nguoi can tao TK"; Type="Text"     },
    @{ Name="user_start_date";      Display="Ngay bat dau lam viec";      Type="Text"     },
    # Backup -- attachments_info da xoa (file luu o SP Attachments va full_data)
    @{ Name="full_data";            Display="Toan bo du lieu JSON";       Type="Note"     }
)

$skipList = @("resolved_time","TicketStatus","TicketPriority")
$missing  = @()

foreach ($col in $RequiredColumns) {
    if (Test-ColumnExists $ListName $col.Name) { Write-OK $col.Name }
    else { Write-Err "$($col.Name) -- THIEU"; $missing += $col }
}

if ($missing.Count -gt 0) {
    Write-Host "`n  Tao $($missing.Count) cot con thieu..." -ForegroundColor Yellow
    foreach ($col in $missing) {
        if ($col.Name -in $skipList) { continue }
        try {
            $ft = if ($col.Type -eq "Note") { "Note" } elseif ($col.Type -eq "DateTime") { "DateTime" } else { "Text" }
            Add-PnPField -List $ListName -DisplayName $col.Display -InternalName $col.Name `
                         -Type $ft -AddToDefaultView $false
            Write-New "$($col.Name) ($ft) -- DA TAO"
        } catch { Write-Err "$($col.Name): $_" }
    }
}

# ============================================================
# BUOC 4 -- Tao View ITD Admin
# ============================================================
Write-Host "`n[BUOC 4] Tao View 'ITD Admin'..." -ForegroundColor Cyan

try {
    $existingView = Get-PnPView -List $ListName -Identity "ITD Admin" -ErrorAction SilentlyContinue
    if (-not $existingView) {
        Add-PnPView -List $ListName -Title "ITD Admin" `
            -Fields @("ticket_id","submitter_name","submitter_email",
                      "TicketGroup","TicketTool","TicketStatus","TicketPriority",
                      "sla","submission_time","ManagerEmail","manager_email","resolved_time") `
            -Query '<OrderBy><FieldRef Name="submission_time" Ascending="FALSE"/></OrderBy>' `
            -RowLimit 100
        Write-New "View 'ITD Admin' -- DA TAO"
    } else {
        Write-Skip "View 'ITD Admin' -- da ton tai"
    }
} catch { Write-Err "View: $_" }

Write-Host "`n============================================" -ForegroundColor Green
Write-Host " HOAN THANH -- Kiem tra lai tai:" -ForegroundColor Green
Write-Host " $SiteUrl/Lists/$ListName/AllItems.aspx" -ForegroundColor Green
Write-Host "============================================`n" -ForegroundColor Green
