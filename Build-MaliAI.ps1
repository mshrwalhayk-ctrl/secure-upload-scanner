#Requires -Version 5.1
<#
================================================================================
  بناء وتركيب إضافة إكسل:  مساعد المحاسب  (MaliAI)
================================================================================
  الاستخدام:
      كليك يمين على ملف "تركيب الإضافة.cmd"  ←  Run as administrator غير مطلوب
      أو من PowerShell:   .\Build-MaliAI.ps1

  شو بيعمل السكربت:
      1) بيفعّل صلاحية الوصول لمشروع VBA (مطلوبة لبناء الإضافة وتشغيل الماكرو)
      2) بيفتح إكسل مخفي، بينشئ الوحدات البرمجية والنماذج
      3) بيحفظ الملف MaliAI.xlam في مجلد إضافات إكسل
      4) بيحقن شريط الأدوات العربي (Ribbon)
      5) بيسجّل الإضافة عشان تشتغل تلقائياً مع كل فتحة إكسل
================================================================================
#>

[CmdletBinding()]
param(
    [switch]$NoInstall,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$script:ExcelApp = $null

# ------------------------------------------------------------------ أدوات عرض
function Say  ($m) { Write-Host $m -ForegroundColor Gray }
function Ok   ($m) { Write-Host "  [OK]   $m" -ForegroundColor Green }
function Warn ($m) { Write-Host "  [!]    $m" -ForegroundColor Yellow }
function Step ($m) { Write-Host "`n>> $m" -ForegroundColor Cyan }
function Die  ($m) { Write-Host "`n[X] $m`n" -ForegroundColor Red; Cleanup; Read-Host "اضغط Enter للخروج"; exit 1 }

function Cleanup {
    if ($script:ExcelApp) {
        try { $script:ExcelApp.DisplayAlerts = $false; $script:ExcelApp.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelApp) } catch {}
        $script:ExcelApp = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "==================================================================" -ForegroundColor White
Write-Host "        مساعد المحاسب  -  MaliAI Excel Add-in  -  Setup" -ForegroundColor White
Write-Host "==================================================================" -ForegroundColor White

# ------------------------------------------------------------------ المسارات
$root    = Split-Path -Parent $PSScriptRoot
$srcDir  = Join-Path $root 'src'
$ribXml  = Join-Path $root 'ribbon\customUI14.xml'

if (-not (Test-Path $srcDir)) { Die "ما لقيت مجلد src في: $root" }
if (-not (Test-Path $ribXml)) { Die "ما لقيت ملف الشريط: $ribXml" }

if ($OutputPath) {
    $addinPath = $OutputPath
} else {
    $addinDir  = Join-Path $env:APPDATA 'Microsoft\AddIns'
    if (-not (Test-Path $addinDir)) { New-Item -ItemType Directory -Path $addinDir -Force | Out-Null }
    $addinPath = Join-Path $addinDir 'MaliAI.xlam'
}
Say "  الوجهة: $addinPath"

# ------------------------------------------------ 1) صلاحية الوصول لمشروع VBA
Step "1/6  تفعيل صلاحية الوصول لمشروع VBA"
$verFound = $false
foreach ($v in '16.0','15.0','14.0') {
    $base = "HKCU:\Software\Microsoft\Office\$v\Excel"
    if (Test-Path $base) {
        $verFound = $true
        $sec = Join-Path $base 'Security'
        if (-not (Test-Path $sec)) { New-Item -Path $sec -Force | Out-Null }
        New-ItemProperty -Path $sec -Name 'AccessVBOM' -Value 1 -PropertyType DWord -Force | Out-Null
        Ok "Office $v  ->  AccessVBOM = 1"
    }
}
if (-not $verFound) { Warn "ما قدرت أحدد نسخة أوفيس من السجل — رح نكمل عادي." }

# ------------------------------------------------------------ 2) تشغيل إكسل
Step "2/6  تشغيل إكسل"
$running = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Warn "إكسل شغّال حالياً. الأفضل تسكّره قبل التركيب."
    $a = Read-Host "  اكتب y واضغط Enter للمتابعة على أي حال، أو أي حرف تاني للإلغاء"
    if ($a -ne 'y') { Die "تم الإلغاء. سكّر إكسل وأعد التشغيل." }
}

try {
    $xl = New-Object -ComObject Excel.Application
} catch {
    Die "ما قدرت أشغّل إكسل. تأكد إنه مايكروسوفت إكسل مثبّت على الجهاز."
}
$script:ExcelApp = $xl
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
Ok "إكسل جاهز (نسخة $($xl.Version))"

$wb = $xl.Workbooks.Add()

# التحقق من صلاحية VBA
try {
    $vbp = $wb.VBProject
    $null = $vbp.VBComponents.Count
} catch {
    Die @"
صلاحية الوصول لمشروع VBA مقفلة.

فعّلها يدوياً:
  إكسل  ->  File  ->  Options  ->  Trust Center  ->  Trust Center Settings
        ->  Macro Settings  ->  ضع علامة على:
            Trust access to the VBA project object model

ثم سكّر إكسل وأعد تشغيل هذا السكربت.
"@
}
Ok "صلاحية VBA مفتوحة"

# ------------------------------------------------------ 3) إضافة الوحدات
Step "3/6  إضافة الوحدات البرمجية"

function Read-Code ($path) {
    $t = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
    $t = $t -replace "`r`n", "`n"
    $t = $t -replace "`n", "`r`n"
    return $t
}

$modFiles = @(
    'modJson.bas','modConfig.bas','modHttp.bas','modApi.bas',
    'modMath.bas','modSchema.bas','modAudit.bas','modTools.bas','modAgent.bas','modUI.bas','modTest.bas'
)

foreach ($f in $modFiles) {
    $p = Join-Path $srcDir $f
    if (-not (Test-Path $p)) { Die "ملف مفقود: $p" }

    $code = Read-Code $p
    $name = [IO.Path]::GetFileNameWithoutExtension($f)

    # حذف سطر Attribute VB_Name (الاسم بنحطه برمجياً)
    $code = ($code -split "`r`n" | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_Name' }) -join "`r`n"

    $c = $vbp.VBComponents.Add(1)      # 1 = وحدة قياسية
    $c.Name = $name
    $c.CodeModule.AddFromString($code)
    Ok "$name  ($([math]::Round($code.Length/1024,1)) كيلوبايت)"
}

# ---------------------------------------------------------- 4) بناء النماذج
Step "4/6  بناء نوافذ الواجهة"

function Set-FormProp($comp, $name, $value) {
    try { $comp.Properties.Item($name).Value = $value } catch { }
}

function Add-Ctl($designer, $progId, $name, $left, $top, $width, $height) {
    $c = $designer.Controls.Add($progId, $name, $true)
    $c.Left = $left; $c.Top = $top; $c.Width = $width; $c.Height = $height
    try { $c.Font.Name = 'Tahoma'; $c.Font.Size = 10 } catch {}
    return $c
}

# --------------------------- frmChat ---------------------------
$fc = $vbp.VBComponents.Add(3)          # 3 = UserForm
$fc.Name = 'frmChat'
Set-FormProp $fc 'Caption' 'مساعد المحاسب'
Set-FormProp $fc 'Width'  620
Set-FormProp $fc 'Height' 476
$dc = $fc.Designer

$t = Add-Ctl $dc 'Forms.TextBox.1' 'txtLog' 8 8 596 300
$t.MultiLine = $true; $t.WordWrap = $true; $t.ScrollBars = 2
$t.Locked = $true; $t.TextAlign = 3; $t.BackColor = 16250871
try { $t.Font.Size = 10 } catch {}

$t = Add-Ctl $dc 'Forms.Label.1' 'lblStatus' 8 312 596 14
$t.TextAlign = 3; $t.Caption = 'جاهز'; $t.ForeColor = 8421504

$t = Add-Ctl $dc 'Forms.TextBox.1' 'txtInput' 8 330 596 62
$t.MultiLine = $true; $t.WordWrap = $true; $t.ScrollBars = 2; $t.TextAlign = 3
try { $t.Font.Size = 11 } catch {}

$t = Add-Ctl $dc 'Forms.CommandButton.1' 'btnSend' 452 400 152 30
$t.Caption = 'إرسال  (Ctrl+Enter)'
try { $t.Font.Bold = $true } catch {}

$t = Add-Ctl $dc 'Forms.CommandButton.1' 'btnUndo'     360 400 86 30 ; $t.Caption = 'تراجع'
$t = Add-Ctl $dc 'Forms.CommandButton.1' 'btnNew'      278 400 76 30 ; $t.Caption = 'محادثة جديدة'
$t = Add-Ctl $dc 'Forms.CommandButton.1' 'btnSettings' 196 400 76 30 ; $t.Caption = 'الإعدادات'
$t = Add-Ctl $dc 'Forms.CommandButton.1' 'btnCopy'     126 400 64 30 ; $t.Caption = 'نسخ السجل'
$t = Add-Ctl $dc 'Forms.CommandButton.1' 'btnHelp'      66 400 54 30 ; $t.Caption = 'مساعدة'

$fc.CodeModule.AddFromString( (Read-Code (Join-Path $srcDir 'frmChat.code.txt')) )
Ok "frmChat"

# --------------------------- frmSettings ---------------------------
$fs = $vbp.VBComponents.Add(3)
$fs.Name = 'frmSettings'
Set-FormProp $fs 'Caption' 'الإعدادات'
Set-FormProp $fs 'Width'  600
Set-FormProp $fs 'Height' 468
$ds = $fs.Designer

function Lbl($d, $name, $top, $text) {
    $l = Add-Ctl $d 'Forms.Label.1' $name 8 $top 574 14
    $l.Caption = $text; $l.TextAlign = 3
    return $l
}

Lbl $ds 'lblProvider' 8 'مزوّد الذكاء الاصطناعي' | Out-Null
$t = Add-Ctl $ds 'Forms.ComboBox.1' 'cboProvider' 8 24 574 20 ; $t.Style = 2 ; $t.TextAlign = 3

Lbl $ds 'lblKey' 50 'مفتاح API  (بينحفظ على جهازك بس)' | Out-Null
$t = Add-Ctl $ds 'Forms.TextBox.1' 'txtKey' 8 66 574 20 ; $t.PasswordChar = '*'

Lbl $ds 'lblBaseUrl' 92 'رابط الـ endpoint  (للمزوّد المخصّص فقط)' | Out-Null
$t = Add-Ctl $ds 'Forms.TextBox.1' 'txtBaseUrl' 8 108 574 20

Lbl $ds 'lblFormat' 134 'نوع الواجهة' | Out-Null
$t = Add-Ctl $ds 'Forms.ComboBox.1' 'cboFormat' 332 150 250 20 ; $t.Style = 2

Lbl $ds 'lblModel' 176 'النموذج' | Out-Null
$t = Add-Ctl $ds 'Forms.ComboBox.1' 'cboModel' 150 192 432 20 ; $t.TextAlign = 3
$t = Add-Ctl $ds 'Forms.CommandButton.1' 'btnRefresh' 8 192 136 20 ; $t.Caption = 'تحديث النماذج'

$t = Add-Ctl $ds 'Forms.CheckBox.1' 'chkFreeOnly' 402 220 180 18
$t.Caption = 'اعرض المجاني فقط ★'
$t = Add-Ctl $ds 'Forms.CheckBox.1' 'chkMacros' 402 242 180 18
$t.Caption = 'اسمح بتشغيل الماكرو'
$t = Add-Ctl $ds 'Forms.CheckBox.1' 'chkBackup' 402 264 180 18
$t.Caption = 'نسخة احتياطية قبل الماكرو'

$t = Add-Ctl $ds 'Forms.Label.1' 'lblSteps' 232 244 160 16
$t.Caption = 'أقصى عدد خطوات للمهمة'; $t.TextAlign = 3
$t = Add-Ctl $ds 'Forms.TextBox.1' 'txtSteps' 180 242 46 20 ; $t.TextAlign = 2

$t = Add-Ctl $ds 'Forms.Label.1' 'lblHint' 8 292 574 34
$t.TextAlign = 3; $t.ForeColor = 8421504

$t = Add-Ctl $ds 'Forms.Label.1' 'lblStatus' 8 332 574 42
$t.TextAlign = 3
try { $t.Font.Bold = $true } catch {}

$t = Add-Ctl $ds 'Forms.CommandButton.1' 'btnSave'   458 384 124 30 ; $t.Caption = 'حفظ'
try { $t.Font.Bold = $true } catch {}
$t = Add-Ctl $ds 'Forms.CommandButton.1' 'btnTest'   330 384 124 30 ; $t.Caption = 'اختبار الاتصال'
$t = Add-Ctl $ds 'Forms.CommandButton.1' 'btnCancel' 222 384 104 30 ; $t.Caption = 'إلغاء'

$fs.CodeModule.AddFromString( (Read-Code (Join-Path $srcDir 'frmSettings.code.txt')) )
Ok "frmSettings"

# ------------------------------------------------ كود ThisWorkbook
$twb = $vbp.VBComponents.Item('ThisWorkbook')
$twb.CodeModule.AddFromString( (Read-Code (Join-Path $srcDir 'ThisWorkbook.code.txt')) )
Ok "ThisWorkbook"

# ------------------------------------------------------------- 5) الحفظ
Step "5/6  حفظ ملف الإضافة"

try { $vbp.Name = 'MaliAI' } catch {}
try { $wb.BuiltinDocumentProperties.Item('Title').Value = 'مساعد المحاسب' } catch {}
try { $wb.BuiltinDocumentProperties.Item('Comments').Value = 'MaliAI Excel Assistant v1.0' } catch {}
$wb.IsAddin = $true

if (Test-Path $addinPath) {
    try { Remove-Item $addinPath -Force } catch { Die "الملف مفتوح أو مقفل: $addinPath — سكّر إكسل وجرّب." }
}
$wb.SaveAs($addinPath, 55)      # 55 = xlOpenXMLAddIn
$wb.Close($false)
Ok "تم الحفظ"

$xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
$script:ExcelApp = $null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep -Milliseconds 1200

# --------------------------------------------------- 6) حقن الشريط العربي
Step "6/6  تركيب شريط الأدوات العربي"

Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$ribbonContent = [IO.File]::ReadAllText($ribXml, [Text.Encoding]::UTF8)

try {
    $zip = [System.IO.Compression.ZipFile]::Open($addinPath, 'Update')

    # (أ) إضافة ملف الشريط
    $old = $zip.GetEntry('customUI/customUI14.xml')
    if ($old) { $old.Delete() }
    $e  = $zip.CreateEntry('customUI/customUI14.xml')
    $sw = New-Object System.IO.StreamWriter($e.Open(), $utf8NoBom)
    $sw.Write($ribbonContent); $sw.Flush(); $sw.Dispose()

    # (ب) ربط الشريط في _rels/.rels
    $relEntry = $zip.GetEntry('_rels/.rels')
    $sr   = New-Object System.IO.StreamReader($relEntry.Open())
    $rels = $sr.ReadToEnd(); $sr.Dispose()

    if ($rels -notmatch 'customUI14\.xml') {
        $newRel = '<Relationship Id="rIdMaliAICustomUI" ' +
                  'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" ' +
                  'Target="customUI/customUI14.xml"/>'
        $rels = $rels -replace '</Relationships>', ($newRel + '</Relationships>')

        $relEntry.Delete()
        $e2 = $zip.CreateEntry('_rels/.rels')
        $sw2 = New-Object System.IO.StreamWriter($e2.Open(), $utf8NoBom)
        $sw2.Write($rels); $sw2.Flush(); $sw2.Dispose()
    }

    # (ج) التأكد من تسجيل امتداد xml في [Content_Types].xml
    $ctEntry = $zip.GetEntry('[Content_Types].xml')
    $sr3 = New-Object System.IO.StreamReader($ctEntry.Open())
    $ct  = $sr3.ReadToEnd(); $sr3.Dispose()
    if ($ct -notmatch 'Extension="xml"') {
        $ct = $ct -replace '(<Types[^>]*>)', ('$1<Default Extension="xml" ContentType="application/xml"/>')
        $ctEntry.Delete()
        $e3 = $zip.CreateEntry('[Content_Types].xml')
        $sw3 = New-Object System.IO.StreamWriter($e3.Open(), $utf8NoBom)
        $sw3.Write($ct); $sw3.Flush(); $sw3.Dispose()
    }

    $zip.Dispose()
    Ok "الشريط جاهز"
}
catch {
    Warn "تعذّر تركيب الشريط ($($_.Exception.Message))."
    Warn "الإضافة رح تشتغل عادي، بس الأوامر رح تظهر في تبويب Add-ins بدل تبويب خاص."
    try { $zip.Dispose() } catch {}
}

# ------------------------------------------------------------ التسجيل
if (-not $NoInstall) {
    Step "تسجيل الإضافة في إكسل"
    try {
        $xl2 = New-Object -ComObject Excel.Application
        $script:ExcelApp = $xl2
        $xl2.Visible = $false
        $xl2.DisplayAlerts = $false
        $null = $xl2.Workbooks.Add()

        $ai = $xl2.AddIns.Add($addinPath, $false)
        $ai.Installed = $true

        $xl2.DisplayAlerts = $false
        $xl2.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
        $script:ExcelApp = $null
        Ok "الإضافة مسجّلة وبتفتح تلقائياً مع إكسل"
    }
    catch {
        Warn "ما قدرت أسجّلها تلقائياً: $($_.Exception.Message)"
        Warn "سجّلها يدوياً: File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse"
        Warn "واختر الملف: $addinPath"
    }
}

[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Green
Write-Host "  تمّ التركيب بنجاح" -ForegroundColor Green
Write-Host "==================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  الملف:  $addinPath"
Write-Host ""
Write-Host "  الخطوة الجاية:" -ForegroundColor White
Write-Host "   1. افتح إكسل"
Write-Host "   2. روح على تبويب:  مساعد المحاسب"
Write-Host "   3. اضغط: الإعدادات  ->  حط المفتاح  ->  تحديث النماذج  ->  اختبار الاتصال  ->  حفظ"
Write-Host "   4. اضغط: افتح المساعد   (أو Ctrl+Shift+A)"
Write-Host ""
Read-Host "اضغط Enter للإغلاق"
