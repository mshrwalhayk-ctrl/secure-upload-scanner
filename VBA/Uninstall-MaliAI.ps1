#Requires -Version 5.1
<#
  إزالة إضافة مساعد المحاسب (MaliAI) من إكسل
  ملاحظة: بيتم كمان مسح المفتاح والإعدادات المحفوظة على الجهاز.
#>
$ErrorActionPreference = 'SilentlyContinue'

Write-Host ""
Write-Host "إزالة إضافة: مساعد المحاسب" -ForegroundColor Cyan
Write-Host ""

$addinPath = Join-Path $env:APPDATA 'Microsoft\AddIns\MaliAI.xlam'

# 1) إلغاء التسجيل من إكسل
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add()
    foreach ($a in $xl.AddIns) {
        if ($a.Name -eq 'MaliAI.xlam') { $a.Installed = $false }
    }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Write-Host "  [OK]   تم إلغاء التسجيل من إكسل" -ForegroundColor Green
} catch {
    Write-Host "  [!]    ما قدرت ألغي التسجيل تلقائياً — ألغيه من: Options > Add-ins" -ForegroundColor Yellow
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep -Milliseconds 800

# 2) حذف الملف
if (Test-Path $addinPath) {
    Remove-Item $addinPath -Force
    Write-Host "  [OK]   تم حذف الملف: $addinPath" -ForegroundColor Green
}

# 3) مسح الإعدادات (المفتاح محفوظ هون)
$reg = 'HKCU:\Software\VB and VBA Program Settings\MaliAI'
if (Test-Path $reg) {
    Remove-Item $reg -Recurse -Force
    Write-Host "  [OK]   تم مسح الإعدادات والمفتاح المحفوظ" -ForegroundColor Green
}

Write-Host ""
Write-Host "خلصت الإزالة." -ForegroundColor Green
Read-Host "اضغط Enter للإغلاق"
