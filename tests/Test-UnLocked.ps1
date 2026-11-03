#requires -Version 5.1
<#
    UnLocked 自動測試（不需 Office）
      - 所有 .ps1 語法檢查
      - FileInspector.cs / OfficeInterop.cs 可由 Windows PowerShell 5.1 (C# 5) 編譯
      - 以 fixtures 內的正常／毀損／受保護範例檔驗證結構檢查結果
      - 共用輔助函式
    用法：powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-UnLocked.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root 'src'
$fixtures = Join-Path $PSScriptRoot 'fixtures'
$script:Failures = 0
$script:Passed = 0

function Assert-That {
    param([bool]$Condition, [string]$Name, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host "  [PASS] $Name" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "  [FAIL] $Name $Detail" -ForegroundColor Red }
}

Write-Host '== 語法檢查 ==' -ForegroundColor Cyan
foreach ($f in Get-ChildItem -Path $src, $PSScriptRoot -Filter *.ps1) {
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
    $detail = ($errors | ForEach-Object { "第 $($_.Extent.StartLineNumber) 行：$($_.Message)" }) -join '；'
    Assert-That ($errors.Count -eq 0) $f.Name $detail
}

Write-Host '== 檔案編碼（Windows PowerShell 5.1 需要 UTF-8 BOM 才能正確讀取中文） ==' -ForegroundColor Cyan
foreach ($f in Get-ChildItem -Path $src, $PSScriptRoot, (Join-Path $root 'build') -Include *.ps1, *.cs -Recurse) {
    $b = [IO.File]::ReadAllBytes($f.FullName)
    Assert-That ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) ("BOM: " + $f.Name)
}

Write-Host '== 編譯 C# 輔助類別 ==' -ForegroundColor Cyan
. (Join-Path $src 'UnLockedEngine.ps1')
try {
    Initialize-UnLocked
    Assert-That $true 'Initialize-UnLocked'
}
catch {
    Assert-That $false 'Initialize-UnLocked' $_.Exception.Message
    exit 1
}
Assert-That ([bool]('UnLocked.FileInspector' -as [type])) 'UnLocked.FileInspector 已載入'
Assert-That ([bool]('UnLocked.Watchdog' -as [type])) 'UnLocked.Watchdog 已載入'

Write-Host '== 結構檢查 ==' -ForegroundColor Cyan
$expected = [IO.File]::ReadAllText((Join-Path $fixtures 'expected.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$fixtureCount = @($expected.PSObject.Properties).Count
foreach ($case in $expected.PSObject.Properties) {
    $r = Invoke-UnLockedInspection (Join-Path $fixtures $case.Name)
    foreach ($k in $case.Value.PSObject.Properties) {
        if ($k.Name -eq 'IssueContains') {
            Assert-That ($r.IssueText.Contains([string]$k.Value)) "$($case.Name) 問題包含「$($k.Value)」" ("實際：" + $r.IssueText)
        }
        else {
            $actual = [string]$r.($k.Name)
            Assert-That ($actual -eq [string]$k.Value) "$($case.Name) $($k.Name) = $($k.Value)" "（實際：${actual}；問題：$($r.IssueText)）"
        }
    }
}
$ok = Invoke-UnLockedInspection (Join-Path $fixtures 'valid.xlsx')
Assert-That ($ok.Issues.Count -eq 0 -and $ok.Warnings.Count -eq 0) 'valid.xlsx 沒有任何問題或警告' ($ok.IssueText + $ok.WarningText)
Assert-That ((Invoke-UnLockedInspection (Join-Path $fixtures 'not-exist.xlsx')).Status -eq 'ERROR') '不存在的檔案 → ERROR'

Write-Host '== 輔助函式 ==' -ForegroundColor Cyan
Assert-That ((ConvertTo-UnLockedGuid '{6F1C2C49-6B5A-4D1E-9A51-7A3C1E0D4B21}') -eq '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21') 'ConvertTo-UnLockedGuid 去除大括號並轉小寫'
Assert-That ((ConvertTo-UnLockedGuid 'abc') -eq '') 'ConvertTo-UnLockedGuid 拒絕非 GUID'
foreach ($ext in @('.xlsx', '.xlsm', '.xlsb', '.xls', '.xltx', '.xltm', '.xlt', '.docx', '.docm', '.doc', '.dotx', '.dotm', '.dot',
        '.pptx', '.pptm', '.ppt', '.ppsx', '.ppsm', '.pps', '.potx', '.potm', '.pot')) {
    Assert-That (Test-UnLockedSupported ('C:\a\報表' + $ext)) "支援 $ext"
    Assert-That ($null -ne $script:SaveFormats[$ext]) "$ext 有對應的另存格式"
}
Assert-That ((Get-UnLockedKind 'a.PPTX') -eq 'PowerPoint' -and (Get-UnLockedKind 'a.Docx') -eq 'Word' -and (Get-UnLockedKind 'a.xlsb') -eq 'Excel') 'Get-UnLockedKind 不分大小寫'
$filter = Get-UnLockedDialogFilter
Assert-That (($filter -split '\|').Count -eq 10 -and $filter.Contains('*.pptx') -and $filter.Contains('*.docx') -and $filter.Contains('*.xlsx')) '檔案選擇視窗篩選條件' $filter
Assert-That (-not (Test-UnLockedSupported 'C:\a\~$報表.xlsx')) '略過 Office 鎖定檔 ~$'
Assert-That (-not (Test-UnLockedSupported 'C:\a\報表.~unlocked.xlsx')) '略過本工具暫存檔'
Assert-That (-not (Test-UnLockedSupported 'C:\a\_UnLocked_Output\報表.xlsx')) '略過輸出資料夾'
Assert-That (-not (Test-UnLockedSupported 'C:\a\報表.pdf')) '不支援 .pdf'
Assert-That ((Get-UnLockedRelativePath 'C:\data\sub\a.xlsx' 'C:\data') -eq 'sub\a.xlsx') '相對路徑保留子資料夾'
Assert-That ((Get-UnLockedRelativePath 'D:\x\a.xlsx' 'C:\data') -eq 'a.xlsx') '不在基準資料夾時只取檔名'
$items = @(Get-UnLockedFileItems -Paths @($fixtures) -Recurse $true)
Assert-That ($items.Count -eq $fixtureCount) "列舉 fixtures 找到 $fixtureCount 個 Office 檔（實際 $($items.Count)）"
Assert-That ((Resolve-UnLockLabelName '' '' @()) -eq '(無標籤)') 'Resolve-UnLockLabelName 無標籤'
$known = @(New-UnLockedLabel -Name 'Public\Public-Edit' -Id '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' -SiteId '')
Assert-That ((Resolve-UnLockLabelName '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' 'Public-Edit' $known) -eq 'Public\Public-Edit') 'Resolve-UnLockLabelName 使用已知名稱'
$dest = Get-UnLockDestination -Item ([pscustomobject]@{ Path = 'C:\data\sub\a.xlsx'; BaseDir = 'C:\data' }) -Options @{ OutputMode = 'Folder'; OutputFolder = 'E:\out' }
Assert-That ($dest -eq 'E:\out\sub\a.xlsx') "輸出路徑（實際 ${dest}）"

Write-Host '== 範例檔標籤讀取（不需 Office） ==' -ForegroundColor Cyan
$l = Read-LabelFromSampleFile -Path (Join-Path $fixtures 'labeled.xlsx')
Assert-That ($l.Id -eq '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' -and $l.SiteId -eq '0e6c3c7a-1b2d-4c5e-8f90-a1b2c3d4e5f6' -and $l.Name -eq 'Public-Edit') 'Read-LabelFromSampleFile'

Write-Host '== 檢查模式批次（不需 Office） ==' -ForegroundColor Cyan
$sync = New-UnLockSync
$res = @(Invoke-UnLockBatch -Items $items -Options @{ Mode = 'Check'; DeepCheck = $false } -Sync $sync)
Assert-That ($res.Count -eq $fixtureCount) "批次檢查回傳 $fixtureCount 筆結果"
Assert-That ($sync.Done) '批次完成旗標'
$byName = @{}
foreach ($x in $res) { $byName[[IO.Path]::GetFileName($x.File)] = $x }
Assert-That ($byName['valid.xlsx'].Result -eq '正常') 'valid.xlsx → 正常'
Assert-That ($byName['crc.xlsx'].Result -eq '需修復') 'crc.xlsx → 需修復'
Assert-That ($byName['password.xlsx'].Result -eq '略過') 'password.xlsx → 略過'
Assert-That ($byName['irm.xlsx'].IrmBefore -eq '受限制') 'irm.xlsx → 受限制'
$csv = Join-Path ([IO.Path]::GetTempPath()) ('unlocked_test_{0}.csv' -f [guid]::NewGuid().ToString('N'))
Export-UnLockReport -Results $res -Path $csv
Assert-That (@(Import-Csv -LiteralPath $csv -Encoding UTF8).Count -eq $fixtureCount) '匯出 CSV 報表'
Remove-Item -LiteralPath $csv -Force

Write-Host ''
Write-Host ("通過 {0}，失敗 {1}" -f $script:Passed, $script:Failures) -ForegroundColor $(if ($script:Failures) { 'Red' } else { 'Green' })
exit $script:Failures
