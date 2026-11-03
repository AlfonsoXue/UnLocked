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
$lab = Format-InspectionSummary (Invoke-UnLockedInspection (Join-Path $fixtures 'irm_label.xlsx'))
Assert-That ($lab.Contains('Internal - Internal (Contoso)') -and $lab.Contains('加密') -and $lab.Contains('Azure RMS')) 'IRM 標籤加密來源摘要' $lab
$adr = Format-InspectionSummary (Invoke-UnLockedInspection (Join-Path $fixtures 'irm_adrms.xlsx'))
Assert-That ($adr.Contains('AD RMS') -and $adr.Contains('rms.contoso.example')) 'AD RMS 來源摘要' $adr
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

$comEx = New-Object System.Runtime.InteropServices.COMException('無法取得類別 Workbooks 的 Open 屬性', -2146827284)
$friendly = Get-FriendlyError (New-Object System.Management.Automation.ErrorRecord($comEx, 'Open', 'NotSpecified', $null))
Assert-That ($friendly.Contains('[0x800A03EC]') -and $friendly.Contains('Workbooks')) 'Get-FriendlyError 附上錯誤碼' $friendly
$openParams = (Get-Command Open-OfficeDocument).Parameters
Assert-That ($openParams.ContainsKey('UsePassword') -and $openParams.ContainsKey('Interactive')) 'Open-OfficeDocument 支援 UsePassword／Interactive'
Assert-That ([bool](Get-UnLockedDefaultConfig).IrmInteractive) '預設開啟 IRM 互動重試'

Write-Host '== 範例檔標籤讀取（不需 Office） ==' -ForegroundColor Cyan
$l = Read-LabelFromSampleFile -Path (Join-Path $fixtures 'labeled.xlsx')
Assert-That ($l.Id -eq '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' -and $l.SiteId -eq '0e6c3c7a-1b2d-4c5e-8f90-a1b2c3d4e5f6' -and $l.Name -eq 'Public-Edit') 'Read-LabelFromSampleFile'
Assert-That (-not $l.Encrypts) 'Read-LabelFromSampleFile：不加密的標籤'
$l2 = Read-LabelFromSampleFile -Path (Join-Path $fixtures 'irm_label.xlsx')
Assert-That ($l2.Id -eq '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' -and $l2.Encrypts -and $l2.Name -eq 'Internal - Internal (Contoso)') 'Read-LabelFromSampleFile：IRM 檔案直接讀出標籤（不需 Office）並標示會加密'
$retryCmd = Get-Command Invoke-UnLockFile
Assert-That ($null -ne $retryCmd -and $null -ne (Get-Command Invoke-UnLockFileOnce)) '自動第二次處理的包裝函式存在'

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
Assert-That ($byName['irm_label.xlsx'].LabelBefore -eq 'Internal - Internal (Contoso)') 'IRM 檔案的原標籤取自容器內的標籤資訊' $byName['irm_label.xlsx'].LabelBefore
$csv = Join-Path ([IO.Path]::GetTempPath()) ('unlocked_test_{0}.csv' -f [guid]::NewGuid().ToString('N'))
Export-UnLockReport -Results $res -Path $csv
Assert-That (@(Import-Csv -LiteralPath $csv -Encoding UTF8).Count -eq $fixtureCount) '匯出 CSV 報表'
Remove-Item -LiteralPath $csv -Force

Write-Host '== 檔案鎖定與不需 Office 的處理 ==' -ForegroundColor Cyan
$work = Join-Path ([IO.Path]::GetTempPath()) ('unlocked_test_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null
try {
    $f = Join-Path $work 'locked.xlsx'
    Copy-Item (Join-Path $fixtures 'valid.xlsx') $f
    Assert-That (-not (Test-UnLockedFileLocked $f)) '未開啟的檔案不算鎖定'
    $hold = [IO.File]::Open($f, 'Open', 'Read', 'Read')
    try {
        Assert-That (Test-UnLockedFileLocked $f) '被其他程式開啟的檔案偵測為鎖定'
        $src = Join-Path $work 'new.xlsx'
        Copy-Item (Join-Path $fixtures 'valid.docx') $src
        Assert-That (-not (Move-UnLockFileOver -Source $src -Destination $f -Retries 2)) '目標檔被鎖住時取代失敗並回傳 false'
        Assert-That (Test-Path $src) '取代失敗時保留處理結果'
    }
    finally { $hold.Close() }
    Assert-That ((Get-UnLockAlternatePath $f) -like '*locked_無限制存取.xlsx') '另存檔名' (Get-UnLockAlternatePath $f)

    $out = Join-Path $work 'out'
    $items2 = @(Get-UnLockedFileItems -Paths @((Join-Path $fixtures 'valid.xlsx'), (Join-Path $fixtures 'valid.docx'), (Join-Path $fixtures 'valid.pptx')))
    $res2 = @(Invoke-UnLockBatch -Items $items2 -Options @{ Mode = 'Process'; SetLabel = $false; RemoveIrm = $true; Repair = $true; OutputMode = 'Folder'; OutputFolder = $out } -Sync (New-UnLockSync))
    Assert-That (@($res2 | Where-Object { $_.Result -eq '無需變更' -and $_.Output -and (Test-Path $_.Output) }).Count -eq 3) '已無限制且正常的檔案不開 Office，直接複製到匯出資料夾' (($res2 | ForEach-Object { $_.Result + ':' + $_.Message }) -join ' | ')

    $known = New-UnLockedLabel -Name 'Public\Public-Edit' -Id '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' -SiteId ''
    $res3 = @(Invoke-UnLockBatch -Items @(Get-UnLockedFileItems -Paths @(Join-Path $fixtures 'labeled.xlsx')) -Options @{ Mode = 'Process'; SetLabel = $true; Label = $known; RemoveIrm = $true; OutputMode = 'Folder'; OutputFolder = $out } -Sync (New-UnLockSync))
    Assert-That ($res3[0].Result -eq '無需變更' -and $res3[0].LabelAfter -eq 'Public\Public-Edit') '標籤已是目標標籤時不開 Office' ($res3[0].Result + ':' + $res3[0].Message)

    $irmCopy = Join-Path $work 'irm_copy.xlsx'
    Copy-Item (Join-Path $fixtures 'irm.xlsx') $irmCopy
    $hold = [IO.File]::Open($irmCopy, 'Open', 'Read', 'Read')
    try {
        $res4 = @(Invoke-UnLockBatch -Items @(Get-UnLockedFileItems -Paths @($irmCopy)) -Options @{ Mode = 'Process'; RemoveIrm = $true; OutputMode = 'InPlace'; Backup = $true } -Sync (New-UnLockSync))
        Assert-That ($res4[0].Result -eq '略過' -and $res4[0].Message -like '*正被其他程式使用*') '覆寫模式下原檔被開啟時先略過，不啟動 Office' ($res4[0].Result + ':' + $res4[0].Message)
    }
    finally { $hold.Close() }

    Add-LabelNameToCache -Id '11111111-2222-3333-4444-555555555555' -Name '測試標籤'
    $script:LabelNameCache = $null
    Assert-That ((Resolve-UnLockLabelName '11111111-2222-3333-4444-555555555555' '' @()) -eq '測試標籤') '標籤名稱快取會保存並讀回'
}
finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host ("通過 {0}，失敗 {1}" -f $script:Passed, $script:Failures) -ForegroundColor $(if ($script:Failures) { 'Red' } else { 'Green' })
exit $script:Failures
