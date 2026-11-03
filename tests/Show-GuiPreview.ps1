#requires -Version 5.1
<#
    建立 UnLocked 主視窗、載入 fixtures 範例檔並填入「僅檢查」結果，然後截圖存檔。
    用途：介面冒煙測試（建立視窗不能出錯）＋產生畫面預覽（CI 會上傳成 artifact）。
    用法：powershell -NoProfile -ExecutionPolicy Bypass -File tests\Show-GuiPreview.ps1 -OutFile gui.png
#>
param([string]$OutFile = (Join-Path ([IO.Path]::GetTempPath()) 'unlocked-gui.png'))
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$fixtures = Join-Path $PSScriptRoot 'fixtures'

. (Join-Path $root 'src\UnLockedEngine.ps1')
Initialize-UnLocked
. (Join-Path $root 'src\UnLockedGui.ps1')
[System.Windows.Forms.Application]::EnableVisualStyles()
function Save-CurrentConfig { }   # 預覽用：不要覆寫使用者真正的設定檔

$script:Cfg = Get-UnLockedDefaultConfig
$script:Cfg.OutputFolder = 'D:\資料攜出\UnLocked_Output'
$script:Cfg.Labels = @(New-UnLockedLabel -Name 'Public\Public-Edit' -Id '6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21' -SiteId '0e6c3c7a-1b2d-4c5e-8f90-a1b2c3d4e5f6')
$script:Cfg.SelectedLabelId = $script:Cfg.Labels[0].Id

$form = New-UnLockedForm
Update-LabelCombo
Set-UiFromConfig
Update-CountLabel
$fixedTime = Get-Date '2026-01-15 09:30:00'
Add-UiLog 'UnLocked 已就緒（介面預覽）' 'Info' $fixedTime
Add-PathsToList -Paths @($fixtures)

$items = @(foreach ($lvi in $script:ui.List.Items) { $lvi.Tag })
$sync = New-UnLockSync
$results = @(Invoke-UnLockBatch -Items $items -Options @{ Mode = 'Check' } -Sync $sync)
foreach ($r in $results) { Update-UiRow $r; $script:LastResults.Add($r) }
$evt = $null
while ($sync.Queue.TryDequeue([ref]$evt)) { if ($evt.Type -eq 'log') { Add-UiLog $evt.Message $evt.Level $fixedTime } }
$script:ui.Progress.Maximum = [Math]::Max(1, $items.Count)
$script:ui.Progress.Value = $script:ui.Progress.Maximum
$script:ui.Status.Text = '完成：' + ((@($results | Group-Object Result | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '、')

# ---- 檢查介面狀態 -------------------------------------------------------
$problems = @()
if ($script:ui.List.Columns.Count -ne 8) { $problems += "清單欄位數 $($script:ui.List.Columns.Count)" }
if ($script:ui.List.Items.Count -ne $items.Count -or $items.Count -eq 0) { $problems += "清單項目數 $($script:ui.List.Items.Count)" }
foreach ($kind in @('Excel', 'Word', 'PowerPoint')) {
    if ($script:ui.CountLabel.Text -notmatch $kind) { $problems += "檔案數量標籤缺少 ${kind}：$($script:ui.CountLabel.Text)" }
}
if (-not $script:ui.OptFolder.Checked -or $script:ui.OptInPlace.Checked) { $problems += '預設應為「匯出到資料夾」' }
$script:ui.OptInPlace.Checked = $true
if ($script:ui.OptFolder.Checked -or $script:ui.OutputBox.Enabled) { $problems += '選擇覆寫原檔後，匯出路徑應停用' }
$script:ui.OptFolder.Checked = $true
if ($script:ui.OptInPlace.Checked -or -not $script:ui.OutputBox.Enabled) { $problems += '切回匯出到資料夾後，匯出路徑應啟用' }
if ($script:ui.OutputBox.Text -ne $script:Cfg.OutputFolder) { $problems += '匯出路徑未載入設定值' }

# ---- 截圖 ---------------------------------------------------------------
$form.StartPosition = 'Manual'
$form.Location = New-Object System.Drawing.Point(0, 0)
$form.Show()
for ($i = 0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100 }

# 版面檢查：檔案清單要有足夠高度，選項區不能被撐得過高（曾發生過的問題）
$scale = $script:Scale
$listH = $script:ui.List.Height
$optH = $script:ui.Options.Height
$expH = $script:ui.ExportGroup.Height
Write-Host ("視窗 {0}x{1}；檔案清單高 {2}；處理選項高 {3}；匯出路徑高 {4}" -f $form.Width, $form.Height, $listH, $optH, $expH)
if ($listH -lt 150 * $scale) { $problems += "檔案清單太矮（$listH px）" }
if ($optH -gt 230 * $scale) { $problems += "「② 處理選項」太高（$optH px）" }
if ($expH -gt 150 * $scale) { $problems += "「③ 匯出路徑」太高（$expH px）" }
foreach ($b in @($script:ui.CheckBtn, $script:ui.StartBtn, $script:ui.ExportBtn, $script:ui.BrowseBtn)) {
    $pt = $form.PointToClient($b.PointToScreen((New-Object System.Drawing.Point(0, 0))))
    if ($pt.X + $b.Width -gt $form.ClientSize.Width) { $problems += "按鈕「$($b.Text)」超出視窗寬度" }
}
$bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
$form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
$bmp.Save($OutFile, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
$script:Running = $null
$form.Close()
$form.Dispose()

Write-Host "畫面預覽已儲存：$OutFile"
if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "  [FAIL] $_" -ForegroundColor Red }
    exit 1
}
Write-Host '  [PASS] 介面建立與狀態檢查' -ForegroundColor Green
exit 0
