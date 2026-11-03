#requires -Version 5.1
<#
    UnLocked 圖形介面（Windows Forms）
    處理作業在獨立的 STA 執行空間中進行，介面不會因 Office 處理而停止回應。
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ui = @{}
$script:ItemMap = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
$script:Running = $null
$script:LastResults = New-Object 'System.Collections.Generic.List[object]'
$script:LastLogDir = ''
$script:Scale = 1.0

# 清單欄位位置
$script:Col = @{ Name = 0; Type = 1; Folder = 2; Result = 3; Check = 4; Label = 5; Irm = 6; Message = 7 }

function ScalePx([double]$px) { [int][Math]::Round($px * $script:Scale) }

function New-Ctl {
    param([string]$Type, [hashtable]$Props = @{})
    $c = New-Object ("System.Windows.Forms.$Type")
    # 一般 @{} 的順序不固定；範圍類屬性必須先設定（例如 NumericUpDown 要先設 Maximum 才能設 Value）
    $first = @('Minimum', 'Maximum')
    foreach ($k in $first) { if ($Props.ContainsKey($k)) { $c.$k = $Props[$k] } }
    foreach ($k in $Props.Keys) { if ($first -notcontains $k) { $c.$k = $Props[$k] } }
    $c
}

function New-Btn {
    param([string]$Text, [scriptblock]$OnClick, [int]$Width = 0)
    $b = New-Ctl Button @{ Text = $Text; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; Padding = (New-Object System.Windows.Forms.Padding(6, 2, 6, 2)); Margin = (New-Object System.Windows.Forms.Padding(3)) }
    if ($Width) { $b.AutoSize = $false; $b.Width = (ScalePx $Width); $b.Height = (ScalePx 28) }
    if ($OnClick) { $b.Add_Click($OnClick) }
    $b
}

function New-Flow {
    # 不自動換行：可換行的 FlowLayoutPanel 會以「每個控制項一行」計算偏好高度，使外層群組被撐得過高
    $f = New-Ctl FlowLayoutPanel @{ Dock = 'Fill'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; WrapContents = $false; Margin = (New-Object System.Windows.Forms.Padding(0)) }
    $f
}

function Show-Info([string]$Text, [string]$Icon = 'Information') {
    [void][System.Windows.Forms.MessageBox]::Show($script:ui.Form, $Text, 'UnLocked', 'OK', $Icon)
}

# ----------------------------------------------------------------------------
#  標籤清單
# ----------------------------------------------------------------------------
function Update-LabelCombo {
    $cb = $script:ui.LabelCombo
    $cb.Items.Clear()
    $sel = -1
    $i = 0
    foreach ($l in @($script:Cfg.Labels)) {
        $site = if ($l.SiteId) { '' } else { '（未設定租用戶 ID）' }
        [void]$cb.Items.Add(('{0}    [{1}]{2}' -f $l.Name, $l.Id, $site))
        if ($l.Id -eq $script:Cfg.SelectedLabelId) { $sel = $i }
        $i++
    }
    if ($sel -lt 0 -and $cb.Items.Count -gt 0) {
        # 預設挑選名稱含 Public-Edit 的標籤
        $i = 0
        foreach ($l in @($script:Cfg.Labels)) { if ($l.Name -match 'Public.?Edit') { $sel = $i; break }; $i++ }
        if ($sel -lt 0) { $sel = 0 }
    }
    if ($sel -ge 0) { $cb.SelectedIndex = $sel }
}

function Get-SelectedLabel {
    $idx = $script:ui.LabelCombo.SelectedIndex
    $labels = @($script:Cfg.Labels)
    if ($idx -lt 0 -or $idx -ge $labels.Count) { return $null }
    $labels[$idx]
}

function Show-LabelDialog {
    <# 手動輸入／確認標籤資訊。回傳標籤物件或 $null。 #>
    param($Initial)
    $dlg = New-Ctl Form @{ Text = '敏感度標籤資訊'; FormBorderStyle = 'FixedDialog'; MaximizeBox = $false; MinimizeBox = $false; StartPosition = 'CenterParent'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; Font = $script:ui.Form.Font; Padding = (New-Object System.Windows.Forms.Padding(10)) }
    $t = New-Ctl TableLayoutPanel @{ ColumnCount = 2; AutoSize = $true; Dock = 'Fill' }
    $fields = @(
        @('顯示名稱（例：Public\Public-Edit）', 'Name'),
        @('標籤 ID（LabelId，GUID）', 'Id'),
        @('租用戶 ID（SiteId，GUID）', 'SiteId')
    )
    $boxes = @{}
    foreach ($f in $fields) {
        [void]$t.Controls.Add((New-Ctl Label @{ Text = $f[0]; AutoSize = $true; Anchor = 'Left'; Margin = (New-Object System.Windows.Forms.Padding(3, 8, 3, 3)) }))
        $tb = New-Ctl TextBox @{ Width = (ScalePx 330) }
        if ($Initial) { $tb.Text = [string]$Initial.($f[1]) }
        $boxes[$f[1]] = $tb
        [void]$t.Controls.Add($tb)
    }
    $hint = New-Ctl Label @{ AutoSize = $true; MaximumSize = (New-Object System.Drawing.Size((ScalePx 560), 0)); ForeColor = [System.Drawing.Color]::DimGray
        Text = '提示：最簡單的方式是先手動將一個檔案設為目標標籤，再用「從範例檔讀取」自動取得這些值。' }
    [void]$t.Controls.Add($hint)
    $t.SetColumnSpan($hint, 2)
    $btns = New-Ctl FlowLayoutPanel @{ FlowDirection = 'RightToLeft'; AutoSize = $true; Dock = 'Fill' }
    $ok = New-Btn '確定' $null 90
    $cancel = New-Btn '取消' $null 90
    $ok.DialogResult = 'OK'
    $cancel.DialogResult = 'Cancel'
    [void]$btns.Controls.Add($cancel)
    [void]$btns.Controls.Add($ok)
    [void]$t.Controls.Add($btns)
    $t.SetColumnSpan($btns, 2)
    $dlg.Controls.Add($t)
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $cancel

    while ($dlg.ShowDialog($script:ui.Form) -eq 'OK') {
        $label = New-UnLockedLabel -Name $boxes.Name.Text -Id $boxes.Id.Text -SiteId $boxes.SiteId.Text
        if (-not $label.Name) { Show-Info '請輸入顯示名稱。' 'Warning'; continue }
        if (-not $label.Id) { Show-Info '標籤 ID 必須是 GUID 格式，例如 1a2b3c4d-1234-5678-9abc-def012345678。' 'Warning'; continue }
        if ($boxes.SiteId.Text.Trim() -and -not $label.SiteId) { Show-Info '租用戶 ID 必須是 GUID 格式（可留空）。' 'Warning'; continue }
        $dlg.Dispose()
        return $label
    }
    $dlg.Dispose()
    return $null
}

function Add-LabelAndSelect {
    param($Label)
    if (-not $Label.SiteId) {
        $withSite = @($script:Cfg.Labels | Where-Object { $_.SiteId }) | Select-Object -First 1
        if ($withSite) { $Label.SiteId = $withSite.SiteId }
    }
    Add-UnLockedLabelToConfig -Config $script:Cfg -Label $Label
    Update-LabelCombo
    Save-CurrentConfig
}

function Invoke-ReadSampleLabel {
    $labelExts = @(Get-UnLockedExtensions | Where-Object { $script:LegacyExtensions -notcontains $_ } | ForEach-Object { '*' + $_ }) -join ';'
    $ofd = New-Ctl OpenFileDialog @{ Title = '選擇已套用目標標籤（例如 Public\Public-Edit）的範例檔'; Filter = ('Office 檔案 (xlsx/docx/pptx…)|{0}|所有檔案 (*.*)|*.*' -f $labelExts) }
    if ($ofd.ShowDialog($script:ui.Form) -ne 'OK') { return }
    $script:ui.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $label = Read-LabelFromSampleFile -Path $ofd.FileName
    }
    catch {
        Show-Info ('讀取失敗：' + (Get-FriendlyError $_)) 'Error'
        return
    }
    finally { $script:ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    if (-not $label -or -not $label.Id) {
        Show-Info '這個檔案沒有敏感度標籤。請先在 Office 中手動設定標籤（常用 → 敏感度）並儲存後再試一次。' 'Warning'
        return
    }
    if (-not $label.Name) { $label.Name = 'Public\Public-Edit' }
    $label = Show-LabelDialog -Initial $label     # 讓使用者確認／修改顯示名稱
    if ($label) {
        Add-LabelAndSelect $label
        Add-UiLog ('已新增標籤：' + $label.Name + ' [' + $label.Id + ']')
    }
}

function Invoke-LoadCachedLabels {
    $script:ui.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { $found = @(Find-OfficeCachedLabels) } finally { $script:ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    if ($found.Count -eq 0) {
        Show-Info '在 Office 標籤快取中找不到標籤資訊。請改用「從範例檔讀取」。' 'Warning'
        return
    }
    foreach ($l in $found) { Add-LabelAndSelect $l }
    $pick = @($script:Cfg.Labels | Where-Object { $_.Name -match 'Public.?Edit' }) | Select-Object -First 1
    if ($pick) { $script:Cfg.SelectedLabelId = $pick.Id; Update-LabelCombo }
    Show-Info ("已從 Office 快取載入 {0} 個標籤。`n此功能為實驗性質，若套用失敗請改用「從範例檔讀取」。" -f $found.Count)
}

# ----------------------------------------------------------------------------
#  檔案清單
# ----------------------------------------------------------------------------
function Add-PathsToList {
    param([string[]]$Paths)
    $exclude = @()
    if ($script:ui.OptFolder.Checked -and $script:ui.OutputBox.Text) { $exclude += $script:ui.OutputBox.Text }
    $items = @(Get-UnLockedFileItems -Paths $Paths -Recurse $script:ui.Recurse.Checked -ExcludeFolders $exclude)
    $lv = $script:ui.List
    $lv.BeginUpdate()
    $added = 0
    foreach ($it in $items) {
        if ($script:ItemMap.ContainsKey($it.Path)) { continue }
        $lvi = New-Object System.Windows.Forms.ListViewItem([IO.Path]::GetFileName($it.Path))
        $kind = Get-UnLockedKind $it.Path
        $typeText = '{0} ({1})' -f $kind, [IO.Path]::GetExtension($it.Path).TrimStart('.').ToLowerInvariant()
        foreach ($t in @($typeText, (Split-Path -Parent $it.Path), '', '', '', '', '')) { [void]$lvi.SubItems.Add($t) }
        $lvi.Tag = $it
        [void]$lv.Items.Add($lvi)
        $script:ItemMap[$it.Path] = $lvi
        $added++
    }
    $lv.EndUpdate()
    Update-CountLabel
    if ($items.Count -eq 0 -and $Paths.Count -gt 0) { Add-UiLog ('沒有找到支援的 Office 檔案。支援：' + ((Get-UnLockedExtensions) -join ' ')) 'Warn' }
    elseif ($added -gt 0) { Add-UiLog ("已加入 {0} 個檔案" -f $added) }
}

function Update-CountLabel {
    $n = $script:ui.List.Items.Count
    if ($n -eq 0) {
        $script:ui.CountLabel.Text = '尚未加入檔案：按「選擇檔案…」，或把檔案／資料夾拖曳到下方清單'
        return
    }
    $counts = [ordered]@{ Excel = 0; Word = 0; PowerPoint = 0 }
    foreach ($lvi in $script:ui.List.Items) {
        $k = Get-UnLockedKind $lvi.Tag.Path
        if ($counts.Contains($k)) { $counts[$k]++ }
    }
    $parts = @(foreach ($k in $counts.Keys) { if ($counts[$k]) { '{0} {1}' -f $k, $counts[$k] } })
    $script:ui.CountLabel.Text = '共 {0} 個檔案（{1}）' -f $n, ($parts -join '、')
}

function Remove-SelectedItems {
    $lv = $script:ui.List
    foreach ($lvi in @($lv.SelectedItems)) {
        [void]$script:ItemMap.Remove($lvi.Tag.Path)
        $lv.Items.Remove($lvi)
    }
    Update-CountLabel
}

function Clear-AllItems {
    $script:ui.List.Items.Clear()
    $script:ItemMap.Clear()
    Update-CountLabel
}

function Update-UiRow {
    param($Res)
    if (-not $script:ItemMap.ContainsKey($Res.File)) { return }
    $lvi = $script:ItemMap[$Res.File]
    $label = $Res.LabelBefore
    if ($Res.LabelAfter -and $Res.LabelAfter -ne $Res.LabelBefore) { $label = "$($Res.LabelBefore) → $($Res.LabelAfter)" }
    $irm = $Res.IrmBefore
    if ($Res.IrmAfter -and $Res.IrmAfter -ne $Res.IrmBefore) { $irm = "$($Res.IrmBefore) → $($Res.IrmAfter)" }
    $check = $Res.Check
    if ($Res.Repaired) { $check = '已修復；' + $check }
    $lvi.SubItems[$script:Col.Result].Text = $Res.Result
    $lvi.SubItems[$script:Col.Check].Text = $check
    $lvi.SubItems[$script:Col.Label].Text = $label
    $lvi.SubItems[$script:Col.Irm].Text = $irm
    $lvi.SubItems[$script:Col.Message].Text = $Res.Message
    $color = switch ($Res.Result) {
        '成功' { [System.Drawing.Color]::DarkGreen }
        '正常' { [System.Drawing.Color]::DarkGreen }
        '無需變更' { [System.Drawing.Color]::DarkGreen }
        '警告' { [System.Drawing.Color]::DarkOrange }
        '需修復' { [System.Drawing.Color]::DarkOrange }
        '待確認' { [System.Drawing.Color]::DarkOrange }
        '失敗' { [System.Drawing.Color]::Firebrick }
        '無法開啟' { [System.Drawing.Color]::Firebrick }
        default { [System.Drawing.Color]::DimGray }
    }
    $lvi.ForeColor = $color
    $lvi.EnsureVisible()
}

function Add-UiLog {
    param([string]$Message, [string]$Level = 'Info', [datetime]$Time = (Get-Date))
    $box = $script:ui.LogBox
    if (-not $box) { return }
    if ($box.TextLength -gt 400000) { $box.Text = $box.Text.Substring(200000) }
    $prefix = @{ Info = ''; Warn = '⚠ '; Error = '✖ '; OK = '✔ ' }[$Level]
    $box.AppendText(('{0:HH:mm:ss}  {1}{2}' -f $Time, $prefix, $Message) + "`r`n")
}

# ----------------------------------------------------------------------------
#  設定 <-> 介面
# ----------------------------------------------------------------------------
function Save-CurrentConfig {
    $c = $script:Cfg
    if ($script:ui.ContainsKey('OptLabel')) {
        $sel = Get-SelectedLabel
        if ($sel) { $c.SelectedLabelId = $sel.Id }
        $c.SetLabel = $script:ui.OptLabel.Checked
        $c.Justification = $script:ui.Justification.Text
        $c.RemoveIrm = $script:ui.OptIrm.Checked
        $c.Repair = $script:ui.OptRepair.Checked
        $c.IrmInteractive = $script:ui.OptInteractive.Checked
        $c.OutputMode = if ($script:ui.OptInPlace.Checked) { 'InPlace' } else { 'Folder' }
        $c.OutputFolder = $script:ui.OutputBox.Text.Trim()
        $c.Backup = $script:ui.OptBackup.Checked
        $c.Recurse = $script:ui.Recurse.Checked
        $c.TimeoutSeconds = [int]$script:ui.Timeout.Value
        $c.DeepCheck = $script:ui.OptDeep.Checked
    }
    try { [void](Save-UnLockedConfig -Config $c) } catch { Add-UiLog ('設定儲存失敗：' + $_.Exception.Message) 'Warn' }
}

function Set-UiFromConfig {
    $c = $script:Cfg
    $script:ui.OptLabel.Checked = [bool]$c.SetLabel
    $script:ui.Justification.Text = [string]$c.Justification
    $script:ui.OptIrm.Checked = [bool]$c.RemoveIrm
    $script:ui.OptRepair.Checked = [bool]$c.Repair
    $script:ui.OptInteractive.Checked = [bool]$c.IrmInteractive
    $script:ui.OptInPlace.Checked = ($c.OutputMode -eq 'InPlace')
    $script:ui.OptFolder.Checked = ($c.OutputMode -ne 'InPlace')
    $script:ui.OutputBox.Text = [string]$c.OutputFolder
    $script:ui.OptBackup.Checked = [bool]$c.Backup
    $script:ui.Recurse.Checked = [bool]$c.Recurse
    $t = [int]$c.TimeoutSeconds
    if ($t -lt $script:ui.Timeout.Minimum -or $t -gt $script:ui.Timeout.Maximum) { $t = 180 }
    $script:ui.Timeout.Value = $t
    $script:ui.OptDeep.Checked = [bool]$c.DeepCheck
    Update-OutputControls
}

function Update-OutputControls {
    $folder = $script:ui.OptFolder.Checked
    $script:ui.OutputBox.Enabled = $folder
    $script:ui.BrowseBtn.Enabled = $folder
    $script:ui.OptBackup.Enabled = -not $folder
}

function Set-UiBusy {
    param([bool]$Busy)
    foreach ($k in @('FileBar', 'Options', 'ExportGroup', 'CheckBtn', 'StartBtn', 'ExportBtn')) { $script:ui[$k].Enabled = -not $Busy }
    $script:ui.List.AllowDrop = -not $Busy
    $script:ui.StopBtn.Enabled = $Busy
    if (-not $Busy) { Update-OutputControls }
}

# ----------------------------------------------------------------------------
#  介面動作
# ----------------------------------------------------------------------------
function Invoke-AddFiles {
    $ofd = New-Ctl OpenFileDialog @{ Multiselect = $true; Title = '選擇要處理的 Office 檔案（可按住 Ctrl 或 Shift 一次選多個）'; Filter = (Get-UnLockedDialogFilter) }
    if ($ofd.ShowDialog($script:ui.Form) -eq 'OK') { Add-PathsToList -Paths $ofd.FileNames }
}

function Invoke-AddFolder {
    $fbd = New-Ctl FolderBrowserDialog @{ Description = '選擇要處理的資料夾'; ShowNewFolderButton = $false }
    if ($fbd.ShowDialog($script:ui.Form) -eq 'OK') { Add-PathsToList -Paths @($fbd.SelectedPath) }
}

function Invoke-BrowseOutput {
    $fbd = New-Ctl FolderBrowserDialog @{ Description = '選擇匯出路徑（處理後的檔案會存到這個資料夾）'; ShowNewFolderButton = $true }
    if ($script:ui.OutputBox.Text -and (Test-Path -LiteralPath $script:ui.OutputBox.Text)) { $fbd.SelectedPath = $script:ui.OutputBox.Text }
    if ($fbd.ShowDialog($script:ui.Form) -eq 'OK') { $script:ui.OutputBox.Text = $fbd.SelectedPath }
}

function Invoke-OpenOutput {
    $p = if ($script:ui.OptFolder.Checked) { $script:ui.OutputBox.Text.Trim() } else { $script:LastLogDir }
    if ($p -and (Test-Path -LiteralPath $p)) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $p) }
    else { Show-Info '資料夾尚未建立（請先執行處理）。' }
}

function Invoke-ShowInExplorer {
    $lvi = @($script:ui.List.SelectedItems) | Select-Object -First 1
    if (-not $lvi) { return }
    $res = @($script:LastResults | Where-Object { $_.File -eq $lvi.Tag.Path -and $_.Output }) | Select-Object -Last 1
    $p = if ($res) { $res.Output } else { $lvi.Tag.Path }
    if (Test-Path -LiteralPath $p) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $p) }
}

function Invoke-ExportReport {
    if ($script:LastResults.Count -eq 0) { Show-Info '目前沒有可匯出的結果。'; return }
    $sfd = New-Ctl SaveFileDialog @{ Title = '匯出報表'; Filter = 'CSV 檔案 (*.csv)|*.csv'; FileName = ('UnLocked_報表_{0:yyyyMMdd_HHmmss}.csv' -f (Get-Date)) }
    if ($sfd.ShowDialog($script:ui.Form) -ne 'OK') { return }
    try {
        Export-UnLockReport -Results $script:LastResults.ToArray() -Path $sfd.FileName
        Add-UiLog ('報表已匯出：' + $sfd.FileName) 'OK'
    }
    catch { Show-Info ('匯出失敗：' + $_.Exception.Message) 'Error' }
}

# ----------------------------------------------------------------------------
#  背景執行
# ----------------------------------------------------------------------------
function Start-UnLockedRun {
    param([ValidateSet('Process', 'Check')][string]$Mode)
    if ($script:Running) { return }
    if ($script:ui.List.Items.Count -eq 0) { Show-Info '請先加入要處理的檔案。' 'Warning'; return }
    $label = Get-SelectedLabel
    if ($Mode -eq 'Process') {
        if (-not ($script:ui.OptLabel.Checked -or $script:ui.OptIrm.Checked -or $script:ui.OptRepair.Checked)) {
            Show-Info '請至少勾選一個處理項目。' 'Warning'; return
        }
        if ($script:ui.OptLabel.Checked -and -not $label) {
            Show-Info "尚未設定目標敏感度標籤。`n`n請先在 Office 中把一個檔案手動設為「Public\Public-Edit」並儲存，再按「從範例檔讀取…」選擇該檔案。" 'Warning'; return
        }
        if ($script:ui.OptIrm.Checked -and -not ($script:ui.OptLabel.Checked -and $label)) {
            $msg = "尚未設定目標敏感度標籤。`n`n" +
            "‧ 受敏感度標籤（例如 Internal）加密的檔案，只移除限制存取無法解除加密。`n" +
            "‧ 沒有標籤的檔案存檔時，Office 可能依公司原則自動套用預設標籤並重新加密。`n`n" +
            "建議先在 Office 把一個檔案設為 Public\Public-Edit，再按「從範例檔讀取…」設定目標標籤。`n`n仍要繼續嗎？"
            $ans = [System.Windows.Forms.MessageBox]::Show($script:ui.Form, $msg, 'UnLocked', 'YesNo', 'Warning')
            if ($ans -ne 'Yes') { return }
        }
        if ($script:ui.OptFolder.Checked) {
            $out = $script:ui.OutputBox.Text.Trim()
            if (-not $out) { Show-Info '請先在「③ 匯出路徑」選擇匯出資料夾。' 'Warning'; return }
            try { if (-not (Test-Path -LiteralPath $out)) { New-Item -ItemType Directory -Path $out -Force | Out-Null } }
            catch { Show-Info ('無法建立匯出資料夾：' + $_.Exception.Message) 'Error'; return }
        }
        else {
            $bk = if ($script:ui.OptBackup.Checked) { '（原檔會先備份到各資料夾下的 _UnLocked_Backup_日期時間 資料夾）' } else { '（注意：未勾選備份！）' }
            $ans = [System.Windows.Forms.MessageBox]::Show($script:ui.Form, "將直接覆寫原始檔案$bk`n`n確定要繼續嗎？", 'UnLocked', 'YesNo', 'Warning')
            if ($ans -ne 'Yes') { return }
        }
    }
    Save-CurrentConfig

    $opts = @{
        Mode             = $Mode
        SetLabel         = $script:ui.OptLabel.Checked
        Label            = $label
        Justification    = $script:ui.Justification.Text
        AssignmentMethod = [int]$script:Cfg.AssignmentMethod
        RemoveIrm        = $script:ui.OptIrm.Checked
        Repair           = $script:ui.OptRepair.Checked
        IrmInteractive   = $script:ui.OptInteractive.Checked
        DeepCheck        = $script:ui.OptDeep.Checked
        OutputMode       = if ($script:ui.OptInPlace.Checked) { 'InPlace' } else { 'Folder' }
        OutputFolder     = $script:ui.OutputBox.Text.Trim()
        Backup           = $script:ui.OptBackup.Checked
        BackupFolder     = [string]$script:Cfg.BackupFolder
        TimeoutSeconds   = [int]$script:ui.Timeout.Value
        KnownLabels      = @($script:Cfg.Labels)
        ForceResave      = $false
    }

    $sync = New-UnLockSync
    $logDir = if ($Mode -eq 'Process' -and $opts.OutputMode -eq 'Folder') { $opts.OutputFolder } else { Join-Path (Split-Path -Parent (Get-UnLockedUserConfigPath)) 'Logs' }
    try { if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null } } catch { }
    $prefix = if ($Mode -eq 'Check') { 'UnLocked_檢查' } else { 'UnLocked_處理' }
    $sync.LogFile = Join-Path $logDir ('{0}_{1}.log' -f $prefix, $sync.RunStamp)
    $opts.ReportPath = Join-Path $logDir ('{0}_{1}.csv' -f $prefix, $sync.RunStamp)
    $script:LastLogDir = $logDir
    $script:LastResults.Clear()

    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($lvi in $script:ui.List.Items) {
        $items.Add($lvi.Tag)
        for ($i = $script:Col.Result; $i -le $script:Col.Message; $i++) { $lvi.SubItems[$i].Text = '' }
        $lvi.SubItems[$script:Col.Result].Text = '等待中'
        $lvi.ForeColor = [System.Drawing.SystemColors]::WindowText
    }

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch { }
    $rs = [runspacefactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('sync', $sync)
    $rs.SessionStateProxy.SetVariable('items', $items.ToArray())
    $rs.SessionStateProxy.SetVariable('options', $opts)
    $rs.SessionStateProxy.SetVariable('enginePath', (Join-Path $script:SrcDir 'UnLockedEngine.ps1'))
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript('. $enginePath; Invoke-UnLockBatch -Items $items -Options $options -Sync $sync | Out-Null')
    $handle = $ps.BeginInvoke()

    $script:Running = @{ PS = $ps; RS = $rs; Handle = $handle; Sync = $sync; Mode = $Mode }
    $script:ui.Progress.Maximum = [Math]::Max(1, $items.Count)
    $script:ui.Progress.Value = 0
    $script:ui.Status.Text = '準備中…'
    Set-UiBusy $true
    $script:ui.Timer.Start()
}

function Update-UnLockedRun {
    $run = $script:Running
    if (-not $run) { $script:ui.Timer.Stop(); return }
    $evt = $null
    $n = 0
    while ($n -lt 300 -and $run.Sync.Queue.TryDequeue([ref]$evt)) {
        $n++
        switch ($evt.Type) {
            'log' { Add-UiLog $evt.Message $evt.Level $evt.Time }
            'result' { Update-UiRow $evt.Data; $script:LastResults.Add($evt.Data) }
            'progress' {
                $idx = [int]$evt.Data.Index
                $script:ui.Progress.Value = [Math]::Min($script:ui.Progress.Maximum, [Math]::Max(0, $idx - 1))
                $script:ui.Status.Text = '處理中 {0} / {1}' -f $idx, $run.Sync.Total
                if ($script:ItemMap.ContainsKey($evt.Data.Path)) {
                    $lvi = $script:ItemMap[$evt.Data.Path]
                    $lvi.SubItems[$script:Col.Result].Text = '處理中…'
                    $lvi.EnsureVisible()
                }
            }
        }
    }
    if ($run.Handle.IsCompleted -and $run.Sync.Queue.IsEmpty) { Complete-UnLockedRun }
}

function Complete-UnLockedRun {
    $run = $script:Running
    $script:ui.Timer.Stop()
    try { [void]$run.PS.EndInvoke($run.Handle) } catch { Add-UiLog ('背景作業錯誤：' + (Get-FriendlyError $_)) 'Error' }
    foreach ($e in @($run.PS.Streams.Error)) { Add-UiLog ('背景作業錯誤：' + $e.ToString()) 'Error' }
    if (-not $run.Sync.Done) { Add-UiLog '背景作業未正常結束，請檢查上方的錯誤訊息。' 'Error' }
    try { $run.PS.Dispose(); $run.RS.Dispose() } catch { }
    $script:Running = $null
    foreach ($lvi in $script:ui.List.Items) {
        if ($lvi.SubItems[$script:Col.Result].Text -in @('等待中', '處理中…')) { $lvi.SubItems[$script:Col.Result].Text = '未處理' }
    }
    $script:ui.Progress.Value = $script:ui.Progress.Maximum
    $summary = @($script:LastResults | Group-Object Result | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join '、'
    $script:ui.Status.Text = '完成：' + $summary
    Set-UiBusy $false
    [System.Media.SystemSounds]::Asterisk.Play()
}

function Stop-UnLockedRun {
    if (-not $script:Running) { return }
    $script:Running.Sync.Cancel = $true
    $script:ui.StopBtn.Enabled = $false
    Add-UiLog '已要求停止，將在目前的檔案處理完成後停止…' 'Warn'
}

# ----------------------------------------------------------------------------
#  建立視窗
# ----------------------------------------------------------------------------
function New-UnLockedForm {
    $pad = { param($a, $b, $c, $d) New-Object System.Windows.Forms.Padding((ScalePx $a), (ScalePx $b), (ScalePx $c), (ScalePx $d)) }
    $form = New-Ctl Form @{ Text = 'UnLocked － Office 檔案批次解除權限與修復'; StartPosition = 'CenterScreen'; AutoScaleMode = 'None'
        Font = (New-Object System.Drawing.Font('Microsoft JhengHei UI', 9)) }
    # 預設 1180×860，但不超過螢幕可用範圍（例如 1024×768 的筆電）
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $form.Size = New-Object System.Drawing.Size([Math]::Min((ScalePx 1180), $wa.Width), [Math]::Min((ScalePx 860), $wa.Height))
    $form.MinimumSize = New-Object System.Drawing.Size([Math]::Min((ScalePx 900), $wa.Width), [Math]::Min((ScalePx 640), $wa.Height))
    $script:ui.Form = $form

    $root = New-Ctl TableLayoutPanel @{ Dock = 'Fill'; ColumnCount = 1; RowCount = 5; Padding = (& $pad 8 6 8 6) }
    [void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 75)))   # ① 檔案
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))     # ② 選項
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))     # ③ 匯出路徑
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))     # ④ 執行／進度
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))   # 紀錄
    $form.Controls.Add($root)

    # ---- ① 檔案 ---------------------------------------------------------
    $g1 = New-Ctl GroupBox @{ Text = '① 選擇檔案　支援 Excel（xlsx、xlsm、xlsb、xls…）、Word（docx、docm、doc…）、PowerPoint（pptx、pptm、ppt…）；可直接拖曳檔案或資料夾到清單'; Dock = 'Fill'; Padding = (& $pad 6 4 6 6) }
    $t1 = New-Ctl TableLayoutPanel @{ Dock = 'Fill'; ColumnCount = 1; RowCount = 2 }
    [void]$t1.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$t1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$t1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $bar = New-Flow
    $script:ui.FileBar = $bar
    $script:ui.Recurse = New-Ctl CheckBox @{ Text = '包含子資料夾'; AutoSize = $true; Checked = $true; Margin = (& $pad 8 8 8 3) }
    $script:ui.CountLabel = New-Ctl Label @{ Text = '共 0 個檔案'; AutoSize = $true; Margin = (& $pad 16 9 3 3); ForeColor = [System.Drawing.Color]::DimGray }
    foreach ($c in @(
            (New-Btn '選擇檔案…' { Invoke-AddFiles }),
            (New-Btn '新增資料夾…' { Invoke-AddFolder }),
            $script:ui.Recurse,
            (New-Btn '移除選取' { Remove-SelectedItems }),
            (New-Btn '清除全部' { Clear-AllItems }),
            $script:ui.CountLabel)) { [void]$bar.Controls.Add($c) }
    $lv = New-Ctl ListView @{ Dock = 'Fill'; View = 'Details'; FullRowSelect = $true; GridLines = $true; AllowDrop = $true; HideSelection = $false }
    foreach ($col in @(@('檔案名稱', 230), @('類型', 120), @('所在資料夾', 200), @('結果', 80), @('結構檢查', 220), @('敏感度標籤', 190), @('限制存取', 130), @('說明', 360))) {
        [void]$lv.Columns.Add($col[0], (ScalePx $col[1]))
    }
    $lv.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy } })
    $lv.Add_DragDrop({
            $dropped = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($dropped -and -not $script:Running) { Add-PathsToList -Paths ([string[]]$dropped) }
        })
    $lv.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete -and -not $script:Running) { Remove-SelectedItems } })
    $lv.Add_DoubleClick({ Invoke-ShowInExplorer })
    $script:ui.List = $lv
    [void]$t1.Controls.Add($bar, 0, 0)
    [void]$t1.Controls.Add($lv, 0, 1)
    $g1.Controls.Add($t1)
    [void]$root.Controls.Add($g1, 0, 0)

    # ---- ② 選項 ---------------------------------------------------------
    $g2 = New-Ctl GroupBox @{ Text = '② 處理選項'; Dock = 'Fill'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; Padding = (& $pad 6 4 6 6) }
    $script:ui.Options = $g2
    $t2 = New-Ctl TableLayoutPanel @{ Dock = 'Fill'; ColumnCount = 1; AutoSize = $true; AutoSizeMode = 'GrowAndShrink' }
    [void]$t2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $rowA = New-Flow
    $script:ui.OptLabel = New-Ctl CheckBox @{ Text = '設定敏感度標籤為：'; AutoSize = $true; Margin = (& $pad 3 8 3 3) }
    $script:ui.LabelCombo = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = (ScalePx 430); Margin = (& $pad 3 5 3 3) }
    foreach ($c in @($script:ui.OptLabel, $script:ui.LabelCombo,
            (New-Btn '從範例檔讀取…' { Invoke-ReadSampleLabel }),
            (New-Btn '從 Office 快取載入' { Invoke-LoadCachedLabels }),
            (New-Btn '手動輸入…' { $l = Show-LabelDialog -Initial (Get-SelectedLabel); if ($l) { Add-LabelAndSelect $l } }))) { [void]$rowA.Controls.Add($c) }

    $rowB = New-Flow
    $script:ui.Justification = New-Ctl TextBox @{ Width = (ScalePx 430); Margin = (& $pad 3 5 3 3) }
    foreach ($c in @((New-Ctl Label @{ Text = '　　調降標籤的理由：'; AutoSize = $true; Margin = (& $pad 3 8 3 3) }), $script:ui.Justification,
            (New-Ctl Label @{ Text = '（公司政策要求調降理由時使用）'; AutoSize = $true; ForeColor = [System.Drawing.Color]::DimGray; Margin = (& $pad 3 8 3 3) }))) { [void]$rowB.Controls.Add($c) }

    $rowC = New-Flow
    $script:ui.OptIrm = New-Ctl CheckBox @{ Text = '移除限制存取（IRM）→ 無限制存取'; AutoSize = $true; Margin = (& $pad 3 6 3 3) }
    $script:ui.OptRepair = New-Ctl CheckBox @{ Text = '偵測毀損並自動修復'; AutoSize = $true; Margin = (& $pad 20 6 3 3) }
    $script:ui.OptInteractive = New-Ctl CheckBox @{ Text = 'IRM 檔案無法在背景開啟時，顯示 Office 視窗重試'; AutoSize = $true; Margin = (& $pad 20 6 3 3) }
    foreach ($c in @($script:ui.OptIrm, $script:ui.OptRepair, $script:ui.OptInteractive)) { [void]$rowC.Controls.Add($c) }
    $tips = New-Object System.Windows.Forms.ToolTip
    $tips.AutoPopDelay = 20000
    $tips.SetToolTip($script:ui.OptIrm, '等同 Office「保護 → 限制存取 → 無限制存取」。需要您是檔案擁有者或具完整控制權限。')
    $tips.SetToolTip($script:ui.OptRepair, '先檢查檔案結構；有毀損時以 Office「開啟並修復」開啟並重新儲存（Excel 最後會嘗試只擷取資料）。')
    $tips.SetToolTip($script:ui.OptInteractive, "受 IRM 保護的檔案第一次開啟時，Office 可能需要登入或選擇權限帳戶。`n勾選後，背景開啟失敗時會顯示 Office 視窗讓您完成操作，之後自動切回背景繼續處理。")
    $script:ui.ToolTip = $tips

    $rowF = New-Flow
    $script:ui.Timeout = New-Ctl NumericUpDown @{ Minimum = 30; Maximum = 3600; Increment = 30; Value = 180; Width = (ScalePx 70); Margin = (& $pad 3 5 3 3) }
    $script:ui.OptDeep = New-Ctl CheckBox @{ Text = '「僅檢查」時也用 Office 實際開啟驗證（較慢，可找出只有 Excel 才會發現的內容錯誤）'; AutoSize = $true; Margin = (& $pad 24 8 3 3) }
    foreach ($c in @((New-Ctl Label @{ Text = '單檔逾時：'; AutoSize = $true; Margin = (& $pad 3 8 3 3) }), $script:ui.Timeout,
            (New-Ctl Label @{ Text = '秒'; AutoSize = $true; Margin = (& $pad 0 8 3 3) }), $script:ui.OptDeep)) { [void]$rowF.Controls.Add($c) }

    foreach ($c in @($rowA, $rowB, $rowC, $rowF)) { [void]$t2.Controls.Add($c) }
    $g2.Controls.Add($t2)
    [void]$root.Controls.Add($g2, 0, 1)

    # ---- ③ 匯出路徑 -----------------------------------------------------
    $g3 = New-Ctl GroupBox @{ Text = '③ 匯出路徑（處理完成的檔案以相同檔名存到這裡，原始檔案不會被修改；以資料夾加入的檔案會保留子資料夾結構）'; Dock = 'Fill'; AutoSize = $true; AutoSizeMode = 'GrowAndShrink'; Padding = (& $pad 6 4 6 6) }
    $script:ui.ExportGroup = $g3
    $t3 = New-Ctl TableLayoutPanel @{ Dock = 'Fill'; ColumnCount = 1; AutoSize = $true; AutoSizeMode = 'GrowAndShrink' }
    [void]$t3.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $rowE1 = New-Flow
    $script:ui.OptFolder = New-Ctl RadioButton @{ Text = '匯出到資料夾：'; AutoSize = $true; Margin = (& $pad 3 8 3 3) }
    $script:ui.OutputBox = New-Ctl TextBox @{ Width = (ScalePx 430); Margin = (& $pad 3 5 3 3) }
    $script:ui.BrowseBtn = New-Btn '選擇匯出路徑…' { Invoke-BrowseOutput }
    $script:ui.OpenOutBtn = New-Btn '開啟資料夾' { Invoke-OpenOutput }
    $script:ui.ExportBtn = New-Btn '匯出報表…' { Invoke-ExportReport }
    foreach ($c in @($script:ui.OptFolder, $script:ui.OutputBox, $script:ui.BrowseBtn, $script:ui.OpenOutBtn, $script:ui.ExportBtn)) { [void]$rowE1.Controls.Add($c) }

    $rowE2 = New-Flow
    $script:ui.OptInPlace = New-Ctl RadioButton @{ Text = '不匯出，直接覆寫原始檔案'; AutoSize = $true; Margin = (& $pad 3 6 3 3) }
    $script:ui.OptBackup = New-Ctl CheckBox @{ Text = '覆寫前先備份原檔'; AutoSize = $true; Margin = (& $pad 12 6 3 3) }
    foreach ($c in @($script:ui.OptInPlace, $script:ui.OptBackup)) { [void]$rowE2.Controls.Add($c) }

    # 兩個選項放在不同的容器，需自行維持「只能選一個」
    $script:ui.OptFolder.Add_CheckedChanged({ if ($script:ui.OptFolder.Checked) { $script:ui.OptInPlace.Checked = $false }; Update-OutputControls })
    $script:ui.OptInPlace.Add_CheckedChanged({ if ($script:ui.OptInPlace.Checked) { $script:ui.OptFolder.Checked = $false }; Update-OutputControls })

    foreach ($c in @($rowE1, $rowE2)) { [void]$t3.Controls.Add($c) }
    $g3.Controls.Add($t3)
    [void]$root.Controls.Add($g3, 0, 2)

    # ---- ④ 執行 ---------------------------------------------------------
    $rowG = New-Flow
    $script:ui.CheckBtn = New-Btn '僅檢查' { Start-UnLockedRun -Mode Check } 110
    $script:ui.StartBtn = New-Btn '開始處理' { Start-UnLockedRun -Mode Process } 130
    $script:ui.StartBtn.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
    $script:ui.StopBtn = New-Btn '停止' { Stop-UnLockedRun } 90
    $script:ui.StopBtn.Enabled = $false
    $script:ui.Progress = New-Ctl ProgressBar @{ Width = (ScalePx 240); Height = (ScalePx 20); Margin = (& $pad 12 8 3 3) }
    $script:ui.Status = New-Ctl Label @{ Text = '就緒'; AutoSize = $true; Margin = (& $pad 6 10 3 3) }
    foreach ($c in @((New-Ctl Label @{ Text = '④ 執行：'; AutoSize = $true; Margin = (& $pad 3 10 0 3) }),
            $script:ui.CheckBtn, $script:ui.StartBtn, $script:ui.StopBtn, $script:ui.Progress, $script:ui.Status)) { [void]$rowG.Controls.Add($c) }
    [void]$root.Controls.Add($rowG, 0, 3)

    # ---- 紀錄 -----------------------------------------------------------
    $script:ui.LogBox = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; WordWrap = $true; BackColor = [System.Drawing.SystemColors]::Window }
    [void]$root.Controls.Add($script:ui.LogBox, 0, 4)

    $script:ui.Timer = New-Object System.Windows.Forms.Timer
    $script:ui.Timer.Interval = 250
    $script:ui.Timer.Add_Tick({ Update-UnLockedRun })

    $form.Add_FormClosing({
            if ($script:Running) {
                $ans = [System.Windows.Forms.MessageBox]::Show($script:ui.Form, "仍在處理中，確定要結束嗎？`n（正在處理中的 Office 程式可能需要從工作管理員關閉）", 'UnLocked', 'YesNo', 'Warning')
                if ($ans -ne 'Yes') { $_.Cancel = $true; return }
                $script:Running.Sync.Cancel = $true
            }
            Save-CurrentConfig
        })
    $form
}

function Show-UnLockedGui {
    param([string[]]$InitialPaths)
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [UnLocked.NativeMethods]::EnableDpiAwareness()
    $g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    $script:Scale = [Math]::Max(1.0, $g.DpiX / 96.0)
    $g.Dispose()

    $script:Cfg = Read-UnLockedConfig
    if (-not $script:Cfg.OutputFolder) {
        $script:Cfg.OutputFolder = Join-Path ([Environment]::GetFolderPath('Desktop')) 'UnLocked_Output'
    }
    $form = New-UnLockedForm
    Update-LabelCombo
    Set-UiFromConfig
    Update-CountLabel

    Add-UiLog 'UnLocked 已就緒。使用步驟：'
    Add-UiLog '  ① 按「選擇檔案…」或「新增資料夾…」加入 Excel / Word / PowerPoint 檔案（也可以直接拖曳）。'
    Add-UiLog '  ② 選擇要處理的項目。第一次使用：先在 Office 中把一個檔案設為 Public\Public-Edit 並儲存，再按「從範例檔讀取…」取得標籤。'
    Add-UiLog '  ③ 按「選擇匯出路徑…」指定處理後檔案要存放的資料夾。'
    Add-UiLog '  ④ 可先按「僅檢查」查看毀損與權限狀態，再按「開始處理」。'
    Add-UiLog '  建議處理前先關閉所有 Excel / Word / PowerPoint 視窗。'
    if (@($script:Cfg.Labels).Count -eq 0) { Add-UiLog '尚未設定目標標籤（步驟 ②）。' 'Warn' }
    if ($InitialPaths) { Add-PathsToList -Paths $InitialPaths }

    [UnLocked.NativeMethods]::HideConsole()
    [void]$form.ShowDialog()
    $form.Dispose()
}
