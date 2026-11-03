#requires -Version 5.1
<#
.SYNOPSIS
    UnLocked － Office 檔案批次解除權限與修復工具

.DESCRIPTION
    批次將 Excel / Word / PowerPoint 檔案：
      1. 設定為指定的敏感度標籤（例如 Public\Public-Edit）
      2. 移除「限制存取」(IRM)，設為「無限制存取」
      3. 偵測檔案是否毀損，並以 Office「開啟並修復」後重新儲存

    不帶參數執行會開啟圖形介面；加上 -NoGui 則以命令列模式執行。

.EXAMPLE
    .\UnLocked.ps1
    開啟圖形介面。

.EXAMPLE
    .\UnLocked.ps1 -NoGui -CheckOnly 'D:\RAW DATA'
    只檢查資料夾內所有檔案的毀損與權限狀態，不做任何修改。

.EXAMPLE
    .\UnLocked.ps1 -NoGui 'D:\RAW DATA' -OutputFolder 'D:\攜出'
    以圖形介面中儲存的標籤設定處理資料夾，結果另存到 D:\攜出。
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,
    [switch]$NoGui,
    [switch]$CheckOnly,
    [switch]$DeepCheck,
    [string]$OutputFolder,
    [switch]$InPlace,
    [switch]$NoBackup,
    [string]$LabelId,
    [string]$LabelName,
    [string]$SiteId,
    [switch]$NoLabel,
    [switch]$NoIrm,
    [switch]$NoRepair,
    [switch]$NoInteractive,
    [string]$Justification,
    [int]$TimeoutSeconds = 0,
    [switch]$NoRecurse,
    [switch]$ForceResave,
    [string]$ReportPath
)

. (Join-Path $PSScriptRoot 'UnLockedEngine.ps1')

function Show-FatalError {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
    if (-not $NoGui) {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [void][System.Windows.Forms.MessageBox]::Show($Message, 'UnLocked', 'OK', 'Error')
        }
        catch { }
    }
}

try {
    Initialize-UnLocked
}
catch {
    Show-FatalError ("UnLocked 初始化失敗：`n" + $_.Exception.Message + "`n`n請確認使用 Windows PowerShell 5.1，且公司政策允許執行 PowerShell 指令碼。")
    exit 2
}

if (-not $NoGui) {
    try {
        . (Join-Path $PSScriptRoot 'UnLockedGui.ps1')
        Show-UnLockedGui -InitialPaths $Path
        exit 0
    }
    catch {
        Show-FatalError ("UnLocked 發生錯誤：`n" + $_.Exception.Message + "`n`n" + $_.ScriptStackTrace)
        exit 2
    }
}

# ---------------------------------------------------------------------------
#  命令列模式
# ---------------------------------------------------------------------------
if (-not $Path) {
    Write-Host '請指定要處理的檔案或資料夾，例如：.\UnLocked.ps1 -NoGui "D:\RAW DATA"' -ForegroundColor Yellow
    exit 1
}

$cfg = Read-UnLockedConfig
$label = $null
if ($LabelId) {
    $site = if ($SiteId) { $SiteId } else { (@($cfg.Labels | Where-Object { $_.SiteId }) | Select-Object -First 1).SiteId }
    $name = if ($LabelName) { $LabelName } else { 'Public\Public-Edit' }
    $label = New-UnLockedLabel -Name $name -Id $LabelId -SiteId $site
    if (-not $label.Id) { Write-Host "-LabelId 不是有效的 GUID：$LabelId" -ForegroundColor Red; exit 1 }
}
else {
    $label = @($cfg.Labels | Where-Object { $_.Id -eq $cfg.SelectedLabelId }) | Select-Object -First 1
    if (-not $label) { $label = @($cfg.Labels) | Select-Object -First 1 }
}

$mode = if ($CheckOnly) { 'Check' } else { 'Process' }
$setLabel = (-not $NoLabel) -and ($null -ne $label)
if ($mode -eq 'Process' -and -not $NoLabel -and -not $label) {
    Write-Warning '尚未設定目標標籤，將略過標籤設定。請以 -LabelId 指定，或先在圖形介面中「從範例檔讀取」。'
}

$outMode = if ($InPlace) { 'InPlace' } else { 'Folder' }
$out = if ($OutputFolder) { $OutputFolder } else { [string]$cfg.OutputFolder }
if ($outMode -eq 'Folder' -and -not $out) { $out = Join-Path ([Environment]::GetFolderPath('Desktop')) 'UnLocked_Output' }
if ($out) { $out = [IO.Path]::GetFullPath($out) }

$opts = @{
    Mode             = $mode
    SetLabel         = $setLabel
    Label            = $label
    Justification    = if ($Justification) { $Justification } else { [string]$cfg.Justification }
    AssignmentMethod = [int]$cfg.AssignmentMethod
    RemoveIrm        = -not $NoIrm
    Repair           = -not $NoRepair
    IrmInteractive   = (-not $NoInteractive) -and [bool]$cfg.IrmInteractive
    DeepCheck        = [bool]$DeepCheck
    OutputMode       = $outMode
    OutputFolder     = $out
    Backup           = -not $NoBackup
    BackupFolder     = [string]$cfg.BackupFolder
    TimeoutSeconds   = if ($TimeoutSeconds -gt 0) { $TimeoutSeconds } else { [int]$cfg.TimeoutSeconds }
    KnownLabels      = @($cfg.Labels)
    ForceResave      = [bool]$ForceResave
}

$items = @(Get-UnLockedFileItems -Paths $Path -Recurse (-not $NoRecurse) -ExcludeFolders @($out))
if ($items.Count -eq 0) {
    Write-Host '找不到支援的 Office 檔案。' -ForegroundColor Yellow
    exit 1
}

$sync = New-UnLockSync -Console
$logDir = if ($mode -eq 'Process' -and $outMode -eq 'Folder') { $out } else { Join-Path (Split-Path -Parent (Get-UnLockedUserConfigPath)) 'Logs' }
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$prefix = if ($mode -eq 'Check') { 'UnLocked_檢查' } else { 'UnLocked_處理' }
if (-not $ReportPath) { $ReportPath = Join-Path $logDir ('{0}_{1}.csv' -f $prefix, $sync.RunStamp) }
$opts.ReportPath = $ReportPath
$sync.LogFile = [IO.Path]::ChangeExtension($ReportPath, '.log')

$results = @(Invoke-UnLockBatch -Items $items -Options $opts -Sync $sync)
$failed = @($results | Where-Object { $_.Result -eq '失敗' }).Count
if ($failed -gt 0) { exit 1 }
exit 0
