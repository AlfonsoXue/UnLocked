#requires -Version 5.1
<#
    UnLocked 核心函式庫
    - 設定檔讀寫、檔案列舉
    - 檔案結構檢查（FileInspector.cs，不需 Office）
    - Office COM 自動化：開啟／修復、敏感度標籤、限制存取 (IRM)、另存
#>

$script:SrcDir = $PSScriptRoot
# 工具根目錄（放共用 config.json 的地方）：以 UnLocked.exe 執行時由啟動程式設定為 exe 所在資料夾
$script:UnLockedRoot = if ($env:UNLOCKED_HOME -and (Test-Path -LiteralPath $env:UNLOCKED_HOME)) { $env:UNLOCKED_HOME } else { Split-Path -Parent $PSScriptRoot }
$script:DummyPassword = 'UnLocked#NoPassword#9f2c'
$script:GuidPattern = '^\{?([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\}?$'

# 支援的副檔名 → 處理用的 Office 程式
$script:ExtensionKinds = @{
    # Excel 活頁簿／範本
    '.xlsx' = 'Excel'; '.xlsm' = 'Excel'; '.xlsb' = 'Excel'; '.xls' = 'Excel'
    '.xltx' = 'Excel'; '.xltm' = 'Excel'; '.xlt' = 'Excel'
    # Word 文件／範本
    '.docx' = 'Word'; '.docm' = 'Word'; '.doc' = 'Word'
    '.dotx' = 'Word'; '.dotm' = 'Word'; '.dot' = 'Word'
    # PowerPoint 簡報／放映檔／範本
    '.pptx' = 'PowerPoint'; '.pptm' = 'PowerPoint'; '.ppt' = 'PowerPoint'
    '.ppsx' = 'PowerPoint'; '.ppsm' = 'PowerPoint'; '.pps' = 'PowerPoint'
    '.potx' = 'PowerPoint'; '.potm' = 'PowerPoint'; '.pot' = 'PowerPoint'
}

$script:KindDisplayNames = @{ Excel = 'Excel'; Word = 'Word'; PowerPoint = 'PowerPoint' }

# 另存格式：XlFileFormat / WdSaveFormat / PpSaveAsFileType
$script:SaveFormats = @{
    '.xlsx' = 51; '.xlsm' = 52; '.xlsb' = 50; '.xls' = 56; '.xltx' = 54; '.xltm' = 53; '.xlt' = 17
    '.docx' = 12; '.docm' = 13; '.doc' = 0; '.dotx' = 14; '.dotm' = 15; '.dot' = 1
    '.pptx' = 24; '.pptm' = 25; '.ppt' = 1; '.ppsx' = 28; '.ppsm' = 29; '.pps' = 7; '.potx' = 26; '.potm' = 27; '.pot' = 5
}

# 只有 Open XML 格式支援敏感度標籤
$script:LegacyExtensions = @('.xls', '.xlt', '.doc', '.dot', '.ppt', '.pps', '.pot')

$script:StatusText = @{
    OK        = '正常'
    CORRUPT   = '損毀'
    PROTECTED = '受 IRM 保護'
    PASSWORD  = '有開啟密碼'
    UNKNOWN   = '無法辨識'
    ERROR     = '無法讀取'
}

# ----------------------------------------------------------------------------
#  初始化
# ----------------------------------------------------------------------------
function Initialize-UnLocked {
    if (-not ('UnLocked.FileInspector' -as [type])) {
        $code = [IO.File]::ReadAllText((Join-Path $script:SrcDir 'FileInspector.cs'), [Text.Encoding]::UTF8)
        # 以實際路徑參照 System.Xml，避免 Windows PowerShell 5.1 以部分名稱解析失敗
        $xmlAsm = [System.Xml.XmlDocument].Assembly.Location
        Add-Type -TypeDefinition $code -Language CSharp -ReferencedAssemblies $xmlAsm -ErrorAction Stop
    }
    if (-not ('UnLocked.Watchdog' -as [type])) {
        $code = [IO.File]::ReadAllText((Join-Path $script:SrcDir 'OfficeInterop.cs'), [Text.Encoding]::UTF8)
        Add-Type -TypeDefinition $code -Language CSharp -ErrorAction Stop
    }
}

# ----------------------------------------------------------------------------
#  設定檔
# ----------------------------------------------------------------------------
function ConvertTo-UnLockedGuid {
    param([string]$Value)
    if ($Value -and $Value.Trim() -match $script:GuidPattern) { return $Matches[1].ToLowerInvariant() }
    return ''
}

function New-UnLockedLabel {
    param([string]$Name, [string]$Id, [string]$SiteId)
    [pscustomobject]@{
        Name   = "$Name".Trim()
        Id     = ConvertTo-UnLockedGuid $Id
        SiteId = ConvertTo-UnLockedGuid $SiteId
    }
}

function Get-UnLockedDefaultConfig {
    [pscustomobject]@{
        Labels           = @()
        SelectedLabelId  = ''
        Justification    = '檔案攜出／分析用途，依公司政策解除權限'
        AssignmentMethod = 1          # MsoAssignmentMethod.PRIVILEGED（等同使用者手動設定）
        SetLabel         = $true
        RemoveIrm        = $true
        Repair           = $true
        IrmInteractive   = $true      # IRM 檔案背景開啟失敗時，顯示 Office 視窗重試
        OutputMode       = 'Folder'   # Folder = 另存到輸出資料夾；InPlace = 覆寫原檔
        OutputFolder     = ''
        Backup           = $true
        BackupFolder     = ''
        Recurse          = $true
        TimeoutSeconds   = 180
        DeepCheck        = $false
    }
}

function Get-UnLockedUserConfigPath {
    $base = $env:APPDATA
    if (-not $base) { $base = [Environment]::GetFolderPath('ApplicationData') }
    Join-Path $base 'UnLocked\config.json'
}

function Read-UnLockedConfig {
    $cfg = Get-UnLockedDefaultConfig
    $candidates = @((Get-UnLockedUserConfigPath), (Join-Path $script:UnLockedRoot 'config.json'))
    foreach ($p in $candidates) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        try {
            $json = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) {
                if ($cfg.PSObject.Properties[$prop.Name]) { $cfg.($prop.Name) = $prop.Value }
            }
            break
        }
        catch {
            Write-Warning "無法讀取設定檔 $p ：$($_.Exception.Message)"
        }
    }
    $cfg.Labels = @(foreach ($l in @($cfg.Labels)) {
            if (-not $l) { continue }
            $n = New-UnLockedLabel -Name $l.Name -Id $l.Id -SiteId $l.SiteId
            if ($n.Id) { $n }      # 略過範例檔中的佔位文字或格式錯誤的 GUID
        })
    $cfg
}

function Save-UnLockedConfig {
    param([Parameter(Mandatory)]$Config)
    $path = Get-UnLockedUserConfigPath
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $json = $Config | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($path, $json, (New-Object Text.UTF8Encoding($true)))
    $path
}

function Add-UnLockedLabelToConfig {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)]$Label)
    $list = @($Config.Labels | Where-Object { $_.Id -ne $Label.Id })
    $Config.Labels = @($list + $Label)
    $Config.SelectedLabelId = $Label.Id
}

# ----------------------------------------------------------------------------
#  檔案列舉與結構檢查
# ----------------------------------------------------------------------------
function Get-UnLockedKind {
    param([string]$Path)
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($script:ExtensionKinds.ContainsKey($ext)) { return $script:ExtensionKinds[$ext] }
    return ''
}

function Get-UnLockedExtensions {
    <# 回傳某個 Office 程式（或全部）支援的副檔名，依字母排序。 #>
    param([string]$Kind = '')
    @($script:ExtensionKinds.Keys | Where-Object { -not $Kind -or $script:ExtensionKinds[$_] -eq $Kind } | Sort-Object)
}

function Get-UnLockedDialogFilter {
    <# 檔案選擇視窗的篩選條件：全部支援的 Office 檔案／Excel／Word／PowerPoint／所有檔案。 #>
    $pattern = { param($exts) (@($exts | ForEach-Object { '*' + $_ }) -join ';') }
    $all = & $pattern (Get-UnLockedExtensions)
    $parts = @("所有支援的 Office 檔案|$all")
    foreach ($k in @(@('Excel', 'Excel 活頁簿'), @('Word', 'Word 文件'), @('PowerPoint', 'PowerPoint 簡報'))) {
        $p = & $pattern (Get-UnLockedExtensions $k[0])
        $parts += ('{0} ({1})|{1}' -f $k[1], $p)
    }
    $parts += '所有檔案 (*.*)|*.*'
    $parts -join '|'
}

function Test-UnLockedSupported {
    param([string]$Path)
    $name = [IO.Path]::GetFileName($Path)
    if ($name.StartsWith('~$')) { return $false }          # Office 鎖定檔
    if ($name -like '*.~unlocked.*') { return $false }     # 本工具的暫存檔
    if ($Path -match '[\\/](_UnLocked_Output|_UnLocked_Backup_[^\\/]*)[\\/]') { return $false }   # 本工具的輸出／備份資料夾
    return [bool](Get-UnLockedKind $name)
}

function Get-UnLockedFileItems {
    <# 回傳 @{ Path; BaseDir }，BaseDir 用來在輸出資料夾中保留子資料夾結構。 #>
    param([string[]]$Paths, [bool]$Recurse = $true, [string[]]$ExcludeFolders = @())
    foreach ($p in $Paths) {
        if (-not $p) { continue }
        if (Test-Path -LiteralPath $p -PathType Container) {
            $base = (Resolve-Path -LiteralPath $p).ProviderPath
            $files = if ($Recurse) {
                Get-ChildItem -LiteralPath $base -File -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                Get-ChildItem -LiteralPath $base -File -Force -ErrorAction SilentlyContinue
            }
            $skip = @($ExcludeFolders | Where-Object { $_ -and -not [string]::Equals([IO.Path]::GetFullPath($_).TrimEnd('\'), $base.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) })
            foreach ($f in $files) {
                $excluded = $false
                foreach ($x in $skip) { if (Test-UnLockedPathUnder $f.FullName $x) { $excluded = $true; break } }
                if ($excluded) { continue }
                if (Test-UnLockedSupported $f.FullName) {
                    [pscustomobject]@{ Path = $f.FullName; BaseDir = $base }
                }
            }
        }
        elseif (Test-Path -LiteralPath $p -PathType Leaf) {
            $full = (Resolve-Path -LiteralPath $p).ProviderPath
            if (Test-UnLockedSupported $full) {
                [pscustomobject]@{ Path = $full; BaseDir = (Split-Path -Parent $full) }
            }
        }
    }
}

function Invoke-UnLockedInspection {
    param([Parameter(Mandatory)][string]$Path)
    [UnLocked.FileInspector]::Inspect($Path)
}

function Format-InspectionSummary {
    param($Inspection)
    if (-not $Inspection) { return '' }
    $text = $script:StatusText[$Inspection.Status]
    if (-not $text) { $text = $Inspection.Status }
    if ($Inspection.Issues.Count -gt 0) { $text += '：' + $Inspection.IssueText }
    elseif ($Inspection.Warnings.Count -gt 0) { $text += '（' + $Inspection.WarningText + '）' }
    elseif ($Inspection.Status -eq 'PROTECTED') {
        $text += Format-ProtectionSource $Inspection
    }
    $text
}

function Format-ProtectionSource {
    <# 例：（標籤「Internal - Internal (EVA)」加密，Azure RMS）或（AD RMS：rms.contoso.com） #>
    param($Inspection)
    $d = @()
    if ($Inspection.LabelId) {
        $n = if ($Inspection.LabelName) { $Inspection.LabelName } else { $Inspection.LabelId }
        $d += ('標籤「{0}」{1}' -f $n, $(if ($Inspection.LabelEncrypts) { '加密' } else { '' }))
        if ($Inspection.RmsKind) { $d += $Inspection.RmsKind }
    }
    elseif ($Inspection.RmsKind) {
        $d += ('{0}：{1}' -f $Inspection.RmsKind, $Inspection.RmsServer)
    }
    if ($d.Count) { return '（' + ($d -join '，') + '）' }
    return ''
}

function Get-UnLockedRelativePath {
    param([string]$Path, [string]$BaseDir)
    $base = $BaseDir.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if ($Path.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) { return $Path.Substring($base.Length) }
    return [IO.Path]::GetFileName($Path)
}

function Test-UnLockedPathUnder {
    param([string]$Path, [string]$Folder)
    if (-not $Folder) { return $false }
    $f = [IO.Path]::GetFullPath($Folder).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    return $Path.StartsWith($f, [StringComparison]::OrdinalIgnoreCase)
}

# ----------------------------------------------------------------------------
#  Office 修復紀錄（Excel 修復時會在暫存資料夾寫出 error*.xml）
# ----------------------------------------------------------------------------
function Get-RepairLogFolders {
    $dirs = @([IO.Path]::GetTempPath(), $env:TEMP, $env:TMP)
    if ($env:USERPROFILE) { $dirs += (Join-Path $env:USERPROFILE '.dsclient-tmp') }   # 公司 DRM 用戶端的暫存資料夾
    if ($env:LOCALAPPDATA) { $dirs += (Join-Path $env:LOCALAPPDATA 'Temp') }
    $seen = @{}
    foreach ($d in $dirs) {
        if (-not $d) { continue }
        try { $full = [IO.Path]::GetFullPath($d).TrimEnd('\') } catch { continue }
        if ($seen.ContainsKey($full.ToLowerInvariant())) { continue }
        $seen[$full.ToLowerInvariant()] = $true
        if (Test-Path -LiteralPath $full -PathType Container) { $full }
    }
}

function Get-RepairLogSnapshot {
    $snap = @{}
    foreach ($d in Get-RepairLogFolders) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d -Filter 'error*.xml' -File -ErrorAction SilentlyContinue)) {
            $snap[$f.FullName] = $f.LastWriteTimeUtc.Ticks
        }
    }
    $snap
}

function Get-NewRepairLogs {
    <# 比對快照，回傳新產生的修復紀錄內容（字串陣列）。 #>
    param([hashtable]$Before)
    foreach ($d in Get-RepairLogFolders) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d -Filter 'error*.xml' -File -ErrorAction SilentlyContinue)) {
            if ($Before.ContainsKey($f.FullName) -and $Before[$f.FullName] -eq $f.LastWriteTimeUtc.Ticks) { continue }
            $lines = @()
            try {
                [xml]$x = [IO.File]::ReadAllText($f.FullName)
                foreach ($n in $x.SelectNodes('//*')) {
                    if ($n.LocalName -in @('summary', 'repairedRecord', 'removedRecord', 'removedPart', 'removedFeature', 'repairedPart')) {
                        $t = $n.InnerText.Trim()
                        if ($t) { $lines += $t }
                    }
                }
            }
            catch { }
            [pscustomobject]@{ File = $f.FullName; Details = ($lines -join ' | ') }
        }
    }
}

# ----------------------------------------------------------------------------
#  Office 應用程式
# ----------------------------------------------------------------------------
function Get-ComHResult {
    param($ErrorRecord)
    $ex = $ErrorRecord
    if ($ex -is [System.Management.Automation.ErrorRecord]) { $ex = $ex.Exception }
    while ($ex) {
        if ($ex -is [System.Runtime.InteropServices.COMException]) { return $ex.ErrorCode }
        if ($ex.HResult -and $ex.HResult -ne -2146233087 -and $ex.HResult -ne -2146233088) { $hr = $ex.HResult }
        $ex = $ex.InnerException
    }
    if ($hr) { return $hr }
    return 0
}

function Test-OfficeDisconnected {
    param($ErrorRecord)
    $hr = Get-ComHResult $ErrorRecord
    # RPC_S_SERVER_UNAVAILABLE / RPC_S_CALL_FAILED / RPC_E_DISCONNECTED / CO_E_OBJNOTCONNECTED
    return ($hr -in @(-2147023174, -2147023170, -2147417848, -2147220995))
}

function Get-FriendlyError {
    param($ErrorRecord)
    $ex = $ErrorRecord
    if ($ex -is [System.Management.Automation.ErrorRecord]) { $ex = $ex.Exception }
    while ($ex.InnerException -and ($ex -is [System.Management.Automation.MethodInvocationException] -or
            $ex -is [System.Management.Automation.SetValueInvocationException] -or
            $ex -is [System.Reflection.TargetInvocationException])) {
        $ex = $ex.InnerException
    }
    $msg = ($ex.Message -replace '\s+', ' ').Trim()
    $hr = Get-ComHResult $ErrorRecord
    if ($hr) { $msg += ' [0x{0:X8}]' -f $hr }
    if (Test-OfficeDisconnected $ErrorRecord) { $msg = 'Office 程式已中斷或被關閉（' + $msg + '）' }
    $msg
}

function New-OfficeApp {
    param([ValidateSet('Excel', 'Word', 'PowerPoint')][string]$Kind)
    $progId = @{ Excel = 'Excel.Application'; Word = 'Word.Application'; PowerPoint = 'PowerPoint.Application' }[$Kind]
    $exe = @{ Excel = 'EXCEL'; Word = 'WINWORD'; PowerPoint = 'POWERPNT' }[$Kind]

    $before = @(Get-Process -Name $exe -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $app = New-Object -ComObject $progId
    $procId = 0
    if ($Kind -eq 'Excel') {
        try { $procId = [UnLocked.NativeMethods]::GetPidFromHwnd([long]$app.Hwnd) } catch { $procId = 0 }
    }
    if (-not $procId) {
        Start-Sleep -Milliseconds 300
        $new = @(Get-Process -Name $exe -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
        if ($new.Count -eq 1) { $procId = $new[0].Id }
    }

    $owned = $true
    switch ($Kind) {
        'Excel' {
            $app.Visible = $false
            $app.DisplayAlerts = $false
            $app.ScreenUpdating = $false
            $app.AskToUpdateLinks = $false
            $app.EnableEvents = $false
            $app.AutomationSecurity = 3      # msoAutomationSecurityForceDisable：不執行巨集
        }
        'Word' {
            $app.Visible = $false
            $app.DisplayAlerts = 0           # wdAlertsNone
            $app.AutomationSecurity = 3
            try { $app.Options.ConfirmConversions = $false } catch { }
            try { $app.Options.UpdateLinksAtOpen = $false } catch { }
        }
        'PowerPoint' {
            $app.DisplayAlerts = 1           # ppAlertsNone
            $app.AutomationSecurity = 3
            # PowerPoint 只有單一執行個體：若使用者已開著 PowerPoint，不可強制關閉它
            if (-not $procId -or $app.Presentations.Count -gt 0) { $owned = $false }
        }
    }
    [pscustomobject]@{ Kind = $Kind; App = $app; Pid = $procId; Owned = $owned }
}

function Get-OfficeApp {
    param([hashtable]$Session, [string]$Kind)
    $entry = $Session[$Kind]
    if ($entry) {
        try { $null = $entry.App.Name; return $entry }   # 健康檢查
        catch { $Session.Remove($Kind) }
    }
    $entry = New-OfficeApp -Kind $Kind
    $Session[$Kind] = $entry
    $entry
}

function Remove-OfficeApp {
    <# 放棄（並在必要時強制結束）一個已失去回應的 Office 執行個體。 #>
    param([hashtable]$Session, [string]$Kind)
    $entry = $Session[$Kind]
    if (-not $entry) { return }
    $Session.Remove($Kind)
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($entry.App) } catch { }
    if ($entry.Owned -and $entry.Pid) {
        Stop-Process -Id $entry.Pid -Force -ErrorAction SilentlyContinue
    }
}

function Close-OfficeSession {
    param([hashtable]$Session)
    foreach ($kind in @($Session.Keys)) {
        $entry = $Session[$kind]
        if ($entry.Owned) {
            try { $entry.App.Quit() } catch { }
        }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($entry.App) } catch { }
        $Session.Remove($kind)
        if ($entry.Owned -and $entry.Pid) {
            # Quit 後若程序仍殘留，等候片刻再強制結束
            $p = Get-Process -Id $entry.Pid -ErrorAction SilentlyContinue
            if ($p) { [void]$p.WaitForExit(5000); if (-not $p.HasExited) { Stop-Process -Id $entry.Pid -Force -ErrorAction SilentlyContinue } }
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ----------------------------------------------------------------------------
#  文件開啟／儲存
# ----------------------------------------------------------------------------
function Open-OfficeDocument {
    <#
        RepairMode：Normal = 一般開啟；Repair = 開啟並修復；Extract = 只擷取資料（僅 Excel）
    #>
    param($AppEntry, [string]$Path, [ValidateSet('Normal', 'Repair', 'Extract')][string]$RepairMode = 'Normal', [bool]$ReadOnly = $false,
        [bool]$UsePassword = $true, [bool]$Interactive = $false)
    $m = [Type]::Missing
    # 假密碼：避免有密碼的檔案跳出看不見的輸入視窗而卡住。
    # 受 IRM 保護的檔案不可傳入密碼（Office 會改用密碼解密而開啟失敗），由呼叫端以 -UsePassword $false 關閉。
    $pw = if ($UsePassword) { $script:DummyPassword } else { $m }
    $app = $AppEntry.App
    switch ($AppEntry.Kind) {
        'Excel' {
            $corrupt = @{ Normal = 0; Repair = 1; Extract = 2 }[$RepairMode]   # XlCorruptLoad
            # Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended,
            #      Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
            return $app.Workbooks.Open($Path, 0, $ReadOnly, $m, $pw, $pw, $true, $m, $m, $false, $false, $m, $false, $m, $corrupt)
        }
        'Word' {
            if ($RepairMode -eq 'Extract') { throw 'Word 不支援僅擷取資料模式' }
            $repair = ($RepairMode -eq 'Repair')
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert,
            #      WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog)
            return $app.Documents.Open($Path, $false, $ReadOnly, $false, $pw, $pw, $false, $pw, $pw, $m, $m, $Interactive, $repair, $m, $true)
        }
        'PowerPoint' {
            if ($RepairMode -eq 'Extract') { throw 'PowerPoint 不支援僅擷取資料模式' }
            $repair = if ($RepairMode -eq 'Repair') { -1 } else { 0 }       # MsoTriState
            $ro = if ($ReadOnly) { -1 } else { 0 }
            $win = if ($Interactive) { -1 } else { 0 }
            # Open2007(FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair)
            return $app.Presentations.Open2007($Path, $ro, 0, $win, $repair)
        }
    }
}

function Set-OfficeInteractive {
    <#
        切換 Office 為「互動模式」（顯示視窗與提示）或「背景模式」。
        受 IRM 保護的檔案第一次開啟時，Office 可能需要登入或選擇權限帳戶，背景模式無法顯示這些視窗。
    #>
    param($AppEntry, [bool]$On)
    $app = $AppEntry.App
    switch ($AppEntry.Kind) {
        'Excel' {
            $app.ScreenUpdating = $On
            $app.Visible = $On
            $app.DisplayAlerts = $On
            if ($On) { try { [UnLocked.NativeMethods]::BringToFront([long]$app.Hwnd) } catch { } }
        }
        'Word' {
            $app.Visible = $On
            $app.DisplayAlerts = if ($On) { -1 } else { 0 }     # wdAlertsAll / wdAlertsNone
            if ($On) { try { $app.Activate() } catch { } }
        }
        'PowerPoint' {
            $app.DisplayAlerts = if ($On) { 2 } else { 1 }      # ppAlertsAll / ppAlertsNone
            if ($On) { try { $app.Activate() } catch { } }
        }
    }
}

function Close-OfficeDocument {
    param($AppEntry, $Doc)
    if (-not $Doc) { return }
    try {
        switch ($AppEntry.Kind) {
            'Excel' { $Doc.Close($false) }
            'Word' { $Doc.Close(0) }          # wdDoNotSaveChanges
            'PowerPoint' { $Doc.Close() }
        }
    }
    catch { }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Doc) } catch { }
}

function Save-OfficeDocumentAs {
    param($AppEntry, $Doc, [string]$Target)
    $ext = [IO.Path]::GetExtension($Target).ToLowerInvariant()
    $fmt = $script:SaveFormats[$ext]
    $m = [Type]::Missing
    switch ($AppEntry.Kind) {
        'Excel' {
            # SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru)
            $Doc.SaveAs($Target, $fmt, $m, $m, $false, $false, 1, 2, $false)
        }
        'Word' {
            $Doc.SaveAs2($Target, $fmt)
        }
        'PowerPoint' {
            $Doc.SaveAs($Target, $fmt)
        }
    }
}

# ----------------------------------------------------------------------------
#  敏感度標籤與 IRM
# ----------------------------------------------------------------------------
function Get-DocumentLabel {
    <# 回傳目前的敏感度標籤；Office 版本不支援時回傳 $null。 #>
    param($Doc)
    try { $sl = $Doc.SensitivityLabel } catch { return $null }
    if (-not $sl) { return $null }
    try { $li = $sl.GetLabel() } catch { return $null }
    if (-not $li) { return $null }
    $id = ConvertTo-UnLockedGuid ([string]$li.LabelId)
    $enabled = $true
    try { $enabled = [bool]$li.IsEnabled } catch { }
    if (-not $enabled) { $id = '' }
    [pscustomobject]@{
        Id     = $id
        Name   = if ($id) { [string]$li.LabelName } else { '' }
        SiteId = ConvertTo-UnLockedGuid ([string]$li.SiteId)
    }
}

function Set-DocumentLabel {
    param($Doc, [Parameter(Mandatory)]$Label, [string]$Justification, [int]$AssignmentMethod = 1)
    $sl = $null
    try { $sl = $Doc.SensitivityLabel } catch { }
    if (-not $sl) { throw '此 Office 版本或檔案格式不支援以程式設定敏感度標籤（需 Microsoft 365 版 Office 與 Open XML 格式）' }
    $li = $sl.CreateLabelInfo()
    $li.AssignmentMethod = $AssignmentMethod
    $li.LabelId = $Label.Id
    $leaf = ($Label.Name -split '\\')[-1]
    if ($leaf) { $li.LabelName = $leaf }
    if ($Label.SiteId) { $li.SiteId = $Label.SiteId }
    if ($Justification) { $li.Justification = $Justification }
    try { $li.IsEnabled = $true } catch { }
    try { $li.SetDate = Get-Date } catch { }
    $sl.SetLabel($li, $li)
}

function Test-ExcelAutoRepaired {
    <#
        Excel 在一般模式下有時會「自動修復」而不報錯，此時視窗標題會多出「[已修復]／[Repaired]」之類的字樣。
        回傳：'Repaired'（確定）、'Maybe'（標題有其他括號標示）、''（無）。
    #>
    param($Workbook)
    try {
        if ($Workbook.ReadOnly) { return '' }
        $caption = [string]$Workbook.Windows.Item(1).Caption
        $name = [string]$Workbook.Name
    }
    catch { return '' }
    $rest = $caption
    if ($caption.StartsWith($name)) { $rest = $caption.Substring($name.Length) }
    if ($rest -match 'Repair|修復|修复') { return 'Repaired' }
    if ($rest -match '\[') { return 'Maybe' }
    return ''
}

function Get-DocumentIrm {
    param($Doc)
    try { return [bool]$Doc.Permission.Enabled } catch { return $null }
}

function Remove-DocumentIrm {
    param($Doc)
    $Doc.Permission.Enabled = $false     # 等同「限制存取 → 無限制存取」
}

# ----------------------------------------------------------------------------
#  取得標籤 ID：範例檔／Office 原則快取
# ----------------------------------------------------------------------------
function Read-LabelFromSampleFile {
    <#
        從「已手動套用目標標籤」的範例檔讀出標籤 ID 與租用戶 ID。
        先直接解析檔案（不需 Office）；若檔案受保護或缺少名稱，再用 Office 讀取。
    #>
    param([Parameter(Mandatory)][string]$Path)
    Initialize-UnLocked
    $ins = Invoke-UnLockedInspection $Path
    $label = $null
    if ($ins.LabelId) { $label = New-UnLockedLabel -Name $ins.LabelName -Id $ins.LabelId -SiteId $ins.LabelSiteId }
    if ($label -and $label.Name -and $label.SiteId) { return $label }

    $kind = Get-UnLockedKind $Path
    $session = @{}
    try {
        $entry = Get-OfficeApp -Session $session -Kind $kind
        $doc = Open-OfficeDocument -AppEntry $entry -Path $Path -ReadOnly $true
        try {
            $l = Get-DocumentLabel $doc
            if ($l -and $l.Id) {
                $name = $l.Name
                if (-not $name -and $label) { $name = $label.Name }
                $site = $l.SiteId
                if (-not $site -and $label) { $site = $label.SiteId }
                $label = New-UnLockedLabel -Name $name -Id $l.Id -SiteId $site
            }
        }
        finally { Close-OfficeDocument -AppEntry $entry -Doc $doc }
    }
    finally { Close-OfficeSession $session }
    $label
}

function Find-OfficeCachedLabels {
    <#
        實驗性：掃描 Office 內建標籤功能的原則快取 (%LOCALAPPDATA%\Microsoft\Office\CLP)，
        找出「含 GUID id 與名稱」的標籤元素。格式未公開，找不到時請改用範例檔。
    #>
    $root = if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'Microsoft\Office\CLP' } else { '' }
    if (-not $root -or -not (Test-Path -LiteralPath $root)) { return }
    $found = @{}
    foreach ($f in @(Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue)) {
        if ($f.Length -gt 20MB) { continue }
        try { [xml]$x = [IO.File]::ReadAllText($f.FullName) } catch { continue }
        foreach ($n in @($x.SelectNodes('//*[@id]'))) {
            $id = ConvertTo-UnLockedGuid $n.GetAttribute('id')
            if (-not $id) { continue }
            $isLabel = ($n.LocalName -match 'label') -or ($n.ParentNode -and $n.ParentNode.LocalName -match 'label')
            if (-not $isLabel) { continue }
            $name = Get-XmlLabelName $n
            if (-not $name) { continue }
            $p = $n.ParentNode
            while ($p -is [System.Xml.XmlElement]) {
                if ((ConvertTo-UnLockedGuid $p.GetAttribute('id')) -and ($pn = Get-XmlLabelName $p)) { $name = "$pn\$name"; break }
                $p = $p.ParentNode
            }
            $site = ''
            foreach ($a in @('siteId', 'tenantId', 'SiteId', 'TenantId')) { if (-not $site) { $site = ConvertTo-UnLockedGuid $n.GetAttribute($a) } }
            if (-not $found.ContainsKey($id)) { $found[$id] = New-UnLockedLabel -Name $name -Id $id -SiteId $site }
        }
    }
    $found.Values | Sort-Object Name
}

function Get-XmlLabelName {
    param([System.Xml.XmlElement]$Node)
    foreach ($a in @('name', 'displayName', 'DisplayName', 'Name')) {
        $v = $Node.GetAttribute($a)
        if ($v) { return $v }
    }
    foreach ($c in @($Node.ChildNodes)) {
        if ($c -is [System.Xml.XmlElement] -and $c.LocalName -in @('name', 'displayName') -and $c.InnerText) { return $c.InnerText.Trim() }
    }
    return ''
}
