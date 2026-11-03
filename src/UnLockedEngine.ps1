#requires -Version 5.1
<#
    UnLocked 批次處理引擎
    每個檔案的流程：
      1. 結構檢查（不需 Office）：ZIP/CRC/XML/OPC 關聯、OLE 結構、IRM/密碼判斷
      2. 以 Office 開啟；若偵測到損毀則使用「開啟並修復」，Excel 最後還可「僅擷取資料」
      3. 設定敏感度標籤（例如 Public\Public-Edit）
      4. 移除限制存取 (IRM)
      5. 以 Office 另存（重新寫出即完成修復），再做一次結構檢查驗證輸出檔
#>

. (Join-Path $PSScriptRoot 'UnLockedCore.ps1')

function New-UnLockSync {
    param([switch]$Console)
    $s = [hashtable]::Synchronized(@{})
    $s.Queue = if ($Console) { $null } else { New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]' }
    $s.Console = [bool]$Console
    $s.Cancel = $false
    $s.Done = $false
    $s.Index = 0
    $s.Total = 0
    $s.LogFile = ''
    $s.RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $s
}

function Send-UnLockEvent {
    param([hashtable]$Sync, [string]$Type, [string]$Message = '', $Data = $null, [ValidateSet('Info', 'Warn', 'Error', 'OK')][string]$Level = 'Info')
    if (-not $Sync) { return }
    $evt = [pscustomobject]@{ Time = (Get-Date); Type = $Type; Level = $Level; Message = $Message; Data = $Data }
    if ($Sync.Queue) { $Sync.Queue.Enqueue($evt) }
    if ($Type -eq 'log') {
        $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f $evt.Time, $Level, $Message
        if ($Sync.LogFile) {
            try { [IO.File]::AppendAllText($Sync.LogFile, $line + "`r`n", [Text.Encoding]::UTF8) } catch { }
        }
        if ($Sync.Console) {
            $color = @{ Info = 'Gray'; Warn = 'Yellow'; Error = 'Red'; OK = 'Green' }[$Level]
            Write-Host $line -ForegroundColor $color
        }
    }
}

function Resolve-UnLockLabelName {
    param([string]$Id, [string]$Name, $Known)
    if (-not $Id) { return '(無標籤)' }
    foreach ($l in @($Known)) { if ($l -and $l.Id -eq $Id -and $l.Name) { return $l.Name } }
    if ($Name) { Add-LabelNameToCache -Id $Id -Name $Name; return $Name }
    $cache = Get-LabelNameCache
    if ($cache.ContainsKey($Id)) { return $cache[$Id] }
    return $Id
}

function Get-UnLockDestination {
    param($Item, [hashtable]$Options)
    if ($Options.OutputMode -eq 'InPlace') { return $Item.Path }
    $root = $Options.OutputFolder
    if (-not $root) { $root = [IO.Path]::Combine($Item.BaseDir, '_UnLocked_Output') }
    [IO.Path]::Combine($root, (Get-UnLockedRelativePath $Item.Path $Item.BaseDir))
}

function Get-UnLockBackupPath {
    param($Item, [hashtable]$Options, [string]$RunStamp)
    $root = $Options.BackupFolder
    if (-not $root) { $root = [IO.Path]::Combine($Item.BaseDir, '_UnLocked_Backup_' + $RunStamp) }
    [IO.Path]::Combine($root, (Get-UnLockedRelativePath $Item.Path $Item.BaseDir))
}

function Initialize-ParentFolder {
    param([string]$Path)
    $dir = [IO.Path]::GetDirectoryName($Path)
    if ($dir -and -not [IO.Directory]::Exists($dir)) { [void][IO.Directory]::CreateDirectory($dir) }
}

function Move-UnLockFileOver {
    <# 以 Source 取代 Destination；目標檔被其他程式鎖住時重試數次，仍失敗則回傳 $false。 #>
    param([string]$Source, [string]$Destination, [int]$Retries = 5)
    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            if ([IO.File]::Exists($Destination)) {
                [IO.File]::SetAttributes($Destination, [IO.FileAttributes]::Normal)
                [IO.File]::Delete($Destination)
            }
            [IO.File]::Move($Source, $Destination)
            return $true
        }
        catch {
            if (-not (Test-UnLockedSharingViolation $_)) { throw }
            Start-Sleep -Milliseconds 800
        }
    }
    return $false
}

function Get-UnLockAlternatePath {
    <# 原檔無法取代時的另存檔名：名稱_無限制存取.xlsx、名稱_無限制存取 (2).xlsx … #>
    param([string]$Path)
    $dir = [IO.Path]::GetDirectoryName($Path)
    $base = [IO.Path]::GetFileNameWithoutExtension($Path) + '_無限制存取'
    $ext = [IO.Path]::GetExtension($Path)
    $candidate = [IO.Path]::Combine($dir, $base + $ext)
    $n = 2
    while ([IO.File]::Exists($candidate)) { $candidate = [IO.Path]::Combine($dir, ('{0} ({1}){2}' -f $base, $n, $ext)); $n++ }
    $candidate
}

function New-UnLockResult {
    param([string]$Path, [string]$Kind)
    [pscustomobject]@{
        File         = $Path
        Kind         = $Kind
        Check        = ''
        LabelBefore  = ''
        LabelAfter   = ''
        IrmBefore    = ''
        IrmAfter     = ''
        Repaired     = $false
        RepairDetail = ''
        Result       = ''
        Message      = ''
        Output       = ''
    }
}

function Invoke-UnLockFileOnce {
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][hashtable]$Options,
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)]$Watchdog,
        [hashtable]$Sync
    )
    $path = $Item.Path
    $kind = Get-UnLockedKind $path
    $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
    $isCheck = ($Options.Mode -eq 'Check')
    $target = $Options.Label
    $r = New-UnLockResult -Path $path -Kind $kind
    $notes = New-Object 'System.Collections.Generic.List[string]'
    $warn = $false

    # ---- 1. 結構檢查 ----------------------------------------------------
    $pre = Invoke-UnLockedInspection $path
    if ($pre.LabelId -and $pre.LabelName) { Add-LabelNameToCache -Id $pre.LabelId -Name $pre.LabelName }
    $r.Check = Format-InspectionSummary $pre
    if ($pre.Container -in @('OOXML', 'CFB')) { $r.IrmBefore = if ($pre.IrmProtected) { '受限制' } else { '無限制' } }
    if ($pre.LabelId) { $r.LabelBefore = Resolve-UnLockLabelName $pre.LabelId $pre.LabelName $Options.KnownLabels }
    elseif ($pre.Container -eq 'OOXML' -and $pre.Status -eq 'OK') { $r.LabelBefore = '(無標籤)' }
    $lvl = if ($pre.IsCorrupt) { 'Warn' } else { 'Info' }
    Send-UnLockEvent $Sync 'log' ('  結構檢查：' + $r.Check) -Level $lvl

    if ($pre.Status -eq 'ERROR') {
        $r.Result = '失敗'
        $r.Message = $pre.IssueText
        return $r
    }
    if ($pre.Status -eq 'PASSWORD') {
        $r.Result = '略過'
        $r.Message = '檔案設有開啟密碼，請先手動移除密碼後再處理'
        return $r
    }
    if ($isCheck -and -not $Options.DeepCheck) {
        $r.Result = if ($pre.IsCorrupt) { '需修復' } elseif ($pre.Status -eq 'UNKNOWN') { '待確認' } else { '正常' }
        return $r
    }

    if (-not $isCheck) {
        $dest0 = Get-UnLockDestination -Item $Item -Options $Options
        $inPlace0 = [string]::Equals([IO.Path]::GetFullPath($dest0), [IO.Path]::GetFullPath($path), [StringComparison]::OrdinalIgnoreCase)

        # 結構正常的 Open XML 檔（ZIP 格式不可能帶 IRM 或開啟密碼）且標籤已符合 → 不需要開 Office
        $labelOk = (-not ($Options.SetLabel -and $target)) -or ($pre.LabelId -and $pre.LabelId -eq $target.Id)
        if ($pre.Container -eq 'OOXML' -and $pre.Status -eq 'OK' -and $labelOk -and -not $Options.ForceResave) {
            $r.IrmAfter = $r.IrmBefore
            if ($Options.SetLabel -and $target) { $r.LabelAfter = $target.Name }
            if ($inPlace0) {
                $r.Message = '檔案已是無限制存取、結構正常，標籤也不需變更'
            }
            else {
                Initialize-ParentFolder $dest0
                Copy-Item -LiteralPath $path -Destination $dest0 -Force
                $r.Output = $dest0
                $r.Message = '檔案已是無限制存取、結構正常，標籤也不需變更，已直接複製到匯出資料夾'
            }
            $r.Result = '無需變更'
            return $r
        }

        # 覆寫原檔前先確認檔案沒有被其他程式（例如自己開著的 Excel）鎖住
        if ($inPlace0 -and (Test-UnLockedFileLocked $path)) {
            $r.Result = '略過'
            $r.Message = '檔案正被其他程式使用（可能已在 Excel 中開啟），請先關閉這個檔案再處理'
            return $r
        }
    }

    # ---- 2. 以 Office 開啟（必要時修復） ---------------------------------
    $entry = $null
    $doc = $null
    $tmp = $null
    $interactive = $false
    try {
        $entry = Get-OfficeApp -Session $Session -Kind $kind
        $Watchdog.Arm($entry.Pid, [int]$Options.TimeoutSeconds)
        $snapshot = Get-RepairLogSnapshot

        # 開啟順序：一般 →（受 IRM 保護時）顯示 Office 視窗重試 → 修復 → 僅擷取資料（Excel）
        # 受 IRM 保護的檔案不傳假密碼，否則 Office 會改用密碼解密而開啟失敗
        $protectedLike = $pre.IrmProtected -or $pre.Status -eq 'UNKNOWN'
        # 假密碼只用在舊版二進位格式（xls/doc/ppt，無法事先判斷是否有密碼）；
        # Open XML 與 IRM 檔案傳入密碼會讓 Office 開啟失敗（0x800A03EC）
        $usePw = ($pre.Container -eq 'CFB') -and -not $pre.IrmProtected
        $canInteract = [bool]$Options.IrmInteractive
        $attempts = New-Object 'System.Collections.Generic.List[object]'
        $repairFirst = (-not $isCheck) -and $Options.Repair -and $pre.IsCorrupt
        if (-not $repairFirst) { $attempts.Add(@{ Mode = 'Normal'; Interactive = $false }) }
        if ($protectedLike -and $canInteract) { $attempts.Add(@{ Mode = 'Normal'; Interactive = $true }) }
        if (-not $isCheck -and $Options.Repair) {
            $attempts.Add(@{ Mode = 'Repair'; Interactive = ($protectedLike -and $canInteract) })
            if ($kind -eq 'Excel') { $attempts.Add(@{ Mode = 'Extract'; Interactive = ($protectedLike -and $canInteract) }) }
        }
        # 一般檔案：背景方式都失敗時，最後顯示 Office 視窗再試一次，讓使用者看到 Office 的實際訊息
        if (-not $protectedLike -and $canInteract) { $attempts.Add(@{ Mode = 'Normal'; Interactive = $true }) }

        $usedMode = $null
        $openError = $null
        foreach ($a in $attempts) {
            if ($a.Interactive -and -not $interactive) {
                Set-OfficeInteractive -AppEntry $entry -On $true
                $interactive = $true
                $Watchdog.Arm($entry.Pid, [Math]::Max([int]$Options.TimeoutSeconds, 600))
                Send-UnLockEvent $Sync 'log' ('  背景開啟失敗，已顯示 {0} 視窗重試：若出現登入、選擇帳戶或權限視窗，請在 Office 視窗中完成操作' -f $kind) -Level Warn
            }
            try {
                $doc = Open-OfficeDocument -AppEntry $entry -Path $path -RepairMode $a.Mode -ReadOnly $isCheck -UsePassword $usePw -Interactive $a.Interactive
                if ($doc) { $usedMode = $a.Mode; break }
            }
            catch {
                if ($Watchdog.Fired -or (Test-OfficeDisconnected $_)) { throw }
                $openError = $_
                $modeName = @{ Normal = '一般'; Repair = '修復'; Extract = '擷取資料' }[$a.Mode]
                if ($a.Interactive) { $modeName += '（顯示視窗）' }
                Send-UnLockEvent $Sync 'log' ("  以{0}模式開啟失敗：{1}" -f $modeName, (Get-FriendlyError $_)) -Level Warn
            }
        }
        if ($interactive) {
            # 已取得授權（或確定無法開啟），切回背景模式繼續處理
            try { Set-OfficeInteractive -AppEntry $entry -On $false } catch { }
            $interactive = $false
            $Watchdog.Arm($entry.Pid, [int]$Options.TimeoutSeconds)
        }
        if (-not $doc) {
            $why = if ($openError) { Get-FriendlyError $openError } else { '未知原因' }
            if ($pre.IrmProtected) {
                $src = Format-ProtectionSource $pre
                $why += ('。受 IRM 保護的檔案{0}無法由程式開啟，可能原因：' -f $src) +
                '(1) 尚未取得此檔案的授權——請先用 Excel 手動開啟一次此檔案（取得授權）並關閉，再用本工具處理；' +
                '(2) 您對此檔案的權限不包含「以程式設計方式存取」，Office 會拒絕程式開啟——請手動開啟後到「檔案 → 資訊 → 檢視權限」確認；' +
                '(3) 您的帳戶沒有開啟權限；(4) 公司的 DRM／資安軟體禁止程式自動開啟受保護的檔案。(2)(3)(4) 需請檔案擁有者或資訊部門協助'
            }
            if ($isCheck) {
                $r.Result = if ($pre.IrmProtected) { '無法開啟' } else { '需修復' }
                $r.Message = 'Office 無法開啟：' + $why
                return $r
            }
            throw ('Office 無法開啟此檔案：' + $why)
        }

        $logs = @(Get-NewRepairLogs -Before $snapshot)
        $r.Repaired = ($usedMode -ne 'Normal') -or ($logs.Count -gt 0)
        $maybeRepaired = $false
        if (-not $r.Repaired -and $kind -eq 'Excel' -and $script:LegacyExtensions -notcontains $ext) {
            $hint = Test-ExcelAutoRepaired $doc
            if ($hint -eq 'Repaired') { $r.Repaired = $true }
            elseif ($hint -eq 'Maybe') { $maybeRepaired = $true }
        }
        if ($logs.Count -gt 0) { $r.RepairDetail = (@($logs | ForEach-Object { $_.Details }) -join ' / ') }
        if ($usedMode -eq 'Extract') { $notes.Add('檔案嚴重損毀，僅擷取出儲存格資料（格式、圖表等可能遺失）'); $warn = $true }
        if ($r.Repaired) {
            $detail = if ($r.RepairDetail) { '：' + $r.RepairDetail } else { '' }
            Send-UnLockEvent $Sync 'log' ('  Office 已修復檔案內容' + $detail) -Level Warn
        }

        # ---- 3. 讀取目前狀態 -------------------------------------------
        $labelBefore = Get-DocumentLabel $doc
        if ($labelBefore -and $labelBefore.Id) {
            $name = if ($labelBefore.Name) { $labelBefore.Name } else { $pre.LabelName }
            $r.LabelBefore = Resolve-UnLockLabelName $labelBefore.Id $name $Options.KnownLabels
        }
        elseif ($labelBefore -and -not $pre.LabelId) { $r.LabelBefore = '(無標籤)' }
        $currentLabelId = if ($labelBefore -and $labelBefore.Id) { $labelBefore.Id } else { $pre.LabelId }
        $irmBefore = Get-DocumentIrm $doc
        if ($null -ne $irmBefore) { $r.IrmBefore = if ($irmBefore) { '受限制' } else { '無限制' } }

        if ($isCheck) {
            $r.Result = if ($r.Repaired -or $pre.IsCorrupt) { '需修復' } elseif ($pre.Status -eq 'UNKNOWN') { '正常（Office 可開啟）' } else { '正常' }
            if ($r.RepairDetail) { $r.Message = 'Office 修復紀錄：' + $r.RepairDetail }
            return $r
        }

        $changed = $false

        # ---- 4. 敏感度標籤 ---------------------------------------------
        if ($Options.SetLabel -and $target) {
            if ($script:LegacyExtensions -contains $ext) {
                $notes.Add("舊版格式 ($ext) 不支援敏感度標籤，已略過標籤設定")
            }
            elseif ($currentLabelId -and $currentLabelId -eq $target.Id) {
                $r.LabelAfter = $target.Name
                Send-UnLockEvent $Sync 'log' ('  標籤已是「' + $target.Name + '」')
            }
            else {
                try {
                    Set-DocumentLabel -Doc $doc -Label $target -Justification $Options.Justification -AssignmentMethod $Options.AssignmentMethod
                    $changed = $true
                    $after = Get-DocumentLabel $doc
                    $r.LabelAfter = if ($after -and $after.Id -eq $target.Id) { $target.Name } else { $target.Name + '（待驗證）' }
                    Send-UnLockEvent $Sync 'log' ('  標籤：' + $r.LabelBefore + ' → ' + $target.Name)
                }
                catch {
                    if ($Watchdog.Fired -or (Test-OfficeDisconnected $_)) { throw }
                    $notes.Add('設定標籤失敗：' + (Get-FriendlyError $_))
                    $warn = $true
                }
            }
        }

        # ---- 5. 限制存取 (IRM) -----------------------------------------
        if ($Options.RemoveIrm) {
            $irmNow = Get-DocumentIrm $doc
            if ($irmNow -and $changed -and $irmBefore -eq $false) {
                # 保護是剛套用的標籤本身帶來的（例如 User Defined 類標籤），不予移除
                $notes.Add('目標標籤本身含有加密保護，保留其權限設定')
                $r.IrmAfter = '受限制（標籤保護）'
            }
            elseif ($irmNow) {
                try {
                    Remove-DocumentIrm $doc
                    $changed = $true
                    Send-UnLockEvent $Sync 'log' '  已移除限制存取 → 無限制存取'
                }
                catch {
                    if ($Watchdog.Fired -or (Test-OfficeDisconnected $_)) { throw }
                    $notes.Add('移除限制存取失敗（需要此檔案的完整控制／擁有者權限）：' + (Get-FriendlyError $_))
                    $warn = $true
                }
                $irmAfter = Get-DocumentIrm $doc
                $r.IrmAfter = if ($irmAfter) { '仍受限制' } else { '無限制' }
            }
            elseif ($irmNow -eq $false) {
                $r.IrmAfter = '無限制'
            }
        }

        # ---- 6. 儲存 ---------------------------------------------------
        $needSave = $changed -or $r.Repaired -or $pre.IsCorrupt -or $Options.ForceResave -or ($maybeRepaired -and $Options.Repair)
        $dest = Get-UnLockDestination -Item $Item -Options $Options
        $inPlace = [string]::Equals([IO.Path]::GetFullPath($dest), [IO.Path]::GetFullPath($path), [StringComparison]::OrdinalIgnoreCase)
        if ($needSave) {
            Initialize-ParentFolder $dest
            $tmp = [IO.Path]::Combine([IO.Path]::GetDirectoryName($dest), [IO.Path]::GetFileNameWithoutExtension($dest) + '.~unlocked' + $ext)
            if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
            Save-OfficeDocumentAs -AppEntry $entry -Doc $doc -Target $tmp
            Close-OfficeDocument -AppEntry $entry -Doc $doc
            $doc = $null
            $Watchdog.Disarm()

            if ($inPlace -and $Options.Backup) {
                $bak = Get-UnLockBackupPath -Item $Item -Options $Options -RunStamp $Sync.RunStamp
                Initialize-ParentFolder $bak
                Copy-Item -LiteralPath $path -Destination $bak -Force
                Send-UnLockEvent $Sync 'log' ('  原檔已備份至 ' + $bak)
            }
            if (Move-UnLockFileOver -Source $tmp -Destination $dest) {
                $r.Output = $dest
            }
            else {
                # 原檔被鎖住：不要丟掉處理結果，改存在原檔旁邊
                $alt = Get-UnLockAlternatePath $dest
                [void](Move-UnLockFileOver -Source $tmp -Destination $alt -Retries 1)
                $r.Output = $alt
                $notes.Add(('原檔正被其他程式使用（可能已在 Excel 中開啟），無法取代；處理後的檔案已另存為「{0}」' -f [IO.Path]::GetFileName($alt)))
                $warn = $true
            }
            $tmp = $null
        }
        else {
            Close-OfficeDocument -AppEntry $entry -Doc $doc
            $doc = $null
            $Watchdog.Disarm()
            if (-not $inPlace) {
                Initialize-ParentFolder $dest
                Copy-Item -LiteralPath $path -Destination $dest -Force
                $r.Output = $dest
                $notes.Add('檔案無需變更，已直接複製到輸出資料夾')
            }
            else {
                $notes.Add('檔案無需變更')
            }
        }

        # ---- 7. 驗證輸出檔 ---------------------------------------------
        if ($r.Output) {
            $post = Invoke-UnLockedInspection $r.Output
            if ($post.Status -eq 'CORRUPT') {
                $notes.Add('輸出檔結構檢查仍有問題：' + $post.IssueText)
                $warn = $true
            }
            elseif ($post.Status -eq 'UNKNOWN') {
                $notes.Add('輸出檔無法直接驗證（可能已由公司 DRM 用戶端加密）')
            }
            else {
                if ($Options.SetLabel -and $target -and $post.Container -eq 'OOXML' -and -not ($script:LegacyExtensions -contains $ext)) {
                    if ($post.LabelId -eq $target.Id) { $r.LabelAfter = $target.Name }
                    elseif ($post.LabelId) {
                        $r.LabelAfter = Resolve-UnLockLabelName $post.LabelId $post.LabelName $Options.KnownLabels
                        $notes.Add('輸出檔的標籤不是目標標籤')
                        $warn = $true
                    }
                }
                if ($Options.RemoveIrm -and $post.IrmProtected -and $r.IrmAfter -notlike '*標籤保護*') {
                    $r.IrmAfter = '仍受限制'
                    if ($post.LabelId -and $post.LabelEncrypts) {
                        $pname = Resolve-UnLockLabelName $post.LabelId $post.LabelName $Options.KnownLabels
                        $r.LabelAfter = $pname
                        if (-not ($Options.SetLabel -and $target)) {
                            $notes.Add(('輸出檔又被加密：存檔時 Office 依公司原則自動套用了會加密的標籤「{0}」（檔案沒有標籤時會套用預設標籤）。' +
                                    '請在「② 處理選項」設定目標標籤（例如 Public\Public-Edit，可用「從範例檔讀取…」取得）後再處理一次') -f $pname)
                        }
                        else {
                            $notes.Add(('輸出檔的標籤「{0}」會加密，請確認目標標籤是否為不加密的標籤（例如 Public\Public-Edit）' -f $pname))
                        }
                    }
                    else {
                        $notes.Add('輸出檔仍受 IRM 保護' + (Format-ProtectionSource $post))
                    }
                    $warn = $true
                }
            }
        }

        $r.Result = if ($warn) { '警告' } elseif ($needSave) { '成功' } else { '無需變更' }
    }
    catch {
        $msg = Get-FriendlyError $_
        if ($Watchdog.Fired) {
            $msg = "處理逾時（超過 $($Options.TimeoutSeconds) 秒），已強制關閉 ${kind}。檔案可能需要手動開啟確認（例如出現登入或權限視窗）"
            Remove-OfficeApp -Session $Session -Kind $kind
            $doc = $null
        }
        elseif (Test-OfficeDisconnected $_) {
            Remove-OfficeApp -Session $Session -Kind $kind
            $doc = $null
        }
        $r.Result = '失敗'
        $notes.Insert(0, $msg)
    }
    finally {
        $Watchdog.Disarm()
        if ($interactive -and $entry) { try { Set-OfficeInteractive -AppEntry $entry -On $false } catch { } }
        if ($doc) { Close-OfficeDocument -AppEntry $entry -Doc $doc }
        if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    if ($r.RepairDetail -and $r.Result -ne '失敗') { $notes.Add('Office 修復紀錄：' + $r.RepairDetail) }
    $r.Message = ($notes -join '；')
    return $r
}

function Invoke-UnLockFile {
    <#
        處理一個檔案。若移除限制存取後存檔，Office 依公司原則自動套用了「會加密的預設標籤」而使輸出檔又被加密，
        自動對輸出檔再處理一次：此時檔案已有標籤，Office 不會再套用預設標籤，移除保護後即可維持「無限制存取」。
    #>
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][hashtable]$Options,
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)]$Watchdog,
        [hashtable]$Sync
    )
    $r = Invoke-UnLockFileOnce -Item $Item -Options $Options -Session $Session -Watchdog $Watchdog -Sync $Sync
    $eligible = ($Options.Mode -ne 'Check') -and $Options.RemoveIrm -and (-not $Options.IsRetry) -and
        $r.Output -and $r.Result -eq '警告' -and $r.IrmAfter -eq '仍受限制'
    if (-not $eligible) { return $r }

    $post = Invoke-UnLockedInspection $r.Output
    if (-not ($post.IrmProtected -and $post.LabelId -and $post.LabelEncrypts)) { return $r }
    $lname = Resolve-UnLockLabelName $post.LabelId $post.LabelName $Options.KnownLabels
    Send-UnLockEvent $Sync 'log' ('  存檔時 Office 自動套用了會加密的標籤「{0}」，自動再處理一次以移除保護…' -f $lname) -Level Warn

    $retryOptions = @{}
    foreach ($k in $Options.Keys) { $retryOptions[$k] = $Options[$k] }
    $retryOptions.OutputMode = 'InPlace'      # 直接處理剛寫出的輸出檔
    $retryOptions.Backup = $false
    $retryOptions.Repair = $false
    $retryOptions.SetLabel = $false           # 保留 Office 剛套用的標籤，只移除保護
    $retryOptions.ForceResave = $false
    $retryOptions.IsRetry = $true
    $retryItem = [pscustomobject]@{ Path = $r.Output; BaseDir = [IO.Path]::GetDirectoryName($r.Output) }
    $r2 = Invoke-UnLockFileOnce -Item $retryItem -Options $retryOptions -Session $Session -Watchdog $Watchdog -Sync $Sync

    $post2 = Invoke-UnLockedInspection $r.Output
    if ($post2.Status -ne 'PROTECTED' -and $post2.Status -ne 'CORRUPT' -and $r2.Result -ne '失敗') {
        $r.Result = '成功'
        $r.IrmAfter = '無限制'
        $r.LabelAfter = $lname
        $r.Message = ('第一次存檔時 Office 依公司原則自動套用了標籤「{0}」並重新加密；已自動再處理一次，保護已移除（標籤保留為「{0}」）' -f $lname)
        Send-UnLockEvent $Sync 'log' '  第二次處理完成：已改為無限制存取' -Level OK
    }
    else {
        $detail = if ($r2.Message) { $r2.Message } else { $r2.Result }
        $r.Message = $r.Message + '；已自動再處理一次仍無法移除保護：' + $detail
    }
    return $r
}

function Export-UnLockReport {
    param([object[]]$Results, [string]$Path)
    Initialize-ParentFolder $Path
    $Results | Select-Object @{ n = '檔案'; e = { $_.File } },
    @{ n = '類型'; e = { $_.Kind } },
    @{ n = '結果'; e = { $_.Result } },
    @{ n = '結構檢查'; e = { $_.Check } },
    @{ n = '原標籤'; e = { $_.LabelBefore } },
    @{ n = '新標籤'; e = { $_.LabelAfter } },
    @{ n = '原限制存取'; e = { $_.IrmBefore } },
    @{ n = '新限制存取'; e = { $_.IrmAfter } },
    @{ n = '已修復'; e = { if ($_.Repaired) { '是' } else { '' } } },
    @{ n = '說明'; e = { $_.Message } },
    @{ n = '輸出檔'; e = { $_.Output } } |
    Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Invoke-UnLockBatch {
    <#
        Options（hashtable）：
          Mode            Process | Check
          SetLabel        $true/$false；Label = New-UnLockedLabel 物件；Justification；AssignmentMethod
          RemoveIrm       $true/$false
          Repair          $true/$false
          DeepCheck       （Check 模式）是否以 Office 實際開啟驗證
          OutputMode      Folder | InPlace；OutputFolder；Backup；BackupFolder
          TimeoutSeconds  單檔逾時秒數
          KnownLabels     已知標籤清單（用來顯示名稱）
          ReportPath      CSV 報表路徑（可省略）
    #>
    param([object[]]$Items, [Parameter(Mandatory)][hashtable]$Options, [hashtable]$Sync)
    Initialize-UnLocked
    if (-not $Sync) { $Sync = New-UnLockSync -Console }
    $Items = @($Items)
    $Sync.Total = $Items.Count
    $Sync.Index = 0
    $Sync.Done = $false
    if (-not $Options.TimeoutSeconds) { $Options.TimeoutSeconds = 180 }
    if (-not $Options.ContainsKey('AssignmentMethod')) { $Options.AssignmentMethod = 1 }
    if (-not $Options.ContainsKey('IrmInteractive')) { $Options.IrmInteractive = $true }

    $session = @{}
    $watchdog = New-Object UnLocked.Watchdog
    $filter = $false
    try { $filter = [UnLocked.MessageFilter]::Register() } catch { $filter = $false }
    $results = New-Object 'System.Collections.Generic.List[object]'
    $modeText = if ($Options.Mode -eq 'Check') { '檢查' } else { '處理' }
    try {
        Send-UnLockEvent $Sync 'log' ("開始{0} {1} 個檔案" -f $modeText, $Items.Count)
        if ($Options.Mode -ne 'Check') {
            if ($Options.SetLabel -and $Options.Label) { Send-UnLockEvent $Sync 'log' ('目標標籤：' + $Options.Label.Name + ' (' + $Options.Label.Id + ')') }
            if ($Options.RemoveIrm -and -not ($Options.SetLabel -and $Options.Label)) {
                Send-UnLockEvent $Sync 'log' '注意：未設定目標標籤。若公司原則會為沒有標籤的檔案自動套用預設標籤（例如會加密的 Internal），存檔後檔案可能又被加密。' -Level Warn
            }
            $outText = if ($Options.OutputMode -eq 'InPlace') { '覆寫原檔' } else { '另存至 ' + $Options.OutputFolder }
            Send-UnLockEvent $Sync 'log' ('輸出方式：' + $outText)
        }
        foreach ($item in $Items) {
            if ($Sync.Cancel) {
                Send-UnLockEvent $Sync 'log' '使用者已停止作業' -Level Warn
                break
            }
            $Sync.Index++
            Send-UnLockEvent $Sync 'progress' -Data ([pscustomobject]@{ Index = $Sync.Index; Path = $item.Path })
            Send-UnLockEvent $Sync 'log' ("[{0}/{1}] {2}" -f $Sync.Index, $Items.Count, $item.Path)
            $res = Invoke-UnLockFile -Item $item -Options $Options -Session $session -Watchdog $watchdog -Sync $Sync
            $results.Add($res)
            Send-UnLockEvent $Sync 'result' -Data $res
            $lvl = switch ($res.Result) { '失敗' { 'Error' } '無法開啟' { 'Error' } '警告' { 'Warn' } '需修復' { 'Warn' } '略過' { 'Warn' } default { 'OK' } }
            $tail = if ($res.Message) { '：' + $res.Message } else { '' }
            Send-UnLockEvent $Sync 'log' ('  → ' + $res.Result + $tail) -Level $lvl
        }
    }
    finally {
        try { Close-OfficeSession $session } catch { }
        $watchdog.Dispose()
        if ($filter) { try { [UnLocked.MessageFilter]::Revoke() } catch { } }
        $groups = @($results | Group-Object Result | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join '、'
        if ($Options.ReportPath -and $results.Count -gt 0) {
            try {
                Export-UnLockReport -Results $results.ToArray() -Path $Options.ReportPath
                Send-UnLockEvent $Sync 'log' ('報表已儲存：' + $Options.ReportPath)
            }
            catch { Send-UnLockEvent $Sync 'log' ('報表儲存失敗：' + $_.Exception.Message) -Level Error }
        }
        Send-UnLockEvent $Sync 'log' ("完成，共 {0} 個檔案。{1}" -f $results.Count, $groups) -Level OK
        $Sync.Done = $true
        Send-UnLockEvent $Sync 'done'
    }
    $results.ToArray()
}
