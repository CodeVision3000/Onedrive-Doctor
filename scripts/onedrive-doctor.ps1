<#
.SYNOPSIS
  Detects OneDrive sync trouble and applies safe repairs.
.DESCRIPTION
  Detect-only by default. -Fix applies tier 1-2 repairs (start client, bounce a
  wedged client). Tier 3 (/reset, unlink, renaming files) is never applied here:
  a reset re-syncs every library and must stay a human decision.

  Writes a JSON report and appends a log under %LOCALAPPDATA%\onedrive-doctor.

  Exit codes: 0 healthy  1 findings, nothing fixed  2 repairs applied  3 failed
#>
[CmdletBinding()]
param(
    [switch]$Fix,
    [switch]$Deep,
    [int]$StaleMinutes   = 15,
    [int]$MinFreeGB      = 10,
    [int]$PathWarn       = 380,
    [int]$ScanHours      = 24,
    [int]$ScanSecsPerLib = 20,
    [int]$MaxItems       = 300000,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$StateDir = Join-Path $env:LOCALAPPDATA 'onedrive-doctor'
if (-not (Test-Path $StateDir)) { New-Item -ItemType Directory -Path $StateDir -Force | Out-Null }
$StateFile  = Join-Path $StateDir 'state.json'
$ReportFile = Join-Path $StateDir 'last-report.json'
$LogFile    = Join-Path $StateDir 'doctor.log'
# Deep findings are expensive (minutes across large libraries), so they run on
# their own schedule and the fast pass re-reports the cached result.
$DeepFile   = Join-Path $StateDir 'deep-cache.json'

$Findings = [System.Collections.Generic.List[object]]::new()
$Actions  = [System.Collections.Generic.List[string]]::new()

function Add-Finding {
    param([string]$Id, [ValidateSet('info','warn','error')][string]$Severity,
          [int]$Tier, [string]$Message, $Detail = $null)
    $Findings.Add([pscustomobject]@{
        id = $Id; severity = $Severity; tier = $Tier; message = $Message; detail = $Detail
    })
}

function Write-Line { param([string]$Text) if (-not $Quiet) { Write-Host $Text } }

function Get-OneDriveExe {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')
    )
    foreach ($p in $candidates) { if ($p -and (Test-Path $p)) { return $p } }
    return $null
}

function Get-Mounts {
    $key = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
    if (-not (Test-Path $key)) { return @() }
    Get-ChildItem $key |
        ForEach-Object { Get-ItemProperty $_.PSPath } |
        Where-Object { $_.MountPoint } |
        Group-Object MountPoint |
        ForEach-Object {
            [pscustomobject]@{
                MountPoint = $_.Name
                Type       = ($_.Group | Select-Object -First 1).LibraryType
                Exists     = Test-Path -LiteralPath $_.Name
            }
        }
}

# The availability-status column index moves between Windows builds, so find it
# by name instead of hard-coding one.
# There are several columns literally named "Status" and they come before the
# one we want (311 on this build), so take an exact "Availability status" match
# across the whole range first and only then fall back.
function Get-AvailabilityColumn {
    param($Folder)
    $fallback = -1
    for ($i = 0; $i -lt 400; $i++) {
        $name = $Folder.GetDetailsOf($null, $i)
        if ($name -eq 'Availability status') { return $i }
        if ($name -eq 'Status' -and $fallback -lt 0) { $fallback = $i }
    }
    return $fallback
}

function Get-MountStatuses {
    param($Mounts)
    $shell = New-Object -ComObject Shell.Application
    # Column indices are per folder view, not global, so cache one per parent.
    $colCache = @{}
    $out = @()
    foreach ($m in $Mounts) {
        if (-not $m.Exists) {
            $out += [pscustomobject]@{ path = $m.MountPoint; type = $m.Type; status = 'MISSING' }
            continue
        }
        $parent = Split-Path -Parent $m.MountPoint
        $leaf   = Split-Path -Leaf   $m.MountPoint
        $ns     = $shell.NameSpace($parent)
        if (-not $ns) {
            $out += [pscustomobject]@{ path = $m.MountPoint; type = $m.Type; status = 'UNKNOWN' }
            continue
        }
        if (-not $colCache.ContainsKey($parent)) {
            $colCache[$parent] = Get-AvailabilityColumn -Folder $ns
        }
        $col  = $colCache[$parent]
        $item = $ns.ParseName($leaf)
        $st = ''
        if ($item -and $col -ge 0) { $st = $ns.GetDetailsOf($item, $col) }
        if ([string]::IsNullOrWhiteSpace($st)) { $st = 'UNKNOWN' }
        $out += [pscustomobject]@{ path = $m.MountPoint; type = $m.Type; status = $st }
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    return $out
}

function Get-HeartbeatAgeMinutes {
    $logRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\logs'
    if (-not (Test-Path $logRoot)) { return $null }
    # Only the sync engine's own logs count. Common/ListSync/OD4/Personal keep
    # writing while a wedged SyncEngine stays silent, which would mask the wedge.
    $newest = Get-ChildItem $logRoot -Recurse -File -Filter 'SyncEngine-*.odl' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $newest) { return $null }
    return [math]::Round(((Get-Date) - $newest.LastWriteTime).TotalMinutes, 1)
}

# OneDrive refuses these leaf names outright; over-long paths stall silently.
function Get-ProblemFiles {
    param($Mounts, [int]$Hours, [int]$SecsPerLib, [int]$MaxLen)
    $cutoff   = (Get-Date).AddHours(-$Hours)
    $badChars = '[<>:"|?*]'
    $conflict = '-' + $env:COMPUTERNAME
    $hits = [System.Collections.Generic.List[object]]::new()
    $truncated = @()

    foreach ($m in ($Mounts | Where-Object { $_.Exists })) {
        $sw  = [System.Diagnostics.Stopwatch]::StartNew()
        $cut = $false
        foreach ($f in (Get-ChildItem -LiteralPath $m.MountPoint -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            if ($sw.Elapsed.TotalSeconds -gt $SecsPerLib) { $cut = $true; break }
            if ($f.LastWriteTime -lt $cutoff) { continue }
            $reason = $null
            if     ($f.Name -match $badChars)        { $reason = 'invalid-character' }
            elseif ($f.Name -match '[ .]$')          { $reason = 'trailing-space-or-dot' }
            elseif ($f.BaseName -like "*$conflict*") { $reason = 'conflict-copy' }
            elseif ($f.FullName.Length -gt $MaxLen)  { $reason = 'path-too-long' }
            if ($reason) {
                $hits.Add([pscustomobject]@{ reason = $reason; path = $f.FullName; length = $f.FullName.Length })
            }
        }
        if ($cut) { $truncated += $m.MountPoint }
    }
    return [pscustomobject]@{ hits = $hits; truncated = $truncated }
}

# Outlook holds .pst/.ost open for as long as it runs, so a data file inside a
# sync root can never finish uploading. Microsoft does not support this layout.
# Detect the configuration rather than the symptom: the client surfaces "this
# file is in use" only in its flyout, which is not readable from here.
function Find-OutlookDataFiles {
    param($Mounts, [switch]$Full, [int]$SecsPerLib = 20)
    $hits = [System.Collections.Generic.List[object]]::new()

    # Fast path: the handful of places Outlook actually puts data files.
    $quick = foreach ($m in ($Mounts | Where-Object { $_.Exists })) {
        Join-Path $m.MountPoint 'Documents\Outlook Files'
        Join-Path $m.MountPoint 'Outlook Files'
        $m.MountPoint
    }
    $roots = if ($Full) { ($Mounts | Where-Object { $_.Exists }).MountPoint } else { $quick }

    foreach ($r in $roots) {
        if (-not (Test-Path -LiteralPath $r)) { continue }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $args = @{ LiteralPath = $r; Force = $true; File = $true; ErrorAction = 'SilentlyContinue' }
        if ($Full) { $args['Recurse'] = $true }
        foreach ($ext in '*.pst','*.ost') {
            foreach ($f in (Get-ChildItem @args -Filter $ext)) {
                if ($Full -and $sw.Elapsed.TotalSeconds -gt $SecsPerLib) { break }
                $offline = ($f.Attributes -band [IO.FileAttributes]::Offline) -ne 0
                $locked  = $true
                try { $s = [IO.File]::Open($f.FullName,'Open','ReadWrite','None'); $s.Close(); $locked = $false } catch { }
                $hits.Add([pscustomobject]@{
                    path = $f.FullName; mb = [math]::Round($f.Length/1MB,1)
                    cloudOnly = $offline; locked = $locked
                })
            }
        }
    }
    # Same file can appear via two roots; key on path.
    return @($hits | Sort-Object path -Unique)
}

# Item count drives everything slow on a big machine: scan time, database size,
# and how long a restart costs. Counting is expensive, hence deep-only.
function Measure-SyncedItems {
    param($Mounts)
    $per = @()
    foreach ($m in ($Mounts | Where-Object { $_.Exists })) {
        $n = (Get-ChildItem -LiteralPath $m.MountPoint -Recurse -Force -ErrorAction SilentlyContinue |
              Measure-Object).Count
        $per += [pscustomobject]@{ path = $m.MountPoint; items = $n }
    }
    return [pscustomobject]@{
        total = ($per | Measure-Object items -Sum).Sum
        perLibrary = @($per | Sort-Object items -Descending)
    }
}

function Read-DeepCache {
    if (Test-Path $DeepFile) {
        try { return Get-Content $DeepFile -Raw | ConvertFrom-Json } catch { }
    }
    return $null
}

function Read-State {
    if (Test-Path $StateFile) {
        try { return Get-Content $StateFile -Raw | ConvertFrom-Json } catch { }
    }
    return [pscustomobject]@{ staleStreak = 0; lastFix = $null }
}

function Write-State {
    param($State)
    $State | ConvertTo-Json -Depth 4 | Set-Content $StateFile -Encoding UTF8
}

# ------------------------------------------------------------------ run

try {
    $exe   = Get-OneDriveExe
    $proc  = @(Get-Process OneDrive -ErrorAction SilentlyContinue)
    $state = Read-State

    if (-not $exe) {
        Add-Finding 'client-missing' 'error' 3 'OneDrive.exe not found - the client is not installed.'
    }

    # --- tier 1: client not running
    if ($proc.Count -eq 0) {
        Add-Finding 'not-running' 'error' 1 'OneDrive client is not running.'
        if ($Fix -and $exe) {
            Start-Process $exe -ArgumentList '/background'
            $Actions.Add('started OneDrive /background')
            Start-Sleep -Seconds 5
            $proc = @(Get-Process OneDrive -ErrorAction SilentlyContinue)
        }
    }

    # --- tier 2: engine wedged (log heartbeat stale while the process is alive)
    $hb = Get-HeartbeatAgeMinutes
    if ($null -ne $hb -and $hb -gt $StaleMinutes -and $proc.Count -gt 0) {
        $state.staleStreak = [int]$state.staleStreak + 1
        Add-Finding 'engine-stale' 'error' 2 `
            "Sync engine wrote no log for $hb min (streak $($state.staleStreak)). Client is up but wedged." `
            @{ heartbeatMinutes = $hb; streak = $state.staleStreak }
        # A restart throws away scan progress. On a machine with a large item
        # count that costs hours, so an automatic bounce does more harm than the
        # wedge it is guessing at. Report and let a human decide instead.
        $deepNow = Read-DeepCache
        $tooBigToBounce = ($deepNow -and $deepNow.items.total -gt $MaxItems)
        if ($tooBigToBounce) {
            Add-Finding 'bounce-suppressed' 'warn' 3 (
                "Not restarting automatically: {0:N0} synced items means a restart forces a full re-scan costing hours. Restart only deliberately." -f $deepNow.items.total
            ) @{ items = $deepNow.items.total; threshold = $MaxItems }
        }
        # Hysteresis: require two consecutive stale checks, so one long upload
        # does not get a restart dropped on it.
        if ($Fix -and $exe -and -not $tooBigToBounce -and $state.staleStreak -ge 2) {
            Start-Process $exe -ArgumentList '/shutdown' -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 10
            Get-Process OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Start-Process $exe -ArgumentList '/background'
            $Actions.Add('bounced OneDrive (/shutdown then /background)')
            $state.staleStreak = 0
            $state.lastFix = (Get-Date).ToString('o')
        }
    } else {
        $state.staleStreak = 0
    }

    # --- per-library status
    $mounts   = @(Get-Mounts)
    $statuses = @()
    if ($mounts.Count -eq 0) {
        Add-Finding 'no-libraries' 'warn' 3 'No synced libraries are registered for this user.'
    } else {
        $statuses = Get-MountStatuses -Mounts $mounts
        foreach ($s in $statuses) {
            if ($s.status -eq 'MISSING') {
                Add-Finding 'mount-missing' 'error' 3 "Synced library folder is gone: $($s.path)" $s
            } elseif ($s.status -match 'rror') {
                Add-Finding 'library-error' 'error' 3 "Library reports an error: $($s.path) [$($s.status)]" $s
            }
        }
    }

    # --- disk
    $free = [math]::Round((Get-PSDrive C).Free / 1GB, 1)
    if ($free -lt $MinFreeGB) {
        Add-Finding 'low-disk' 'error' 3 "Only $free GB free on C: - sync stalls as the disk fills." @{ freeGB = $free }
    }

    # --- Outlook data files: always checked, cheaply, because one blocks sync
    #     permanently and silently. Full recursive sweep only on a deep run.
    $pst = Find-OutlookDataFiles -Mounts $mounts -Full:$Deep -SecsPerLib $ScanSecsPerLib

    # --- deep-only work, cached for the fast pass to re-report
    $deepCache = Read-DeepCache
    if ($Deep) {
        $pf = Get-ProblemFiles -Mounts $mounts -Hours $ScanHours -SecsPerLib $ScanSecsPerLib -MaxLen $PathWarn
        $counts = Measure-SyncedItems -Mounts $mounts
        $deepCache = [pscustomobject]@{
            timestamp = (Get-Date).ToString('o')
            items     = $counts
            files     = @($pf.hits)
            truncated = @($pf.truncated)
            outlook   = @($pst)
        }
        $deepCache | ConvertTo-Json -Depth 6 | Set-Content $DeepFile -Encoding UTF8
    } elseif ($deepCache -and $deepCache.outlook) {
        # The quick sweep only knows Outlook's default locations. Carry the deep
        # sweep's finds forward, re-checking lock state live since that changes.
        $known = @($pst | ForEach-Object { $_.path })
        foreach ($c in @($deepCache.outlook)) {
            if ($known -contains $c.path) { continue }
            if (-not (Test-Path -LiteralPath $c.path)) { continue }
            $fi = Get-Item -LiteralPath $c.path -Force
            $off = ($fi.Attributes -band [IO.FileAttributes]::Offline) -ne 0
            $lk = $true
            try { $s = [IO.File]::Open($c.path,'Open','ReadWrite','None'); $s.Close(); $lk = $false } catch { }
            $pst = @($pst) + [pscustomobject]@{
                path = $c.path; mb = [math]::Round($fi.Length/1MB,1); cloudOnly = $off; locked = $lk
            }
        }
    }

    $blocking = @($pst | Where-Object { -not $_.cloudOnly })
    if ($blocking.Count -gt 0) {
        Add-Finding 'outlook-data-file' 'error' 3 (
            "$($blocking.Count) Outlook data file(s) inside a synced folder. Outlook holds these open, so they can never sync. Move them outside OneDrive."
        ) $blocking
    }
    $dormantPst = @($pst | Where-Object { $_.cloudOnly })
    if ($dormantPst.Count -gt 0) {
        Add-Finding 'outlook-data-file-cloud' 'info' 0 (
            "$($dormantPst.Count) Outlook data file(s) present but cloud-only. They sync fine and block nothing - moving them would force a download first."
        ) $dormantPst
    }

    if ($deepCache) {
        $age = [math]::Round(((Get-Date) - [datetime]$deepCache.timestamp).TotalHours, 1)
        if ($deepCache.items.total -gt $MaxItems) {
            Add-Finding 'too-many-items' 'warn' 3 (
                "{0:N0} synced items, over the {1:N0} Microsoft recommends. This is what makes sync slow and restarts expensive. Reduce scope with Choose folders." -f $deepCache.items.total, $MaxItems
            ) @{ total = $deepCache.items.total; perLibrary = $deepCache.items.perLibrary }
        }
        foreach ($g in (@($deepCache.files) | Group-Object reason)) {
            if (-not $g.Name) { continue }
            Add-Finding "files-$($g.Name)" 'warn' 3 `
                "$($g.Count) file(s) with $($g.Name) - OneDrive will not sync these." `
                ($g.Group | Select-Object -First 25)
        }
        if ($age -gt 48) {
            Add-Finding 'deep-scan-stale' 'info' 0 "Deep scan results are $age h old; run with -Deep to refresh."
        }
    } else {
        Add-Finding 'no-deep-scan' 'info' 0 'No deep scan has run yet - item count and filename checks are unknown. Run with -Deep.'
    }

    # --- report-only context
    $xl = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if ($xl.Count -gt 0) {
        Add-Finding 'excel-open' 'info' 0 `
            "$($xl.Count) Excel process(es) open - open workbooks hold locks and will not upload." @{ pids = $xl.Id }
    }

    foreach ($acct in (Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue)) {
        $v = Get-ItemProperty $acct.PSPath
        if ($v.ResetRequested -eq 1) {
            # No UserFolder means the account syncs nothing, so it is a dormant
            # leftover, not a live problem. Keeping it a warning would make every
            # scheduled run exit non-zero forever.
            $sev = if ([string]::IsNullOrWhiteSpace($v.UserFolder)) { 'info' } else { 'warn' }
            Add-Finding 'reset-requested' $sev 3 `
                "Account $($acct.PSChildName) has ResetRequested=1 - the client wants a reset. Human decision." `
                @{ account = $acct.PSChildName; userFolder = $v.UserFolder }
        }
    }

    Write-State $state

    $errors = @($Findings | Where-Object { $_.severity -eq 'error' })
    $warns  = @($Findings | Where-Object { $_.severity -eq 'warn' })
    $report = [pscustomobject]@{
        timestamp        = (Get-Date).ToString('o')
        mode             = if ($Fix) { 'fix' } else { 'detect' }
        deep             = [bool]$Deep
        itemCount        = $(if ($deepCache) { $deepCache.items.total } else { $null })
        healthy          = ($errors.Count -eq 0)
        heartbeatMinutes = $hb
        freeGB           = $free
        processCount     = $proc.Count
        libraries        = $statuses
        findings         = $Findings
        actions          = $Actions
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content $ReportFile -Encoding UTF8

    Add-Content $LogFile ("{0} healthy={1} errors={2} warns={3} actions={4} hb={5}min" -f `
        (Get-Date -Format 's'), $report.healthy, $errors.Count, $warns.Count, $Actions.Count, $hb)

    if (-not $Quiet) {
        $verdict = if ($report.healthy) { 'HEALTHY' } else { 'PROBLEMS FOUND' }
        Write-Line "OneDrive doctor - $verdict"
        Write-Line "  processes=$($proc.Count)  heartbeat=${hb}min  freeC=${free}GB  libraries=$($statuses.Count)"
        foreach ($f in $Findings) { Write-Line ("  [{0}] {1}" -f $f.severity.ToUpper(), $f.message) }
        foreach ($a in $Actions)  { Write-Line "  FIXED: $a" }
        Write-Line "  report: $ReportFile"
    }

    if ($Actions.Count -gt 0) { exit 2 }
    # Errors only. Warnings are standing advisories (item count, a reset the
    # client wants) that persist until someone chooses to act, and failing the
    # scheduled task forever on those would make the automation look broken and
    # teach the user to ignore it. Warnings still appear in the report and log.
    if ($errors.Count -gt 0) { exit 1 }
    exit 0
}
catch {
    $msg = "$(Get-Date -Format 's') DOCTOR FAILED: $($_.Exception.Message)"
    Add-Content $LogFile $msg
    Write-Error $msg
    exit 3
}
