<#
.SYNOPSIS
  Registers the OneDrive Doctor scheduled tasks.
.DESCRIPTION
  Two tasks, because the checks have very different costs:

    "OneDrive Doctor"       every 15 min + at logon - the fast pass (seconds)
    "OneDrive Doctor Deep"  once a day              - item counts and full
                                                      filename/PST sweeps,
                                                      which take minutes and
                                                      compete with OneDrive
                                                      for disk

  Both run as the current user with an interactive logon type: the doctor reads
  HKCU and talks to the OneDrive client in the user's own session, and neither
  works from SYSTEM. No elevation required. They start through wscript.exe and
  run-hidden.vbs, which run pwsh with no window and hand back its exit code;
  pwsh started straight from such a task opens a window on screen every run.

  Run with -Uninstall to remove both.
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    [int]$IntervalMinutes = 15,
    [string]$DeepAt = '05:30'
)

$ErrorActionPreference = 'Stop'
$Fast   = 'OneDrive Doctor'
$DeepT  = 'OneDrive Doctor Deep'
$Script = Join-Path $PSScriptRoot 'onedrive-doctor.ps1'
$Hidden = Join-Path $PSScriptRoot 'run-hidden.vbs'

if ($Uninstall) {
    foreach ($n in $Fast, $DeepT) {
        Unregister-ScheduledTask -TaskName $n -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Removed scheduled task '$n'."
    }
    return
}

if (-not (Test-Path $Script)) { throw "Doctor script not found at $Script" }
if (-not (Test-Path $Hidden)) { throw "Hidden launcher not found at $Hidden" }
# Store PowerShell: use its app alias, which survives updates. Get-Command can return the
# versioned package folder instead, which the Store deletes on update, and the tasks stop launching.
$Alias = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\Microsoft.PowerShell_8wekyb3d8bbwe\pwsh.exe'
$Pwsh = if (Test-Path $Alias) { $Alias } else { (Get-Command pwsh -ErrorAction SilentlyContinue).Source }
if (-not $Pwsh) { $Pwsh = (Get-Command powershell).Source }

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Limited

function New-DoctorAction {
    param([string]$ExtraArgs)
    New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\wscript.exe" -Argument (
        '//B //NoLogo "{0}" "{1}" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{2}" {3}' -f $Hidden, $Pwsh, $Script, $ExtraArgs)
}

# --- fast pass -------------------------------------------------------------
Unregister-ScheduledTask -TaskName $Fast -Confirm:$false -ErrorAction SilentlyContinue

$fastTriggers = @(
    (New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)),
    (New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME)
)
$fastSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew -StartWhenAvailable

Register-ScheduledTask -TaskName $Fast -Action (New-DoctorAction '-Fix -Quiet') `
    -Trigger $fastTriggers -Principal $principal -Settings $fastSettings `
    -Description 'Fast OneDrive health check. Applies tier 1-2 repairs only; never resets.' | Out-Null
Write-Host "Registered '$Fast' - every $IntervalMinutes minutes and at logon."

# --- daily deep scan -------------------------------------------------------
Unregister-ScheduledTask -TaskName $DeepT -Confirm:$false -ErrorAction SilentlyContinue

# Deep runs are read-only on purpose: no -Fix. They walk every library, so they
# are given a wide time limit and are skipped on battery.
$deepSettings = New-ScheduledTaskSettingsSet -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew -StartWhenAvailable

Register-ScheduledTask -TaskName $DeepT -Action (New-DoctorAction '-Deep -Quiet') `
    -Trigger (New-ScheduledTaskTrigger -Daily -At $DeepAt) `
    -Principal $principal -Settings $deepSettings `
    -Description 'Daily deep OneDrive scan: item counts, filename problems, Outlook data files. Read-only.' | Out-Null
Write-Host "Registered '$DeepT' - daily at $DeepAt (read-only)."

Write-Host ""
Write-Host "Reports: $env:LOCALAPPDATA\onedrive-doctor\last-report.json"
Write-Host "Check:   Get-ScheduledTaskInfo -TaskName '$Fast'"
