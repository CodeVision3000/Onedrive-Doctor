<#
.SYNOPSIS
  Registers the "OneDrive Doctor" scheduled task: every 15 minutes, plus at logon.
.DESCRIPTION
  Runs as the current user with an interactive logon type - the doctor reads HKCU
  and talks to the same-session OneDrive client, neither of which works from SYSTEM.
  No elevation required.

  Run with -Uninstall to remove the task.
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    [int]$IntervalMinutes = 15
)

$ErrorActionPreference = 'Stop'
$TaskName = 'OneDrive Doctor'
$Script   = Join-Path $PSScriptRoot 'onedrive-doctor.ps1'

if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    return
}

if (-not (Test-Path $Script)) { throw "Doctor script not found at $Script" }
$Pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
if (-not $Pwsh) { $Pwsh = (Get-Command powershell).Source }

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction -Execute $Pwsh -Argument (
    '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Fix -Quiet -ScanSecsPerLib 8' -f $Script)

$triggers = @(
    (New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)),
    (New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME)
)

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew -StartWhenAvailable

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers `
    -Principal $principal -Settings $settings `
    -Description 'Detects OneDrive sync problems and applies safe repairs (tier 1-2 only). Never resets.' | Out-Null

Write-Host "Registered '$TaskName' - every $IntervalMinutes minutes and at logon."
Write-Host "Reports: $env:LOCALAPPDATA\onedrive-doctor\last-report.json"
Write-Host "Check:   Get-ScheduledTaskInfo -TaskName '$TaskName'"
