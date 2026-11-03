---
name: onedrive-doctor
description: Diagnose and repair local OneDrive sync problems on this machine - client not running, sync engine wedged, library errors, files OneDrive refuses (invalid characters, trailing spaces, over-long paths, conflict copies), low disk, pending account resets. Use when OneDrive is stuck, files are not uploading, a bid folder shows out of date, sync icons are spinning or red, or the scheduled onedrive-doctor task reported a problem. Not for Microsoft Graph or cloud-side file operations.
---

# OneDrive Doctor

Repairs the **local OneDrive sync client**. Cloud-side file operations (Graph /
Composio) are a different toolkit entirely - do not reach for this skill for
those.

## How it runs

A scheduled task (`OneDrive Doctor`) runs the script every 15 minutes with
`-Fix`, so tier 1-2 repairs happen without anyone present. This skill is the
escalation layer: read what the task found, explain it, and own the tier 3
decisions the task is deliberately not allowed to make.

## Step 1 - read the last report

```powershell
Get-Content "$env:LOCALAPPDATA\onedrive-doctor\last-report.json" -Raw | ConvertFrom-Json
```

Trend over time: `Get-Content "$env:LOCALAPPDATA\onedrive-doctor\doctor.log" -Tail 30`

## Step 2 - run it fresh

```powershell
& "$env:USERPROFILE\.claude\skills\onedrive-doctor\scripts\onedrive-doctor.ps1"
```

Detect-only. Add `-Fix` to apply tier 1-2. Exit codes: `0` healthy, `1` findings
not fixed, `2` repairs applied, `3` the doctor itself failed.

Useful switches: `-StaleMinutes`, `-MinFreeGB`, `-PathWarn`, `-ScanHours`,
`-ScanSecsPerLib` (per-library scan budget; raise it when the report says the
scan was truncated and you need full coverage of a specific library).

## What it checks

| Finding | Tier | Handling |
|---|---|---|
| `not-running` | 1 | auto: `OneDrive.exe /background` |
| `engine-stale` | 2 | auto after two consecutive stale checks: `/shutdown` then `/background` |
| `client-missing` | 3 | OneDrive is not installed - needs IT |
| `mount-missing` | 3 | a synced library folder is gone |
| `library-error` | 3 | the shell reports an error state on a library |
| `low-disk` | 3 | free space on C: below threshold |
| `files-invalid-character`, `files-trailing-space-or-dot`, `files-path-too-long`, `files-conflict-copy` | 3 | **report only** |
| `reset-requested` | 3 | the client has asked for a reset |
| `excel-open`, `scan-truncated` | info | context, not a problem |

The heartbeat check reads the modification time of the newest `.odl` log, not
its contents. Those logs are obfuscated binary; decoding them is not worth it,
and the file timestamp alone is a reliable liveness signal - the engine writes
one every minute or two even when idle.

## Tier 3 - ask first, every time

**Never run these from automation, and never without saying what they cost.**

- **`OneDrive.exe /reset`** re-syncs *every* library on the machine. On a host
  with large team sites attached that is hours of re-download, and those
  libraries are unreliable until it finishes. Ask, give the time estimate, and
  only then: `& "C:\Program Files\Microsoft OneDrive\OneDrive.exe" /reset` (then
  `/background` if it does not relaunch itself).
- **Renaming files OneDrive refuses.** Report the paths; do not rename. A file
  in a shared library belongs to whoever put it there, and the rename propagates
  to everyone else syncing it. Say which files and what is wrong with each; let
  the user decide.
- **Unlinking or re-adding an account.** Requires signing back in.

## Known state on this machine

`LOCAL-NOTES.md` next to this file records what is specific to this host: which
libraries sync, which accounts are dormant, and why renaming and `/reset` are
handled the way they are here. Read it before acting on a tier 3 finding. It is
gitignored and never leaves the machine.

Normal library statuses are `Available when online` and `Sync pending`. The
latter is routine during active upload and is not a fault on its own.

## Manage the schedule

Install or remove it:

```powershell
& "$env:USERPROFILE\.claude\skills\onedrive-doctor\scripts\install-schedule.ps1"
& "$env:USERPROFILE\.claude\skills\onedrive-doctor\scripts\install-schedule.ps1" -Uninstall
```

```powershell
Get-ScheduledTaskInfo -TaskName 'OneDrive Doctor'
Start-ScheduledTask   -TaskName 'OneDrive Doctor'
Disable-ScheduledTask -TaskName 'OneDrive Doctor'
```

**While the task is enabled, quitting OneDrive does not stick** - the doctor
treats a stopped client as a fault and restarts it within 15 minutes. To stop
OneDrive deliberately, disable the task first.

## Optional: true per-library sync state

The shell's availability column reports file materialization, not sync health,
so a **paused** client can still look fine. The community `OneDriveLib.dll`
(`Get-ODStatus`) returns real per-library states (UpToDate / Syncing / Error /
Paused). It is a third-party binary and is not installed; mention it if paused
state matters, do not add it unasked.
