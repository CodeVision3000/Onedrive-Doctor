# OneDrive Doctor

Detects OneDrive sync problems on a Windows machine and repairs the safe ones
without being asked. A scheduled task runs every 15 minutes; a Claude Code skill
handles anything the task is deliberately not allowed to touch.

This repairs the **local OneDrive sync client**. It is not a Microsoft Graph
tool and does nothing cloud-side.

## Why

OneDrive fails quietly. The client exits and nothing says so; the sync engine
wedges while the icon still looks fine; a file with a colon in its name or a
420-character path sits unsynced indefinitely. You find out when a file you
needed was never uploaded.

The doctor watches for those states on a timer and fixes the ones that are
genuinely safe to fix.

## Install

Requires Windows 10/11 and PowerShell 7 (`pwsh`); falls back to Windows
PowerShell if `pwsh` is absent. No admin rights needed.

```powershell
git clone https://github.com/CodeVision3000/Onedrive-Doctor.git `
  "$env:USERPROFILE\.claude\skills\onedrive-doctor"

& "$env:USERPROFILE\.claude\skills\onedrive-doctor\scripts\install-schedule.ps1"
```

The install registers a task named **OneDrive Doctor** that runs every 15
minutes and at logon, as the current user with an interactive logon type. That
matters: the doctor reads `HKCU` and talks to the OneDrive client in the user's
own session, and neither works from `SYSTEM`.

Remove it with `install-schedule.ps1 -Uninstall`.

Cloning into `.claude\skills\` also registers it as a Claude Code skill. The
script is standalone, so skip that path if you only want the scheduled task.

## Use

```powershell
# detect only, change nothing
.\scripts\onedrive-doctor.ps1

# detect and apply tier 1-2 repairs
.\scripts\onedrive-doctor.ps1 -Fix
```

```
OneDrive doctor - HEALTHY
  processes=1  heartbeat=0.3min  freeC=311.4GB  libraries=9
  [INFO] 1 Excel process(es) open - open workbooks hold locks and will not upload.
  report: C:\Users\<you>\AppData\Local\onedrive-doctor\last-report.json
```

Exit codes: `0` healthy, `1` findings nothing fixed, `2` repairs applied,
`3` the doctor itself failed.

Every run writes `last-report.json` and appends one line to `doctor.log`, both
under `%LOCALAPPDATA%\onedrive-doctor`.

## What it detects

| Finding | Meaning |
|---|---|
| `not-running` | the client process is gone |
| `engine-stale` | the process is alive but the sync engine stopped logging |
| `client-missing` | OneDrive is not installed |
| `mount-missing` | a registered library folder no longer exists |
| `library-error` | the shell reports an error state on a library |
| `low-disk` | free space below the threshold; sync stalls as a disk fills |
| `files-invalid-character` | `< > : " \| ? *` in a filename |
| `files-trailing-space-or-dot` | a name ending in a space or dot |
| `files-path-too-long` | a full path near the limit |
| `files-conflict-copy` | a `-COMPUTERNAME` conflict copy |
| `reset-requested` | the client has flagged itself for a reset |
| `excel-open`, `scan-truncated` | context, not faults |

## The fix ladder

**Tier 1 — automatic.** Client not running, so start it with `/background`.

**Tier 2 — automatic, with hysteresis.** The sync engine writes a log every
minute or two even when idle, so silence past `-StaleMinutes` means a wedge.
The doctor requires **two consecutive** stale checks before acting, then issues
`/shutdown` and restarts. Without that hysteresis a single long upload would get
a restart dropped on top of it.

**Tier 3 — never automatic.** `/reset` re-syncs every attached library, which on
a machine with large team sites is hours of re-download. Renaming files OneDrive
refuses means editing files that belong to other people in shared libraries.
Unlinking an account forces a new sign-in. The scheduled task cannot reach any
of these; they are surfaced for a human, with their cost stated.

That boundary is the point of the design. An auto-healer that is willing to
trigger a multi-hour resync is more dangerous than the problem it fixes.

## Options

| Switch | Default | Purpose |
|---|---|---|
| `-Fix` | off | apply tier 1-2 repairs |
| `-StaleMinutes` | 15 | heartbeat silence before a wedge is called |
| `-MinFreeGB` | 10 | low-disk threshold |
| `-PathWarn` | 380 | path length that triggers a warning |
| `-ScanHours` | 24 | only inspect files modified this recently |
| `-ScanSecsPerLib` | 20 | per-library scan budget |
| `-Quiet` | off | suppress console output; still writes report and log |

The file scan is time-boxed per library because a synced team site can hold
millions of files. When the budget runs out the report says so rather than
pretending the scan was complete. Raise `-ScanSecsPerLib` when you need full
coverage of a specific library.

## How it works

**Libraries** come from `HKCU:\Software\SyncEngines\Providers\OneDrive`, which
lists every mount including SharePoint team sites. The `mysite` and `personal`
entries point at the same path on business tenants, so mounts are deduplicated.

**Per-library status** comes from the shell's *Availability status* column via
`Shell.Application`. Two details matter and both caused bugs during development:
several columns are literally named `Status` and appear *before* the one that
carries a value, so the lookup matches `Availability status` exactly across the
whole range first; and the column index is per folder view rather than global,
so it is discovered and cached per parent folder.

**Liveness** is the modification time of the newest `SyncEngine-*.odl` log. The
filter is deliberate: the `Common`, `ListSync` and `OD4` log folders keep
writing while a wedged sync engine stays silent, so watching all `.odl` files
would mask the exact failure the check exists to catch. The logs themselves are
obfuscated binary and are never parsed — the timestamp alone is the signal.

## Limitations

The availability column reports file materialization, not sync health, so a
deliberately **paused** client still looks fine. Detecting pause reliably needs
the third-party `OneDriveLib.dll` (`Get-ODStatus`), which this does not bundle.

Column names are read in English; a non-English Windows install needs the
`Availability status` string localized in `Get-AvailabilityColumn`.

While the task is enabled, quitting OneDrive does not stick — the doctor treats
a stopped client as a fault and restarts it within 15 minutes. Run
`Disable-ScheduledTask -TaskName 'OneDrive Doctor'` before stopping it on
purpose.

## Files

```
scripts/onedrive-doctor.ps1     detection and tier 1-2 repair
scripts/install-schedule.ps1    registers/removes the scheduled task
SKILL.md                        Claude Code skill: triage, tier 3 escalation
```

## License

MIT. See `LICENSE`.
