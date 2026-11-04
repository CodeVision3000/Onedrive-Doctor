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

The install registers two tasks:

- **OneDrive Doctor** — every 15 minutes and at logon, the fast pass, with
  `-Fix`
- **OneDrive Doctor Deep** — once a day (`-DeepAt`, default 05:30), the
  expensive sweeps, read-only

Both run as the current user with an interactive logon type. That matters: the
doctor reads `HKCU` and talks to the OneDrive client in the user's own session,
and neither works from `SYSTEM`. They start through `wscript.exe` and
`scripts/run-hidden.vbs`, which run `pwsh` with no window and pass its exit code
back to Task Scheduler: `pwsh` started straight from such a task opens a
PowerShell window on screen every 15 minutes.

Remove both with `install-schedule.ps1 -Uninstall`.

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

Exit codes: `0` healthy or advisory-only, `1` errors needing action, `2` repairs applied,
`3` the doctor itself failed.

Every run writes `last-report.json` and appends one line to `doctor.log`, both
under `%LOCALAPPDATA%\onedrive-doctor`.

## What it detects

| Finding | Meaning |
|---|---|
| `not-running` | the client process is gone |
| `engine-stale` | the process is alive but the sync engine stopped logging |
| `outlook-data-file` | a `.pst`/`.ost` inside a synced folder — Outlook holds these open, so they can **never** sync |
| `client-missing` | OneDrive is not installed |
| `mount-missing` | a registered library folder no longer exists |
| `library-error` | the shell reports an error state on a library |
| `low-disk` | free space below the threshold; sync stalls as a disk fills |
| `too-many-items` | synced item count over `-MaxItems`; the usual cause of permanent high CPU |
| `bounce-suppressed` | a restart was warranted but withheld, because at this item count it would cost hours |
| `files-invalid-character` | `< > : " \| ? *` in a filename |
| `files-trailing-space-or-dot` | a name ending in a space or dot |
| `files-path-too-long` | a full path near the limit |
| `files-conflict-copy` | a `-COMPUTERNAME` conflict copy |
| `reset-requested` | the client has flagged itself for a reset |
| `outlook-data-file-cloud` | a data file that is cloud-only — syncs fine, blocks nothing, leave it |
| `excel-open`, `scan-truncated`, `no-deep-scan`, `deep-scan-stale` | context, not faults |

### Fast checks vs deep checks

Process, heartbeat, library status, disk and Outlook data files in their usual
locations take seconds and run every time. Item counts and full filename sweeps
take **minutes** on a large machine and compete with OneDrive for the same
disk, so they run only under `-Deep`, cache their results, and the fast pass
re-reports the cache. The installer schedules them separately.

## The fix ladder

**Tier 1 — automatic.** Client not running, so start it with `/background`.

**Tier 2 — automatic, with two guards.** The sync engine writes a log every
minute or two even when idle, so silence past `-StaleMinutes` means a wedge.
The doctor requires **two consecutive** stale checks before acting, then issues
`/shutdown` and restarts. Without that hysteresis a single long upload would get
a restart dropped on top of it.

The second guard matters more. A restart discards scan progress, and on a
machine with a large item count that scan takes **hours** — so an automatic
bounce inflicts far more damage than the wedge it is guessing at. Above
`-MaxItems` the doctor refuses to restart and reports `bounce-suppressed`
instead. A repair that is correct on a normal machine is not automatically
correct on a big one.

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
| `-Deep` | off | run item counts and full filename/PST sweeps, then cache them |
| `-MaxItems` | 300000 | item count above which sync degrades and auto-restart is withheld |
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

**Outlook data files** are detected as a *configuration* problem, not a symptom.
OneDrive reports "this file is in use" only in its tray flyout, which no API
exposes. But the cause is deterministic: Outlook holds `.pst`/`.ost` open for
its whole session, so one inside a sync root can never finish uploading.
Finding the file is reliable where reading the error is not. Files that are
cloud-only placeholders are reported separately as harmless — they already
synced, and moving one out would force a full download first.

### Two things that are not progress signals

Both of these were tried, and both are wrong:

- **`LastDAUEventTimestamp_*` in the account registry key.** DAU is Daily Active
  Usage. These fire once per day as telemetry, so they sit frozen for hours
  while uploads complete normally. Anything built on them reports a healthy
  client as hung.
- **Sustained high CPU.** On a machine near or past the item ceiling, 100% of a
  core for an hour is the steady state, not a fault.

Treat "is it making progress" as unanswerable from the outside, and detect
causes instead: dead process, silent log, blocking file, item count.

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
scripts/run-hidden.vbs          starts the tasks' pwsh with no window
SKILL.md                        Claude Code skill: triage, tier 3 escalation
```

## License

MIT. See `LICENSE`.
