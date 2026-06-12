# SQL Server Migration — Technical Procedure

Automated migration of a SQL Server instance using parameter snapshot and
unattended installation. Developed for the Vault Migration 2024 → 2026 project.

**Toolchain:** `Get-SqlParameters.ps1` (source) → `Install-SqlFromSnapshot.ps1` (target)

---

## Overview

```
SOURCE SERVER                          TARGET SERVER
─────────────────────────────────      ─────────────────────────────────
Get-SqlParameters.ps1                  Install-SqlFromSnapshot.ps1
  │                                      │
  ├─ Connects to SQL instance            ├─ Phase 1: setup.exe /Q
  ├─ Collects 14 parameter blocks        │    collation, auth, FTS, tempdb
  ├─ Writes sql_params.json    ─────────►│
  └─ Writes sql_params.txt               ├─ Phase 2: post-install T-SQL
                                         │    sp_configure, SA, TCP port
                                         │
                                         └─ Phase 3: verification
                                              compare → exit 0 / exit 2
```

The JSON file is the single transfer artifact. No agent, no remote session,
no shared storage required — copy the file by any means (USB, fileshare, SFTP).

---

## Prerequisites

### Source server

| Requirement | Detail |
|---|---|
| PowerShell | 5.1+ or `pwsh` |
| SQL auth | Windows Auth **or** SQL login with `VIEW SERVER STATE`, `VIEW ANY DEFINITION` |
| Network | TCP access to the SQL instance (local or remote) |

### Target server

| Requirement | Detail |
|---|---|
| PowerShell | 5.1+ |
| OS | Windows Server 2022 (same or higher than source) |
| Disk layout | Target drives must exist (e.g. `D:\`). Directories are **auto-created** from the snapshot — no manual pre-configuration needed unless you want to override the derived paths. |
| SQL Server media | `setup.exe` accessible on the target (local path or mounted ISO) |
| Permissions | **Local Administrator** (`#Requires -RunAsAdministrator`) |
| SQL instance | **No existing instance** with the same name — the script exits 1 if found |

---

## Step-by-step Procedure

### Step 1 — Collect parameters on the source server

```powershell
# Run interactively — prompts for instance name and credentials
.\Get-SqlParameters.ps1
```

The script opens a connection prompt, collects all parameters, and saves two files
to the current user's Desktop:

```
sql_params_<hostname>_<yyyyMMdd_HHmmss>.json   ← machine-readable, used by install script
sql_params_<hostname>_<yyyyMMdd_HHmmss>.txt    ← human-readable report for review
```

Review the `.txt` report before proceeding. Verify:
- Authentication Mode (Mixed Mode required for Vault)
- `max server memory` — note the value vs. target RAM
- tempdb file count and sizes
- TCP port (non-1433 requires firewall rule on target)

### Step 2 — Transfer the JSON file

Copy `sql_params_*.json` to the target server by any means. Suggested destination:
`C:\Temp\sql_params.json`

The file contains no credentials, no passwords, and no secrets — it is safe to
transfer over any channel.

### Step 3 — Install on the target server

```powershell
# Standard invocation — disk layout is derived automatically from the snapshot
$pw = Read-Host -AsSecureString 'SA Password'
.\Install-SqlFromSnapshot.ps1 `
    -SetupExe     'D:\SQLSetup\setup.exe' `
    -SnapshotJson 'C:\Temp\sql_params.json' `
    -SaPassword   $pw
```

The script starts with an **interactive prompt** for hostname/year string replacement:

```
  String to replace in snapshot values (e.g. 2023) — ENTER to skip: 2023
  Replace '2023' with: 2025
```

This renames any snapshot-derived string containing the old value — e.g.
`AZVDB2023` → `AZVDB2025`, `ADMS-AZVDB2023` → `ADMS-AZVDB2025`.
Press ENTER twice to skip replacement.

The script then derives and logs the disk paths it will use:

```
  DataDir (derived): D:\Vault\DB
  LogDir  (derived): D:\Vault\DB
  BackupDir derived: D:\Vault\Backup
```

To override any path explicitly:

```powershell
.\Install-SqlFromSnapshot.ps1 `
    -SetupExe     'D:\SQLSetup\setup.exe' `
    -SnapshotJson 'C:\Temp\sql_params.json' `
    -SaPassword   $pw `
    -DataDir      'E:\Vault\DB' `
    -BackupDir    'F:\Vault\Backup'
```

**Dry run first:**
```powershell
.\Install-SqlFromSnapshot.ps1 ... -WhatIf
```
WhatIf prints all resolved parameters — features, collation, derived paths, tempdb layout —
without touching the disk or registry.

### Step 4 — Restore user databases

After the script exits 0, restore databases from backup:

```powershell
# Example — adjust paths and database names
Restore-SqlDatabase `
    -ServerInstance '.' `
    -Database 'VaultDB' `
    -BackupFile 'G:\SQLBackup\VaultDB_Full.bak' `
    -RelocateFile @(
        New-Object Microsoft.SqlServer.Management.Smo.RelocateFile('VaultDB',      'E:\SQLData\VaultDB.mdf'),
        New-Object Microsoft.SqlServer.Management.Smo.RelocateFile('VaultDB_log',  'F:\SQLLogs\VaultDB_log.ldf')
    )
```

### Step 5 — Post-restore checklist

- [ ] Verify `VaultSys` login exists and is enabled (`sql_params.txt` → VaultSys section)
- [ ] Set `VaultSys` password (not captured in snapshot — coordinate with Vault team)
- [ ] Check database compatibility levels match source (see `sql_params.txt` block 3.6)
- [ ] Verify SQL Agent jobs if applicable (re-create manually — jobs are not migrated by this script)
- [ ] Confirm TCP port is reachable from application servers (firewall rule)
- [ ] Run a full backup to initialize backup chain on target

---

## What the Install Script Configures

### Phase 1 — setup.exe arguments (derived from snapshot)

| setup.exe argument | Source in snapshot | Default if missing |
|---|---|---|
| `/SQLCOLLATION` | `3_5_Server_Collation.Server_Collation` | `SQL_Latin1_General_CP1_CI_AS` |
| `/SECURITYMODE=SQL` | `3_7_Auth_Mode.WindowsAuthOnly_Value == 0` | Mixed Mode |
| `/SAPWD` | `-SaPassword` parameter | — |
| `/TCPENABLED=1` | always enabled | — |
| `/FEATURES` includes `FULLTEXT` | `3_9a_FTS_Installed.FTS_Installed == 1` | omitted |
| `/SQLUSERDBDIR` | directory of first ROWS file in `DB_Files_All` (non-system DB) | `D:\Vault\DB` |
| `/SQLUSERDBLOGDIR` | directory of first LOG file in `DB_Files_All` (non-system DB) | same as DataDir |
| `/SQLBACKUPDIR` | derived from DataDir (`\DB` → `\Backup`) | `D:\Vault\Backup` |
| `/SQLTEMPDBFILECOUNT` | ROWS count in `DB_Files_All` where `Database = tempdb` | 1 |
| `/SQLTEMPDBFILESIZE` | `SizeMB` of first tempdb data file | 8 MB |
| `/SQLTEMPDBLOGFILESIZE` | `SizeMB` of tempdb log file | 8 MB |
| `/SQLTEMPDBDIR` | **only passed if `-TempDir` is explicitly set** — otherwise setup uses the instance DATA folder (matches source behaviour) | not passed |

> **tempdb detection note:** The `3_12_TempDB` snapshot block is produced by an
> unfiltered `sys.master_files` query and therefore contains files from **all**
> databases (Vault DBs, master, model, msdb, tempdb). The install script
> cross-references with `DB_Files_All` (which has a `Database` column) to isolate
> only the true tempdb files (`tempdev`, `temp2`, `templog`) before computing
> the file count and sizes passed to setup.exe.

### Phase 2 — Post-install T-SQL configuration

| Setting | Source block | Notes |
|---|---|---|
| `sp_configure` values | `3_11_SP_Configure` | All 7 captured parameters applied via `RECONFIGURE WITH OVERRIDE` |
| `max server memory` | `3_11_SP_Configure` | Capped at 85 % of target RAM if source value exceeds target |
| SA login password | `-SaPassword` | Applied and enabled if source SA was active |
| tempdb autogrowth | `3_12_TempDB.Autogrowth_*` | Applied per file (data + log); skips gracefully if file name differs |
| TCP static port | `3_8_TCP_Listener.TCP_Port` | Set via registry `IPAll\TcpPort`; clears `TcpDynamicPorts` |
| Service restart | automatic | Triggered if `remote access` or TCP port changed |

**sp_configure parameters captured and applied:**

| Parameter | Typical source value |
|---|---|
| `max server memory (MB)` | environment-specific |
| `min server memory (MB)` | usually 0 |
| `max degree of parallelism` | 0–8 |
| `cost threshold for parallelism` | 5–50 |
| `optimize for ad hoc workloads` | 0 or 1 |
| `backup compression default` | 0 or 1 |
| `remote access` | 1 |

### Phase 3 — Verification checks

| Check | Pass condition |
|---|---|
| Server collation | Matches snapshot exactly |
| Authentication mode | `IsIntegratedSecurityOnly` matches `WindowsAuthOnly_Value` |
| tempdb data file count | Matches snapshot count |
| All sp_configure values | `sys.configurations.value` matches snapshot `Configured` |

Verification failures exit with code 2 and detail each mismatch in the log.
The instance is functional — mismatches require manual correction.

---

## JSON Snapshot Structure

```json
{
  "Meta": {
    "CollectedAt":  "12.06.2026 13:27:20",
    "CollectedOn":  "AZVDB2023",
    "SqlInstance":  "AZVDB2023\\AUTODESKVAULT",
    "AuthMode":     "SQL Authentication (sa)",
    "ScriptVersion":"1.1 / JM Consulting"
  },
  "Blocks": {
    "3_5_Server_Collation":  [{ "Server_Collation": "SQL_Latin1_General_CP1_CI_AS" }],
    "3_7_Auth_Mode":         [{ "WindowsAuthOnly_Value": 0, "Authentication_Mode": "Mixed Mode ..." }],
    "3_8_TCP_Listener":      null,
    "3_9a_FTS_Installed":    [{ "FTS_Installed": 1 }],
    "3_11_SP_Configure":     [{ "Parameter": "max server memory (MB)", "Configured": 2147483647, ... }],
    "3_12_TempDB":           [
      { "File_Name": "Aquaflex",   "Type": "ROWS", "Path": "D:\\Vault\\DB\\Aquaflex.mdf",  ... },
      { "File_Name": "tempdev",    "Type": "ROWS", "Path": "C:\\...\\DATA\\tempdb.mdf",    ... },
      { "File_Name": "temp2",      "Type": "ROWS", "Path": "C:\\...\\DATA\\tempdb_mssql_2.ndf", ... },
      { "File_Name": "templog",    "Type": "LOG",  "Path": "C:\\...\\DATA\\templog.ldf",   ... }
    ],
    "DB_Files_All":          [
      { "Database": "Aquaflex",  "Type": "ROWS", "Logical_Name": "Aquaflex",  "Physical_Path": "D:\\Vault\\DB\\Aquaflex.mdf",  "SizeMB": 35457 },
      { "Database": "tempdb",    "Type": "ROWS", "Logical_Name": "tempdev",   "Physical_Path": "C:\\...\\DATA\\tempdb.mdf",    "SizeMB": 8 },
      { "Database": "tempdb",    "Type": "ROWS", "Logical_Name": "temp2",     "Physical_Path": "C:\\...\\DATA\\tempdb_mssql_2.ndf", "SizeMB": 8 },
      { "Database": "tempdb",    "Type": "LOG",  "Logical_Name": "templog",   "Physical_Path": "C:\\...\\DATA\\templog.ldf",   "SizeMB": 8 }
    ],
    "VaultSys_SA_Status":    [{ "Login": "sa", "Disabled": false, ... }],
    ...
  }
}
```

**Key block notes:**

| Block | Used for | Caution |
|---|---|---|
| `3_12_TempDB` | tempdb autogrowth settings (has `Autogrowth_Raw`, `Autogrowth_Percent` fields) | Contains **all** database files — `sys.master_files` is unfiltered. Do not use for tempdb file count. |
| `DB_Files_All` | disk path derivation + tempdb file count/size | Has `Database` column — filter on `Database = 'tempdb'` for tempdb files, exclude system DBs for user data paths. |

All block keys use underscores (`_`) — PowerShell's `ConvertFrom-Json` replaces
dots and spaces with underscores when creating property names.

---

## Exit Codes

| Code | Meaning | Action |
|---|---|---|
| `0` | Full success, all verification checks passed | Proceed to Step 4 (database restore) |
| `1` | Pre-condition failure | Fix the reported issue (wrong path, existing instance, bad JSON) |
| `2` | Setup failed **or** verification mismatch | Check log and SQL Server Setup Bootstrap log |
| `3` | Unhandled exception | Check log for stack trace |

Log file location: `<script directory>\Install-SqlFromSnapshot_<yyyyMMdd_HHmmss>.log`
SQL Server setup log: `%ProgramFiles%\Microsoft SQL Server\<ver>\Setup Bootstrap\Log\Summary.txt`

---

## Known Limitations

| Item | Detail |
|---|---|
| SQL Agent jobs | Captured in snapshot (block `3_13_SQL_Agent_Jobs`) for reference only — not re-created automatically |
| User databases | Not migrated — restore from backup separately (Step 4) |
| Additional logins | Snapshot lists all logins; only SA is configured — recreate others manually |
| Named instances | Supported via `-InstanceName`; service names and registry paths resolve automatically |
| SQL Server edition | setup.exe determines edition (Developer, Standard, Enterprise); the script does not select edition |
| Collation on existing DBs | `ALTER DATABASE ... COLLATE` is not performed — user databases inherit their own collation from backup |

---

## Security Notes

- The SA password is passed to `setup.exe` as a plain-text command-line argument.
  This is a SQL Server setup constraint and cannot be avoided. The value is **never
  written to the log file**.
- The JSON snapshot contains **no credentials** — it is safe to store and transfer.
- The script requires Local Administrator rights and must run in an elevated session.
- After migration, review SQL logins (`sql_params.txt` → SQL Logins section) and
  disable or remove any accounts that should not exist on the target.

---

*Document version: 1.1 — 2026-06-12*
*Author: (c) JM and Claude Code*
