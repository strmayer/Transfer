# SQL Server Migration — Technical Procedure

Automated creation of a SQL Server instance with identical parameters to the
source environment. Developed for the Vault Migration 2024 → 2026 project.

**Scope:** This procedure covers SQL Server installation and configuration only.
Database backup and restore is performed separately by the **Autodesk Vault
migration helper** — that tool expects a correctly configured SQL instance to
already exist before it runs.

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

                                         ↓
                                         Autodesk Vault migration helper
                                         (backup / restore — out of scope here)
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

### Step 4 — Hand off to Autodesk Vault migration helper

Once the script exits with code `0`, the SQL Server instance is ready.
Database backup and restore is handled entirely by the **Autodesk Vault
migration helper** — run it according to the Autodesk migration guide.

**Pre-conditions the Autodesk tool expects (verified by our script):**

- [ ] SQL Server instance running with correct collation (`SQL_Latin1_General_CP1_CI_AS`)
- [ ] Mixed Mode authentication enabled, SA account active
- [ ] `VaultSys` login present — set password before starting the Autodesk tool (not captured in snapshot)
- [ ] TCP port reachable from application servers (check firewall if non-1433)
- [ ] Data directory exists and is writable (`D:\Vault\DB\` or as derived from snapshot)

### Step 5 — Post-migration checklist

After the Autodesk migration helper completes:

- [ ] Verify `Aquaflex` and `KnowledgeVaultMaster` databases are online
- [ ] Check database compatibility levels match source (see `sql_params.txt` block 3.6 — source: level 110)
- [ ] Verify SQL Agent jobs if applicable (re-create manually — not migrated by either script)
- [ ] Run a full backup to initialize backup chain on the new instance

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
| `/AGTSVCSTARTUPTYPE` | `SQL_Services` — entry matching `*Agent*` → `Startup_Type` | `Automatic` |

> **tempdb detection note:** The `3_12_TempDB` snapshot block is produced by an
> unfiltered `sys.master_files` query and therefore contains files from **all**
> databases (Vault DBs, master, model, msdb, tempdb). The install script
> cross-references with `DB_Files_All` (which has a `Database` column) to isolate
> only the true tempdb files (`tempdev`, `temp2`, `templog`) before computing
> the file count and sizes passed to setup.exe.

**Service account note:** The script reads the SQL Engine service account from
`SQL_Services` and logs it. If the source runs under a **domain account**
(e.g. `appl_admin@aquaflex.de`) rather than a built-in NT account, the script
emits a WARN and the operator must supply `-SqlSvcAccount` and `-SqlSvcPassword`
explicitly. Without this, setup uses `NT AUTHORITY\SYSTEM`.

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
    "SQL_Services":          [
      { "Service": "SQL Server (AUTODESKVAULT)",       "Service_Account": "appl_admin@aquaflex.de", "Startup_Type": "Automatic", "Status": "Running" },
      { "Service": "SQL Server-Agent (AUTODESKVAULT)", "Service_Account": "appl_admin@aquaflex.de", "Startup_Type": "Automatic", "Status": "Running" }
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

## Observed Differences Between Source Systems

Two source snapshots have been collected and tested (AZVDB2023, AZVDB2024).
Key differences relevant to the install script:

| Parameter | AZVDB2023 | AZVDB2024 | Handled by script |
|---|---|---|---|
| Server collation | `SQL_Latin1_General_CP1_CI_AS` | `Latin1_General_CI_AS` | Yes — read from snapshot |
| SQL Agent startup | `Manual` / Stopped | `Automatic` / Running | Yes — read from `SQL_Services` |
| Service account | `appl_admin@aquaflex.de` | `appl_admin@aquaflex.de` | Warning logged — set `-SqlSvcAccount` manually |
| `BUILTIN\Administrators` sysadmin | not present | present | No — recreate manually if needed |
| Aquaflex data size | ~35 GB | ~41 GB | No impact on install |
| KnowledgeVaultMaster log | 4 MB | 2.1 GB | No impact on install |
| Backup path | `E:\Backups\` | `F:\VaultBackups\2023\` | Out of scope (Autodesk helper) |
| sp_configure values | identical | identical | — |

**Collation note:** The collation difference (`SQL_Latin1_General_CP1_CI_AS` vs.
`Latin1_General_CI_AS`) is significant — it affects sort order and string comparison
behaviour in SQL. The install script applies whichever value the snapshot contains.
Verify that the Vault application is compatible with the target collation before go-live.

---

## Known Limitations

| Item | Detail |
|---|---|
| SQL Agent jobs | Captured in snapshot (block `3_13_SQL_Agent_Jobs`) for reference only — not re-created automatically |
| User databases | Backup/restore handled by Autodesk Vault migration helper (out of scope) |
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

*Document version: 1.3 — 2026-06-12*
*Author: (c) JM and Claude Code*
