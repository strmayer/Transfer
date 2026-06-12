#Requires -Version 5.1
<#
.SYNOPSIS
    Automated collection of all SQL Server parameters from the source environment for Vault migration 2024 → 2026.

.DESCRIPTION
    This script connects to the AUTODESKVAULT SQL instance of the existing Vault 2024 environment
    and collects all parameters required for setting up the target environment and data migration.

    Output:
      - Readable text report (TXT)
      - Structured JSON (for machine processing/documentation)

    Requires: .NET SqlClient (always available since Windows Server 2012 / PS 5.1)
    No external modules required (no SQLPS / no dbatools).

.EXAMPLE
    PS C:\> .\get-sqlparameters.ps1
    Runs the script interactively, prompting for SQL instance, authentication, and output path.

.NOTES
    V1.0/Created  - 2026-05-08 - Initial version
    V1.1/Modified - 2026-05-20 - Translated to English; fixed connection/resource disposal; added warn on sp_configure failure

    Author: (c) JM and Claude Code
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Helper Functions
# ─────────────────────────────────────────────────────────────────────────────

function Write-Banner {
    $banner = @"

  ╔══════════════════════════════════════════════════════════════════════════╗
  ║   JM Consulting  ·  Vault Migration 2024 → 2026                        ║
  ║   SQL Parameter Collection – Source Environment                         ║
  ║   Version 1.1  ·  Dr. Jörg Mayer                                       ║
  ╚══════════════════════════════════════════════════════════════════════════╝
"@
    Write-Host $banner -ForegroundColor Cyan
}

function Write-Step {
    param([string]$Text)
    Write-Host "`n  ► $Text" -ForegroundColor Yellow
}

function Write-Ok {
    param([string]$Text)
    Write-Host "    ✓ $Text" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Text)
    Write-Host "    ⚠ $Text" -ForegroundColor DarkYellow
}

function Write-Fail {
    param([string]$Text)
    Write-Host "    ✗ $Text" -ForegroundColor Red
}

function Invoke-SqlQuery {
    <#
    Executes a T-SQL query and returns the result as an array of PSCustomObjects.
    Returns $null on error and writes a warning.
    #>
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Query,
        [string]$Label
    )
    $cmd     = $null
    $adapter = $null
    try {
        $cmd                = $Connection.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = 60

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        [void]$adapter.Fill($dataset)

        $rows = [System.Collections.Generic.List[object]]::new()
        if ($dataset.Tables.Count -gt 0) {
            foreach ($row in $dataset.Tables[0].Rows) {
                $obj = [ordered]@{}
                foreach ($col in $dataset.Tables[0].Columns) {
                    $val = $row[$col.ColumnName]
                    $obj[$col.ColumnName] = if ($val -is [DBNull]) { $null } else { $val }
                }
                $rows.Add([PSCustomObject]$obj)
            }
        }
        Write-Ok $Label
        return , $rows.ToArray()
    }
    catch {
        Write-Warn "$Label – Error: $($_.Exception.Message)"
        return $null
    }
    finally {
        if ($null -ne $adapter) { $adapter.Dispose() }
        if ($null -ne $cmd)     { $cmd.Dispose() }
    }
}

function Format-Table-Text {
    param([object[]]$Data, [string]$Title)
    if (-not $Data -or $Data.Count -eq 0) {
        return "  [$Title]`n  (no data)`n"
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("  [$Title]")

    # Calculate column widths
    $cols   = $Data[0].PSObject.Properties.Name
    $widths = @{}
    foreach ($col in $cols) {
        $widths[$col] = $col.Length
        foreach ($row in $Data) {
            $val = if ($null -eq $row.$col) { 'NULL' } else { "$($row.$col)" }
            if ($val.Length -gt $widths[$col]) { $widths[$col] = $val.Length }
        }
        if ($widths[$col] -gt 80) { $widths[$col] = 80 }
    }

    # Header row
    $header  = '  '
    $divider = '  '
    foreach ($col in $cols) {
        $header  += $col.PadRight($widths[$col] + 2)
        $divider += ('-' * $widths[$col]) + '  '
    }
    [void]$sb.AppendLine($header)
    [void]$sb.AppendLine($divider)

    # Data rows
    foreach ($row in $Data) {
        $line = '  '
        foreach ($col in $cols) {
            $val = if ($null -eq $row.$col) { 'NULL' } else { "$($row.$col)" }
            if ($val.Length -gt 80) { $val = $val.Substring(0, 77) + '...' }
            $line += $val.PadRight($widths[$col] + 2)
        }
        [void]$sb.AppendLine($line)
    }
    [void]$sb.AppendLine()
    return $sb.ToString()
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Connection Parameters
# ─────────────────────────────────────────────────────────────────────────────

Clear-Host
Write-Banner

Write-Host "`n  Please enter connection parameters." -ForegroundColor White
Write-Host "  Press ENTER to accept the default value shown in brackets.`n" -ForegroundColor Gray

# SQL Server instance
$defaultInstance = "$($env:COMPUTERNAME)\AUTODESKVAULT"
$inputInstance   = Read-Host "  SQL Server Instance [$defaultInstance]"
$sqlInstance     = if ([string]::IsNullOrWhiteSpace($inputInstance)) { $defaultInstance } else { $inputInstance.Trim() }

# Authentication mode
Write-Host ""
Write-Host "  Authentication:" -ForegroundColor White
Write-Host "    [1] SQL Authentication (recommended: sa account)" -ForegroundColor Gray
Write-Host "    [2] Windows Authentication (current Windows user)" -ForegroundColor Gray
$authChoice = Read-Host "  Choice [1]"
if ([string]::IsNullOrWhiteSpace($authChoice)) { $authChoice = '1' }

$useWindowsAuth = $authChoice -eq '2'

$sqlUser     = $null
$sqlPassword = $null

if (-not $useWindowsAuth) {
    $defaultUser = 'sa'
    $inputUser   = Read-Host "  SQL Login [$defaultUser]"
    $sqlUser     = if ([string]::IsNullOrWhiteSpace($inputUser)) { $defaultUser } else { $inputUser.Trim() }
    $secPwd      = Read-Host "  Password for '$sqlUser'" -AsSecureString
    $sqlPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                       [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPwd))
}

# Output path
$defaultOutput = "$($env:USERPROFILE)\Desktop\VaultSQL_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmm')"
$inputOutput   = Read-Host "  Output file (without extension) [$defaultOutput]"
$outputBase    = if ([string]::IsNullOrWhiteSpace($inputOutput)) { $defaultOutput } else { $inputOutput.Trim() }
$outputTxt     = "$outputBase.txt"
$outputJson    = "$outputBase.json"


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Connect to SQL Server
# ─────────────────────────────────────────────────────────────────────────────

Write-Step "Connecting to '$sqlInstance' ..."

$connString = if ($useWindowsAuth) {
    "Server=$sqlInstance;Database=master;Integrated Security=True;Connection Timeout=15;"
} else {
    "Server=$sqlInstance;Database=master;User Id=$sqlUser;Password=$sqlPassword;Connection Timeout=15;"
}

try {
    $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
    $conn.Open()
    Write-Ok "Connection established."
}
catch {
    Write-Fail "Connection failed: $($_.Exception.Message)"
    Write-Host "`n  Please verify instance name, credentials, and firewall rules." -ForegroundColor Red
    Read-Host "`n  Press ENTER to exit"
    exit 1
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Data Collection
# ─────────────────────────────────────────────────────────────────────────────

Write-Step "Data collection running – please wait ..."

$inventory = [ordered]@{
    Meta = [ordered]@{
        CollectedAt   = (Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
        CollectedOn   = $env:COMPUTERNAME
        SqlInstance   = $sqlInstance
        AuthMode      = if ($useWindowsAuth) { 'Windows Authentication' } else { "SQL Authentication ($sqlUser)" }
        ScriptVersion = '1.1 / JM Consulting'
    }
    Blocks = [ordered]@{}
}

try {

# ── 3.1 / 3.2  Version and Edition ───────────────────────────────────────────
$q_version = @"
SELECT
    @@VERSION                                           AS [SQL_FullVersion],
    CAST(SERVERPROPERTY('ProductVersion')  AS NVARCHAR) AS [ProductVersion],
    CAST(SERVERPROPERTY('ProductLevel')    AS NVARCHAR) AS [ProductLevel],
    CAST(SERVERPROPERTY('Edition')         AS NVARCHAR) AS [Edition],
    CAST(SERVERPROPERTY('EngineEdition')   AS NVARCHAR) AS [EngineEdition],
    @@SERVERNAME                                        AS [ServerName],
    CAST(SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS NVARCHAR) AS [ComputerName]
"@
$r_version = Invoke-SqlQuery -Connection $conn -Query $q_version -Label '3.1/3.2  Version and Edition'

# ── 3.3  Server name / instance name (already included in version query) ──────

# ── 3.5 / 3.6  Collation (server + per database) ─────────────────────────────
$q_collation_server = "SELECT CAST(SERVERPROPERTY('Collation') AS NVARCHAR) AS [Server_Collation]"
$r_collation_server = Invoke-SqlQuery -Connection $conn -Query $q_collation_server -Label '3.5     Server Collation'

$q_collation_db = @"
SELECT
    name                                         AS [Database],
    collation_name                               AS [Collation],
    recovery_model_desc                          AS [Recovery_Model],
    state_desc                                   AS [Status],
    compatibility_level                          AS [Compat_Level],
    is_read_only                                 AS [ReadOnly],
    CAST(create_date AS NVARCHAR)                AS [Created_At]
FROM sys.databases
ORDER BY name
"@
$r_collation_db = Invoke-SqlQuery -Connection $conn -Query $q_collation_db -Label '3.6     Collation per Database'

# ── 3.7  Authentication Mode ──────────────────────────────────────────────────
$q_auth = @"
SELECT
    CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS INT) AS [WindowsAuthOnly_Value],
    CASE CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS INT)
        WHEN 0 THEN 'Mixed Mode (SQL + Windows) – CORRECT for Vault'
        WHEN 1 THEN 'Windows Auth ONLY – PROBLEM: Vault requires Mixed Mode!'
        ELSE 'Unknown'
    END AS [Authentication_Mode]
"@
$r_auth = Invoke-SqlQuery -Connection $conn -Query $q_auth -Label '3.7     Authentication Mode'

# ── 3.8  TCP Port / Network Listener ─────────────────────────────────────────
$q_tcp = @"
SELECT
    local_net_address   AS [IP_Address],
    local_tcp_port      AS [TCP_Port],
    is_listening        AS [Listening],
    ip_address          AS [Client_IP]
FROM sys.dm_tcp_listener_states
WHERE is_ipv4 = 1
ORDER BY local_tcp_port
"@
$r_tcp = Invoke-SqlQuery -Connection $conn -Query $q_tcp -Label '3.8     TCP Listener / Port'

# ── 3.9  Full-Text Search ─────────────────────────────────────────────────────
$q_fts_installed = "SELECT CAST(SERVERPROPERTY('IsFullTextInstalled') AS INT) AS [FTS_Installed]"
$r_fts_installed = Invoke-SqlQuery -Connection $conn -Query $q_fts_installed -Label '3.9a    Full-Text Search installed?'

$q_fts_catalogs = @"
SELECT
    fc.name                         AS [Catalog_Name],
    DB_NAME(fc.database_id)         AS [Database],
    fc.is_default                   AS [IsDefault],
    fc.is_accent_sensitivity_on     AS [AccentSensitive],
    CAST(fc.create_date AS NVARCHAR) AS [Created_At]
FROM sys.fulltext_catalogs fc
"@
$r_fts_catalogs = Invoke-SqlQuery -Connection $conn -Query $q_fts_catalogs -Label '3.9b    Full-Text Catalogs'

# ── 3.10  Recovery Model (already included in DB collation query) ─────────────

# ── 3.11  Memory Configuration ───────────────────────────────────────────────
$q_advanced = @"
EXEC sp_configure 'show advanced options', 1
RECONFIGURE
"@
$cmdAdv = $null
try {
    $cmdAdv              = $conn.CreateCommand()
    $cmdAdv.CommandText  = $q_advanced
    [void]$cmdAdv.ExecuteNonQuery()
} catch {
    Write-Warn "sp_configure 'show advanced options' failed: $($_.Exception.Message)"
} finally {
    if ($null -ne $cmdAdv) { $cmdAdv.Dispose() }
}

$q_sp_config = @"
SELECT
    name                AS [Parameter],
    value               AS [Configured],
    value_in_use        AS [Active],
    minimum             AS [Min],
    maximum             AS [Max],
    description         AS [Description]
FROM sys.configurations
WHERE name IN (
    'max server memory (MB)',
    'min server memory (MB)',
    'max degree of parallelism',
    'cost threshold for parallelism',
    'optimize for ad hoc workloads',
    'backup compression default',
    'remote access'
)
ORDER BY name
"@
$r_sp_config = Invoke-SqlQuery -Connection $conn -Query $q_sp_config -Label '3.11    sp_configure (Memory, MAXDOP, ...)'

# ── 3.12  tempdb Configuration ────────────────────────────────────────────────
$q_tempdb = @"
SELECT
    name                            AS [File_Name],
    type_desc                       AS [Type],
    physical_name                   AS [Path],
    size * 8 / 1024                 AS [SizeMB],
    growth                          AS [Autogrowth_Raw],
    is_percent_growth               AS [Autogrowth_Percent],
    CASE is_percent_growth
        WHEN 1 THEN CAST(growth AS NVARCHAR) + '%'
        ELSE CAST(growth * 8 / 1024 AS NVARCHAR) + ' MB'
    END AS [Autogrowth_Readable]
FROM tempdb.sys.master_files
ORDER BY type, name
"@
$r_tempdb = Invoke-SqlQuery -Connection $conn -Query $q_tempdb -Label '3.12    tempdb Configuration'

# ── Database Files (all Vault-relevant DBs) ───────────────────────────────────
$q_dbfiles = @"
SELECT
    DB_NAME(database_id)            AS [Database],
    type_desc                       AS [Type],
    name                            AS [Logical_Name],
    physical_name                   AS [Physical_Path],
    size * 8 / 1024                 AS [SizeMB],
    CASE is_percent_growth
        WHEN 1 THEN CAST(growth AS NVARCHAR) + '%'
        ELSE CAST(growth * 8 / 1024 AS NVARCHAR) + ' MB'
    END AS [Autogrowth]
FROM sys.master_files
ORDER BY DB_NAME(database_id), type_desc
"@
$r_dbfiles = Invoke-SqlQuery -Connection $conn -Query $q_dbfiles -Label '        Database Files (all DBs)'

# ── 3.13  SQL Agent Jobs ──────────────────────────────────────────────────────
$q_jobs = @"
SELECT
    j.name                           AS [Job_Name],
    j.enabled                        AS [Active],
    j.description                    AS [Description],
    CAST(j.date_created AS NVARCHAR) AS [Created_At],
    c.name                           AS [Category],
    ISNULL(
        (SELECT TOP 1 CAST(run_date AS NVARCHAR)
         FROM msdb.dbo.sysjobhistory h
         WHERE h.job_id = j.job_id AND h.step_id = 0
         ORDER BY run_date DESC, run_time DESC),
        'Never run'
    )                                AS [Last_Run]
FROM msdb.dbo.sysjobs j
LEFT JOIN msdb.dbo.syscategories c ON j.category_id = c.category_id
ORDER BY j.name
"@
$r_jobs = Invoke-SqlQuery -Connection $conn -Query $q_jobs -Label '3.13    SQL Agent Jobs'

# ── Logins / Service Account ──────────────────────────────────────────────────
$q_logins = @"
SELECT
    sp.name                         AS [Login_Name],
    sp.type_desc                    AS [Type],
    sp.is_disabled                  AS [Disabled],
    sp.default_database_name        AS [Default_DB],
    sp.create_date                  AS [Created_At],
    STUFF((
        SELECT ', ' + sr.name
        FROM sys.server_role_members srm
        JOIN sys.server_principals sr ON srm.role_principal_id = sr.principal_id
        WHERE srm.member_principal_id = sp.principal_id
        FOR XML PATH('')
    ), 1, 2, '')                    AS [Server_Roles]
FROM sys.server_principals sp
WHERE sp.type IN ('S', 'U', 'G')
  AND sp.name NOT LIKE '##%'
ORDER BY sp.type, sp.name
"@
$r_logins = Invoke-SqlQuery -Connection $conn -Query $q_logins -Label '        SQL Logins and Server Roles'

$q_services = @"
SELECT
    servicename                     AS [Service],
    service_account                 AS [Service_Account],
    startup_type_desc               AS [Startup_Type],
    status_desc                     AS [Status],
    process_id                      AS [ProcessID]
FROM sys.dm_server_services
"@
$r_services = Invoke-SqlQuery -Connection $conn -Query $q_services -Label '        SQL Server Services and Accounts'

# ── Backup History (last 20 entries) ─────────────────────────────────────────
$q_backups = @"
SELECT TOP 20
    bs.database_name                AS [Database],
    CASE bs.type
        WHEN 'D' THEN 'Full'
        WHEN 'I' THEN 'Differential'
        WHEN 'L' THEN 'Log'
        ELSE bs.type
    END                             AS [Backup_Type],
    CAST(bs.backup_start_date  AS NVARCHAR) AS [Start],
    CAST(bs.backup_finish_date AS NVARCHAR) AS [End],
    DATEDIFF(MINUTE, bs.backup_start_date, bs.backup_finish_date) AS [Duration_Min],
    CAST(bs.backup_size / 1048576 AS BIGINT) AS [SizeMB],
    bmf.physical_device_name        AS [Target_Path]
FROM msdb.dbo.backupset bs
JOIN msdb.dbo.backupmediafamily bmf ON bs.media_set_id = bmf.media_set_id
ORDER BY bs.backup_start_date DESC
"@
$r_backups = Invoke-SqlQuery -Connection $conn -Query $q_backups -Label '        Backup History (last 20 entries)'

# ── Check VaultSys account ────────────────────────────────────────────────────
$q_vaultsys = @"
SELECT
    name                            AS [Login],
    type_desc                       AS [Type],
    is_disabled                     AS [Disabled],
    is_policy_checked               AS [Policy_Check],
    is_expiration_checked           AS [Expiration_Check]
FROM sys.sql_logins
WHERE name IN ('VaultSys', 'sa')
"@
$r_vaultsys = Invoke-SqlQuery -Connection $conn -Query $q_vaultsys -Label '        VaultSys and SA Login Status'

}
finally {
    # Close connection regardless of success or failure
    $conn.Close()
    Write-Ok "Connection closed."
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Build Inventory Object
# ─────────────────────────────────────────────────────────────────────────────

$inventory.Blocks['3_1_2_Version_Edition']  = $r_version
$inventory.Blocks['3_5_Server_Collation']   = $r_collation_server
$inventory.Blocks['3_6_DB_Collation']       = $r_collation_db
$inventory.Blocks['3_7_Auth_Mode']          = $r_auth
$inventory.Blocks['3_8_TCP_Listener']       = $r_tcp
$inventory.Blocks['3_9a_FTS_Installed']     = $r_fts_installed
$inventory.Blocks['3_9b_FTS_Catalogs']      = $r_fts_catalogs
$inventory.Blocks['3_11_SP_Configure']      = $r_sp_config
$inventory.Blocks['3_12_TempDB']            = $r_tempdb
$inventory.Blocks['DB_Files_All']           = $r_dbfiles
$inventory.Blocks['3_13_SQL_Agent_Jobs']    = $r_jobs
$inventory.Blocks['SQL_Logins']             = $r_logins
$inventory.Blocks['SQL_Services']           = $r_services
$inventory.Blocks['Backup_History']         = $r_backups
$inventory.Blocks['VaultSys_SA_Status']     = $r_vaultsys


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Write Text Report
# ─────────────────────────────────────────────────────────────────────────────

Write-Step "Writing report ..."

$report = [System.Text.StringBuilder]::new()

[void]$report.AppendLine('=' * 80)
[void]$report.AppendLine('  JM Consulting – Vault Migration 2024 → 2026')
[void]$report.AppendLine('  SQL Parameter Collection – Source Environment')
[void]$report.AppendLine('=' * 80)
[void]$report.AppendLine("  Collected at  : $($inventory.Meta.CollectedAt)")
[void]$report.AppendLine("  Collected on  : $($inventory.Meta.CollectedOn)")
[void]$report.AppendLine("  SQL Instance  : $($inventory.Meta.SqlInstance)")
[void]$report.AppendLine("  Auth Mode     : $($inventory.Meta.AuthMode)")
[void]$report.AppendLine('=' * 80)
[void]$report.AppendLine()

# Section 3.1/3.2 – Version
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.1/3.2 – SQL Server Version and Edition')
[void]$report.AppendLine('─' * 80)
if ($r_version) {
    [void]$report.AppendLine("  ProductVersion : $($r_version[0].ProductVersion)")
    [void]$report.AppendLine("  ProductLevel   : $($r_version[0].ProductLevel)")
    [void]$report.AppendLine("  Edition        : $($r_version[0].Edition)")
    [void]$report.AppendLine("  EngineEdition  : $($r_version[0].EngineEdition)")
    [void]$report.AppendLine("  ServerName     : $($r_version[0].ServerName)")
    [void]$report.AppendLine("  ComputerName   : $($r_version[0].ComputerName)")
    [void]$report.AppendLine()
    [void]$report.AppendLine("  SQL_FullVersion:")
    [void]$report.AppendLine("    $($r_version[0].SQL_FullVersion -replace "`n", "`n    ")")

    # Vault 2026 compatibility check
    $ver   = $r_version[0].ProductVersion
    $major = [int]($ver.Split('.')[0])
    [void]$report.AppendLine()
    if ($major -ge 15) {
        [void]$report.AppendLine('  ✓ SQL version is compatible with Vault 2026 (SQL 2019 = v15, SQL 2022 = v16)')
    } else {
        [void]$report.AppendLine('  ⚠ WARNING: SQL version may not be compatible with Vault 2026!')
        [void]$report.AppendLine('    Vault 2026 supports: SQL 2019 (min. CU29) and SQL 2022 (min. CU16)')
    }
}
[void]$report.AppendLine()

# Section 3.5 – Server Collation
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.5 – Server Collation')
[void]$report.AppendLine('─' * 80)
if ($r_collation_server) {
    [void]$report.AppendLine("  Server_Collation : $($r_collation_server[0].Server_Collation)")
    [void]$report.AppendLine()
    [void]$report.AppendLine('  ► This value MUST be configured identically on the target instance (SQL 2022)!')
}
[void]$report.AppendLine()

# Section 3.6 – DB Collation
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.6 – Collation and Recovery Model per Database')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_collation_db -Title 'Databases'))

# Section 3.7 – Auth Mode
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.7 – Authentication Mode')
[void]$report.AppendLine('─' * 80)
if ($r_auth) {
    [void]$report.AppendLine("  IsIntegratedSecurityOnly : $($r_auth[0].WindowsAuthOnly_Value)")
    [void]$report.AppendLine("  Assessment               : $($r_auth[0].Authentication_Mode)")
}
[void]$report.AppendLine()

# Section 3.8 – TCP Port
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.8 – TCP Listener and Port')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_tcp -Title 'TCP Listeners'))
[void]$report.AppendLine('  ► Verify in SQL Server Configuration Manager whether the TCP port is static or dynamic!')
[void]$report.AppendLine()

# Section 3.9 – Full-Text Search
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.9 – Full-Text Search')
[void]$report.AppendLine('─' * 80)
if ($r_fts_installed) {
    $ftsVal = $r_fts_installed[0].FTS_Installed
    [void]$report.AppendLine("  Installed : $ftsVal $(if ($ftsVal -eq 1) {'(Yes – must be installed on the target instance!)'} else {'(No)'})")
}
[void]$report.Append((Format-Table-Text -Data $r_fts_catalogs -Title 'Full-Text Catalogs'))

# Section 3.11 – sp_configure
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.11 – sp_configure (Memory, MAXDOP, miscellaneous parameters)')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_sp_config -Title 'sp_configure Values'))

# Section 3.12 – tempdb
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.12 – tempdb Configuration')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_tempdb -Title 'tempdb Files'))

# Database Files
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  Database Files (MDF/LDF for all databases)')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_dbfiles -Title 'Database Files'))

# SQL Agent Jobs
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  BLOCK 3.13 – SQL Agent Jobs')
[void]$report.AppendLine('─' * 80)
if (-not $r_jobs -or $r_jobs.Count -eq 0) {
    [void]$report.AppendLine('  (No jobs found or SQL Agent not active – normal for Express Edition)')
    [void]$report.AppendLine()
} else {
    [void]$report.Append((Format-Table-Text -Data $r_jobs -Title 'SQL Agent Jobs'))
}

# Logins
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  SQL Logins and Server Roles')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_logins -Title 'SQL Logins'))

# Services
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  SQL Server Services and Service Accounts')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_services -Title 'SQL Server Services'))

# Backup History
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  Backup History (last 20 entries)')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_backups -Title 'Backup History'))

# VaultSys / SA Status
[void]$report.AppendLine('─' * 80)
[void]$report.AppendLine('  VaultSys and SA Login Status')
[void]$report.AppendLine('─' * 80)
[void]$report.Append((Format-Table-Text -Data $r_vaultsys -Title 'VaultSys / SA'))

# Notes
[void]$report.AppendLine('=' * 80)
[void]$report.AppendLine('  NOTES AND NEXT STEPS')
[void]$report.AppendLine('=' * 80)
[void]$report.AppendLine()
[void]$report.AppendLine('  1) Collation of the target instance (SQL 2022) MUST be identical to the Server Collation above.')
[void]$report.AppendLine('  2) Authentication mode on target instance: Mixed Mode is mandatory (IsIntegratedSecurityOnly = 0).')
[void]$report.AppendLine('  3) Full-Text Search: If installed (value = 1), the feature must be explicitly included during')
[void]$report.AppendLine('     SQL 2022 setup.')
[void]$report.AppendLine('  4) Max Server Memory: Adjust on target VM (recommended: 70% of available RAM).')
[void]$report.AppendLine('  5) SA password and VaultSys password: Verify whether they are default or changed (section 4')
[void]$report.AppendLine('     of the data collection sheet). Passwords ONLY in the password vault.')
[void]$report.AppendLine('  6) TCP port of the AUTODESKVAULT instance: Verify in SQL Server Configuration Manager')
[void]$report.AppendLine('     whether static or dynamic – static is recommended for firewall rules.')
[void]$report.AppendLine('  7) Note the backup path from the backup history – relevant for downtime planning.')
[void]$report.AppendLine()
[void]$report.AppendLine('=' * 80)
[void]$report.AppendLine("  Report generated by: JM Consulting – Dr. Jörg Mayer")
[void]$report.AppendLine("  Confidential – for project team members only")
[void]$report.AppendLine('=' * 80)


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Write Output Files
# ─────────────────────────────────────────────────────────────────────────────

$report.ToString() | Out-File -FilePath $outputTxt -Encoding UTF8 -Force

# JSON – passwords are not included in the inventory, only query results
$inventory | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputJson -Encoding UTF8 -Force

Write-Ok "Text report : $outputTxt"
Write-Ok "JSON export : $outputJson"

Write-Host "`n  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║   Data collection complete. Files saved to Desktop.          ║" -ForegroundColor Green
Write-Host "  ╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Green

Read-Host "  Press ENTER to exit"
exit 0
