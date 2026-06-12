#Requires -Version 5.1
#Requires -RunAsAdministrator
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

<#
.SYNOPSIS
    Installs SQL Server unattended from setup.exe and applies all parameters captured by Get-SqlParameters.ps1.

.DESCRIPTION
    Three-phase automation for SQL Server migration / disaster recovery:

    Phase 1 - Unattended Setup:
      Reads collation, auth mode, FTS flag, and tempdb layout from the source
      snapshot JSON and maps them to setup.exe command-line arguments.

    Phase 2 - Post-install configuration:
      Applies sp_configure values (max memory, MAXDOP, cost threshold, etc.),
      tempdb autogrowth settings, static TCP port, and SA account state.
      Warns and caps max server memory if the target has less RAM than the source.

    Phase 3 - Verification:
      Connects to the new instance and compares critical settings against the
      snapshot. Exits 2 if any critical mismatch is detected.

    User databases are NOT restored — that is a separate operation.

.PARAMETER SetupExe
    Full path to setup.exe on the target server (e.g. D:\SQLSetup\setup.exe).

.PARAMETER SnapshotJson
    Path to the JSON file produced by Get-SqlParameters.ps1 on the source server.

.PARAMETER SaPassword
    SA account password as SecureString.
    Prompt: (Read-Host -AsSecureString 'SA Password')

.PARAMETER InstanceName
    SQL Server instance name. Default: MSSQLSERVER (default instance).

.PARAMETER SqlSvcAccount
    Service account for SQL Server Engine. Default: NT AUTHORITY\SYSTEM.

.PARAMETER SqlSvcPassword
    Password for SqlSvcAccount as SecureString. Required for domain accounts.

.PARAMETER AgentSvcAccount
    Service account for SQL Agent. Default: NT AUTHORITY\SYSTEM.

.PARAMETER DataDir
    Root directory for SQL data files. Default: C:\SQLData.

.PARAMETER LogDir
    Root directory for SQL log files. Default: C:\SQLLogs.

.PARAMETER TempDir
    Root directory for tempdb files. Default: C:\SQLTemp.

.PARAMETER BackupDir
    Root directory for SQL backups. Default: C:\SQLBackup.

.PARAMETER SysAdminAccounts
    Windows account(s) for sysadmin role. Default: BUILTIN\Administrators.

.PARAMETER SkipVerification
    Skips the post-install verification phase.

.PARAMETER LogFile
    Path for the installation log. Default: script directory with timestamp.

.EXAMPLE
    PS C:\> $pw = Read-Host -AsSecureString 'SA Password'
    PS C:\> .\Install-SqlFromSnapshot.ps1 -SetupExe 'D:\Setup\setup.exe' -SnapshotJson 'C:\Temp\sql_params.json' -SaPassword $pw -WhatIf
    Shows what would be executed without making any changes.

.EXAMPLE
    PS C:\> $pw = Read-Host -AsSecureString 'SA Password'
    PS C:\> .\Install-SqlFromSnapshot.ps1 -SetupExe 'D:\Setup\setup.exe' -SnapshotJson 'C:\Temp\sql_params.json' -SaPassword $pw -DataDir 'E:\Data' -LogDir 'F:\Logs' -BackupDir 'G:\Backup'
    Full installation with custom disk layout matching the target disk structure.

.NOTES
    V1.0/Created - 2026-06-12 - Initial version

    Author: (c) JM and Claude Code

    Required Modules: none (uses System.Data.SqlClient for post-install T-SQL)

    Required permissions:
      - Local Administrator on the target server (#Requires -RunAsAdministrator)
      - Read access to -SetupExe and -SnapshotJson paths

    Output files:
      Default log: $PSScriptRoot\Install-SqlFromSnapshot_<yyyyMMdd_HHmmss>.log
      SQL Server setup log: %ProgramFiles%\Microsoft SQL Server\...\Setup Bootstrap\Log\

    Exit codes:
      0  Full success
      1  Pre-condition failure (bad paths, existing instance, malformed JSON)
      2  Core operation error (setup failure or post-install mismatch)
      3  Unexpected / unhandled exception
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SetupExe,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SnapshotJson,

    [Parameter(Mandatory)]
    [SecureString]$SaPassword,

    [string]$InstanceName     = 'MSSQLSERVER',
    [string]$SqlSvcAccount    = 'NT AUTHORITY\SYSTEM',
    [SecureString]$SqlSvcPassword,
    [string]$AgentSvcAccount  = 'NT AUTHORITY\SYSTEM',

    [string]$DataDir          = 'C:\SQLData',
    [string]$LogDir           = 'C:\SQLLogs',
    [string]$TempDir          = 'C:\SQLTemp',
    [string]$BackupDir        = 'C:\SQLBackup',

    [string]$SysAdminAccounts = 'BUILTIN\Administrators',

    [switch]$SkipVerification,

    [string]$LogFile = (Join-Path $PSScriptRoot "Install-SqlFromSnapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').log")
)


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Functions
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'CRITICAL')][string]$Level = 'INFO'
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "$ts [$Level] $Message"
    $line | Out-File -FilePath $LogFile -Append -Encoding UTF8 -WhatIf:$false
    Write-Verbose $line
    if ($Level -in 'WARN', 'ERROR', 'CRITICAL') { Write-Warning $Message }
    else { Write-Host "  $line" }
}

function ConvertTo-PlainText {
    param([Parameter(Mandatory)][SecureString]$SecureString)
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($SecureString)
    try   { return [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ptr) }
}

function Get-SnapBlock {
    param([Parameter(Mandatory)]$Snapshot, [Parameter(Mandatory)][string]$Key)
    $block = $Snapshot.Blocks.$Key
    if ($null -eq $block) {
        Write-Log "Snapshot block '$Key' not found — defaulting to empty." 'WARN'
        return @()
    }
    return @($block)
}

function Invoke-SqlNonQuery {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$Query,
        [string]$Label = ''
    )
    $cmd = $null
    try {
        $cmd                = $Connection.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = 300
        [void]$cmd.ExecuteNonQuery()
        if ($Label) { Write-Log "Applied: $Label" }
    }
    finally { if ($null -ne $cmd) { $cmd.Dispose() } }
}

function Invoke-SqlScalar {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][string]$Query
    )
    $cmd = $null
    try {
        $cmd                = $Connection.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = 30
        return $cmd.ExecuteScalar()
    }
    finally { if ($null -ne $cmd) { $cmd.Dispose() } }
}

function Get-SqlRegistryRoot {
    param([string]$Instance = 'MSSQLSERVER')
    $nameMap = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL' -ErrorAction Stop
    $key     = $nameMap.$Instance
    if (-not $key) { throw "Instance '$Instance' not found in registry — is SQL Server installed?" }
    return "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$key"
}

function Set-TcpStaticPort {
    param([string]$Instance, [int]$Port)
    $root   = Get-SqlRegistryRoot -Instance $Instance
    $ipAll  = "$root\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"
    Set-ItemProperty -Path $ipAll -Name 'TcpPort'        -Value "$Port" -Type String
    Set-ItemProperty -Path $ipAll -Name 'TcpDynamicPorts' -Value ''     -Type String
    Write-Log "TCP static port configured: $Port"
}

function Wait-SqlServiceReady {
    param([string]$Instance, [int]$TimeoutSeconds = 180)
    $svcName = if ($Instance -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$Instance" }
    Write-Log "Waiting for service '$svcName' to reach Running state..."
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') { Write-Log "Service is Running."; return }
        Start-Sleep -Seconds 5
    }
    throw "SQL Server service '$svcName' did not start within $TimeoutSeconds seconds."
}

function Restart-SqlService {
    param([string]$Instance)
    $svcName = if ($Instance -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$Instance" }
    Write-Log "Restarting '$svcName' to apply settings..."
    Restart-Service -Name $svcName -Force -ErrorAction Stop
    Wait-SqlServiceReady -Instance $Instance
}

function Open-SqlConnectionWinAuth {
    param([string]$SqlInstance)
    $conn = New-Object System.Data.SqlClient.SqlConnection(
        "Server=$SqlInstance;Database=master;Integrated Security=True;Connection Timeout=30;")
    $conn.Open()
    return $conn
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Load and validate snapshot
# ─────────────────────────────────────────────────────────────────────────────

Write-Log "=== Install-SqlFromSnapshot v1.0 starting ==="
Write-Log "Snapshot : $SnapshotJson"
Write-Log "Setup    : $SetupExe"
Write-Log "Instance : $InstanceName"

try {
    $snap = Get-Content $SnapshotJson -Raw -Encoding UTF8 | ConvertFrom-Json
}
catch {
    Write-Log "Failed to parse snapshot JSON: $($_.Exception.Message)" 'CRITICAL'
    exit 1
}

# Extract values — use safe defaults when blocks are absent
$collationRows  = Get-SnapBlock $snap '3_5_Server_Collation'
$authRows       = Get-SnapBlock $snap '3_7_Auth_Mode'
$tcpRows        = Get-SnapBlock $snap '3_8_TCP_Listener'
$ftsRows        = Get-SnapBlock $snap '3_9a_FTS_Installed'
$spConfigRows   = Get-SnapBlock $snap '3_11_SP_Configure'
$tempdbRows     = Get-SnapBlock $snap '3_12_TempDB'
$vaultSysRows   = Get-SnapBlock $snap 'VaultSys_SA_Status'

$collation    = if ($collationRows.Count -gt 0) { $collationRows[0].Server_Collation } else { 'SQL_Latin1_General_CP1_CI_AS' }
$winAuthOnly  = if ($authRows.Count -gt 0)      { [int]$authRows[0].WindowsAuthOnly_Value } else { 0 }
$ftsInstalled = if ($ftsRows.Count -gt 0)       { [int]$ftsRows[0].FTS_Installed } else { 0 }
$mixedMode    = ($winAuthOnly -eq 0)

$tempdbDataFiles = @($tempdbRows | Where-Object { $_.Type -eq 'ROWS' })
$tempdbLogFile   = $tempdbRows | Where-Object { $_.Type -eq 'LOG' } | Select-Object -First 1
$tempdbFileCount = [Math]::Max(1, $tempdbDataFiles.Count)
$tempdbFileSize  = if ($tempdbDataFiles.Count -gt 0) { [Math]::Max(8, [int]$tempdbDataFiles[0].SizeMB) } else { 8 }
$tempdbLogSize   = if ($null -ne $tempdbLogFile)     { [Math]::Max(8, [int]$tempdbLogFile.SizeMB)      } else { 8 }

$tcpPortSource = ($tcpRows | Where-Object { [int]$_.TCP_Port -gt 0 } | Select-Object -First 1).TCP_Port
$tcpPortSource = if ($tcpPortSource) { [int]$tcpPortSource } else { 1433 }

# SA status on source (is_disabled = 0 means active)
$saRow      = $vaultSysRows | Where-Object { $_.Login -eq 'sa' } | Select-Object -First 1
$saEnabled  = ($null -ne $saRow -and [int]$saRow.Disabled -eq 0)

Write-Log "Source collation : $collation"
Write-Log "Mixed Mode       : $mixedMode"
Write-Log "FTS installed    : $($ftsInstalled -eq 1)"
Write-Log "TCP port         : $tcpPortSource"
Write-Log "tempdb data files: $tempdbFileCount x ${tempdbFileSize} MB"
Write-Log "SA active on src : $saEnabled"


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Pre-flight checks
# ─────────────────────────────────────────────────────────────────────────────

# Check for existing instance
$existingKey = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'
if (Test-Path $existingKey) {
    $existing = (Get-ItemProperty $existingKey -ErrorAction SilentlyContinue).$InstanceName
    if ($existing) {
        Write-Log "SQL Server instance '$InstanceName' already exists (registry key: $existing)." 'ERROR'
        Write-Log "Use a different -InstanceName or uninstall the existing instance first." 'ERROR'
        exit 1
    }
}

# Create directories (idempotent)
foreach ($dir in @($DataDir, $LogDir, $TempDir, $BackupDir)) {
    if (-not (Test-Path $dir)) {
        if ($PSCmdlet.ShouldProcess($dir, 'Create directory')) {
            New-Item $dir -ItemType Directory -Force | Out-Null
            Write-Log "Created directory: $dir"
        }
    }
}

# Build features list
$features = 'SQLENGINE,REPLICATION'
if ($ftsInstalled -eq 1) { $features += ',FULLTEXT' }
Write-Log "Features         : $features"


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Phase 1 — Unattended SQL Server installation
# ─────────────────────────────────────────────────────────────────────────────

Write-Log "--- Phase 1: Unattended installation ---"

$saPlain   = $null
$svcPlain  = $null

try {
    $saPlain = ConvertTo-PlainText $SaPassword

    $setupArgs = [System.Collections.Generic.List[string]]@(
        '/Q'
        '/ACTION=Install'
        "/FEATURES=$features"
        "/INSTANCENAME=$InstanceName"
        "/SQLCOLLATION=$collation"
        '/TCPENABLED=1'
        "/SQLUSERDBDIR=$DataDir"
        "/SQLUSERDBLOGDIR=$LogDir"
        "/SQLTEMPDBDIR=$TempDir"
        "/SQLBACKUPDIR=$BackupDir"
        "/SQLSYSADMINACCOUNTS=$SysAdminAccounts"
        "/SQLTEMPDBFILECOUNT=$tempdbFileCount"
        "/SQLTEMPDBFILESIZE=$tempdbFileSize"
        "/SQLTEMPDBLOGFILESIZE=$tempdbLogSize"
        "/SQLSVCACCOUNT=$SqlSvcAccount"
        "/AGTSVCACCOUNT=$AgentSvcAccount"
        '/AGTSVCSTARTUPTYPE=Automatic'
        '/SQLSVCSTARTUPTYPE=Automatic'
        '/BROWSERSVCSTARTUPTYPE=Disabled'
        '/IACCEPTSQLSERVERLICENSETERMS'   # SQL Server 2022 / 2025
    )

    if ($mixedMode) {
        $setupArgs.Add('/SECURITYMODE=SQL')
        $setupArgs.Add("/SAPWD=$saPlain")  # plain text required by setup.exe; not logged
    }

    if ($SqlSvcPassword) {
        $svcPlain = ConvertTo-PlainText $SqlSvcPassword
        $setupArgs.Add("/SQLSVCPASSWORD=$svcPlain")
    }

    Write-Log "setup.exe arguments built (passwords omitted from log)."

    if ($PSCmdlet.ShouldProcess("SQL Server '$InstanceName' on $($env:COMPUTERNAME)", 'Run unattended installation')) {
        Write-Log "Launching setup.exe — this takes 5-15 minutes..."
        $proc = Start-Process -FilePath $SetupExe -ArgumentList ($setupArgs.ToArray()) -Wait -PassThru -NoNewWindow

        if ($proc.ExitCode -ne 0) {
            Write-Log "setup.exe exited with code $($proc.ExitCode)." 'CRITICAL'
            Write-Log "Review: $env:ProgramFiles\Microsoft SQL Server\*\Setup Bootstrap\Log\Summary.txt" 'CRITICAL'
            exit 2
        }
        Write-Log "SQL Server installation completed. Exit code: 0"
    }
    else {
        Write-Log "WhatIf: would run setup.exe with $($setupArgs.Count) arguments."
        Write-Log "WhatIf: features=$features | collation=$collation | tempdb=$tempdbFileCount x ${tempdbFileSize}MB"
    }
}
finally {
    $saPlain  = $null
    $svcPlain = $null
    [System.GC]::Collect()
}

if ($WhatIfPreference) {
    Write-Log "WhatIf mode: skipping post-install phases."
    exit 0
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Phase 2 — Post-install configuration
# ─────────────────────────────────────────────────────────────────────────────

Write-Log "--- Phase 2: Post-install configuration ---"

Wait-SqlServiceReady -Instance $InstanceName

$sqlConnStr = if ($InstanceName -eq 'MSSQLSERVER') { '.' } else { ".\$InstanceName" }
$restartNeeded = $false
$conn = $null

try {
    $conn = Open-SqlConnectionWinAuth -SqlInstance $sqlConnStr
    Write-Log "Connected to $sqlConnStr (Windows Auth)."

    # SA: set password hash and enable if source had SA active
    if ($mixedMode) {
        $saPlain3 = ConvertTo-PlainText $SaPassword
        try {
            $escapedPw = $saPlain3.Replace("'", "''")
            Invoke-SqlNonQuery $conn "ALTER LOGIN [sa] WITH PASSWORD = N'$escapedPw', CHECK_EXPIRATION = OFF, CHECK_POLICY = OFF;" 'SA password set'
            if ($saEnabled) {
                Invoke-SqlNonQuery $conn 'ALTER LOGIN [sa] ENABLE;' 'SA login enabled'
            }
        }
        finally { $saPlain3 = $null; [System.GC]::Collect() }
    }

    # sp_configure — enable advanced options first, then apply all captured values
    Invoke-SqlNonQuery $conn "EXEC sp_configure 'show advanced options', 1; RECONFIGURE WITH OVERRIDE;" 'Show advanced options'

    # Check available RAM to guard against memory overcommit
    $targetRamMB = [int]((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB)

    foreach ($row in $spConfigRows) {
        $paramName = $row.Parameter
        $configVal = [long]$row.Configured

        if ($paramName -eq 'max server memory (MB)' -and $configVal -gt $targetRamMB) {
            $cappedVal = [int]($targetRamMB * 0.85)
            Write-Log "Source max memory ($configVal MB) > target RAM ($targetRamMB MB). Capping to $cappedVal MB." 'WARN'
            $configVal = $cappedVal
        }

        Invoke-SqlNonQuery $conn "EXEC sp_configure N'$paramName', $configVal; RECONFIGURE WITH OVERRIDE;" "'$paramName' = $configVal"

        # remote access requires service restart to take effect
        if ($paramName -eq 'remote access') { $restartNeeded = $true }
    }

    # tempdb autogrowth — apply per-file settings from snapshot
    foreach ($file in $tempdbDataFiles) {
        $growthSpec = if ([int]$file.Autogrowth_Percent -eq 1) {
            "$([int]$file.Autogrowth_Raw)%"
        } else {
            "$([Math]::Max(64, [int]$file.Autogrowth_Raw * 8 / 1024))MB"
        }
        try {
            Invoke-SqlNonQuery $conn "ALTER DATABASE tempdb MODIFY FILE (NAME = N'$($file.File_Name)', FILEGROWTH = $growthSpec);" "tempdb '$($file.File_Name)' autogrowth=$growthSpec"
        }
        catch { Write-Log "tempdb autogrowth skipped for '$($file.File_Name)' (file may use a different name after setup): $($_.Exception.Message)" 'WARN' }
    }

    if ($null -ne $tempdbLogFile) {
        $logGrowth = if ([int]$tempdbLogFile.Autogrowth_Percent -eq 1) {
            "$([int]$tempdbLogFile.Autogrowth_Raw)%"
        } else {
            "$([Math]::Max(64, [int]$tempdbLogFile.Autogrowth_Raw * 8 / 1024))MB"
        }
        try {
            Invoke-SqlNonQuery $conn "ALTER DATABASE tempdb MODIFY FILE (NAME = N'$($tempdbLogFile.File_Name)', FILEGROWTH = $logGrowth);" "tempdb log autogrowth=$logGrowth"
        }
        catch { Write-Log "tempdb log autogrowth skipped: $($_.Exception.Message)" 'WARN' }
    }

    $conn.Close()
}
finally {
    if ($null -ne $conn -and $conn.State -ne 'Closed') { $conn.Close() }
    $conn = $null
}

# TCP static port — configure via registry if source used non-default or non-dynamic port
if ($tcpPortSource -ne 1433) {
    Write-Log "Source used non-default TCP port $tcpPortSource — applying via registry."
    Set-TcpStaticPort -Instance $InstanceName -Port $tcpPortSource
    $restartNeeded = $true
}

if ($restartNeeded) {
    Write-Log "Service restart required for pending settings."
    Restart-SqlService -Instance $InstanceName
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Phase 3 — Verification
# ─────────────────────────────────────────────────────────────────────────────

if ($SkipVerification) {
    Write-Log "Verification skipped (-SkipVerification)." 'WARN'
    Write-Log "=== Installation complete (unverified) ==="
    exit 0
}

Write-Log "--- Phase 3: Verification ---"

$errors = 0
$conn   = $null

try {
    $conn = Open-SqlConnectionWinAuth -SqlInstance $sqlConnStr

    # Collation
    $actualCollation = Invoke-SqlScalar $conn "SELECT CAST(SERVERPROPERTY('Collation') AS NVARCHAR)"
    if ($actualCollation -ne $collation) {
        Write-Log "MISMATCH collation: expected '$collation', got '$actualCollation'" 'ERROR'
        $errors++
    } else { Write-Log "OK  collation: $actualCollation" }

    # Auth mode
    $actualWinOnly = [int](Invoke-SqlScalar $conn "SELECT CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS INT)")
    if ($actualWinOnly -ne $winAuthOnly) {
        Write-Log "MISMATCH auth mode: expected WindowsAuthOnly=$winAuthOnly, got $actualWinOnly" 'ERROR'
        $errors++
    } else { Write-Log "OK  auth mode: WindowsAuthOnly=$actualWinOnly" }

    # tempdb file count
    $actualTempDbCount = [int](Invoke-SqlScalar $conn "SELECT COUNT(*) FROM tempdb.sys.master_files WHERE type = 0")
    if ($actualTempDbCount -ne $tempdbFileCount) {
        Write-Log "MISMATCH tempdb data files: expected $tempdbFileCount, got $actualTempDbCount" 'WARN'
    } else { Write-Log "OK  tempdb data files: $actualTempDbCount" }

    # Key sp_configure values
    foreach ($row in $spConfigRows) {
        $pName  = $row.Parameter
        $pExpect = [long]$row.Configured

        # Re-apply memory cap for verification comparison
        if ($pName -eq 'max server memory (MB)' -and $pExpect -gt $targetRamMB) {
            $pExpect = [int]($targetRamMB * 0.85)
        }

        $actualVal = [long](Invoke-SqlScalar $conn "SELECT value FROM sys.configurations WHERE name = N'$pName'")
        if ($actualVal -ne $pExpect) {
            Write-Log "MISMATCH sp_configure '$pName': expected $pExpect, got $actualVal" 'WARN'
        } else { Write-Log "OK  sp_configure '$pName' = $actualVal" }
    }

    $conn.Close()
}
catch {
    Write-Log "Verification connection failed: $($_.Exception.Message)" 'ERROR'
    $errors++
}
finally {
    if ($null -ne $conn -and $conn.State -ne 'Closed') { $conn.Close() }
}

# Summary
Write-Log "=== Installation summary ==="
Write-Log "Instance   : $InstanceName on $($env:COMPUTERNAME)"
Write-Log "Collation  : $collation"
Write-Log "Mixed Mode : $mixedMode"
Write-Log "TCP port   : $tcpPortSource"
Write-Log "Features   : $features"
Write-Log "Log file   : $LogFile"

if ($errors -gt 0) {
    Write-Log "$errors verification check(s) failed — review log for details." 'ERROR'
    exit 2
}

Write-Log "All verification checks passed."
exit 0
