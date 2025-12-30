#requires -Version 5.1
<#
Bepoz Toolkit Tool: Create Store + Assign to Venue
- Safe by default (DryRun = true)
- Writes Report.json to the RunDir
- Uses Integrated Security to connect to SQL Server
- Attempts to auto-discover a Bepoz database containing dbo.Venue + dbo.Store if Database not provided

Config precedence:
1) Explicit parameters
2) Config JSON (default: CreateStore.config.json in RunDir, or -ConfigPath)
3) Environment variables:
   BEPOZ_SQLSERVER, BEPOZ_DATABASE, BEPOZ_VENUEID, BEPOZ_STORENAME, BEPOZ_STOREGROUP, BEPOZ_DRYRUN, BEPOZ_FORCE
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

param(
    [string]$ConfigPath,

    [string]$SqlServer,
    [string]$Database,

    [int]$VenueId,
    [string]$StoreName,
    [int]$StoreGroup,

    [Nullable[bool]]$DryRun,
    [switch]$Force
)

# ----------------------------
# Paths + report scaffolding
# ----------------------------
$RunDir = $PSScriptRoot
$ReportPath = Join-Path $RunDir "Report.json"

function Write-Log {
    param(
        [ValidateSet("INFO","WARN","ERROR","DEBUG")]
        [string]$Level,
        [string]$Message
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts][$Level] $Message"
}

function Save-Report {
    param([hashtable]$Report)
    try {
        $ReportJson = ($Report | ConvertTo-Json -Depth 10)
        $ReportJson | Set-Content -Path $ReportPath -Encoding UTF8
        Write-Log INFO "Wrote report: $ReportPath"
    } catch {
        Write-Log ERROR "Failed to write Report.json: $($_.Exception.Message)"
    }
}

function Test-IsInteractive {
    # Conservative: only treat as interactive if console host and UserInteractive.
    try {
        if (-not [Environment]::UserInteractive) { return $false }
        if ($Host -and $Host.Name -eq "ConsoleHost") { return $true }
        return $false
    } catch {
        return $false
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not $Path) { return $null }
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return ($raw | ConvertFrom-Json -ErrorAction Stop)
}

function Get-EnvBool {
    param([string]$Name)
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    switch -Regex ($v.Trim().ToLowerInvariant()) {
        '^(1|true|yes|y)$' { return $true }
        '^(0|false|no|n)$' { return $false }
        default { return $null }
    }
}

function Get-EnvInt {
    param([string]$Name)
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    $out = 0
    if ([int]::TryParse($v.Trim(), [ref]$out)) { return $out }
    return $null
}

function Get-EnvString {
    param([string]$Name)
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    return $v.Trim()
}

function New-SqlConnection {
    param(
        [string]$Server,
        [string]$Db
    )
    if ([string]::IsNullOrWhiteSpace($Server)) {
        throw "SqlServer is empty."
    }
    if ([string]::IsNullOrWhiteSpace($Db)) {
        throw "Database is empty."
    }

    Add-Type -AssemblyName System.Data

    $cs = "Server=$Server;Database=$Db;Integrated Security=SSPI;Connect Timeout=8;Application Name=BepozToolkit-CreateStore;"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    return $conn
}

function Invoke-SqlDataTable {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql,
        [hashtable]$Parameters
    )
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = 30

    if ($Parameters) {
        foreach ($k in $Parameters.Keys) {
            $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::Variant)
            $p.Value = $Parameters[$k]
        }
    }

    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return $dt
}

function Invoke-SqlScalar {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Sql,
        [hashtable]$Parameters
    )
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = 30

    if ($Parameters) {
        foreach ($k in $Parameters.Keys) {
            $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::Variant)
            $p.Value = $Parameters[$k]
        }
    }
    return $cmd.ExecuteScalar()
}

function Test-DbHasVenueAndStoreTables {
    param(
        [string]$Server,
        [string]$DbName
    )
    try {
        $c = New-SqlConnection -Server $Server -Db $DbName
        try {
            $sql = @"
SELECT
  CASE
    WHEN EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE s.name='dbo' AND t.name='Venue')
     AND EXISTS (SELECT 1 FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE s.name='dbo' AND t.name='Store')
    THEN 1 ELSE 0
  END;
"@
            $ok = Invoke-SqlScalar -Connection $c -Sql $sql -Parameters @{}
            return ([int]$ok -eq 1)
        } finally {
            $c.Close()
            $c.Dispose()
        }
    } catch {
        return $false
    }
}

function Find-BepozDatabases {
    param([string]$Server)

    $dbs = @()
    $master = New-SqlConnection -Server $Server -Db "master"
    try {
        $dt = Invoke-SqlDataTable -Connection $master -Sql "SELECT name FROM sys.databases WHERE state = 0 AND database_id > 4 ORDER BY name;" -Parameters @{}
        foreach ($row in $dt.Rows) {
            $name = [string]$row["name"]
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if (Test-DbHasVenueAndStoreTables -Server $Server -DbName $name) {
                $dbs += $name
            }
        }
    } finally {
        $master.Close()
        $master.Dispose()
    }
    return $dbs
}

function Get-NonNullableColumnsWithoutDefault {
    param([System.Data.SqlClient.SqlConnection]$Conn)

    $sql = @"
SELECT
  c.name AS ColumnName,
  t.name AS TypeName,
  c.max_length AS MaxLength,
  c.precision AS [Precision],
  c.scale AS [Scale],
  c.is_identity AS IsIdentity,
  c.is_computed AS IsComputed,
  c.is_nullable AS IsNullable,
  dc.definition AS DefaultDefinition
FROM sys.columns c
JOIN sys.types t ON t.user_type_id = c.user_type_id
LEFT JOIN sys.default_constraints dc ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
WHERE c.object_id = OBJECT_ID('dbo.Store')
ORDER BY c.column_id;
"@
    $dt = Invoke-SqlDataTable -Connection $Conn -Sql $sql -Parameters @{}

    $required = @()
    foreach ($r in $dt.Rows) {
        $isIdentity = [bool]$r["IsIdentity"]
        $isComputed = [bool]$r["IsComputed"]
        $isNullable = [bool]$r["IsNullable"]
        $defaultDef = $r["DefaultDefinition"]

        if ($isIdentity -or $isComputed) { continue }
        if ($isNullable) { continue }
        if ($null -ne $defaultDef -and -not [string]::IsNullOrWhiteSpace([string]$defaultDef)) { continue }

        $required += [pscustomobject]@{
            ColumnName = [string]$r["ColumnName"]
            TypeName   = [string]$r["TypeName"]
            MaxLength  = [int]$r["MaxLength"]
            Precision  = [int]$r["Precision"]
            Scale      = [int]$r["Scale"]
        }
    }
    return $required
}

function Get-FallbackValueForType {
    param(
        [string]$TypeName
    )

    switch ($TypeName.ToLowerInvariant()) {
        "bit" { return 0 }
        "tinyint" { return 0 }
        "smallint" { return 0 }
        "int" { return 0 }
        "bigint" { return 0 }
        "decimal" { return 0 }
        "numeric" { return 0 }
        "money" { return 0 }
        "smallmoney" { return 0 }
        "float" { return 0 }
        "real" { return 0 }
        "date" { return (Get-Date).Date }
        "datetime" { return Get-Date }
        "datetime2" { return Get-Date }
        "smalldatetime" { return Get-Date }
        "time" { return [TimeSpan]::Zero }
        default { return "" } # strings + anything unknown
    }
}

function Build-InsertPlan {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [int]$VenueIdValue,
        [string]$StoreNameValue,
        [Nullable[int]]$StoreGroupValue,
        [hashtable]$Overrides
    )

    # Determine required columns in dbo.Store that do NOT have defaults.
    $required = Get-NonNullableColumnsWithoutDefault -Conn $Conn

    # Always set these explicitly if present in table.
    $values = @{}
    $values["VenueID"] = $VenueIdValue
    $values["Name"]    = $StoreNameValue

    if ($StoreGroupValue -ne $null) {
        $values["StoreGroup"] = [int]$StoreGroupValue
    }

    if ($Overrides) {
        foreach ($k in $Overrides.Keys) {
            $values[$k] = $Overrides[$k]
        }
    }

    # Ensure all required columns have a value.
    foreach ($col in $required) {
        $cn = $col.ColumnName
        if (-not $values.ContainsKey($cn)) {
            # Provide a safe fallback to satisfy NOT NULL.
            $values[$cn] = Get-FallbackValueForType -TypeName $col.TypeName
        }
    }

    # Ensure we only insert into columns that exist in dbo.Store
    # (in case overrides include unknown keys)
    $existingSql = "SELECT c.name FROM sys.columns c WHERE c.object_id = OBJECT_ID('dbo.Store');"
    $existingDt = Invoke-SqlDataTable -Connection $Conn -Sql $existingSql -Parameters @{}
    $existing = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($r in $existingDt.Rows) { [void]$existing.Add([string]$r["name"]) }

    $final = @{}
    foreach ($k in $values.Keys) {
        if ($existing.Contains($k)) {
            $final[$k] = $values[$k]
        }
    }

    $cols = @($final.Keys) | Sort-Object
    if ($cols.Count -eq 0) { throw "Insert plan produced no columns. Cannot continue." }

    $colList = ($cols | ForEach-Object { "[" + $_ + "]" }) -join ", "
    $paramList = ($cols | ForEach-Object { "@" + $_ }) -join ", "

    $sql = "INSERT INTO dbo.Store ($colList) VALUES ($paramList); SELECT CAST(SCOPE_IDENTITY() AS int) AS NewStoreID;"
    return [pscustomobject]@{
        Sql = $sql
        Values = $final
        Columns = $cols
        RequiredColumnsWithoutDefaults = $required
    }
}

function Confirm-Apply {
    param([string]$Message)

    if ($Force.IsPresent) { return $true }

    if (-not (Test-IsInteractive)) { return $false }

    Write-Host ""
    Write-Host $Message
    Write-Host "Type CREATE to proceed, anything else to cancel: " -NoNewline
    $resp = Read-Host
    return ($resp -eq "CREATE")
}

# ----------------------------
# Main
# ----------------------------
$report = @{
    toolId     = "store-create-assign-venue"
    startedAt  = (Get-Date).ToString("o")
    runDir     = $RunDir
    inputs     = @{}
    resolved   = @{}
    discovery  = @{}
    checks     = @()
    actions    = @()
    result     = @{
        status = "Unknown"
        message = ""
        storeId = $null
        database = $null
        venueId = $null
        storeName = $null
    }
}

try {
    Write-Log INFO "Create Store + Assign to Venue starting..."

    # Resolve config file
    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        $ConfigPath = Join-Path $RunDir "CreateStore.config.json"
    }
    $cfg = Read-JsonFile -Path $ConfigPath
    if ($cfg) {
        Write-Log INFO "Loaded config: $ConfigPath"
    } else {
        Write-Log INFO "No config loaded (looked for: $ConfigPath)"
    }

    # Environment values
    $envSqlServer  = Get-EnvString "BEPOZ_SQLSERVER"
    $envDatabase   = Get-EnvString "BEPOZ_DATABASE"
    $envVenueId    = Get-EnvInt "BEPOZ_VENUEID"
    $envStoreName  = Get-EnvString "BEPOZ_STORENAME"
    $envStoreGroup = Get-EnvInt "BEPOZ_STOREGROUP"
    $envDryRun     = Get-EnvBool "BEPOZ_DRYRUN"
    $envForce      = Get-EnvBool "BEPOZ_FORCE"

    # Apply precedence: params > config > env
    if ([string]::IsNullOrWhiteSpace($SqlServer)) {
        if ($cfg -and $cfg.SqlServer) { $SqlServer = [string]$cfg.SqlServer }
        elseif ($envSqlServer) { $SqlServer = $envSqlServer }
        else { $SqlServer = ".\SQLEXPRESS" } # sensible default for many installs
    }

    if ([string]::IsNullOrWhiteSpace($Database)) {
        if ($cfg -and $cfg.Database) { $Database = [string]$cfg.Database }
        elseif ($envDatabase) { $Database = $envDatabase }
    }

    if (-not $PSBoundParameters.ContainsKey("VenueId") -or $VenueId -eq 0) {
        if ($cfg -and $cfg.VenueId) { $VenueId = [int]$cfg.VenueId }
        elseif ($envVenueId -ne $null) { $VenueId = [int]$envVenueId }
    }

    if ([string]::IsNullOrWhiteSpace($StoreName)) {
        if ($cfg -and $cfg.StoreName) { $StoreName = [string]$cfg.StoreName }
        elseif ($envStoreName) { $StoreName = $envStoreName }
    }

    if (-not $PSBoundParameters.ContainsKey("StoreGroup")) {
        if ($cfg -and $cfg.StoreGroup -ne $null) { $StoreGroup = [int]$cfg.StoreGroup }
        elseif ($envStoreGroup -ne $null) { $StoreGroup = [int]$envStoreGroup }
    }

    if ($DryRun -eq $null) {
        if ($cfg -and $cfg.DryRun -ne $null) { $DryRun = [bool]$cfg.DryRun }
        elseif ($envDryRun -ne $null) { $DryRun = [bool]$envDryRun }
        else { $DryRun = $true } # SAFE BY DEFAULT
    }

    if (-not $Force.IsPresent) {
        if ($cfg -and $cfg.Force -eq $true) { $Force = $true }
        elseif ($envForce -eq $true) { $Force = $true }
    }

    $overrides = $null
    if ($cfg -and $cfg.StoreOverrides) {
        $overrides = @{}
        foreach ($p in $cfg.StoreOverrides.PSObject.Properties) {
            $overrides[$p.Name] = $p.Value
        }
    }

    $report.inputs = @{
        configPath = $ConfigPath
    }
    $report.resolved = @{
        sqlServer = $SqlServer
        database  = $Database
        venueId   = $VenueId
        storeName = $StoreName
        storeGroup = $StoreGroup
        dryRun    = $DryRun
        force     = [bool]$Force.IsPresent
        hasOverrides = [bool]($overrides -ne $null -and $overrides.Count -gt 0)
    }

    Write-Log INFO "SQL Server: $SqlServer"
    Write-Log INFO ("DryRun: " + $DryRun.ToString() + " (Force: " + ([bool]$Force.IsPresent).ToString() + ")")

    # Auto-discover database if missing
    if ([string]::IsNullOrWhiteSpace($Database)) {
        Write-Log WARN "Database not provided. Attempting to discover Bepoz databases (dbo.Venue + dbo.Store)..."
        $candidates = Find-BepozDatabases -Server $SqlServer
        $report.discovery = @{
            candidates = $candidates
        }

        if ($candidates.Count -eq 0) {
            throw "Could not discover any databases containing dbo.Venue and dbo.Store. Provide -Database or set BEPOZ_DATABASE."
        }

        if ($candidates.Count -eq 1) {
            $Database = $candidates[0]
            Write-Log INFO "Discovered database: $Database"
        } else {
            # Multiple candidates; require selection if interactive, else fail safely.
            if (Test-IsInteractive) {
                Write-Host ""
                Write-Host "Multiple candidate databases found:"
                for ($i=0; $i -lt $candidates.Count; $i++) {
                    Write-Host ("[{0}] {1}" -f $i, $candidates[$i])
                }
                $idx = Read-Host "Enter the index of the database to use"
                $n = 0
                if (-not [int]::TryParse($idx, [ref]$n)) { throw "Invalid index." }
                if ($n -lt 0 -or $n -ge $candidates.Count) { throw "Index out of range." }
                $Database = $candidates[$n]
                Write-Log INFO "Selected database: $Database"
            } else {
                throw "Multiple candidate databases found but tool is not interactive. Provide -Database or BEPOZ_DATABASE."
            }
        }
    }

    $report.result.database = $Database

    # Connect to target database
    $conn = New-SqlConnection -Server $SqlServer -Db $Database
    try {
        # If VenueId missing, allow interactive selection
        if ($VenueId -le 0) {
            if (-not (Test-IsInteractive)) {
                throw "VenueId not supplied and tool is not interactive. Provide -VenueId or BEPOZ_VENUEID."
            }
            $venues = Invoke-SqlDataTable -Connection $conn -Sql "SELECT VenueID, Name, SiteCode FROM dbo.Venue ORDER BY VenueID;" -Parameters @{}
            if ($venues.Rows.Count -eq 0) { throw "No venues found in dbo.Venue." }

            Write-Host ""
            Write-Host "Select a VenueID:"
            foreach ($r in $venues.Rows) {
                $vid = $r["VenueID"]
                $vn  = $r["Name"]
                $sc  = $r["SiteCode"]
                Write-Host ("  {0}  -  {1}  (SiteCode: {2})" -f $vid, $vn, $sc)
            }
            $vIn = Read-Host "Enter VenueID"
            $tmp = 0
            if (-not [int]::TryParse($vIn, [ref]$tmp)) { throw "Invalid VenueID." }
            $VenueId = $tmp
            Write-Log INFO "Selected VenueID: $VenueId"
        }

        # If StoreName missing, allow interactive prompt
        if ([string]::IsNullOrWhiteSpace($StoreName)) {
            if (-not (Test-IsInteractive)) {
                throw "StoreName not supplied and tool is not interactive. Provide -StoreName or BEPOZ_STORENAME."
            }
            $StoreName = Read-Host "Enter new Store Name"
            if ([string]::IsNullOrWhiteSpace($StoreName)) { throw "StoreName cannot be empty." }
        }

        $report.result.venueId = $VenueId
        $report.result.storeName = $StoreName

        # Validate venue exists
        $venueExists = Invoke-SqlScalar -Connection $conn -Sql "SELECT CASE WHEN EXISTS (SELECT 1 FROM dbo.Venue WHERE VenueID=@VenueId) THEN 1 ELSE 0 END;" -Parameters @{ VenueId = $VenueId }
        $report.checks += @{
            name = "VenueExists"
            venueId = $VenueId
            ok = ([int]$venueExists -eq 1)
        }
        if ([int]$venueExists -ne 1) {
            throw "VenueID $VenueId does not exist in dbo.Venue."
        }

        # Check if store name already exists for that venue
        $existingStoreId = Invoke-SqlScalar -Connection $conn -Sql "SELECT TOP 1 StoreID FROM dbo.Store WHERE VenueID=@VenueId AND Name=@Name ORDER BY StoreID;" -Parameters @{ VenueId = $VenueId; Name = $StoreName }
        $report.checks += @{
            name = "StoreNameUniqueWithinVenue"
            ok = ($null -eq $existingStoreId -or $existingStoreId -eq [DBNull]::Value)
            existingStoreId = $existingStoreId
        }
        if ($existingStoreId -ne $null -and $existingStoreId -ne [DBNull]::Value) {
            throw "A store named '$StoreName' already exists for VenueID $VenueId (StoreID: $existingStoreId)."
        }

        # Build insert plan (handles NOT NULL columns without defaults)
        $plan = Build-InsertPlan -Conn $conn -VenueIdValue $VenueId -StoreNameValue $StoreName -StoreGroupValue ($(if ($PSBoundParameters.ContainsKey("StoreGroup")) { [Nullable[int]]$StoreGroup } else { $null })) -Overrides $overrides

        $report.actions += @{
            name = "InsertPlan"
            dryRun = [bool]$DryRun
            columns = $plan.Columns
            usedOverrides = $overrides
            requiredColumnsWithoutDefaults = @($plan.RequiredColumnsWithoutDefaults | ForEach-Object { $_.ColumnName })
            sql = $plan.Sql
        }

        Write-Log INFO "Planned INSERT columns: $($plan.Columns -join ', ')"
        if ($DryRun) {
            Write-Log WARN "DryRun is enabled. No changes will be made."
            $report.result.status = "DryRun"
            $report.result.message = "DryRun only. Re-run with -DryRun:\$false and confirm to create."
            return
        }

        # Confirm apply (unless -Force)
        $ok = Confirm-Apply -Message "This will CREATE a new store '$StoreName' for VenueID $VenueId in database '$Database'."
        if (-not $ok) {
            $report.result.status = "Cancelled"
            $report.result.message = "User did not confirm CREATE (or Force not provided in non-interactive mode)."
            Write-Log WARN $report.result.message
            return
        }

        # Execute insert within a transaction
        $tx = $conn.BeginTransaction()
        try {
            $cmd = $conn.CreateCommand()
            $cmd.Transaction = $tx
            $cmd.CommandText = $plan.Sql
            $cmd.CommandTimeout = 30

            foreach ($col in $plan.Columns) {
                $p = $cmd.Parameters.Add("@$col", [System.Data.SqlDbType]::Variant)
                $p.Value = $plan.Values[$col]
            }

            $newId = $cmd.ExecuteScalar()
            if ($newId -eq $null -or $newId -eq [DBNull]::Value) {
                throw "Insert succeeded but did not return a new StoreID."
            }

            $tx.Commit()

            $report.result.status = "Success"
            $report.result.message = "Created StoreID $newId for VenueID $VenueId."
            $report.result.storeId = [int]$newId

            Write-Log INFO $report.result.message
        } catch {
            try { $tx.Rollback() } catch {}
            throw
        }
    } finally {
        $conn.Close()
        $conn.Dispose()
    }

} catch {
    $report.result.status = "Error"
    $report.result.message = $_.Exception.Message
    Write-Log ERROR $report.result.message
} finally {
    $report.finishedAt = (Get-Date).ToString("o")
    Save-Report -Report $report
}

# Exit codes:
# 0 = success or dryrun/cancelled (toolkit-friendly)
# 1 = error
if ($report.result.status -eq "Error") { exit 1 } else { exit 0 }
