param(
  [Parameter(Mandatory=$true)]
  [string]$RunDir
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $RunDir | Out-Null

function Try-GetBackofficeSqlConfig {
  $regPath = "HKCU:\SOFTWARE\Backoffice"
  if (-not (Test-Path $regPath)) { return $null }

  try {
    $p = Get-ItemProperty -Path $regPath -ErrorAction Stop
    if (-not $p.SQL_Server -or -not $p.SQL_DSN) { return $null }
    return [ordered]@{
      SQL_Server = [string]$p.SQL_Server
      SQL_DSN    = [string]$p.SQL_DSN
      RegPath    = $regPath
    }
  } catch { return $null }
}

function Try-TestLocalSql {
  param(
    [Parameter(Mandatory=$true)][string]$SqlServer,
    [Parameter(Mandatory=$true)][string]$Database
  )

  try {
    Add-Type -AssemblyName System.Data | Out-Null
    $connString = "Server=$SqlServer;Database=$Database;Integrated Security=True;TrustServerCertificate=True;Connection Timeout=5;"
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = @"
SELECT
  @@SERVERNAME AS AtAtServerName,
  CAST(SERVERPROPERTY('ServerName') AS nvarchar(256)) AS ServerProperty_ServerName,
  CAST(SERVERPROPERTY('MachineName') AS nvarchar(256)) AS ServerProperty_MachineName,
  CAST(@@SERVICENAME AS nvarchar(256)) AS ServiceName,
  DB_NAME() AS DatabaseName,
  GETDATE() AS ServerTime;
"@
    $r = $cmd.ExecuteReader()
    $r.Read() | Out-Null

    $result = [ordered]@{
      Ok = $true
      AtAtServerName = [string]$r["AtAtServerName"]
      ServerName     = [string]$r["ServerProperty_ServerName"]
      MachineName    = [string]$r["ServerProperty_MachineName"]
      ServiceName    = [string]$r["ServiceName"]
      Database       = [string]$r["DatabaseName"]
      ServerTime     = [string]$r["ServerTime"]
    }

    $r.Close()
    $conn.Close()
    return $result
  }
  catch {
    return [ordered]@{ Ok = $false; Error = $_.Exception.Message }
  }
}

function Get-DotNet48Release {
  $key = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
  if (-not (Test-Path $key)) { return $null }
  try { return (Get-ItemProperty -Path $key -Name Release -ErrorAction Stop).Release }
  catch { return $null }
}

Start-Transcript -Path (Join-Path $RunDir "transcript.log") -Append | Out-Null
try {
  $extra = [ordered]@{}

  $os = Get-CimInstance Win32_OperatingSystem
  $cs = Get-CimInstance Win32_ComputerSystem
  $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"

  $extra["OSCaption"]     = $os.Caption
  $extra["OSVersion"]     = $os.Version
  $extra["InstallDate"]   = [string]$os.InstallDate
  $extra["UptimeHours"]   = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalHours, 1)
  $extra["TotalRAM_GB"]   = [math]::Round(($cs.TotalPhysicalMemory / 1GB), 2)
  $extra["CDriveFree_GB"] = if ($disk) { [math]::Round(($disk.FreeSpace / 1GB), 2) } else { $null }
  $extra["CDriveSize_GB"] = if ($disk) { [math]::Round(($disk.Size / 1GB), 2) } else { $null }
  $extra["DotNet48Release"] = Get-DotNet48Release

  $sqlCfg = Try-GetBackofficeSqlConfig
  if ($null -eq $sqlCfg) {
    $extra["BackofficeRegistryFound"] = $false
    $extra["SqlConfig"] = $null
    $extra["SqlTest"]   = [ordered]@{ Ok = $false; Error = "HKCU:\SOFTWARE\Backoffice missing or SQL_Server/SQL_DSN not set for this user." }
    $exitCode = 1
  } else {
    $extra["BackofficeRegistryFound"] = $true
    $extra["SqlConfig"] = $sqlCfg

    $sqlTest = Try-TestLocalSql -SqlServer $sqlCfg.SQL_Server -Database $sqlCfg.SQL_DSN
    $extra["SqlTest"] = $sqlTest
    $exitCode = if ($sqlTest.Ok) { 0 } else { 1 }

    # Warning if SQL thinks it has a different name (helps catch renamed/restored instances)
    if ($sqlTest.Ok) {
      $connectedTo = $sqlCfg.SQL_Server
      if ($sqlTest.ServerName -and ($connectedTo -notlike "$($sqlTest.ServerName)*")) {
        $extra["Warnings"] = @("Connected instance '$connectedTo' reports ServerName '$($sqlTest.ServerName)'. This can happen after SQL renames/restores.")
      }
    }
  }

  # Write files
  ($extra | ConvertTo-Json -Depth 10) | Out-File (Join-Path $RunDir "doctor-report.json") -Encoding utf8

  $summary = @()
  $summary += "Doctor Report - $(Get-Date -Format s)"
  $summary += "Machine: $env:COMPUTERNAME  User: $env:USERNAME"
  $summary += "OS: $($extra.OSCaption) ($($extra.OSVersion))  UptimeHours: $($extra.UptimeHours)"
  $summary += "RAM(GB): $($extra.TotalRAM_GB)  C:\ Free(GB): $($extra.CDriveFree_GB) / Size(GB): $($extra.CDriveSize_GB)"
  $summary += ".NET Release: $($extra.DotNet48Release)"
  $summary += "Backoffice Registry Found: $($extra.BackofficeRegistryFound)"
  if ($extra.SqlConfig) {
    $summary += "SQL_Server: $($extra.SqlConfig.SQL_Server)"
    $summary += "SQL_DSN:    $($extra.SqlConfig.SQL_DSN)"
  }
  $summary += "SQL Test OK: $($extra.SqlTest.Ok)"
  if (-not $extra.SqlTest.Ok) { $summary += "SQL Error: $($extra.SqlTest.Error)" }
  if ($extra.Warnings) { $summary += "Warnings: $($extra.Warnings -join '; ')" }

  $summary | Out-File (Join-Path $RunDir "doctor-summary.txt") -Encoding utf8
  $summary | ForEach-Object { Write-Host $_ }

  exit $exitCode
}
catch {
  Write-Error $_
  exit 1
}
finally {
  Stop-Transcript | Out-Null
}
