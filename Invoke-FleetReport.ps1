<#
.SYNOPSIS
    Master fleet report: runs SQL Server inventory + Turbonomic VM metrics,
    then merges both outputs into a single beautified HTML dashboard with
    per-section CSV download buttons.

.DESCRIPTION
    1. Invokes sql-server-inventory\Get-SQLServerInventory.ps1
    2. Invokes turbonomic\get_vm_metrics.py  (requires Python 3 on PATH)
    3. Reads both output JSON files
    4. Writes a merged HTML report (fixed top-bar, sidebar, KPI cards,
       SVG charts from both sources, full data tables, inline CSV download)

    All output is written to .\fleet-output\ by default.

.PARAMETER ConfigFile
    Path to the unified JSON configuration file.
    Default: .\config\config.json

.PARAMETER OutputDirectory
    Directory to write the merged report and sub-script outputs.
    Default: .\fleet-output

.PARAMETER SqlInventoryScript
    Path to Get-SqlServerInventory.ps1.
    Default: .\collectors\Get-SqlServerInventory.ps1

.PARAMETER TurboScript
    Path to get_turbonomic_metrics.py.
    Default: .\collectors\get_turbonomic_metrics.py

.PARAMETER PythonExe
    Python executable to use.  Tries 'python' then 'python3' if not specified.

.PARAMETER Days
    Historical period in days passed to the Turbonomic collector (default: 1).

.PARAMETER SkipSqlInventory
    Skip running Get-SQLServerInventory.ps1 (use existing JSON in OutputDirectory).

.PARAMETER SkipTurbo
    Skip running get_vm_metrics.py (use existing JSON in OutputDirectory).

.PARAMETER SkipSqlLoginSetup
    Passed through to Get-SQLServerInventory.ps1.

.EXAMPLE
    .\Invoke-FleetReport.ps1

.EXAMPLE
    .\Invoke-FleetReport.ps1 -Days 7 -SkipSqlLoginSetup
#>

[CmdletBinding()]
param(
    [string]$ConfigFile       = ".\config\config.json",
    [string]$ServersCSV       = ".\config\config.json",
    [string]$OutputDirectory  = ".\fleet-output",
    [string]$TurboConfig      = ".\config\config.json",
    [string]$SqlInventoryScript = ".\collectors\Get-SqlServerInventory.ps1",
    [string]$TurboScript      = ".\collectors\get_turbonomic_metrics.py",
    [string]$PythonExe        = "",
    [int]   $Days             = 1,
    [switch]$SkipSqlInventory,
    [switch]$SkipTurbo,
    [switch]$SkipSqlLoginSetup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

function Write-Step {
    param([string]$Msg)
    Write-Host ""
    Write-Host "  $Msg" -ForegroundColor Cyan
}

function Write-OK   { param([string]$Msg); Write-Host "  [OK]  $Msg" -ForegroundColor Green }
function Write-Warn { param([string]$Msg); Write-Host "  [!!]  $Msg" -ForegroundColor Yellow }
function Write-Fail { param([string]$Msg); Write-Host "  [XX]  $Msg" -ForegroundColor Red }

function HtmlEnc {
    param([object]$v)
    if ($null -eq $v -or [string]$v -eq '') { return '—' }
    [System.Net.WebUtility]::HtmlEncode([string]$v)
}

function Find-Python {
    if ($PythonExe -and (Get-Command $PythonExe -ErrorAction SilentlyContinue)) {
        return $PythonExe
    }
    foreach ($candidate in @('python', 'python3', 'py')) {
        if (Get-Command $candidate -ErrorAction SilentlyContinue) {
            $ver = & $candidate --version 2>&1
            if ($ver -match '^Python 3') { return $candidate }
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# SETUP
# ---------------------------------------------------------------------------

$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$sqlSubDir   = Join-Path $OutputDirectory "sql"
$turboSubDir = Join-Path $OutputDirectory "turbo"
New-Item -ItemType Directory -Path $sqlSubDir   -Force | Out-Null
New-Item -ItemType Directory -Path $turboSubDir -Force | Out-Null

$runStamp   = Get-Date -Format 'yyyyMMdd-HHmmssff'
$mergedHtml = Join-Path $OutputDirectory "fleet-report-$runStamp.html"
$logFile    = Join-Path $OutputDirectory "fleet-report-$runStamp.log"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Msg"
    Write-Host $line
    Add-Content -LiteralPath $logFile -Value $line
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor White
Write-Host "  FLEET REPORT — SQL Inventory + Turbonomic VM Metrics" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor White
Write-Log "Output directory : $OutputDirectory"

# ---------------------------------------------------------------------------
# STEP 1 — SQL SERVER INVENTORY
# ---------------------------------------------------------------------------

$sqlJsonPath = $null

if (-not $SkipSqlInventory) {
    Write-Step "Running SQL Server Inventory..."

    $inputToUse = if ($ConfigFile -and (Test-Path -LiteralPath $ConfigFile)) { $ConfigFile } else { $ServersCSV }
    if (-not (Test-Path -LiteralPath $SqlInventoryScript)) {
        Write-Fail "Not found: $SqlInventoryScript"
        Write-Log "SQL inventory script not found: $SqlInventoryScript" 'ERROR'
    } elseif (-not (Test-Path -LiteralPath $inputToUse)) {
        Write-Fail "Not found: $inputToUse"
        Write-Log "Config/Input file not found: $inputToUse" 'ERROR'
    } else {
        try {
            $sqlArgs = @{
                InputFile           = (Resolve-Path $inputToUse).Path
                OutputDirectory     = $sqlSubDir
                # Always skip the interactive SQL login pre-flight when called from
                # the orchestrator — credentials are already in config.json and the
                # interactive Get-Credential prompt would hang non-interactively.
                # Pass -SkipSqlLoginSetup:$false explicitly when the flag is NOT set
                # so the caller's choice is honoured.
                SkipSqlLoginSetup   = $SkipSqlLoginSetup.IsPresent
                # Suppress the WinRM credential prompt for localhost targets by using
                # the current session's identity (the script is already running as the
                # local admin on the management VM).
                UseCurrentCredential = $true
            }

            & $SqlInventoryScript @sqlArgs
            Write-OK "SQL inventory complete"
            Write-Log "SQL inventory complete"
        } catch {
            # Walk the full exception chain so the root cause is always visible,
            # not just the outermost wrapper message.
            $ex    = $_.Exception
            $chain = @()
            while ($ex) {
                if ($ex.Message) { $chain += $ex.Message.Trim() }
                $ex = $ex.InnerException
            }
            $fullMsg = $chain -join ' --> '
            Write-Warn "SQL inventory failed: $fullMsg"
            Write-Log  "SQL inventory failed: $fullMsg" 'WARN'
            Write-Log  "  ScriptStackTrace: $($_.ScriptStackTrace)" 'WARN'
        }
    }
}

# Find latest JSON in sqlSubDir
$sqlJsonPath = Get-ChildItem -Path $sqlSubDir -Filter 'sql-inventory-*.json' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }

if ($sqlJsonPath) { Write-Log "SQL JSON: $sqlJsonPath" }
else              { Write-Warn "No SQL inventory JSON found — SQL sections will be empty" }

# ---------------------------------------------------------------------------
# STEP 2 — TURBONOMIC VM METRICS
# ---------------------------------------------------------------------------

$turboJsonPath = $null

if (-not $SkipTurbo) {
    Write-Step "Running Turbonomic VM Metrics collector..."

    $pyExe = Find-Python
    if (-not $pyExe) {
        Write-Warn "Python 3 not found on PATH — skipping Turbonomic collection"
        Write-Log "Python 3 not found — Turbonomic skipped" 'WARN'
    } elseif (-not (Test-Path -LiteralPath $TurboScript)) {
        Write-Fail "Not found: $TurboScript"
        Write-Log "Turbonomic script not found: $TurboScript" 'ERROR'
    } else {
        try {
            $turboArgs = @(
                $TurboScript,
                '--config', (Resolve-Path $TurboConfig).Path,
                '--output', $turboSubDir,
                '--days',   $Days
            )
            & $pyExe @turboArgs
            Write-OK "Turbonomic collection complete"
            Write-Log "Turbonomic collection complete"
        } catch {
            Write-Warn "Turbonomic collection failed: $($_.Exception.Message)"
            Write-Log "Turbonomic collection failed: $($_.Exception.Message)" 'WARN'
            Remove-Item -LiteralPath (Join-Path $OutputDirectory 'turbo-config-tmp.json') -ErrorAction SilentlyContinue
        }
    }
}

# Find latest JSON in turboSubDir
$turboJsonPath = Get-ChildItem -Path $turboSubDir -Filter 'turbo-vm-metrics-*.json' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }

if ($turboJsonPath) { Write-Log "Turbo JSON: $turboJsonPath" }
else                { Write-Warn "No Turbonomic JSON found — Turbonomic sections will be empty" }

# ---------------------------------------------------------------------------
# STEP 3 — LOAD DATA
# ---------------------------------------------------------------------------

Write-Step "Loading data for merged report..."

# ── SQL data ─────────────────────────────────────────────────────────────────
$sqlReport      = $null
$sqlServers     = @()   # raw result objects (per VM)
$sqlInstances   = @()   # flattened derived instances
$sqlDatabases   = @()
$sqlDisks       = @()
$oracleInstances = @()  # flattened Oracle instances
$oracleDatabases = @()  # flattened Oracle databases

if ($sqlJsonPath) {
    try {
        $sqlReport = Get-Content -LiteralPath $sqlJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $sqlServers = @($sqlReport.Servers | Where-Object { $_ })

        foreach ($srv in $sqlServers) {
            if (-not ($srv.Success -and $srv.Data)) { continue }
            $vmName = [string]$srv.DisplayName
            $vmData = $srv.Data

            # Generic safe-read helper: returns '' when a property is absent on a
            # PSCustomObject (ConvertFrom-Json in PS 5.1 omits null-valued properties).
            # $obj  = the PSCustomObject to read from
            # $name = property name string
            function ObjProp($obj, [string]$name) {
                if ($null -ne $obj -and $obj.PSObject.Properties.Name -contains $name) {
                    return [string]$obj.$name
                }
                return ''
            }

            # VM-level hardware properties (may be absent when WMI collection failed)
            $vmObj    = $vmData.VM
            $vmOS     = ObjProp $vmObj 'OS'
            $vmMemGB  = if ($null -ne $vmObj -and $vmObj.PSObject.Properties.Name -contains 'TotalPhysicalMemoryGB') { $vmObj.TotalPhysicalMemoryGB } else { $null }
            $vmLCPUs  = if ($null -ne $vmObj -and $vmObj.PSObject.Properties.Name -contains 'LogicalProcessorCount') { $vmObj.LogicalProcessorCount } else { $null }
            $vmPCores = if ($null -ne $vmObj -and $vmObj.PSObject.Properties.Name -contains 'PhysicalCoreCount')     { $vmObj.PhysicalCoreCount     } else { $null }
            $vmSocks  = if ($null -ne $vmObj -and $vmObj.PSObject.Properties.Name -contains 'NumberOfProcessors')    { $vmObj.NumberOfProcessors    } else { $null }
            $vmOSVer  = ObjProp $vmObj 'OSVersion'
            $vmBuild  = ObjProp $vmObj 'OSBuild'

            # Disks
            foreach ($d in @($vmData.Disks)) {
                if (-not $d) { continue }
                $sqlDisks += [PSCustomObject]@{
                    VM          = $vmName
                    Drive       = ObjProp $d 'Drive'
                    VolumeName  = ObjProp $d 'VolumeName'
                    FileSystem  = ObjProp $d 'FileSystem'
                    SizeGB      = if ($d.PSObject.Properties.Name -contains 'SizeGB')      { $d.SizeGB      } else { 0 }
                    FreeGB      = if ($d.PSObject.Properties.Name -contains 'FreeGB')      { $d.FreeGB      } else { 0 }
                    FreePercent = if ($d.PSObject.Properties.Name -contains 'FreePercent') { $d.FreePercent } else { $null }
                }
            }

            # Instances + Databases
            foreach ($inst in @($vmData.SQLInstances)) {
                if (-not $inst) { continue }
                $iprops  = @($inst.PSObject.Properties.Name)

                # SafeProp: reads a string property from $inst safely
                function SafeProp([string]$n) {
                    if ($iprops -contains $n) { return [string]$inst.$n } else { return '' }
                }

                $edition = if (SafeProp 'Edition')        { SafeProp 'Edition'        } else { SafeProp 'RegistryEdition' }
                $version = if (SafeProp 'ProductVersion') { SafeProp 'ProductVersion' } else { SafeProp 'RegistryVersion' }
                $tcpPort = ''
                if (($iprops -contains 'TCP') -and $inst.TCP) {
                    if ($inst.TCP.PSObject.Properties.Name -contains 'Port') { $tcpPort = [string]$inst.TCP.Port }
                }

                $sqlInstances += [PSCustomObject]@{
                    VM                     = $vmName
                    Address                = [string]$srv.InputAddress
                    Instance               = SafeProp 'InstanceName'
                    Endpoint               = SafeProp 'Endpoint'
                    ServiceState           = SafeProp 'ServiceState'
                    Edition                = [string]$edition
                    ProductVersion         = [string]$version
                    ProductLevel           = SafeProp 'ProductLevel'
                    LicenseType            = SafeProp 'LicenseType'
                    NumLicenses            = SafeProp 'NumLicenses'
                    IsClustered            = SafeProp 'IsClustered'
                    ClusterName            = SafeProp 'ClusterName'
                    ClusterNodeRole        = SafeProp 'ClusterNodeRole'
                    IsHadrEnabled          = SafeProp 'IsHadrEnabled'
                    AGNames                = SafeProp 'AGNames'
                    AGLocalRole            = SafeProp 'AGLocalRole'
                    ActiveOrReadReplica    = SafeProp 'ActiveOrReadReplica'
                    PassiveMethod          = SafeProp 'PassiveMethod'
                    IsLogShippingSecondary = SafeProp 'IsLogShippingSecondary'
                    IsMirroringMirror      = SafeProp 'IsMirroringMirror'
                    ForcePassive           = SafeProp 'ForcePassive'
                    HyperThreadRatio       = SafeProp 'HyperThreadRatio'
                    MinVCPUs               = SafeProp 'MinVCPUs'
                    MinRAMGB               = SafeProp 'MinRAMGB'
                    SQLConnectionOK        = SafeProp 'SQLConnectionOK'
                    DatabaseCount          = [int]@($inst.Databases | Where-Object { $_ }).Count
                    TCPPort                = $tcpPort
                    SQLAgentState          = SafeProp 'SQLAgentState'
                    OS                     = $vmOS
                    MemoryGB               = $vmMemGB
                    LogicalCPUs            = $vmLCPUs
                    PhysicalCores          = $vmPCores
                    Sockets                = $vmSocks
                }

                foreach ($db in @($inst.Databases)) {
                    if (-not $db) { continue }
                    $dbName = ObjProp $db 'DatabaseName'
                    $files     = @($inst.DatabaseFiles | Where-Object { (ObjProp $_ 'DatabaseName') -eq $dbName })
                    $sizeFiles = @($files | Where-Object { $null -ne $_.SizeMB -and "$($_.SizeMB)" -match '^\d' })
                    $usedFiles = @($files | Where-Object { $null -ne $_.UsedMB -and "$($_.UsedMB)" -match '^\d' })
                    $sizeMB    = if ($sizeFiles.Count -gt 0) { [double]($sizeFiles | Measure-Object -Property SizeMB -Sum).Sum } else { 0.0 }
                    $usedMB    = if ($usedFiles.Count -gt 0) { [double]($usedFiles | Measure-Object -Property UsedMB -Sum).Sum } else { 0.0 }
                    $sqlDatabases += [PSCustomObject]@{
                        VM            = $vmName
                        Instance      = SafeProp 'InstanceName'
                        Database      = $dbName
                        State         = ObjProp $db 'State'
                        RecoveryModel = ObjProp $db 'RecoveryModel'
                        CompatLevel   = ObjProp $db 'CompatibilityLevel'
                        SizeGB        = [math]::Round($sizeMB / 1024, 2)
                        UsedGB        = [math]::Round($usedMB / 1024, 2)
                        ReadOnly      = ObjProp $db 'IsReadOnly'
                        Encrypted     = ObjProp $db 'Encrypted'
                        Owner         = ObjProp $db 'OwnerName'
                        CreateDate    = ObjProp $db 'CreateDate'
                        LogReuseWait  = ObjProp $db 'LogReuseWait'
                    }
                }
            }

            # Oracle instances from this VM
            foreach ($oi in @($vmData.OracleInstances)) {
                if (-not $oi) { continue }
                function OracleProp([string]$n) {
                    if ($oi.PSObject.Properties.Name -contains $n -and $null -ne $oi.$n) { return [string]$oi.$n }
                    return ''
                }
                # Derive Edition from VersionBanner if collector left it blank
                # (v$instance.edition is empty on some Oracle 21c installs)
                $rawEdition = OracleProp 'Edition'
                $rawBanner  = OracleProp 'VersionBanner'
                $rawVersion = OracleProp 'Version'
                if (-not $rawEdition -and $rawBanner) {
                    if     ($rawBanner -match 'Enterprise') { $rawEdition = 'Enterprise Edition' }
                    elseif ($rawBanner -match 'Standard.*2')  { $rawEdition = 'Standard Edition 2' }
                    elseif ($rawBanner -match 'Standard')  { $rawEdition = 'Standard Edition' }
                    elseif ($rawBanner -match 'Express')   { $rawEdition = 'Express Edition (XE)' }
                    elseif ($rawBanner -match 'Personal')  { $rawEdition = 'Personal Edition' }
                    else                                   { $rawEdition = 'Unknown' }
                }
                # Final fallback: derive from major version number in registry version string
                if (-not $rawEdition -and $rawVersion -match '^(\d+)\.') {
                    $rawEdition = "Oracle $($Matches[1])c (Edition Unknown)"
                }

                $oracleInstances += [PSCustomObject]@{
                    VM               = $vmName
                    InputAddress     = [string]$srv.InputAddress  # VM address from config.json
                    OS               = $vmOS                       # full OS caption from WMI
                    SID              = OracleProp 'SID'
                    Edition          = $rawEdition
                    Version          = $rawVersion
                    ProductLevel     = OracleProp 'ProductLevel'
                    License          = OracleProp 'License'
                    VersionBanner    = OracleProp 'VersionBanner'
                    ServiceState     = OracleProp 'ServiceState'
                    ServiceName      = OracleProp 'ServiceName'
                    ServiceStartMode = OracleProp 'ServiceStartMode'
                    ListenerPort     = OracleProp 'ListenerPort'
                    ListenerName     = OracleProp 'ListenerName'
                    ListenerState    = OracleProp 'ListenerState'
                    ConnectString    = OracleProp 'ConnectString'
                    OracleConnOK     = OracleProp 'OracleConnOK'
                    DatabaseCount    = if ($oi.PSObject.Properties.Name -contains 'DatabaseCount') { [int]$oi.DatabaseCount } else { 0 }
                    TotalDataSizeGB  = OracleProp 'TotalDataSizeGB'
                    OracleHome       = OracleProp 'OracleHome'
                    OracleBase       = OracleProp 'OracleBase'
                    HostName         = OracleProp 'HostName'
                }
                foreach ($db in @($oi.Databases)) {
                    if (-not $db) { continue }
                    $oracleDatabases += [PSCustomObject]@{
                        VM           = $vmName
                        SID          = OracleProp 'SID'
                        DBName       = if ($db.PSObject.Properties.Name -contains 'DBName')       { [string]$db.DBName       } else { '' }
                        DBUniqueName = if ($db.PSObject.Properties.Name -contains 'DBUniqueName') { [string]$db.DBUniqueName } else { '' }
                        OpenMode     = if ($db.PSObject.Properties.Name -contains 'OpenMode')     { [string]$db.OpenMode     } else { '' }
                        LogMode      = if ($db.PSObject.Properties.Name -contains 'LogMode')      { [string]$db.LogMode      } else { '' }
                        IsCDB        = if ($db.PSObject.Properties.Name -contains 'IsCDB')        { [string]$db.IsCDB        } else { '' }
                        PDBName      = if ($db.PSObject.Properties.Name -contains 'PDBName')      { [string]$db.PDBName      } else { '' }
                        PDBOpenMode  = if ($db.PSObject.Properties.Name -contains 'PDBOpenMode')  { [string]$db.PDBOpenMode  } else { '' }
                        DBVersion    = if ($db.PSObject.Properties.Name -contains 'DBVersion')    { [string]$db.DBVersion    } else { '' }
                        Platform     = if ($db.PSObject.Properties.Name -contains 'Platform')     { [string]$db.Platform     } else { '' }
                        CreatedDate  = if ($db.PSObject.Properties.Name -contains 'CreatedDate')  { [string]$db.CreatedDate  } else { '' }
                    }
                }
            }
        }
        Write-OK "SQL data loaded: $($sqlServers.Count) VMs, $($sqlInstances.Count) instances, $($sqlDatabases.Count) databases, $($oracleInstances.Count) Oracle instances"
        Write-Log "SQL data: $($sqlServers.Count) VMs / $($sqlInstances.Count) inst / $($sqlDatabases.Count) db / $($oracleInstances.Count) Oracle"
    } catch {
        Write-Warn "Failed to parse SQL JSON: $($_.Exception.Message)"
        Write-Log "SQL JSON parse failed: $($_.Exception.Message)" 'WARN'
    }
}

# ── Turbonomic data ────────────────────────────────────────────────────────
$turboVMs = @()

if ($turboJsonPath) {
    try {
        # PS 5.1: ConvertFrom-Json on a JSON array returns a single Object[] — wrap via
        # [Object[]] cast so @() enumerates the elements rather than boxing the array.
        $turboVMs = [Object[]](Get-Content -LiteralPath $turboJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json)
        if ($null -eq $turboVMs) { $turboVMs = @() }
        Write-OK "Turbonomic data loaded: $($turboVMs.Count) VMs"
        Write-Log "Turbonomic data: $($turboVMs.Count) VMs"
    } catch {
        Write-Warn "Failed to parse Turbonomic JSON: $($_.Exception.Message)"
        Write-Log "Turbonomic JSON parse failed: $($_.Exception.Message)" 'WARN'
    }
}

# ---------------------------------------------------------------------------
# STEP 4 — BUILD MERGED HTML
# ---------------------------------------------------------------------------

Write-Step "Building merged HTML report..."

# ── KPI aggregates ───────────────────────────────────────────────────────────
$totalVMs        = [math]::Max($sqlServers.Count, $turboVMs.Count)
$sqlSuccessVMs   = @($sqlServers | Where-Object { $_.Success }).Count
$totalInstances  = $sqlInstances.Count
$totalDatabases  = $sqlDatabases.Count
$poweredOn       = @($turboVMs | Where-Object { [string]$_.powerState -eq 'POWERED_ON' }).Count
# actionCount may be absent from PS objects when JSON value was null (PS 5.1 drops null props)
$_mActions       = $turboVMs | ForEach-Object { if ($_.PSObject.Properties.Name -contains 'actionCount') { [int]$_.actionCount } else { 0 } } | Measure-Object -Sum
$totalActions    = [int]   $(if ($null -ne $_mActions.Sum)  { $_mActions.Sum }  else { 0 })
$_mCost          = $turboVMs | ForEach-Object { $_.vmCostPerMonth -as [double] } | Measure-Object -Sum
$totalCostMo     = [double]$(if ($null -ne $_mCost.Sum)     { $_mCost.Sum }     else { 0 })
$_mSavings       = $turboVMs | ForEach-Object { $_.totalActionSavingsPerMonth -as [double] } | Measure-Object -Sum
$totalSavingsMo  = [double]$(if ($null -ne $_mSavings.Sum)  { $_mSavings.Sum }  else { 0 })

# ── HTML table row builders ──────────────────────────────────────────────────

# SQL VM overview rows
# HtmlVmProp defined outside the loop — PS 5.1 valid; (if...) → $(if...) for subexpression
function HtmlVmProp($o, [string]$p) { HtmlEnc $(if ($null -ne $o -and $o.PSObject.Properties.Name -contains $p) { $o.$p } else { '' }) }
$tSqlVm = @(foreach ($srv in $sqlServers) {
    if ($srv.Success) {
        $vm = $srv.Data.VM
        "<tr>" +
        "<td>$(HtmlEnc $srv.DisplayName)</td>" +
        "<td>$(HtmlEnc $srv.InputAddress)</td>" +
        "<td><span class='badge badge-ok'>OK</span></td>" +
        "<td>$(HtmlVmProp $vm 'OS')</td>" +
        "<td>$(HtmlVmProp $vm 'OSVersion')</td>" +
        "<td>$(HtmlVmProp $vm 'OSBuild')</td>" +
        "<td>$(HtmlVmProp $vm 'TotalPhysicalMemoryGB')</td>" +
        "<td>$(HtmlVmProp $vm 'LogicalProcessorCount')</td>" +
        "<td>$(HtmlVmProp $vm 'PhysicalCoreCount')</td>" +
        "<td>$(HtmlVmProp $vm 'NumberOfProcessors')</td>" +
        "<td>$(@($srv.Data.SQLInstances).Count)</td>" +
        "</tr>"
    } else {
        "<tr><td>$(HtmlEnc $srv.InputName)</td><td>$(HtmlEnc $srv.InputAddress)</td>" +
        "<td><span class='badge badge-bad'>FAILED</span></td>" +
        "<td colspan='8'>$(HtmlEnc $srv.Error)</td></tr>"
    }
}) -join "`n"

# SQL instance rows
$tSqlInst = @(foreach ($d in $sqlInstances) {
    "<tr>" +
    "<td>$(HtmlEnc $d.VM)</td>" +
    "<td>$(HtmlEnc $d.Instance)</td>" +
    "<td>$(HtmlEnc $d.Endpoint)</td>" +
    "<td>$(if([string]$d.ServiceState -eq 'Running'){"<span class='badge badge-ok'>$([string]$d.ServiceState)</span>"}else{"<span class='badge badge-warn'>$(HtmlEnc $d.ServiceState)</span>"})</td>" +
    "<td>$(HtmlEnc $d.Edition)</td>" +
    "<td>$(HtmlEnc $d.ProductVersion)</td>" +
    "<td>$(HtmlEnc $d.ProductLevel)</td>" +
    "<td>$(HtmlEnc $d.LicenseType)</td>" +
    "<td>$(HtmlEnc $d.NumLicenses)</td>" +
    "<td>$(HtmlEnc $d.IsClustered)</td>" +
    "<td>$(HtmlEnc $d.IsHadrEnabled)</td>" +
    "<td>$(if([string]$d.SQLConnectionOK -eq 'True'){"<span class='badge badge-ok'>Yes</span>"}elseif([string]$d.SQLConnectionOK -eq ''){"<span class='badge badge-muted'>Skipped</span>"}else{"<span class='badge badge-bad'>No</span>"})</td>" +
    "<td>$(HtmlEnc $d.DatabaseCount)</td>" +
    "<td>$(HtmlEnc $d.TCPPort)</td>" +
    "<td>$(HtmlEnc $d.SQLAgentState)</td>" +
    "<td>$(HtmlEnc $d.OS)</td>" +
    "<td>$(HtmlEnc $d.MemoryGB)</td>" +
    "<td>$(HtmlEnc $d.LogicalCPUs)</td>" +
    "</tr>"
}) -join "`n"

# Database rows
$tDbs = @(foreach ($db in $sqlDatabases) {
    "<tr>" +
    "<td>$(HtmlEnc $db.VM)</td>" +
    "<td>$(HtmlEnc $db.Instance)</td>" +
    "<td>$(HtmlEnc $db.Database)</td>" +
    "<td>$(HtmlEnc $db.State)</td>" +
    "<td>$(HtmlEnc $db.RecoveryModel)</td>" +
    "<td>$(HtmlEnc $db.CompatLevel)</td>" +
    "<td>$(HtmlEnc $db.SizeGB)</td>" +
    "<td>$(HtmlEnc $db.UsedGB)</td>" +
    "<td>$(HtmlEnc $db.ReadOnly)</td>" +
    "<td>$(HtmlEnc $db.Encrypted)</td>" +
    "<td>$(HtmlEnc $db.Owner)</td>" +
    "<td>$(HtmlEnc $db.CreateDate)</td>" +
    "<td>$(HtmlEnc $db.LogReuseWait)</td>" +
    "</tr>"
}) -join "`n"

# Disk rows
$tDisks = @(foreach ($d in $sqlDisks) {
    $pct = if ($d.FreePercent -and [double]$d.FreePercent -lt 20) { "badge-bad" } elseif ($d.FreePercent -and [double]$d.FreePercent -lt 40) { "badge-warn" } else { "badge-ok" }
    "<tr>" +
    "<td>$(HtmlEnc $d.VM)</td>" +
    "<td>$(HtmlEnc $d.Drive)</td>" +
    "<td>$(HtmlEnc $d.VolumeName)</td>" +
    "<td>$(HtmlEnc $d.FileSystem)</td>" +
    "<td>$(HtmlEnc $d.SizeGB)</td>" +
    "<td>$(HtmlEnc $d.FreeGB)</td>" +
    "<td><span class='badge $pct'>$(HtmlEnc $d.FreePercent)%</span></td>" +
    "</tr>"
}) -join "`n"

# Turbonomic VM rows
$tTurboVm = @(foreach ($v in $turboVMs) {
    $psBadge = if ([string]$v.powerState -eq 'POWERED_ON') { "badge-ok" } else { "badge-muted" }
    $sevCls  = switch ([string]$v.severity) { 'CRITICAL'{'badge-bad'} 'MAJOR'{'badge-warn'} 'MINOR'{'badge-warn'} default{'badge-muted'} }
    "<tr>" +
    "<td>$(HtmlEnc $v.name)</td>" +
    "<td><span class='badge $psBadge'>$(HtmlEnc $v.powerState)</span></td>" +
    "<td><span class='badge $sevCls'>$(HtmlEnc $v.severity)</span></td>" +
    "<td>$(HtmlEnc $v.osType)</td>" +
    "<td>$(HtmlEnc $v.cloudTier)</td>" +
    "<td>$(HtmlEnc $v.region)</td>" +
    "<td>$(HtmlEnc $v.resourceGroup)</td>" +
    "<td>$(HtmlEnc $v.numCPUs)</td>" +
    "<td>$(HtmlEnc $v.memCapacityGB)</td>" +
    "<td>$(HtmlEnc $v.storageProvisionedGiB)</td>" +
    "<td>$(HtmlEnc $v.billingType)</td>" +
    "<td>$(if($v.vmCostPerMonth){"$" + [math]::Round([double]$v.vmCostPerMonth,2)}else{"—"})</td>" +
    "<td>$(if($v.totalActionSavingsPerMonth){"$" + [math]::Round([double]$v.totalActionSavingsPerMonth,2)}else{"—"})</td>" +
    "<td>$(HtmlEnc $v.actionCount)</td>" +
    "</tr>"
}) -join "`n"

# Helper for HTML progress bar cells — defined once, used in the utilization loop below
function PctBar($pct) {
    try {
        $p = [double]$pct
        $col = if ($p -lt 70) { '#198038' } elseif ($p -lt 90) { '#f1620a' } else { '#da1e28' }
        "<div class='pbar-wrap'><div class='pbar-track'><div class='pbar-fill' style='width:$([math]::Min($p,100))%;background:$col'></div></div><span class='pbar-lbl'>$([math]::Round($p,1))%</span></div>"
    } catch { "<span class='muted'>—</span>" }
}

# Turbonomic utilization rows
$tTurboUtil = @(foreach ($v in $turboVMs) {
    "<tr>" +
    "<td>$(HtmlEnc $v.name)</td>" +
    "<td>$(PctBar $v.cpuUtilizationPct)</td>" +
    "<td>$(PctBar $v.memUtilizationPct)</td>" +
    "<td>$(PctBar $v.cpuPeakUtilPct)</td>" +
    "<td>$(PctBar $v.memPeakUtilPct)</td>" +
    "<td>$(HtmlEnc $v.iopsUsed)</td>" +
    "<td>$(PctBar $v.netThroughputUtilPct)</td>" +
    "<td>$(PctBar $v.storageLatencyUtilPct)</td>" +
    "</tr>"
}) -join "`n"

# Turbonomic actions rows
$tActions = @(foreach ($v in $turboVMs) {
    foreach ($a in @($v.actions)) {
        if (-not $a) { continue }
        $savStr = if ($a.savingsPerMonth) { '$' + [math]::Round([double]$a.savingsPerMonth,2) + '/mo' } else { '—' }
        $sevA = switch ([string]$a.riskSeverity) { 'CRITICAL'{'badge-bad'} 'MAJOR'{'badge-warn'} 'MINOR'{'badge-warn'} default{'badge-muted'} }
        "<tr>" +
        "<td>$(HtmlEnc $v.name)</td>" +
        "<td><span class='badge badge-info'>$(HtmlEnc $a.actionType)</span></td>" +
        "<td>$(HtmlEnc $a.actionMode)</td>" +
        "<td>$(HtmlEnc $a.details)</td>" +
        "<td><span class='badge $sevA'>$(HtmlEnc $a.riskSeverity)</span></td>" +
        "<td>$(HtmlEnc $a.currentEntity)</td>" +
        "<td>$(HtmlEnc $a.newEntity)</td>" +
        "<td style='color:#198038;font-weight:700'>$savStr</td>" +
        "</tr>"
    }
}) -join "`n"

# Migration Sizing rows — one row per SQL instance, joined with matching Turbonomic VM
$tMigration = @(foreach ($d in $sqlInstances) {
    # Look up matching Turbonomic VM by name (case-insensitive)
    $tv = $turboVMs | Where-Object { [string]$_.name -ieq $d.VM } | Select-Object -First 1

    # Helper: safe read from turbo object
    function TurboVal($prop) {
        if ($null -ne $tv -and $tv.PSObject.Properties.Name -contains $prop -and $null -ne $tv.$prop) {
            return [string]$tv.$prop
        }
        return ''
    }

    $cpuAvg  = TurboVal 'cpuUtilizationPct'
    $cpuPeak = TurboVal 'cpuPeakUtilPct'
    $memAvg  = TurboVal 'memUtilizationPct'
    $memPeak = TurboVal 'memPeakUtilPct'
    $iops    = TurboVal 'iopsUsed'
    $ioMBps  = TurboVal 'ioThroughputUsedMBps'
    $stoUsed = TurboVal 'storageAmountUsedGB'
    $stoProv = TurboVal 'storageProvisionedGiB'
    $region  = TurboVal 'region'
    $ec2Tier = TurboVal 'tierName'

    "<tr>" +
    "<td>$(HtmlEnc $d.VM)</td>" +
    "<td>$(HtmlEnc $d.Instance)</td>" +
    "<td>$(HtmlEnc $d.Edition)</td>" +
    "<td>$(HtmlEnc $d.OS)</td>" +
    "<td>$(HtmlEnc $d.PhysicalCores)</td>" +
    "<td>$(HtmlEnc $d.MemoryGB)</td>" +
    "<td>$(HtmlEnc (TurboVal 'cloudTier'))</td>" +
    "<td>Production</td>" +
    "<td>$(HtmlEnc $region)</td>" +
    "<td>$(HtmlEnc $stoUsed)</td>" +
    "<td>$(HtmlEnc $cpuAvg)</td>" +
    "<td>$(HtmlEnc $cpuPeak)</td>" +
    "<td>$(HtmlEnc (TurboVal 'uptimePct'))</td>" +
    "<td>$(HtmlEnc $memPeak)</td>" +
    "<td>$(HtmlEnc $memAvg)</td>" +
    "<td>$(HtmlEnc $iops)</td>" +
    "<td>$(HtmlEnc $ioMBps)</td>" +
    "<td>$(HtmlEnc $d.Sockets)</td>" +
    "<td>$(HtmlEnc $d.HyperThreadRatio)</td>" +
    "<td>$(HtmlEnc $stoProv)</td>" +
    "<td>$(HtmlEnc $d.ClusterName)</td>" +
    "<td>$(HtmlEnc $d.AGNames)</td>" +
    "<td>$(HtmlEnc $d.ClusterNodeRole)</td>" +
    "<td>$(HtmlEnc $d.ActiveOrReadReplica)</td>" +
    "<td>$(HtmlEnc $d.PassiveMethod)</td>" +
    "<td>$(HtmlEnc $d.ForcePassive)</td>" +
    "<td>$(HtmlEnc $d.ProductVersion)</td>" +
    "<td>$(HtmlEnc $d.LicenseType)</td>" +
    "<td>$(HtmlEnc $d.MinVCPUs)</td>" +
    "<td>$(HtmlEnc $d.MinRAMGB)</td>" +
    "<td></td>" +
    "<td></td>" +
    "<td></td>" +
    "<td></td>" +
    "<td></td>" +
    "<td></td>" +
    "<td>$(HtmlEnc (TurboVal 'hyperVCluster'))</td>" +
    "<td>$(HtmlEnc (TurboVal 'discoveredByName'))</td>" +
    "<td>$(HtmlEnc (TurboVal 'physicalProcessors'))</td>" +
    "<td></td>" +
    "<td>$(HtmlEnc (TurboVal 'physicalCores'))</td>" +
    "<td>$(HtmlEnc $ec2Tier)</td>" +
    "<td></td>" +
    "<td></td>" +
    "<td></td>" +
    "</tr>"
}) -join "`n"

# Master Data rows — one row per SQL instance (DB Type = SQL Server, Oracle cols blank)
$tMaster = @(foreach ($d in $sqlInstances) {
    $tv = $turboVMs | Where-Object { [string]$_.name -ieq $d.VM } | Select-Object -First 1
    function MasterTurboVal($prop) {
        if ($null -ne $tv -and $tv.PSObject.Properties.Name -contains $prop -and $null -ne $tv.$prop) {
            return [string]$tv.$prop
        }
        return ''
    }
    "<tr>" +
    "<td><span class='badge badge-info'>SQL Server</span></td>" +
    # SQL fields
    "<td>$(HtmlEnc $d.VM)</td>" +
    "<td>$(HtmlEnc $d.Address)</td>" +
    "<td>$(HtmlEnc $d.Instance)</td>" +
    "<td>$(HtmlEnc $d.Endpoint)</td>" +
    "<td>$(HtmlEnc $d.ServiceState)</td>" +
    "<td>$(HtmlEnc $d.Edition)</td>" +
    "<td>$(HtmlEnc $d.ProductVersion)</td>" +
    "<td>$(HtmlEnc $d.ProductLevel)</td>" +
    "<td>$(HtmlEnc $d.LicenseType)</td>" +
    "<td>$(HtmlEnc $d.NumLicenses)</td>" +
    "<td>$(HtmlEnc $d.OS)</td>" +
    "<td>$(HtmlEnc $d.MemoryGB)</td>" +
    "<td>$(HtmlEnc $d.LogicalCPUs)</td>" +
    "<td>$(HtmlEnc $d.PhysicalCores)</td>" +
    "<td>$(HtmlEnc $d.Sockets)</td>" +
    "<td>$(HtmlEnc $d.HyperThreadRatio)</td>" +
    "<td>$(HtmlEnc $d.DatabaseCount)</td>" +
    "<td>$(HtmlEnc $d.TCPPort)</td>" +
    "<td>$(HtmlEnc $d.SQLAgentState)</td>" +
    "<td>$(HtmlEnc $d.IsClustered)</td>" +
    "<td>$(HtmlEnc $d.ClusterName)</td>" +
    "<td>$(HtmlEnc $d.ClusterNodeRole)</td>" +
    "<td>$(HtmlEnc $d.IsHadrEnabled)</td>" +
    "<td>$(HtmlEnc $d.AGNames)</td>" +
    "<td>$(HtmlEnc $d.AGLocalRole)</td>" +
    "<td>$(HtmlEnc $d.ActiveOrReadReplica)</td>" +
    "<td>$(HtmlEnc $d.PassiveMethod)</td>" +
    "<td>$(HtmlEnc $d.IsLogShippingSecondary)</td>" +
    "<td>$(HtmlEnc $d.IsMirroringMirror)</td>" +
    "<td>$(HtmlEnc $d.ForcePassive)</td>" +
    "<td>$(HtmlEnc $d.MinVCPUs)</td>" +
    "<td>$(HtmlEnc $d.MinRAMGB)</td>" +
    # Turbonomic fields (VM-level — applies to SQL rows)
    "<td>$(HtmlEnc (MasterTurboVal 'powerState'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'severity'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'cloudProvider'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'cloudTier'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'region'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'resourceGroup'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'numCPUs'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'memCapacityGB'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'storageProvisionedGiB'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'storageAmountUsedGB'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'storageAmountCapGB'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'iopsCapacity'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'ioThroughputCapMBps'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'physicalCores'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'physicalProcessors'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'hyperThreadRatio'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'tierName'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'recTierName'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'vmCostPerHour'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'vmCostPerMonth'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'totalActionSavingsPerMonth'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'billingType'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'riCoveragePct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'uptimePct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'cpuUtilizationPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'cpuPeakUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'cpuUsedMHz'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'cpuPeakMHz'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'memUtilizationPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'memPeakUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'memUsedGB'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'memPeakGB'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'storageAmountUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'storageLatencyMs'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'storageLatencyUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'iopsUsed'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'iopsPeak'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'ioThroughputUsedMBps'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'netThroughputUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'netThroughputKbitps'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'actionCount'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'discoveredByName'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'hyperVCluster'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'ipAddresses'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'environmentType'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'osType'))</td>" +
    "<td>$(HtmlEnc (MasterTurboVal 'discoveredByType'))</td>" +
    # Oracle columns — blank for SQL rows (10 cols: SID, ProductLevel, ServiceState,
    # ListenerPort, ConnectString, OracleConnOK, DatabaseCount, TotalDataSizeGB, Home, Base)
    "<td></td><td></td><td></td><td></td><td></td>" +
    "<td></td><td></td><td></td><td></td><td></td>" +
    "</tr>"
}) -join "`n"

# Master Data rows — one row per Oracle instance (DB Type = Oracle, SQL cols blank)
$tMasterOracle = @(foreach ($ov in $oracleInstances) {
    $tv = $turboVMs | Where-Object { [string]$_.name -ieq $ov.VM } | Select-Object -First 1
    function MasterOraVal($prop) {
        if ($null -ne $ov -and $ov.PSObject.Properties.Name -contains $prop -and $null -ne $ov.$prop) {
            return [string]$ov.$prop
        }
        return ''
    }
    function MasterOraTurboVal($prop) {
        if ($null -ne $tv -and $tv.PSObject.Properties.Name -contains $prop -and $null -ne $tv.$prop) {
            return [string]$tv.$prop
        }
        return ''
    }
    "<tr>" +
    "<td><span class='badge' style='background:#f1620a;color:#fff'>Oracle</span></td>" +
    # SQL-mapped columns for Oracle rows (32 cols total)
    # Col 2: Server Name
    "<td>$(HtmlEnc $ov.VM)</td>" +
    # Col 3: Address — VM IP/address from config.json
    "<td>$(HtmlEnc $ov.InputAddress)</td>" +
    # Cols 4-6: DB Instance Name, Endpoint, Service State — blank
    "<td></td><td></td><td></td>" +
    # Col 7: DB Server Edition — Oracle Edition
    "<td>$(HtmlEnc (MasterOraVal 'Edition'))</td>" +
    # Col 8: DB Server Version — Oracle Version
    "<td>$(HtmlEnc (MasterOraVal 'Version'))</td>" +
    # Col 9: Product Level — blank
    "<td></td>" +
    # Col 10: License Type — Oracle License
    "<td>$(HtmlEnc (MasterOraVal 'License'))</td>" +
    # Col 11: Num Licenses — blank
    "<td></td>" +
    # Col 12: Operating System — full WMI OS caption (same source as SQL rows)
    "<td>$(HtmlEnc $ov.OS)</td>" +
    # Col 13: RAM (GB) — from Turbonomic memCapacityGB
    "<td>$(HtmlEnc (MasterOraTurboVal 'memCapacityGB'))</td>" +
    # Col 14: Logical CPUs — from Turbonomic numCPUs
    "<td>$(HtmlEnc (MasterOraTurboVal 'numCPUs'))</td>" +
    # Col 15: Physical Cores — from Turbonomic physicalCores
    "<td>$(HtmlEnc (MasterOraTurboVal 'physicalCores'))</td>" +
    # Col 16: Physical Processors On Server — from Turbonomic physicalProcessors
    "<td>$(HtmlEnc (MasterOraTurboVal 'physicalProcessors'))</td>" +
    # Cols 17-33: Hyper-Thread Ratio through Min RAM(GB) — blank (17 cols)
    "<td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>" +
    "<td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>" +
    "<td></td>" +
    # Turbonomic fields (VM-level — shared with Oracle rows)
    "<td>$(HtmlEnc (MasterOraTurboVal 'powerState'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'severity'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'cloudProvider'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'cloudTier'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'region'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'resourceGroup'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'numCPUs'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'memCapacityGB'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'storageProvisionedGiB'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'storageAmountUsedGB'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'storageAmountCapGB'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'iopsCapacity'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'ioThroughputCapMBps'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'physicalCores'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'physicalProcessors'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'hyperThreadRatio'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'tierName'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'recTierName'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'vmCostPerHour'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'vmCostPerMonth'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'totalActionSavingsPerMonth'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'billingType'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'riCoveragePct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'uptimePct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'cpuUtilizationPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'cpuPeakUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'cpuUsedMHz'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'cpuPeakMHz'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'memUtilizationPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'memPeakUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'memUsedGB'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'memPeakGB'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'storageAmountUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'storageLatencyMs'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'storageLatencyUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'iopsUsed'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'iopsPeak'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'ioThroughputUsedMBps'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'netThroughputUtilPct'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'netThroughputKbitps'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'actionCount'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'discoveredByName'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'hyperVCluster'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'ipAddresses'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'environmentType'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'osType'))</td>" +
    "<td>$(HtmlEnc (MasterOraTurboVal 'discoveredByType'))</td>" +
    # Oracle fields — filled for Oracle rows (Edition/Version in SQL cols 7-8; License in SQL col 10)
    "<td>$(HtmlEnc (MasterOraVal 'SID'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'ProductLevel'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'ServiceState'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'ListenerPort'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'ConnectString'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'OracleConnOK'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'DatabaseCount'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'TotalDataSizeGB'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'OracleHome'))</td>" +
    "<td>$(HtmlEnc (MasterOraVal 'OracleBase'))</td>" +
    "</tr>"
}) -join "`n"

# Oracle instance rows
$tOracleInst = @(foreach ($oi in $oracleInstances) {
    $connBadge  = if ([string]$oi.OracleConnOK -eq 'True') { "badge-ok'>Connected" } else { "badge-muted'>Not Connected" }
    $stateBadge = if ([string]$oi.ServiceState -eq 'Running') { "badge-ok'>Running" } else { "badge-warn'>$(HtmlEnc $oi.ServiceState)" }
    "<tr>" +
    "<td>$(HtmlEnc $oi.VM)</td>" +
    "<td>$(HtmlEnc $oi.SID)</td>" +
    "<td>$(HtmlEnc $oi.Edition)</td>" +
    "<td>$(HtmlEnc $oi.Version)</td>" +
    "<td>$(HtmlEnc $oi.ProductLevel)</td>" +
    "<td>$(HtmlEnc $oi.License)</td>" +
    "<td><span class='badge $stateBadge</span></td>" +
    "<td>$(HtmlEnc $oi.ServiceName)</td>" +
    "<td>$(HtmlEnc $oi.ServiceStartMode)</td>" +
    "<td>$(HtmlEnc $oi.ListenerPort)</td>" +
    "<td>$(HtmlEnc $oi.ListenerName)</td>" +
    "<td>$(HtmlEnc $oi.ListenerState)</td>" +
    "<td>$(HtmlEnc $oi.ConnectString)</td>" +
    "<td><span class='badge $connBadge</span></td>" +
    "<td>$(HtmlEnc $oi.DatabaseCount)</td>" +
    "<td>$(HtmlEnc $oi.TotalDataSizeGB)</td>" +
    "<td>$(HtmlEnc $oi.OracleHome)</td>" +
    "<td>$(HtmlEnc $oi.OracleBase)</td>" +
    "</tr>"
}) -join "`n"

# Oracle database rows
$tOracleDb = @(foreach ($db in $oracleDatabases) {
    $cdbBadge = if ([string]$db.IsCDB -eq 'YES') { "badge-info'>CDB" } else { "badge-muted'>Non-CDB" }
    "<tr>" +
    "<td>$(HtmlEnc $db.VM)</td>" +
    "<td>$(HtmlEnc $db.SID)</td>" +
    "<td>$(HtmlEnc $db.DBName)</td>" +
    "<td>$(HtmlEnc $db.DBUniqueName)</td>" +
    "<td>$(HtmlEnc $db.OpenMode)</td>" +
    "<td>$(HtmlEnc $db.LogMode)</td>" +
    "<td><span class='badge $cdbBadge</span></td>" +
    "<td>$(HtmlEnc $db.PDBName)</td>" +
    "<td>$(HtmlEnc $db.PDBOpenMode)</td>" +
    "<td>$(HtmlEnc $db.DBVersion)</td>" +
    "<td>$(HtmlEnc $db.Platform)</td>" +
    "<td>$(HtmlEnc $db.CreatedDate)</td>" +
    "</tr>"
}) -join "`n"

if (-not $tSqlVm)      { $tSqlVm      = "<tr><td colspan='11' class='empty'>No SQL VM data collected.</td></tr>" }
if (-not $tSqlInst)    { $tSqlInst    = "<tr><td colspan='19' class='empty'>No SQL instance data collected.</td></tr>" }
if (-not $tDbs)        { $tDbs        = "<tr><td colspan='13' class='empty'>No database data collected.</td></tr>" }
if (-not $tDisks)      { $tDisks      = "<tr><td colspan='7'  class='empty'>No disk data collected.</td></tr>" }
if (-not $tTurboVm)    { $tTurboVm    = "<tr><td colspan='14' class='empty'>No Turbonomic VM data collected.</td></tr>" }
if (-not $tTurboUtil)  { $tTurboUtil  = "<tr><td colspan='8'  class='empty'>No utilization data collected.</td></tr>" }
if (-not $tActions)    { $tActions    = "<tr><td colspan='8'  class='empty'>No pending actions.</td></tr>" }
if (-not $tMigration)  { $tMigration  = "<tr><td colspan='45' class='empty'>No migration sizing data available.</td></tr>" }
if (-not $tMaster)     { $tMaster     = "<tr><td colspan='91' class='empty'>No master data available.</td></tr>" }
if (-not $tOracleInst) { $tOracleInst = "<tr><td colspan='19' class='empty'>No Oracle instances detected on any VM.</td></tr>" }
if (-not $tOracleDb)   { $tOracleDb   = "<tr><td colspan='12' class='empty'>No Oracle databases collected.</td></tr>" }

# ── Chart data ───────────────────────────────────────────────────────────────
$vmLabelsJs = '[' + (($sqlServers | Where-Object { $_.Success } | ForEach-Object {
    '"' + ([string]$_.DisplayName -replace '"', '\"') + '"'
}) -join ',') + ']'
$memJs = '[' + (($sqlServers | Where-Object { $_.Success } | ForEach-Object {
    $v = $_.Data.VM
    $val = if ($null -ne $v -and $v.PSObject.Properties.Name -contains 'TotalPhysicalMemoryGB') { $v.TotalPhysicalMemoryGB } else { $null }
    if ($val) { [math]::Round([double]$val, 2) } else { 0 }
}) -join ',') + ']'
$cpuJs = '[' + (($sqlServers | Where-Object { $_.Success } | ForEach-Object {
    $v = $_.Data.VM
    $val = if ($null -ne $v -and $v.PSObject.Properties.Name -contains 'LogicalProcessorCount') { $v.LogicalProcessorCount } else { $null }
    if ($val) { [int]$val } else { 0 }
}) -join ',') + ']'

$tVmLabelsJs = '[' + (($turboVMs | ForEach-Object { '"' + ([string]$_.name -replace '"','\"') + '"' }) -join ',') + ']'
$cpuAvgJs    = '[' + (($turboVMs | ForEach-Object { if ($_.cpuUtilizationPct)               { [math]::Round([double]$_.cpuUtilizationPct,1) }               else { 0 } }) -join ',') + ']'
$cpuPkJs     = '[' + (($turboVMs | ForEach-Object { if ($_.cpuPeakUtilPct)                  { [math]::Round([double]$_.cpuPeakUtilPct,1) }                  else { 0 } }) -join ',') + ']'
$memAvgJs    = '[' + (($turboVMs | ForEach-Object { if ($_.memUtilizationPct)               { [math]::Round([double]$_.memUtilizationPct,1) }               else { 0 } }) -join ',') + ']'
$memPkJs     = '[' + (($turboVMs | ForEach-Object { if ($_.memPeakUtilPct)                  { [math]::Round([double]$_.memPeakUtilPct,1) }                  else { 0 } }) -join ',') + ']'
$costJs      = '[' + (($turboVMs | ForEach-Object { if ($_.vmCostPerMonth)                  { [math]::Round([double]$_.vmCostPerMonth,2) }                  else { 0 } }) -join ',') + ']'
$savJs       = '[' + (($turboVMs | ForEach-Object { if ($_.totalActionSavingsPerMonth)      { [math]::Round([double]$_.totalActionSavingsPerMonth,2) }      else { 0 } }) -join ',') + ']'

# Edition breakdown
$edGroups = @($sqlInstances | Group-Object Edition | Sort-Object Count -Descending)
$edLabJs  = '[' + (($edGroups | ForEach-Object { '"' + ([string]$_.Name -replace '"','\"') + '"' }) -join ',') + ']'
$edCntJs  = '[' + (($edGroups | ForEach-Object { $_.Count }) -join ',') + ']'

$generatedAt = (Get-Date).ToUniversalTime().ToString('o')

# ── Emit HTML ─────────────────────────────────────────────────────────────────
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fleet Report — SQL + Turbonomic</title>
<style>
:root{--bg:#f0f2f5;--surface:#fff;--border:#e2e6ea;--text:#1a1d23;--muted:#6b7280;--accent:#0f62fe;--al:#eef3ff;--ok:#198038;--warn:#f1620a;--bad:#da1e28;--nav:220px;--hdr:60px}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:-apple-system,"Segoe UI",system-ui,sans-serif;font-size:13px;line-height:1.5}
.topbar{position:fixed;top:0;left:0;right:0;height:var(--hdr);background:#0d1117;color:#fff;display:flex;align-items:center;padding:0 24px;gap:14px;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.4)}
.topbar h1{font-size:16px;font-weight:700;white-space:nowrap}
.topbar .meta{font-size:11px;color:#8b949e;margin-left:auto;white-space:nowrap}
.topbar .badge-src{padding:3px 10px;border-radius:20px;font-size:10.5px;font-weight:700;letter-spacing:.04em}
.badge-sql{background:#1d4ed8;color:#fff}.badge-turbo{background:#7c3aed;color:#fff}
.sidebar{position:fixed;top:var(--hdr);left:0;width:var(--nav);bottom:0;background:#fff;border-right:1px solid var(--border);overflow-y:auto;z-index:100;padding:14px 0}
.nav-group{padding:6px 16px 2px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)}
.nav-item{display:block;padding:7px 20px;color:#374151;text-decoration:none;font-size:12.5px;border-left:3px solid transparent;transition:all .15s}
.nav-item:hover,.nav-item.active{background:var(--al);border-left-color:var(--accent);color:var(--accent)}
.main{margin-left:var(--nav);margin-top:var(--hdr);padding:22px 26px;min-height:calc(100vh - var(--hdr))}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:22px}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px}
.kpi-val{font-size:26px;font-weight:800;line-height:1}
.kpi-lbl{font-size:11px;color:var(--muted);margin-top:4px}
.kpi.ok .kpi-val{color:var(--ok)}.kpi.bad .kpi-val{color:var(--bad)}.kpi.accent .kpi-val{color:var(--accent)}.kpi.purple .kpi-val{color:#7c3aed}
.divider{display:flex;align-items:center;gap:10px;margin:24px 0 14px;color:var(--muted);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em}
.divider::before,.divider::after{content:'';flex:1;height:1px;background:var(--border)}
.charts-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;margin-bottom:20px}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px}
.chart-card h3{font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px}
.section{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px;margin-bottom:20px;scroll-margin-top:calc(var(--hdr)+14px)}
.section-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.section-hdr h2{font-size:13.5px;font-weight:700;display:flex;align-items:center;gap:7px}
.section-hdr h2 .ico{width:20px;height:20px;border-radius:5px;background:var(--accent);display:flex;align-items:center;justify-content:center;color:#fff;font-size:10px;font-weight:800;flex-shrink:0}
.dl-btn{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:11.5px;font-weight:600;cursor:pointer;text-decoration:none;white-space:nowrap}
.dl-btn:hover{background:#0050d8}
.search-bar{margin-bottom:10px}
.search-bar input{padding:7px 11px;border:1px solid var(--border);border-radius:6px;font-size:12px;outline:none;width:min(340px,100%)}
.search-bar input:focus{border-color:var(--accent);box-shadow:0 0 0 2px rgba(15,98,254,.15)}
.tablewrap{overflow-x:auto;border-radius:6px;border:1px solid var(--border)}
table{border-collapse:collapse;width:100%;font-size:12.5px}
thead th{background:#f8f9fb;text-align:left;padding:8px 11px;border-bottom:2px solid var(--border);white-space:nowrap;font-weight:600;color:#374151;font-size:11.5px}
tbody td{padding:7px 11px;border-bottom:1px solid #f0f2f5;vertical-align:middle}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:#f8faff}
td.empty{text-align:center;color:var(--muted);font-style:italic;padding:18px}
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;white-space:nowrap}
.badge-ok{background:#dcf5e6;color:#0a6640}.badge-warn{background:#fff3cd;color:#856404}
.badge-bad{background:#fde8e8;color:#9b1c1c}.badge-info{background:#dbeafe;color:#1e40af}
.badge-muted{background:#f3f4f6;color:#4b5563}
.pbar-wrap{display:flex;align-items:center;gap:6px}
.pbar-track{background:#e5e7eb;border-radius:4px;height:7px;width:80px;flex-shrink:0;overflow:hidden}
.pbar-fill{height:7px;border-radius:4px}
.pbar-lbl{font-size:11px;color:var(--muted);min-width:32px}
.muted{color:var(--muted)}
footer{text-align:center;font-size:11px;color:var(--muted);padding:18px;border-top:1px solid var(--border);margin-top:6px}
@media(max-width:800px){.sidebar{display:none}.main{margin-left:0}}
</style>
</head>
<body>

<div class="topbar">
  <h1>&#9711; Fleet Report</h1>
  <span class="badge-src badge-sql">SQL Server Inventory</span>
  <span class="badge-src badge-turbo">Turbonomic</span>
  <span class="meta">Generated: $(HtmlEnc $generatedAt) &nbsp;|&nbsp; $totalVMs VMs</span>
</div>

<nav class="sidebar">
  <div class="nav-group">Overview</div>
  <a class="nav-item active" href="#kpis">&#128200; KPI Summary</a>
  <a class="nav-item" href="#charts">&#9685; Charts</a>
  <div class="nav-group">SQL Server</div>
  <a class="nav-item" href="#sql-vms">&#128187; VM Inventory</a>
  <a class="nav-item" href="#sql-inst">&#9671; SQL Instances</a>
  <a class="nav-item" href="#sql-db">&#128451; Databases</a>
  <a class="nav-item" href="#sql-disk">&#128190; Disks</a>
  <div class="nav-group">Turbonomic</div>
  <a class="nav-item" href="#turbo-vms">&#9964; VM Overview</a>
  <a class="nav-item" href="#turbo-util">&#128209; Utilization</a>
  <a class="nav-item" href="#turbo-actions">&#9889; Pending Actions</a>
  <div class="nav-group">Oracle</div>
  <a class="nav-item" href="#oracle-inst">&#9678; Oracle Instances</a>
  <a class="nav-item" href="#oracle-db">&#128448; Oracle Databases</a>
  <div class="nav-group">Migration</div>
  <a class="nav-item" href="#migration">&#128640; Sizing Input</a>
  <a class="nav-item" href="#master">&#9783; Master Data</a>
</nav>

<main class="main">

<!-- KPIs -->
<div id="kpis" class="kpi-grid">
  <div class="kpi accent"><div class="kpi-val">$totalVMs</div><div class="kpi-lbl">Total VMs</div></div>
  <div class="kpi ok"><div class="kpi-val">$sqlSuccessVMs</div><div class="kpi-lbl">SQL Collected OK</div></div>
  <div class="kpi ok"><div class="kpi-val">$poweredOn</div><div class="kpi-lbl">Powered On</div></div>
  <div class="kpi"><div class="kpi-val">$totalInstances</div><div class="kpi-lbl">SQL Instances</div></div>
  <div class="kpi"><div class="kpi-val">$totalDatabases</div><div class="kpi-lbl">Databases</div></div>
  <div class="kpi$(if($totalActions -gt 0){' bad'}else{''})"><div class="kpi-val">$totalActions</div><div class="kpi-lbl">Pending Actions</div></div>
  <div class="kpi"><div class="kpi-val" style="font-size:18px">`$$([math]::Round($totalCostMo,2))</div><div class="kpi-lbl">Total Cost / Mo</div></div>
  <div class="kpi ok"><div class="kpi-val" style="font-size:18px">`$$([math]::Round($totalSavingsMo,2))</div><div class="kpi-lbl">Potential Savings / Mo</div></div>
</div>

<!-- CHARTS -->
<div id="charts">
<div class="divider">SQL Server — Host Metrics</div>
<div class="charts-row">
  <div class="chart-card"><h3>Memory per VM (GB)</h3><canvas id="cMem"></canvas></div>
  <div class="chart-card"><h3>Logical CPUs per VM</h3><canvas id="cCpu"></canvas></div>
  <div class="chart-card"><h3>SQL Edition Distribution</h3><canvas id="cEd"></canvas></div>
</div>
<div class="divider">Turbonomic — Performance &amp; Cost</div>
<div class="charts-row">
  <div class="chart-card"><h3>CPU Avg vs Peak Utilization (%)</h3><canvas id="cCpuUtil"></canvas></div>
  <div class="chart-card"><h3>Memory Avg vs Peak Utilization (%)</h3><canvas id="cMemUtil"></canvas></div>
  <div class="chart-card"><h3>VM Cost vs Savings (`$/month)</h3><canvas id="cCost"></canvas></div>
</div>
</div>

<!-- SQL VMs -->
<div class="section" id="sql-vms">
<div class="section-hdr">
  <h2><span class="ico" style="background:#1d4ed8">H</span>VM Inventory <span class="badge badge-sql" style="font-size:10px;margin-left:4px">SQL</span></h2>
  <button class="dl-btn" onclick="dlCsv('tSqlVm','fleet-sql-vms.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f1" placeholder="&#128269; Filter..." oninput="ft('tSqlVm','f1')"></div>
<div class="tablewrap"><table id="tSqlVm">
<thead><tr>
  <th>VM Name</th><th>Address</th><th>Status</th><th>OS</th><th>OS Version</th><th>Build</th>
  <th>Memory (GB)</th><th>Logical CPUs</th><th>Physical Cores</th><th>Sockets</th><th>SQL Instances</th>
</tr></thead>
<tbody>$tSqlVm</tbody>
</table></div>
</div>

<!-- SQL Instances -->
<div class="section" id="sql-inst">
<div class="section-hdr">
  <h2><span class="ico" style="background:#1d4ed8">S</span>SQL Server Instances <span class="badge badge-sql" style="font-size:10px;margin-left:4px">SQL</span></h2>
  <button class="dl-btn" onclick="dlCsv('tSqlInst','fleet-sql-instances.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f2" placeholder="&#128269; Filter..." oninput="ft('tSqlInst','f2')"></div>
<div class="tablewrap"><table id="tSqlInst">
<thead><tr>
  <th>VM</th><th>Instance</th><th>Endpoint</th><th>State</th>
  <th>Edition</th><th>Version</th><th>Product Level</th>
  <th>License Type</th><th>Num Licenses</th>
  <th>Clustered</th><th>HADR</th><th>SQL OK</th>
  <th>DB Count</th><th>TCP Port</th><th>Agent</th>
  <th>OS</th><th>Mem (GB)</th><th>Logical CPUs</th>
</tr></thead>
<tbody>$tSqlInst</tbody>
</table></div>
</div>

<!-- Databases -->
<div class="section" id="sql-db">
<div class="section-hdr">
  <h2><span class="ico" style="background:#1d4ed8">D</span>Databases <span class="badge badge-sql" style="font-size:10px;margin-left:4px">SQL</span></h2>
  <button class="dl-btn" onclick="dlCsv('tDbs','fleet-databases.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f3" placeholder="&#128269; Filter..." oninput="ft('tDbs','f3')"></div>
<div class="tablewrap"><table id="tDbs">
<thead><tr>
  <th>VM</th><th>Instance</th><th>Database</th><th>State</th><th>Recovery</th>
  <th>Compat.</th><th>Size (GB)</th><th>Used (GB)</th><th>Read Only</th>
  <th>Encrypted</th><th>Owner</th><th>Created</th><th>Log Reuse Wait</th>
</tr></thead>
<tbody>$tDbs</tbody>
</table></div>
</div>

<!-- Disks -->
<div class="section" id="sql-disk">
<div class="section-hdr">
  <h2><span class="ico" style="background:#1d4ed8">&#128190;</span>Disk Inventory <span class="badge badge-sql" style="font-size:10px;margin-left:4px">SQL</span></h2>
  <button class="dl-btn" onclick="dlCsv('tDisks','fleet-disks.csv')">&#11123; Download CSV</button>
</div>
<div class="tablewrap"><table id="tDisks">
<thead><tr>
  <th>VM</th><th>Drive</th><th>Volume</th><th>File System</th>
  <th>Size (GB)</th><th>Free (GB)</th><th>Free %</th>
</tr></thead>
<tbody>$tDisks</tbody>
</table></div>
</div>

<!-- Turbonomic VM Overview -->
<div class="section" id="turbo-vms">
<div class="section-hdr">
  <h2><span class="ico" style="background:#7c3aed">V</span>VM Overview <span class="badge badge-turbo" style="font-size:10px;margin-left:4px">Turbonomic</span></h2>
  <button class="dl-btn" onclick="dlCsv('tTurboVm','fleet-turbo-vms.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f4" placeholder="&#128269; Filter..." oninput="ft('tTurboVm','f4')"></div>
<div class="tablewrap"><table id="tTurboVm">
<thead><tr>
  <th>Name</th><th>Power</th><th>Severity</th><th>OS</th><th>Cloud Tier</th>
  <th>Region</th><th>Resource Group</th><th>vCPUs</th><th>Mem (GB)</th>
  <th>Storage (GiB)</th><th>Billing</th><th>Cost (\$/mo)</th><th>Savings (\$/mo)</th><th>Actions</th>
</tr></thead>
<tbody>$tTurboVm</tbody>
</table></div>
</div>

<!-- Turbonomic Utilization -->
<div class="section" id="turbo-util">
<div class="section-hdr">
  <h2><span class="ico" style="background:#7c3aed">U</span>Avg &amp; Peak Utilization <span class="badge badge-turbo" style="font-size:10px;margin-left:4px">Turbonomic</span></h2>
  <button class="dl-btn" onclick="dlCsv('tTurboUtil','fleet-turbo-utilization.csv')">&#11123; Download CSV</button>
</div>
<div class="tablewrap"><table id="tTurboUtil">
<thead><tr>
  <th>Name</th><th>CPU Avg %</th><th>Mem Avg %</th><th>CPU Peak %</th><th>Mem Peak %</th>
  <th>IOPS Used</th><th>Net Throughput %</th><th>Storage Latency %</th>
</tr></thead>
<tbody>$tTurboUtil</tbody>
</table></div>
</div>

<!-- Pending Actions -->
<div class="section" id="turbo-actions">
<div class="section-hdr">
  <h2><span class="ico" style="background:#7c3aed">!</span>Pending Actions <span class="badge badge-turbo" style="font-size:10px;margin-left:4px">Turbonomic</span></h2>
  <button class="dl-btn" onclick="dlCsv('tActions','fleet-turbo-actions.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f5" placeholder="&#128269; Filter..." oninput="ft('tActions','f5')"></div>
<div class="tablewrap"><table id="tActions">
<thead><tr>
  <th>VM</th><th>Action Type</th><th>Mode</th><th>Details</th>
  <th>Risk Severity</th><th>Current Instance</th><th>Recommended Instance</th><th>Savings</th>
</tr></thead>
<tbody>$tActions</tbody>
</table></div>
</div>

<!-- Migration Sizing Input -->
<div class="section" id="migration">
<div class="section-hdr">
  <h2><span class="ico" style="background:#059669">&#128640;</span>Migration Sizing Input <span class="badge" style="background:#d1fae5;color:#065f46;font-size:10px;margin-left:4px">SQL + Turbonomic</span></h2>
  <button class="dl-btn" onclick="dlCsv('tMigration','fleet-migration-sizing.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f6" placeholder="&#128269; Filter..." oninput="ft('tMigration','f6')"></div>
<div class="tablewrap"><table id="tMigration">
<thead><tr>
  <th>Server Name</th>
  <th>DB Instance Name</th>
  <th>DB Server Edition</th>
  <th>Operating System</th>
  <th>Total Physical / Virtual Cores</th>
  <th>RAM (GB)</th>
  <th>Machine Type</th>
  <th>Environment</th>
  <th>AWS Region For Pricing</th>
  <th>Used Storage (GB)</th>
  <th>CPU AVG Util%</th>
  <th>CPU Peak Util%</th>
  <th>Usage% (runtime)</th>
  <th>Peak RAM Utilization (%)</th>
  <th>Average RAM Utilization (%)</th>
  <th>Req'd IOPs</th>
  <th>Req'd Throughput MB/Sec</th>
  <th>Physical Processors On Server</th>
  <th>Hyper-Thread Ratio</th>
  <th>Provisioned Storage (GB)</th>
  <th>SQL Cluster Name</th>
  <th>SQL AG Name</th>
  <th>SQL Cluster Node Role</th>
  <th>Is Node Active or a Read-Replica?</th>
  <th>Passive Instance Using Log Ship / DB Mirroring (Non-Cluster)</th>
  <th>Force This SQL Instance To Be Passive</th>
  <th>DB Server Version</th>
  <th>License DB By Cores / S/CAL / NUPs</th>
  <th>Min vCPUs Req'd From Vendor App</th>
  <th>Min RAM(GB) Req'd From Vendor App</th>
  <th>Vendor App Unsupported Cloud Tenancy</th>
  <th>Vendor App Unsupported Virtualization</th>
  <th>Preferred AWS Deployment Model For OS</th>
  <th>Preferred AWS Deployment Model For Database</th>
  <th>In-Scope for HA/DR Sizing In AWS</th>
  <th>HA/DR Availability Zone/Region Required</th>
  <th>VMWare/Hyper-V Cluster</th>
  <th>Hypervisor Host</th>
  <th>Hypervisor Physical Processors</th>
  <th>Hypervisor Cores per Processor</th>
  <th>Hypervisor Total Physical Cores</th>
  <th>EC2 Instance</th>
  <th>Application Name or Workload Type</th>
  <th>Line of Business</th>
  <th>Migration Wave</th>
</tr></thead>
<tbody>$tMigration</tbody>
</table></div>
</div>

<!-- Master Data Table -->
<div class="section" id="master">
<div class="section-hdr">
  <h2><span class="ico" style="background:#374151">&#9783;</span>Master Data <span class="badge" style="background:#f3f4f6;color:#374151;font-size:10px;margin-left:4px">All Fields</span></h2>
  <button class="dl-btn" onclick="dlCsv('tMaster','fleet-master-data.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f7" placeholder="&#128269; Filter..." oninput="ft('tMaster','f7')"></div>
<div class="tablewrap"><table id="tMaster">
<thead><tr>
  <th>DB Type</th>
  <th>Server Name</th>
  <th>Address</th>
  <th>DB Instance Name</th>
  <th>Endpoint</th>
  <th>Service State</th>
  <th>DB Server Edition</th>
  <th>DB Server Version</th>
  <th>Product Level</th>
  <th>License Type</th>
  <th>Num Licenses</th>
  <th>Operating System</th>
  <th>RAM (GB)</th>
  <th>Logical CPUs</th>
  <th>Physical Cores</th>
  <th>Physical Processors On Server</th>
  <th>Hyper-Thread Ratio</th>
  <th>DB Count</th>
  <th>TCP Port</th>
  <th>SQL Agent State</th>
  <th>Is Clustered</th>
  <th>SQL Cluster Name</th>
  <th>SQL Cluster Node Role</th>
  <th>HADR Enabled</th>
  <th>SQL AG Name</th>
  <th>AG Local Role</th>
  <th>Is Node Active or a Read-Replica?</th>
  <th>Passive Method (Log Ship / Mirroring)</th>
  <th>Is Log Shipping Secondary</th>
  <th>Is Mirroring Mirror</th>
  <th>Force This SQL Instance To Be Passive</th>
  <th>Min vCPUs Req'd</th>
  <th>Min RAM(GB) Req'd</th>
  <th>Power State</th>
  <th>Severity</th>
  <th>Cloud Provider</th>
  <th>Cloud Tier</th>
  <th>AWS Region</th>
  <th>Resource Group</th>
  <th>vCPUs (Cloud)</th>
  <th>Mem Capacity (GB)</th>
  <th>Storage Provisioned (GiB)</th>
  <th>Storage Used (GB)</th>
  <th>Storage Amount Capacity (GB)</th>
  <th>IOPS Capacity</th>
  <th>IO Throughput Capacity (MB/s)</th>
  <th>Physical Cores (Hypervisor)</th>
  <th>Physical Processors (Hypervisor)</th>
  <th>Hyper-Thread Ratio (Hypervisor)</th>
  <th>Current EC2 Tier</th>
  <th>Recommended EC2 Tier</th>
  <th>VM Cost Per Hour</th>
  <th>VM Cost Per Month</th>
  <th>Action Savings Per Month</th>
  <th>Billing Type</th>
  <th>RI Coverage %</th>
  <th>Uptime %</th>
  <th>CPU Avg Util%</th>
  <th>CPU Peak Util%</th>
  <th>CPU Used (MHz)</th>
  <th>CPU Peak (MHz)</th>
  <th>Mem Avg Util%</th>
  <th>Mem Peak Util%</th>
  <th>Mem Used (GB)</th>
  <th>Mem Peak (GB)</th>
  <th>Storage Avg Util%</th>
  <th>Storage Latency (ms)</th>
  <th>Storage Latency Util%</th>
  <th>IOPS Used</th>
  <th>IOPS Peak</th>
  <th>IO Throughput Used (MB/s)</th>
  <th>Net Throughput Util%</th>
  <th>Net Throughput (Kbit/s)</th>
  <th>Pending Actions</th>
  <th>Hypervisor Host</th>
  <th>VMWare/Hyper-V Cluster</th>
  <th>IP Addresses</th>
  <th>Environment Type</th>
  <th>OS Type (Cloud)</th>
  <th>Discovered By</th>
  <th>Oracle SID</th>
  <th>Oracle Product Level</th>
  <th>Oracle Service State</th>
  <th>Oracle Listener Port</th>
  <th>Oracle Connect String</th>
  <th>Oracle SQL*Plus OK</th>
  <th>Oracle DB Count</th>
  <th>Oracle Total Data Size (GB)</th>
  <th>Oracle Home</th>
  <th>Oracle Base</th>
</tr></thead>
<tbody>$tMaster
$tMasterOracle</tbody>
</table></div>
</div>

<!-- Oracle Instances -->
<div class="section" id="oracle-inst">
<div class="section-hdr">
  <h2><span class="ico" style="background:#f1620a">O</span>Oracle Instances <span class="badge" style="background:#fff3cd;color:#856404;font-size:10px;margin-left:4px">Oracle</span></h2>
  <button class="dl-btn" onclick="dlCsv('tOracleInst','fleet-oracle-instances.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f8" placeholder="&#128269; Filter..." oninput="ft('tOracleInst','f8')"></div>
<div class="tablewrap"><table id="tOracleInst">
<thead><tr>
  <th>VM</th><th>SID</th><th>Edition</th><th>Version</th><th>Product Level</th><th>License</th>
  <th>Service State</th><th>Service Name</th><th>Start Mode</th>
  <th>Listener Port</th><th>Listener Name</th><th>Listener State</th>
  <th>Connect String</th><th>SQL*Plus Connected</th>
  <th>DB Count</th><th>Total Data Size (GB)</th>
  <th>Oracle Home</th><th>Oracle Base</th>
</tr></thead>
<tbody>$tOracleInst</tbody>
</table></div>
</div>

<!-- Oracle Databases -->
<div class="section" id="oracle-db">
<div class="section-hdr">
  <h2><span class="ico" style="background:#f1620a">D</span>Oracle Databases / PDBs <span class="badge" style="background:#fff3cd;color:#856404;font-size:10px;margin-left:4px">Oracle</span></h2>
  <button class="dl-btn" onclick="dlCsv('tOracleDb','fleet-oracle-databases.csv')">&#11123; Download CSV</button>
</div>
<div class="search-bar"><input id="f9" placeholder="&#128269; Filter..." oninput="ft('tOracleDb','f9')"></div>
<div class="tablewrap"><table id="tOracleDb">
<thead><tr>
  <th>VM</th><th>SID</th><th>DB Name</th><th>DB Unique Name</th>
  <th>Open Mode</th><th>Log Mode</th><th>CDB Type</th>
  <th>PDB Name</th><th>PDB Open Mode</th>
  <th>DB Version</th><th>Platform</th><th>Created</th>
</tr></thead>
<tbody>$tOracleDb</tbody>
</table></div>
</div>

</main>
<footer>Fleet Report &nbsp;|&nbsp; SQL Server Inventory + Turbonomic VM Metrics &nbsp;|&nbsp; Made with IBM Bob</footer>

<script>
/* ── Table filter ── */
function ft(tId,iId){
  var f=document.getElementById(iId).value.toLowerCase();
  var rows=document.getElementById(tId).getElementsByTagName('tr');
  for(var i=1;i<rows.length;i++)
    rows[i].style.display=rows[i].innerText.toLowerCase().indexOf(f)>=0?'':'none';
}

/* ── CSV download ── */
function dlCsv(tId,fname){
  var tbl=document.getElementById(tId);
  if(!tbl)return;
  var rows=tbl.getElementsByTagName('tr');
  var lines=[];
  for(var i=0;i<rows.length;i++){
    var cells=rows[i].querySelectorAll('th,td');
    var cols=[];
    for(var j=0;j<cells.length;j++){
      var txt=cells[j].innerText.replace(/\s+/g,' ').trim();
      cols.push('"'+txt.replace(/"/g,'""')+'"');
    }
    lines.push(cols.join(','));
  }
  var blob=new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=fname;
  a.click();
}

/* ── Active nav on scroll ── */
(function(){
  var items=document.querySelectorAll('.nav-item[href^="#"]');
  var tgts=[].map.call(items,function(a){return document.querySelector(a.getAttribute('href'))});
  function upd(){var y=window.scrollY+80,cur=-1;tgts.forEach(function(t,i){if(t&&t.offsetTop<=y)cur=i});items.forEach(function(a,i){a.classList.toggle('active',i===cur)})}
  window.addEventListener('scroll',upd,{passive:true});upd();
})();

/* ── SVG chart engine ── */
(function(){
  var C=['#0f62fe','#da1e28','#198038','#f1620a','#8b5cf6','#06b6d4','#f59e0b','#10b981'];

  function mkSvg(canvasId,w,h){
    var c=document.getElementById(canvasId);if(!c)return null;
    var s=document.createElementNS('http://www.w3.org/2000/svg','svg');
    s.setAttribute('viewBox','0 0 '+w+' '+h);s.setAttribute('width','100%');
    s.setAttribute('font-family','inherit');s.setAttribute('font-size','11');
    c.parentNode.replaceChild(s,c);return s;
  }
  function el(tag,a,p){
    var e=document.createElementNS('http://www.w3.org/2000/svg',tag);
    for(var k in a)e.setAttribute(k,a[k]);if(p)p.appendChild(e);return e;
  }

  function hbar(id,labels,vals,color){
    var s=mkSvg(id,520,Math.max(60,10+labels.length*34));if(!s)return;
    var W=520,bH=24,gap=10,pL=170,pR=55,pT=10;
    var maxV=Math.max.apply(null,vals)||1;
    labels.forEach(function(lbl,i){
      var y=pT+i*(bH+gap);
      var bw=Math.max(2,(vals[i]/maxV)*(W-pL-pR));
      el('text',{x:pL-6,y:y+bH/2+4,'text-anchor':'end',fill:'#374151'},s).textContent=lbl;
      el('rect',{x:pL,y:y,width:W-pL-pR,height:bH,rx:4,fill:'#f0f2f5'},s);
      el('rect',{x:pL,y:y,width:bw,height:bH,rx:4,fill:color||C[0]},s);
      el('text',{x:pL+bw+5,y:y+bH/2+4,fill:'#374151'},s).textContent=vals[i];
    });
  }

  function hbarGroup(id,labels,series){
    var ns=series.length,bH=13,gap=3,gGap=10,pL=170,pR=50,pT=26,pB=8;
    var gH=ns*(bH+gap)-gap,W=540,H=pT+labels.length*(gH+gGap)-gGap+pB;
    var s=mkSvg(id,W,Math.max(60,H));if(!s)return;
    var allV=series.reduce(function(a,b){return a.concat(b.vals)},[]);
    var maxV=Math.max.apply(null,allV)||1;
    var lx=pL;
    series.forEach(function(sr){
      el('rect',{x:lx,y:8,width:11,height:11,rx:2,fill:sr.color},s);
      el('text',{x:lx+14,y:18,fill:'#374151'},s).textContent=sr.name;
      lx+=sr.name.length*6.5+28;
    });
    labels.forEach(function(lbl,i){
      var gy=pT+i*(gH+gGap);
      el('text',{x:pL-6,y:gy+gH/2+4,'text-anchor':'end',fill:'#374151'},s).textContent=lbl;
      series.forEach(function(sr,si){
        var y=gy+si*(bH+gap);
        var bw=Math.max(2,(sr.vals[i]/maxV)*(W-pL-pR));
        el('rect',{x:pL,y:y,width:W-pL-pR,height:bH,rx:3,fill:'#f0f2f5'},s);
        el('rect',{x:pL,y:y,width:bw,height:bH,rx:3,fill:sr.color},s);
        el('text',{x:pL+bw+4,y:y+bH-2,fill:'#374151'},s).textContent=sr.vals[i];
      });
    });
  }

  function pie(id,labels,vals){
    var W=400,H=190,cx=95,cy=95,r=82,ir=50;
    var s=mkSvg(id,W,H);if(!s)return;
    var tot=vals.reduce(function(a,b){return a+b},0)||1,angle=-Math.PI/2;
    vals.forEach(function(v,i){
      if(!v)return;
      var sw=2*Math.PI*(v/tot);
      var x1=cx+r*Math.cos(angle),y1=cy+r*Math.sin(angle);
      var x2=cx+r*Math.cos(angle+sw),y2=cy+r*Math.sin(angle+sw);
      var xi1=cx+ir*Math.cos(angle),yi1=cy+ir*Math.sin(angle);
      var xi2=cx+ir*Math.cos(angle+sw),yi2=cy+ir*Math.sin(angle+sw);
      var lg=sw>Math.PI?1:0;
      el('path',{d:'M'+xi1+' '+yi1+' L'+x1+' '+y1+' A'+r+' '+r+' 0 '+lg+' 1 '+x2+' '+y2+' L'+xi2+' '+yi2+' A'+ir+' '+ir+' 0 '+lg+' 0 '+xi1+' '+yi1+'Z',fill:C[i%C.length]},s);
      angle+=sw;
    });
    var ly=16;
    labels.forEach(function(lbl,i){
      if(!vals[i])return;
      el('rect',{x:205,y:ly-10,width:11,height:11,rx:2,fill:C[i%C.length]},s);
      el('text',{x:220,y:ly,fill:'#374151'},s).textContent=lbl+' ('+vals[i]+')';
      ly+=19;
    });
  }

  var VMS    = $vmLabelsJs;
  var MEM    = $memJs;
  var CPU    = $cpuJs;
  var TVMS   = $tVmLabelsJs;
  var CPU_A  = $cpuAvgJs;
  var CPU_P  = $cpuPkJs;
  var MEM_A  = $memAvgJs;
  var MEM_P  = $memPkJs;
  var COST   = $costJs;
  var SAV    = $savJs;
  var ED_L   = $edLabJs;
  var ED_C   = $edCntJs;

  hbar('cMem',VMS,MEM,'#198038');
  hbar('cCpu',VMS,CPU,'#f1620a');
  pie('cEd',ED_L,ED_C);
  hbarGroup('cCpuUtil',TVMS,[{name:'Avg %',color:'#0f62fe',vals:CPU_A},{name:'Peak %',color:'#da1e28',vals:CPU_P}]);
  hbarGroup('cMemUtil',TVMS,[{name:'Avg %',color:'#198038',vals:MEM_A},{name:'Peak %',color:'#f1620a',vals:MEM_P}]);
  hbarGroup('cCost',TVMS,[{name:'Cost \$/mo',color:'#0f62fe',vals:COST},{name:'Savings \$/mo',color:'#198038',vals:SAV}]);
})();
</script>
</body>
</html>
"@

Set-Content -LiteralPath $mergedHtml -Value $html -Encoding UTF8
Write-OK "Merged HTML report written: $mergedHtml"
Write-Log "Merged HTML: $mergedHtml"

# ---------------------------------------------------------------------------
# DONE
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "================================================================" -ForegroundColor White
Write-Host "  FLEET REPORT COMPLETE" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor White
Write-Host "  HTML Report  : $mergedHtml" -ForegroundColor White
Write-Host "  SQL outputs  : $sqlSubDir" -ForegroundColor White
Write-Host "  Turbo outputs: $turboSubDir" -ForegroundColor White
Write-Host "  Log          : $logFile" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor White
Write-Host ""

# Open the report in the default browser
try {
    Start-Process $mergedHtml
} catch {}
