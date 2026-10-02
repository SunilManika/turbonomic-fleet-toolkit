# Database & Workload Collectors

## Files

- `Get-SqlServerInventory.ps1` - remote SQL Server discovery, instance topology, and metadata collector.
- `get_turbonomic_metrics.py` - Turbonomic VM metrics, capacity utilization, and action recommendations collector.
- `../config/config.json` - Unified project input configuration file.
- `../output/` - generated reports are written here.

## Input

VM targets, connection addresses, and credentials are centrally configured in `../config/config.json`:

```json
{
  "targets": {
    "windows": [
      {
        "vm_name": "SQLVM01",
        "nameOrAddress": "10.0.1.10",
        "username": "turbowmi",
        "password": "Password123!",
        "sql_username": "azureuser",
        "sql_password": "AdminPassword123!"
      }
    ]
  }
}
```

`nameOrAddress` can be a hostname or IP address.

## Run

From a central Windows management VM:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Get-SqlServerInventory.ps1 `
    -InputFile ..\config\config.json `
    -OutputDirectory ..\output
```

The script prompts for a credential.

If the current Windows identity already has the required permissions:

```powershell
.\Get-SqlServerInventory.ps1 `
    -InputFile ..\config\config.json `
    -UseCurrentCredential `
    -OutputDirectory ..\output
```

## What it collects

### VM
- hostname
- domain/workgroup
- manufacturer/model
- OS/version/build
- CPU
- RAM
- BIOS/serial/UUID
- disks and free space
- network adapters/IP/DNS/gateway

### SQL Server WMI/CIM
- SQL WMI namespaces
- SQL Server services
- SQL Agent
- SQL Browser
- default/named instances
- service state
- startup mode
- service account
- instance ID
- SQL installation paths
- registry version/edition/patch information
- TCP configuration

### SQL Server engine
- ServerName
- MachineName
- InstanceName
- Edition
- EditionID
- ProductVersion
- ProductLevel
- ProductUpdateLevel
- ProductUpdateReference
- EngineEdition
- clustering
- HADR
- authentication mode
- collation
- resource version

### Databases
- database name
- state
- user access
- recovery model
- compatibility level
- read-only
- auto-close
- auto-shrink
- Service Broker
- trustworthy
- DB chaining
- encryption
- Query Store
- owner
- create date
- log reuse wait
- files
- physical paths
- size
- used space
- growth

## Important architecture note

WMI is used for SQL installation/service/instance discovery and Windows inventory.

WMI does not expose the complete SQL database catalog. After an instance is discovered through WMI, the remote collector connects locally to SQL Server using Windows Integrated Authentication and queries SQL metadata.

## Network prerequisites

The central machine must be able to establish PowerShell Remoting to each target:

```powershell
Test-WSMan SQLVM01
```

The remote account needs:
- PowerShell Remoting/WinRM access
- WMI/CIM access
- SQL Server metadata permissions

For SQL inventory, avoid using `sysadmin` unless your security policy requires it. Grant the least SQL permissions necessary for the queries being used.

## Output

Each run creates:

```text
output/
  sql-inventory-YYYYMMDD-HHMMSS.html
  sql-inventory-YYYYMMDD-HHMMSS.json
  sql-inventory-instances-YYYYMMDD-HHMMSS.csv
  sql-inventory-databases-YYYYMMDD-HHMMSS.csv
  sql-inventory-YYYYMMDD-HHMMSS.log
```

Open the HTML file in a browser for the dashboard.

## Optional

Skip database queries:

```powershell
.\Get-SqlServerInventory.ps1 `
    -InputFile ..\config\config.json `
    -SkipDatabaseDetails
```

Skip network adapter discovery:

```powershell
.\Get-SqlServerInventory.ps1 `
    -InputFile ..\config\config.json `
    -SkipNetworkDetails
```
