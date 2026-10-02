# Turbonomic & SQL Server Fleet Optimization Toolkit

Comprehensive automation toolkit for target VM onboarding into **IBM Turbonomic Application Resource Management (ARM)**, deep **SQL Server inventory & topology discovery**, and **unified fleet reporting & cost optimization**.

---

## Table of Contents

- [Overview](#overview)
- [Key Capabilities](#key-capabilities)
- [Project Architecture](#project-architecture)
- [Directory Layout](#directory-layout)
- [Configuration Specification (`config/config.json`)](#configuration-specification-configconfigjson)
- [Step-by-Step Execution Guide](#step-by-step-execution-guide)
  - [Step 1: Target VM Hardening & Preparation](#step-1-target-vm-hardening--preparation)
  - [Step 2: Turbonomic Target Registration & Discovery](#step-2-turbonomic-target-registration--discovery)
  - [Step 3: Fleet Reporting & Analytics Dashboard](#step-3-fleet-reporting--analytics-dashboard)
- [Standalone Collector Execution](#standalone-collector-execution)
- [Generated Outputs & Artifacts](#generated-outputs--artifacts)
- [Security & Authentication Principles](#security--authentication-principles)
- [Troubleshooting & FAQ](#troubleshooting--faq)

---

## Overview

Enterprises running multi-VM database fleets in AWS/Azure/vSphere often struggle with two major challenges:
1. **Manual Overhead of WMI/WinRM Onboarding:** Turbonomic requires specific WinRM, WMI namespace permissions (`root` and `root\cimv2`), firewall rules, and token filtering policies across every target Windows VM.
2. **Disconnected Insights:** Database licensing, instance topology, Always On AG states, and storage configurations are collected separately from cloud resource sizing, historical utilization metrics, and Turbonomic cost-optimization actions.

This toolkit solves both by providing an end-to-end, automated 3-step pipeline driven by a **single central configuration file**.

---

## Key Capabilities

### 1. Automated Target Hardening (`target-setup/Set-WindowsPrerequisites.ps1`)
- **Automated Service Account Provisioning:** Creates or updates local service account with `PasswordNeverExpires`.
- **Group Membership:** Automatically adds the user to `Remote Management Users` / `WinRMRemoteWMIUsers__`, `Performance Monitor Users`, and `Administrators`.
- **Binary WMI Security Descriptor Configuration:** Automatically grants `Execute Methods` (0x1), `Enable Account` (0x20), and `Remote Enable` (0x80) on both `root` and `root\cimv2` using low-level binary `SetSD` (with `ContainerInherit`). Eliminates all manual GUI clicks in `wmimgmt.msc`.
- **Remote Access Registry Configuration:** Sets `LocalAccountTokenFilterPolicy = 1` in `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System` (essential for remote non-domain admin WMI queries).
- **WinRM & Remote Registry:** Configures WinRM with Negotiate (NTLM) authentication, unencrypted HTTP transport (5985), and enables/starts `RemoteRegistry`.
- **Firewall Provisioning:** Automatically opens inbound TCP ports 5985 (WinRM), 135 (RPC Endpoint Mapper), and 49152-65535 (Dynamic RPC).

### 2. Turbonomic Target Onboarding (`target-setup/register_turbo_targets.py`)
- **Multi-Target Ingestion:** Reads Windows target fleet from `config/config.json` and filters out Linux SSH targets.
- **Dynamic Entity Resolution:** Automatically queries Turbonomic inventory (`/api/v3/search`) to resolve VM entity UUIDs by `vm_name` (with fallback to IP address).
- **Shared Static Scope Management:** Automatically creates and updates a single shared static scope group (`Group-WMI-Windows-Fleet`) containing all resolved Windows VMs (`isStatic: True`), avoiding group proliferation.
- **Resilient Target Creation & Update:** Idempotently creates or updates WMI targets (`PUT /targets` or `POST /targets`) with exponential backoff and 504 Gateway Timeout handling.
- **Automated Rediscovery & Health Polling:** Triggers rediscovery actions and polls until targets achieve `Successful` health state.

### 3. Deep SQL Server & Workload Inventory (`collectors/Get-SqlServerInventory.ps1`)
- **Multi-OS Transport:** Connects to Windows VMs via WinRM and Linux VMs via SSH (using key or password authentication).
- **Non-Interactive SQL Provisioning:** Automatically provisions SQL Server read-only logins (`VIEW SERVER STATE`, `VIEW ANY DATABASE`, `VIEW ANY DEFINITION`) during pre-flight using `sql_username`/`sql_password` without interactive prompts.
- **Comprehensive Metadata Extraction:**
  - **Host & OS:** Hostname, domain/workgroup, CPU model, physical cores, logical processors, RAM, disks & free space, network interfaces.
  - **SQL Instances:** SQL Server services, edition, version, build number, service pack, collation, clustering state.
  - **Databases:** Database name, size (MB), data/log space used, state, compatibility level, recovery model.
  - **High Availability & DR:** Always On Availability Groups, replica roles (Primary/Secondary), availability modes, failover modes, Log Shipping, DB Mirroring.
  - **Multi-Workload Support:** Discovers both Microsoft SQL Server and Oracle database instances.

### 4. Turbonomic Utilization & Sizing Metrics (`collectors/get_turbonomic_metrics.py`)
- **Direct API v3 Extraction:** Extracts live commodity data directly mapped to Turbonomic internal specs (`VCPU`, `CPU`, `VMem`, `Mem`, `StorageAmount`, `StorageAccess` IOPS, `NetThroughput`, `IOThroughput`, `StorageLatency`).
- **Peak & Historical Analysis:** Queries historical statistical series to capture 95th percentile and peak usages over configurable time windows (`--days`).
- **Cost & Action Recommendations:** Captures pending scale/resize actions, current vs. recommended tier specs, and hourly/monthly estimated cost savings.

### 5. Master Fleet Orchestrator (`Invoke-FleetReport.ps1`)
- **Pipeline Automation:** Runs SQL Server inventory and Turbonomic metrics collectors in sequence.
- **Data Correlation:** Correlates OS/database topology with Turbonomic performance data matched on `vm_name` and IP address.
- **Executive HTML Dashboard:** Generates a single, self-contained, responsive dashboard with KPI summary cards, SVG visual charts, full data tables, and client-side CSV export buttons.

---

## Project Architecture

```
                               ┌─────────────────────────────────────────┐
                               │           config/config.json            │
                               │      (Single Source of Truth)           │
                               └────────────────────┬────────────────────┘
                                                    │
                 ┌──────────────────────────────────┼──────────────────────────────────┐
                 │                                  │                                  │
                 ▼                                  ▼                                  ▼
   ┌───────────────────────────┐      ┌───────────────────────────┐      ┌───────────────────────────┐
   │          STEP 1           │      │          STEP 2           │      │          STEP 3           │
   │       Target Setup        │      │    Turbonomic Onboard     │      │   Fleet Report Pipeline   │
   ├───────────────────────────┤      ├───────────────────────────┤      ├───────────────────────────┤
   │ target-setup/             │      │ target-setup/             │      │ Invoke-FleetReport.ps1    │
   │ Set-WindowsPrerequisites  │      │ register_turbo_targets.py │      │   ├─ Get-SqlServerInven.. │
   │                           │      │                           │      │   └─ get_turbonomic_met.. │
   │ • Binary WMI SD (root/cim)│      │ • Resolves VM Entity UUIDs│      │                           │
   │ • LocalAccountTokenFilter │      │ • Builds Static Scope Grp │      │ • Deep SQL Instance Topo  │
   │ • WinRM & Firewall Ports  │      │ • Creates/Updates Targets │      │ • Turbo CPU/Mem/IOPS/Cost │
   │ • Service Account Creation│      │ • Triggers Rediscovery    │      │ • Merged HTML + CSV Report│
   └───────────────────────────┘      └───────────────────────────┘      └───────────────────────────┘
```

---

## Directory Layout

```text
Turbonomic_AWS/
├── config/
│   ├── config.json                     # Active central configuration
│   └── config_sample.json              # Sanitized configuration template
│
├── target-setup/
│   ├── Set-WindowsPrerequisites.ps1    # Step 1: Windows VM prerequisite & WMI automation
│   └── register_turbo_targets.py       # Step 2: Target onboarding & rediscovery into Turbonomic
│
├── collectors/
│   ├── Get-SqlServerInventory.ps1      # Step 3a: Deep SQL Server metadata & instance collector
│   ├── get_turbonomic_metrics.py       # Step 3b: Turbonomic VM utilization & actions collector
│   └── README.md                       # Collectors documentation
│
├── Invoke-FleetReport.ps1               # Step 3: Master orchestrator & dashboard generator
│
├── output/                             # Default directory for generated reports & CSVs
└── docs/                               # Architecture decks, guides, and documentation
```

---

## Configuration Specification (`config/config.json`)

The entire toolkit is driven by [`config/config.json`](config/config.json). Copy [`config/config_sample.json`](config/config_sample.json) to start:

```json
{
  "turbonomic": {
    "host": "tz1.demo.turbonomic.com",
    "api_base_url": "https://tz1.demo.turbonomic.com/api/v3",
    "admin_username": "turboplanuser",
    "admin_password": "turboplanuser"
  },
  "targets": {
    "windows": [
      {
        "display_name": "WMI-turbo-wmi-win-d16",
        "category": "Guest OS Processes",
        "type": "WMI",
        "nameOrAddress": "172.16.0.8",
        "vm_name": "turbo-wmi-win-d16",
        "username": "turbowmid16",
        "password": "TrbWmi#2026!X9@Q7",
        "sql_username": "azureuser",
        "sql_password": "azureuser@123",
        "domain": "WORKGROUP",
        "use_ntlm": true,
        "use_https": false,
        "authenticate_all_db_servers": false,
        "scope_uuid": ""
      }
    ],
    "linux": [
      {
        "display_name": "SSH-turbo-wmi-linux-d8",
        "category": "Guest OS Processes",
        "type": "SSH",
        "nameOrAddress": "172.16.0.10",
        "vm_name": "turbo-wmi-linux-d8",
        "username": "azureuser",
        "password": "",
        "ssh_key_file": "C:\\scripts\\keys\\id_ed25519",
        "sql_username": "sa",
        "sql_password": "azureuser@123",
        "scope_uuid": ""
      }
    ]
  },
  "discovery": {
    "trigger_rediscovery_after_add": true,
    "rediscover_poll_interval_seconds": 15,
    "rediscover_max_wait_seconds": 900
  },
  "logging": {
    "level": "DEBUG",
    "log_file": "wmi_target_setup.log"
  }
}
```

### Parameter Reference

| Section | Parameter | Description |
|---|---|---|
| `turbonomic` | `host` | Turbonomic appliance hostname or IP |
| `turbonomic` | `api_base_url` | Full API base URL (`https://<host>/api/v3`) |
| `turbonomic` | `admin_username` | Turbonomic administrator username |
| `turbonomic` | `admin_password` | Turbonomic administrator password |
| `targets.windows` | `vm_name` | Exact Virtual Machine name as discovered in Turbonomic / Hypervisor |
| `targets.windows` | `nameOrAddress` | IP address or resolvable hostname for remoting |
| `targets.windows` | `username` / `password` | Dedicated service account used by Turbonomic for WMI/WinRM |
| `targets.windows` | `sql_username` / `sql_password` | Local Admin / SQL sysadmin account for pre-flight SQL login provisioning |
| `targets.windows` | `domain` | AD Domain or `WORKGROUP` |
| `targets.linux` | `ssh_key_file` | Path to private SSH key for remote authentication |
| `targets.linux` | `sql_username` / `sql_password` | SQL Server authentication login (e.g. `sa`) |

---

## Step-by-Step Execution Guide

### Step 1: Target VM Hardening & Preparation

Run on each target Windows VM as **Administrator** (or deploy via AWS Systems Manager / Azure Run Command / Ansible):

```powershell
# Set execution policy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Execute setup with target service credentials
.\target-setup\Set-WindowsPrerequisites.ps1 -Username "turbowmid16" -Password "TrbWmi#2026!X9@Q7"
```

*Note: Restart the Windows VM after running to apply the `LocalAccountTokenFilterPolicy` registry key and WMI security descriptors.*

---

### Step 2: Turbonomic Target Registration & Discovery

Run from your central management workstation or bastion host:

```bash
# Verify Python requirements (Linux / macOS)
pip install requests urllib3

# Verify Python requirements (Windows)
python -m pip install requests urllib3

# Run target onboarding (Linux / macOS)
python3 target-setup/register_turbo_targets.py

# Run target onboarding (Windows)
py.exe target-setup\register_turbo_targets.py
```

Optional flags:
```bash
# Specify custom config path
python3 target-setup/register_turbo_targets.py --config config/config.json
```

---

### Step 3: Fleet Reporting & Analytics Dashboard

#### 3a: Install PowerShell 7 (Required for Linux target SSH collection)

PowerShell 7 is required if your `config.json` includes Linux targets (`"OS": "Linux"`).
Windows-only fleets can use PowerShell 5.1, but PS 7 is recommended for all environments.

**Check your current version first:**

```powershell
$PSVersionTable.PSVersion
# Major = 7 → already installed, skip to Step 3b
# Major = 5 → install PS 7 using the steps below
```

**Install PowerShell 7 (no `winget` required) — run in your existing PS 5.1 session:**

```powershell
# Download the latest PS 7 MSI for Windows x64
$msi = "$env:TEMP\pwsh7.msi"
Invoke-WebRequest -Uri "https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/PowerShell-7.4.6-win-x64.msi" `
                  -OutFile $msi -UseBasicParsing

# Install silently (no reboot required)
Start-Process msiexec.exe -ArgumentList "/i `"$msi`" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1" -Wait

# Verify installation
& "C:\Program Files\PowerShell\7\pwsh.exe" -Command '$PSVersionTable.PSVersion'
```

Expected output:
```
Major  Minor  Patch  PreReleaseLabel BuildLabel
-----  -----  -----  --------------- ----------
7      4      6
```

> **No internet access?** Download the MSI from another machine and copy it over:
> `https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/PowerShell-7.4.6-win-x64.msi`
> Then run: `Start-Process msiexec.exe -ArgumentList '/i "C:\path\to\PowerShell-7.4.6-win-x64.msi" /quiet' -Wait`

---

#### 3b: Run the Fleet Report

Run from your central management workstation using **PowerShell 7** (`pwsh.exe`):

```powershell
# Run full fleet analysis (default 1-day historical period)
& "C:\Program Files\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -File .\Invoke-FleetReport.ps1

# Run with custom historical peak window (e.g. 7 days) and custom output folder
& "C:\Program Files\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -File .\Invoke-FleetReport.ps1 -Days 7 -OutputDirectory .\fleet-output
```

Or open a PS 7 session first, then run directly:

```powershell
# Open PS 7 shell
& "C:\Program Files\PowerShell\7\pwsh.exe"

# Then inside the PS 7 prompt:
cd C:\scripts
.\Invoke-FleetReport.ps1
```

Optional switches:
- `-SkipSqlLoginSetup`: Skips the pre-flight SQL login creation step.
- `-SkipSqlInventory`: Runs only the Turbonomic metrics collection.
- `-SkipTurbo`: Runs only the SQL Server inventory collection.

---

#### 3c: Fix SSH Key File Permissions (Required for Linux targets)

PowerShell SSH remoting rejects private key files that have overly permissive ACLs.
Run these commands **once** to lock down your SSH key before invoking the report:

```powershell
# Set the path to your SSH private key (must match ssh_key_file in config.json)
$keyPath = "C:\scripts\keys\id_ed25519"

# 1. Remove all inherited permissions
icacls.exe $keyPath /inheritance:r

# 2. Grant only the current user Read & Write — strip all other accounts
icacls.exe $keyPath /grant:r "$($env:USERNAME):(R,W)"

# 3. Verify — only your account should appear
icacls.exe $keyPath
```

Expected output after step 3:
```
C:\scripts\keys\id_ed25519  DESKTOP-XXXX\azureuser:(R,W)
Successfully processed 1 files; Failed processing 0 files
```

> **Why this is needed:** OpenSSH (used by PowerShell 7 SSH remoting) follows the same
> strict permission rules as Linux — it refuses to use a private key readable by
> anyone other than the owning user. If you skip this step, `Invoke-Command` over SSH
> will fail with `Permission denied (publickey)` or silently fall back to password auth.

---

## Standalone Collector Execution

Both collectors can be executed independently if needed:

### SQL Server Inventory Collector
```powershell
.\collectors\Get-SqlServerInventory.ps1 -InputFile .\config\config.json -OutputDirectory .\output
```

### Turbonomic Metrics Collector
```bash
# Collect metrics for all VMs defined in config.json
python3 collectors/get_turbonomic_metrics.py --config config/config.json --days 7

# Query a single specific VM
python3 collectors/get_turbonomic_metrics.py --vm turbo-wmi-win-d16 --days 30
```

---

## Generated Outputs & Artifacts

All outputs are structured, timestamped, and saved in `output/` (or `-OutputDirectory`):

```text
fleet-output/
├── fleet-report-YYYYMMDD-HHMMSS.html       # Master Unified Executive Dashboard
├── fleet-report-YYYYMMDD-HHMMSS.log        # Master pipeline execution log
│
├── sql/
│   ├── sql-inventory-YYYYMMDD-HHMMSS.json  # Raw SQL Server & host inventory JSON
│   ├── sql-inventory-YYYYMMDD-HHMMSS.html  # Dedicated SQL Server HTML dashboard
│   ├── sql-inventory-instances-*.csv       # SQL instances flattened CSV
│   └── sql-inventory-databases-*.csv       # Databases detailed inventory CSV
│
└── turbo/
    ├── turbo-vm-metrics-YYYYMMDD-HHMMSS.json # Turbonomic commodity stats & actions JSON
    ├── turbo-vm-metrics-YYYYMMDD-HHMMSS.csv  # Sizing, peak utilization & savings CSV
    └── turbo-vm-metrics-YYYYMMDD-HHMMSS.html # Dedicated Turbonomic HTML report
```

---

## Security & Authentication Principles

1. **Principle of Least Privilege:**
   - Turbonomic does not require domain-admin rights. Target VMs use dedicated restricted service accounts with explicit namespace permissions (`Execute Methods`, `Enable Account`, `Remote Enable`).
2. **Credential Separation:**
   - `username` / `password`: Scoped strictly for ongoing WMI/WinRM metrics collection.
   - `sql_username` / `sql_password`: Used exclusively during pre-flight to provision the minimum read-only permissions (`VIEW SERVER STATE`, `VIEW ANY DATABASE`, `VIEW ANY DEFINITION`) for the service account.
3. **Redacted Logs:**
   - All REST API payloads automatically redact sensitive password fields in debug logs.

---

## Troubleshooting & FAQ

### 1. `Set-WindowsPrerequisites.ps1` permissions not showing in `wmimgmt.msc`
The script uses low-level binary security descriptors via `[wmiclass]"\\.\root:__SystemSecurity"::SetSD`. If checking manually in `wmimgmt.msc`, click **Advanced** under the **Security** tab to view inherited permissions for the user.

### 2. Turbonomic Returns 504 Gateway Timeout during target addition
The script includes automatic retry logic with exponential backoff and verifies via `GET /targets` whether the target was already persisted before issuing another request.

### 3. Linux Target Remoting Requirements
Linux targets require PowerShell 7+ on the management workstation when invoking `Get-SqlServerInventory.ps1` via SSH remoting (`-HostName`, `-UserName`, `-KeyFilePath`).

---

## Requirements

- **Operating System:** Windows Server 2012+ / Windows 10+ / Linux / macOS (for Python collectors).
- **PowerShell:** Windows PowerShell 5.1 or PowerShell 7+.
- **Python:** Python 3.9+ with `requests` and `urllib3` (`pip install requests urllib3`).
- **Turbonomic:** Turbonomic 8.x+ with WMI probe enabled.
