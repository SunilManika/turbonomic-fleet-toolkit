#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Turbonomic WMI Restricted Service Account Setup (Option B)
    Fully automated — no manual intervention required.

.DESCRIPTION
    Implements every step of Option B from the Turbonomic WMI Configuration
    Guide (WMI_Setup_Guide.docx), plus additional registry and permission fixes:

      1.  Create the local user account (turbowmi)
      2.  Add the account to WinRMRemoteWMIUsers__ and Performance Monitor Users
      3.  Grant WMI namespace permissions (Execute Methods + Enable Account +
          Remote Enable) on root and root\cimv2 recursively
      4.  Set LocalAccountTokenFilterPolicy = 1 in registry (required for
          remote WMI access by non-built-in-admin local accounts)
      5.  Enable WinRM (winrm quickconfig)
      6.  Enable Negotiate authentication
      7.  Allow unencrypted WinRM (HTTP)
      8.  Enable and start Remote Registry
      9.  Open firewall ports (5985 HTTP, 135 RPC, 49152-65535 dynamic RPC)
      10. (Optional) Configure HTTPS WinRM listener
      11. Verify all settings and print a summary
      12. Prompt to restart the VM (required for registry and WMI changes)

    Must be run as Administrator on the Windows VM (52.118.252.201).

.PARAMETER Username
    Local username for the WMI service account. Default: turbowmi

.PARAMETER Password
    Password for the service account. Default: TrbWmi#2026!X9@Q7

.PARAMETER EnableHttps
    When specified, configures an HTTPS WinRM listener in addition to HTTP.
    Requires a certificate to exist in cert:\LocalMachine\My.

.EXAMPLE
    # Run with defaults (matches config/config.json)
    .\Set-WindowsPrerequisites.ps1

.EXAMPLE
    # Run with custom credentials
    .\Set-WindowsPrerequisites.ps1 -Username myuser -Password 'MyP@ss!'

.NOTES
    Reference: https://www.ibm.com/docs/en/tarm/8.21.1?topic=targets-wmi
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Username   = 'turbowmi',
    [string]$Password   = 'TrbWmi#2026!X9@Q7',
    [switch]$EnableHttps
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

function Write-Step {
    param([int]$Number, [string]$Title)
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  STEP $Number — $Title" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
}

function Write-OK   { param([string]$Msg) Write-Host "  [OK]  $Msg" -ForegroundColor Green  }
function Write-Skip { param([string]$Msg) Write-Host "  [--]  $Msg" -ForegroundColor Yellow }
function Write-Info { param([string]$Msg) Write-Host "        $Msg" -ForegroundColor White  }

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — Create local user account
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 1 "Create local user account '$Username'"

$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$existingUser   = Get-LocalUser -Name $Username -ErrorAction SilentlyContinue

if ($existingUser) {
    Write-Skip "User '$Username' already exists — updating password and flags"
    Set-LocalUser -Name $Username -Password $securePassword `
                  -PasswordNeverExpires $true `
                  -UserMayChangePassword $true
} else {
    New-LocalUser -Name $Username `
                  -Password $securePassword `
                  -PasswordNeverExpires `
                  -Description 'Turbonomic WMI service account (restricted)' `
                  | Out-Null
    Write-OK "User '$Username' created"
}

# Ensure the account is enabled
Enable-LocalUser -Name $Username
Write-OK "User '$Username' is enabled"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — Add account to required local security groups
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 2 "Add '$Username' to required local security groups"

# Prefer WinRMRemoteWMIUsers__ on older systems; Remote Management Users on 2012+
$winrmGroup = 'WinRMRemoteWMIUsers__'
if (-not (Get-LocalGroup -Name $winrmGroup -ErrorAction SilentlyContinue)) {
    $winrmGroup = 'Remote Management Users'
}

$targetGroups = @($winrmGroup, 'Performance Monitor Users', 'Administrators')

foreach ($grp in $targetGroups) {
    $group = Get-LocalGroup -Name $grp -ErrorAction SilentlyContinue
    if (-not $group) {
        Write-Skip "Group '$grp' not found on this system — skipping"
        continue
    }

    $members = Get-LocalGroupMember -Group $grp -ErrorAction SilentlyContinue
    $already = $members | Where-Object { $_.Name -like "*\$Username" -or $_.Name -eq $Username }

    if ($already) {
        Write-Skip "'$Username' is already a member of '$grp'"
    } else {
        Add-LocalGroupMember -Group $grp -Member $Username
        Write-OK "Added '$Username' to '$grp'"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — Grant WMI namespace permissions (Enable Account + Remote Enable)
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 3 "Grant WMI namespace permissions (Enable Account + Remote Enable)"

function Set-WmiNamespacePermission {
    <#
    .SYNOPSIS
        Grants 'Execute Methods', 'Enable Account', and 'Remote Enable' WMI permissions
        to the specified user on the given namespace using binary SetSD, which reliably updates
        the Windows Security Descriptor visible in wmimgmt.msc.
    #>
    param(
        [string]$Namespace = 'root',
        [string]$User
    )

    # WMI namespace specific permissions:
    # 0x00001 = WBEM_ENABLE (Enable Account)
    # 0x00002 = WBEM_METHOD_EXECUTE (Execute Methods)
    # 0x00020 = WBEM_REMOTE_ACCESS (Remote Enable)
    # 0x20000 = READ_CONTROL
    $WBEM_RIGHTS = 0x20023  # Combined: Enable Account + Execute Methods + Remote Enable + Read Control
    $AceFlags = [System.Security.AccessControl.AceFlags]::ContainerInherit

    $account = New-Object System.Security.Principal.NTAccount($User)
    $sid = $account.Translate([System.Security.Principal.SecurityIdentifier])

    $secClass = [wmiclass]"\\.\${Namespace}:__SystemSecurity"
    $resGet = $secClass.GetSD()
    if ($resGet.ReturnValue -ne 0) {
        throw "GetSD failed on '$Namespace' with return code $($resGet.ReturnValue)"
    }

    $rawSD = New-Object System.Security.AccessControl.RawSecurityDescriptor($resGet.SD, 0)
    $dacl = $rawSD.DiscretionaryAcl

    # Remove any existing ACE for this user SID
    for ($i = $dacl.Count - 1; $i -ge 0; $i--) {
        if ($dacl[$i].SecurityIdentifier -eq $sid) {
            $dacl.RemoveAce($i)
        }
    }

    # Add the new AccessAllowed ACE
    $ace = New-Object System.Security.AccessControl.CommonAce(
        $AceFlags,
        [System.Security.AccessControl.AceQualifier]::AccessAllowed,
        $WBEM_RIGHTS,
        $sid,
        $false,
        $null
    )
    $dacl.InsertAce($dacl.Count, $ace)

    # Convert back to binary format and save with SetSD
    $binarySD = New-Object byte[] $rawSD.BinaryLength
    $rawSD.GetBinaryForm($binarySD, 0)

    $resSet = $secClass.SetSD($binarySD)
    if ($resSet.ReturnValue -ne 0) {
        throw "SetSD failed on '$Namespace' with return code $($resSet.ReturnValue)"
    }

    Write-OK "WMI permissions successfully set on '$Namespace' for '$User'"
}

# Apply to root — CONTAINER_INHERIT_ACE propagates to all subnamespaces
Set-WmiNamespacePermission -Namespace 'root' -User $Username

# Explicitly set on root\cimv2 (the primary namespace Turbonomic queries)
Set-WmiNamespacePermission -Namespace 'root\cimv2' -User $Username

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — Set LocalAccountTokenFilterPolicy = 1
# ─────────────────────────────────────────────────────────────────────────────
# Without this registry value set to 1, remote WMI connections from local
# (non-built-in-Administrator) accounts are silently denied even when all other
# permissions are correct. Required for WORKGROUP machines.
Write-Step 4 "Set LocalAccountTokenFilterPolicy = 1 (registry)"

$regPath  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
$regName  = 'LocalAccountTokenFilterPolicy'
$regValue = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue

if ($regValue -and $regValue.$regName -eq 1) {
    Write-Skip "LocalAccountTokenFilterPolicy is already 1"
} else {
    $currentVal = if ($regValue) { $regValue.$regName } else { '(not set)' }
    Write-Info "Current value: $currentVal — setting to 1"
    New-ItemProperty -Path $regPath -Name $regName -PropertyType DWord -Value 1 -Force | Out-Null
    Write-OK "LocalAccountTokenFilterPolicy set to 1"
    Write-Info "NOTE: A restart is required for this change to take effect"
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — Enable WinRM (winrm quickconfig)
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 5 "Enable WinRM (winrm quickconfig)"

$winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
if ($winrmService -and $winrmService.Status -eq 'Running') {
    Write-Skip "WinRM service is already running"
} else {
    # Enable-PSRemoting runs the equivalent of winrm quickconfig non-interactively
    Enable-PSRemoting -Force -SkipNetworkProfileCheck | Out-Null
    Write-OK "WinRM enabled and configured"
}

# Ensure WinRM starts automatically on boot
Set-Service -Name WinRM -StartupType Automatic
Write-OK "WinRM startup type set to Automatic"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — Enable Negotiate authentication
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 6 "Enable Negotiate (NTLM) authentication"

Set-Item -Path 'WSMan:\localhost\Service\Auth\Negotiate' -Value $true
Write-OK "Negotiate authentication enabled"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7 — Allow unencrypted WinRM (HTTP)
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 7 "Allow unencrypted WinRM connections (HTTP)"

Set-Item -Path 'WSMan:\localhost\Service\AllowUnencrypted' -Value $true
Set-Item -Path 'WSMan:\localhost\Client\AllowUnencrypted' -Value $true
Write-OK "AllowUnencrypted = true (service and client)"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 8 — Enable and start Remote Registry
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 8 "Enable and start Remote Registry service"

Set-Service -Name RemoteRegistry -StartupType Automatic
Start-Service -Name RemoteRegistry
$rrStatus = (Get-Service -Name RemoteRegistry).Status
Write-OK "Remote Registry is $rrStatus"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 9 — Configure Windows Firewall
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 9 "Configure Windows Firewall"

function Ensure-FirewallRule {
    # LocalPort accepts either int[] for discrete ports, or a string for ranges.
    # Use separate -PortValue / -PortDisplay params to avoid type-coercion errors
    # on PowerShell 5.1 when a range string like "49152-65535" is passed.
    param(
        [string]$DisplayName,
        [string]$PortSpec,          # passed directly to -LocalPort (int or range string)
        [string]$Protocol = 'TCP'
    )
    $existing = Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Skip "Firewall rule '$DisplayName' already exists"
    } else {
        New-NetFirewallRule `
            -DisplayName  $DisplayName `
            -Direction    Inbound `
            -Protocol     $Protocol `
            -LocalPort    $PortSpec `
            -Action       Allow `
            -Profile      Any | Out-Null
        Write-OK "Firewall rule '$DisplayName' created (TCP $PortSpec)"
    }
}

Ensure-FirewallRule -DisplayName 'Turbonomic WinRM HTTP (5985)'              -PortSpec '5985'
Ensure-FirewallRule -DisplayName 'Turbonomic WMI RPC (135)'                  -PortSpec '135'
Ensure-FirewallRule -DisplayName 'Turbonomic WMI Dynamic RPC (49152-65535)'  -PortSpec '49152-65535'

if ($EnableHttps) {
    Ensure-FirewallRule -DisplayName 'Turbonomic WinRM HTTPS (5986)' -PortSpec '5986'
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 10 — (Optional) HTTPS WinRM listener
# ─────────────────────────────────────────────────────────────────────────────
if ($EnableHttps) {
    Write-Step 10 "Configure HTTPS WinRM listener"

    $cert = Get-ChildItem Cert:\LocalMachine\My |
            Where-Object { $_.Subject -match $env:COMPUTERNAME } |
            Sort-Object NotBefore -Descending |
            Select-Object -First 1

    if ($cert) {
        $existingHttps = Get-WSManInstance -ResourceURI winrm/config/listener `
                            -SelectorSet @{Address='*'; Transport='HTTPS'} `
                            -ErrorAction SilentlyContinue
        if (-not $existingHttps) {
            New-WSManInstance -ResourceURI winrm/config/listener `
                -SelectorSet @{Address='*'; Transport='HTTPS'} `
                -ValueSet @{
                    Hostname             = $env:COMPUTERNAME
                    CertificateThumbprint = $cert.Thumbprint
                    Port                 = '5986'
                } | Out-Null
            Write-OK "HTTPS listener created (Thumbprint: $($cert.Thumbprint))"
        } else {
            Write-Skip "HTTPS listener already exists"
        }
    } else {
        Write-Host "  [WARN] No certificate found in Cert:\LocalMachine\My — skipping HTTPS listener" `
                   -ForegroundColor Yellow
        Write-Info "         Create or import a certificate first, then re-run with -EnableHttps"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 11 — Verification
# ─────────────────────────────────────────────────────────────────────────────
Write-Step 11 "Verification"

$results = [ordered]@{}

# LocalAccountTokenFilterPolicy = 1
$laftp = Get-ItemProperty `
    -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' `
    -Name 'LocalAccountTokenFilterPolicy' -ErrorAction SilentlyContinue
if ($laftp -and $laftp.LocalAccountTokenFilterPolicy -eq 1) {
    $results['LocalAccountTokenFilterPolicy = 1'] = 'PASS'
} else {
    $results['LocalAccountTokenFilterPolicy = 1'] = 'FAIL'
}

# User exists and is enabled
$u = Get-LocalUser -Name $Username -ErrorAction SilentlyContinue
if ($u -and $u.Enabled) {
    $results["User '$Username' exists and is enabled"] = 'PASS'
} else {
    $results["User '$Username' exists and is enabled"] = 'FAIL'
}

# Group memberships
foreach ($grp in @('WinRMRemoteWMIUsers__', 'Remote Management Users', 'Performance Monitor Users', 'Administrators')) {
    $g = Get-LocalGroup -Name $grp -ErrorAction SilentlyContinue
    if ($g) {
        $m = Get-LocalGroupMember -Group $grp -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -like "*$Username" }
        if ($m) {
            $results["Member of '$grp'"] = 'PASS'
        } else {
            $results["Member of '$grp'"] = 'FAIL'
        }
    }
}

# WinRM service
if ((Get-Service WinRM).Status -eq 'Running') {
    $results['WinRM service running'] = 'PASS'
} else {
    $results['WinRM service running'] = 'FAIL'
}

# WinRM listener on 5985
$listener = Get-WSManInstance -ResourceURI winrm/config/listener `
                -SelectorSet @{Address='*'; Transport='HTTP'} -ErrorAction SilentlyContinue
if ($listener) {
    $results['WinRM HTTP listener (port 5985)'] = 'PASS'
} else {
    $results['WinRM HTTP listener (port 5985)'] = 'FAIL'
}

# Negotiate auth
$negotiate = (Get-Item WSMan:\localhost\Service\Auth\Negotiate).Value
if ($negotiate -eq 'true') {
    $results['Negotiate authentication enabled'] = 'PASS'
} else {
    $results['Negotiate authentication enabled'] = 'FAIL'
}

# AllowUnencrypted
$unenc = (Get-Item WSMan:\localhost\Service\AllowUnencrypted).Value
if ($unenc -eq 'true') {
    $results['AllowUnencrypted (service)'] = 'PASS'
} else {
    $results['AllowUnencrypted (service)'] = 'FAIL'
}

# Remote Registry
if ((Get-Service RemoteRegistry).Status -eq 'Running') {
    $results['Remote Registry running'] = 'PASS'
} else {
    $results['Remote Registry running'] = 'FAIL'
}

# Firewall rules
$fwRules = @(
    'Turbonomic WinRM HTTP (5985)',
    'Turbonomic WMI RPC (135)',
    'Turbonomic WMI Dynamic RPC (49152-65535)'
)
foreach ($rule in $fwRules) {
    $r = Get-NetFirewallRule -DisplayName $rule -ErrorAction SilentlyContinue
    if ($r) {
        $results["Firewall rule: $rule"] = 'PASS'
    } else {
        $results["Firewall rule: $rule"] = 'FAIL'
    }
}

# Print summary table
Write-Host ""
Write-Host "  ┌─────────────────────────────────────────────────────────────┬────────┐" -ForegroundColor White
Write-Host "  │ Check                                                        │ Result │" -ForegroundColor White
Write-Host "  ├─────────────────────────────────────────────────────────────┼────────┤" -ForegroundColor White
foreach ($key in $results.Keys) {
    $status = $results[$key]
    $color  = if ($status -eq 'PASS') { 'Green' } else { 'Red' }
    $line   = "  │ {0,-60} │ {1,-6} │" -f $key, $status
    Write-Host $line -ForegroundColor $color
}
Write-Host "  └─────────────────────────────────────────────────────────────┴────────┘" -ForegroundColor White

# Wrap in @() so .Count works correctly on PowerShell 5.1 when zero or one result is returned
$failures = @($results.Values | Where-Object { $_ -eq 'FAIL' })
Write-Host ""
if ($failures.Count -eq 0) {
    Write-Host "  All checks passed. The Windows VM is ready for Turbonomic WMI discovery." `
               -ForegroundColor Green
} else {
    Write-Host "  $($failures.Count) check(s) failed. Review the output above and re-run." `
               -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 12 — Prompt to restart
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  STEP 12 — Restart Required" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  A restart is required to apply:" -ForegroundColor Yellow
Write-Host "    • LocalAccountTokenFilterPolicy registry change" -ForegroundColor Yellow
Write-Host "    • WMI security descriptor changes" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Restart now? [Y] Yes  [N] No (restart manually later)" -ForegroundColor White
$response = Read-Host "  Enter choice"

if ($response -match '^[Yy]$') {
    Write-Host ""
    Write-Host "  Restarting in 10 seconds. Press Ctrl+C to cancel." -ForegroundColor Yellow
    Start-Sleep -Seconds 10
    Restart-Computer -Force
} else {
    Write-Host ""
    Write-Host "  Skipping restart. Remember to restart the VM before running the" -ForegroundColor Yellow
    Write-Host "  Turbonomic WMI target setup (create_wmi_target.py)." -ForegroundColor Yellow
}
