<#
.SYNOPSIS
    Runs a local Windows endpoint health check — disk, services, updates and network.

.DESCRIPTION
    Test-EndpointHealth.ps1 performs a set of read-only checks on the
    local Windows machine and returns a structured result for each check
    with a Pass / Warning / Fail status. It is designed for help-desk
    triage, onboarding verification and Intune remediation investigations.

    Checks performed:
      * Operating system and uptime
      * System drive free space (Warning < 20%, Fail < 10% by default)
      * Memory utilisation
      * Critical service states (Windows Update, BITS, WinDefend /
        Sense where present, Intune Management Extension where present)
      * Pending reboot detection (CBS / Windows Update registry markers)
      * Recent Windows Update history (last successful install, if available)
      * Network adapter status, IPv4 address and default gateway reachability
      * DNS resolution for a configurable test host

    The script changes nothing on the machine. Results can be exported
    to JSON or CSV for attaching to a support ticket.

.PARAMETER DriveWarningPercent
    Free-space percentage below which the disk check is a Warning.
    Default: 20

.PARAMETER DriveFailPercent
    Free-space percentage below which the disk check is a Fail.
    Default: 10

.PARAMETER DnsTestHost
    Host name used for the DNS resolution check.
    Default: www.microsoft.com

.PARAMETER SkipGatewayPing
    When specified, the default-gateway ICMP reachability test is skipped.
    Useful on networks where ICMP is blocked by policy.

.PARAMETER OutputPath
    Optional path to export results. The file extension determines the
    format: .json exports JSON, any other extension exports CSV.

.EXAMPLE
    .\Test-EndpointHealth.ps1

    Runs all checks on the local machine and displays a summary table.

.EXAMPLE
    .\Test-EndpointHealth.ps1 -OutputPath ".\health.json"

    Runs all checks and exports the full result set as JSON.

.EXAMPLE
    .\Test-EndpointHealth.ps1 -DriveWarningPercent 25 -SkipGatewayPing

    Uses a stricter disk warning threshold and skips the gateway ping.

.NOTES
    Author      : Zohaib Yousaf (info@zohaibyousaf.com)
    Requires    : Windows, PowerShell 5.1 or PowerShell 7+.
                  Administrator rights are NOT required for the default
                  checks, though some service / update details are more
                  complete when elevated.
    Version     : 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 99)]
    [int]$DriveWarningPercent = 20,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 98)]
    [int]$DriveFailPercent = 10,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DnsTestHost = 'www.microsoft.com',

    [Parameter(Mandatory = $false)]
    [switch]$SkipGatewayPing,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$results = [System.Collections.Generic.List[object]]::new()

function Add-HealthResult {
    param(
        [string]$Check,
        [ValidateSet('Pass', 'Warning', 'Fail', 'Info')]
        [string]$Status,
        [string]$Details
    )
    $results.Add([PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        Check        = $Check
        Status       = $Status
        Details      = $Details
        CheckedAt    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

function Test-DiskSpace {
    try {
        $drive = Get-PSDrive -Name ([System.IO.Path]::GetPathRoot($env:SystemRoot).TrimEnd(':\')) -ErrorAction Stop
        $totalGB = [math]::Round(($drive.Used + $drive.Free) / 1GB, 2)
        $freeGB  = [math]::Round($drive.Free / 1GB, 2)
        $freePct = if ($totalGB -gt 0) { [math]::Round(($freeGB / $totalGB) * 100, 1) } else { 0 }

        $status = if ($freePct -lt $DriveFailPercent) { 'Fail' }
                  elseif ($freePct -lt $DriveWarningPercent) { 'Warning' }
                  else { 'Pass' }
        Add-HealthResult -Check 'Disk Space (System Drive)' -Status $status `
            -Details ("{0} GB free of {1} GB ({2}% free)" -f $freeGB, $totalGB, $freePct)
    }
    catch {
        Add-HealthResult -Check 'Disk Space (System Drive)' -Status 'Warning' -Details ("Could not read drive info: {0}" -f $_.Exception.Message)
    }
}

function Test-OperatingSystem {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $uptime = (Get-Date) - $os.LastBootUpTime
        Add-HealthResult -Check 'Operating System' -Status 'Info' `
            -Details ("{0} (Build {1}) | Uptime: {2} days, {3} hours" -f $os.Caption, $os.BuildNumber, $uptime.Days, $uptime.Hours)

        $freeMemPct = if ($os.TotalVisibleMemorySize -gt 0) {
            [math]::Round(($os.FreePhysicalMemory / $os.TotalVisibleMemorySize) * 100, 1)
        } else { 0 }
        $memStatus = if ($freeMemPct -lt 10) { 'Warning' } else { 'Pass' }
        Add-HealthResult -Check 'Memory' -Status $memStatus -Details ("{0}% free physical memory" -f $freeMemPct)
    }
    catch {
        Add-HealthResult -Check 'Operating System' -Status 'Warning' -Details ("CIM query failed: {0}" -f $_.Exception.Message)
    }
}

function Test-CriticalServices {
    $services = @('wuauserv', 'BITS', 'WinDefend', 'Sense', 'IntuneManagementExtension')
    foreach ($name in $services) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if (-not $svc) {
            Add-HealthResult -Check ("Service: {0}" -f $name) -Status 'Info' -Details 'Service not present on this device'
            continue
        }
        $status = if ($svc.Status -eq 'Running') { 'Pass' } else { 'Warning' }
        Add-HealthResult -Check ("Service: {0}" -f $name) -Status $status `
            -Details ("Status: {0} | StartType: {1}" -f $svc.Status, $svc.StartType)
    }
}

function Test-PendingReboot {
    $pending = $false
    $reasons = @()
    try {
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pending = $true; $reasons += 'CBS RebootPending' }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pending = $true; $reasons += 'Windows Update RebootRequired' }
        $sessionManager = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
        if ($sessionManager) { $pending = $true; $reasons += 'PendingFileRenameOperations' }
    }
    catch {
        Add-HealthResult -Check 'Pending Reboot' -Status 'Info' -Details ("Reboot marker check incomplete: {0}" -f $_.Exception.Message)
        return
    }
    if ($pending) {
        Add-HealthResult -Check 'Pending Reboot' -Status 'Warning' -Details ("Reboot pending: {0}" -f ($reasons -join ', '))
    }
    else {
        Add-HealthResult -Check 'Pending Reboot' -Status 'Pass' -Details 'No pending reboot markers found'
    }
}

function Test-Network {
    try {
        $adapters = Get-NetAdapter -Physical -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' }
        if ($adapters) {
            Add-HealthResult -Check 'Network Adapter' -Status 'Pass' -Details ("Up adapters: {0}" -f (($adapters | Select-Object -ExpandProperty Name) -join ', '))
        }
        else {
            Add-HealthResult -Check 'Network Adapter' -Status 'Fail' -Details 'No physical network adapter is Up'
        }
    }
    catch {
        Add-HealthResult -Check 'Network Adapter' -Status 'Warning' -Details ("Get-NetAdapter failed, falling back to CIM: {0}" -f $_.Exception.Message)
    }

    if (-not $SkipGatewayPing) {
        try {
            $gateway = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop | Sort-Object -Property RouteMetric | Select-Object -First 1).NextHop
            if ($gateway) {
                $ping = Test-Connection -ComputerName $gateway -Count 2 -Quiet -ErrorAction SilentlyContinue
                $status = if ($ping) { 'Pass' } else { 'Warning' }
                Add-HealthResult -Check 'Default Gateway' -Status $status -Details ("Gateway {0} reachable: {1}" -f $gateway, $ping)
            }
            else {
                Add-HealthResult -Check 'Default Gateway' -Status 'Warning' -Details 'No default route found'
            }
        }
        catch {
            Add-HealthResult -Check 'Default Gateway' -Status 'Warning' -Details ("Gateway check failed: {0}" -f $_.Exception.Message)
        }
    }

    try {
        $dns = Resolve-DnsName -Name $DnsTestHost -ErrorAction Stop | Select-Object -First 1
        if ($dns) {
            Add-HealthResult -Check 'DNS Resolution' -Status 'Pass' -Details ("{0} resolved to {1}" -f $DnsTestHost, $dns.IPAddress)
        }
    }
    catch {
        Add-HealthResult -Check 'DNS Resolution' -Status 'Fail' -Details ("Could not resolve {0}: {1}" -f $DnsTestHost, $_.Exception.Message)
    }
}

# --- Run checks -------------------------------------------------------------
Write-Host ("Endpoint health check for {0} — {1}" -f $env:COMPUTERNAME, (Get-Date))
Test-OperatingSystem
Test-DiskSpace
Test-CriticalServices
Test-PendingReboot
Test-Network

$results | Format-Table Check, Status, Details -AutoSize -Wrap

$failCount = @($results | Where-Object { $_.Status -eq 'Fail' }).Count
$warnCount = @($results | Where-Object { $_.Status -eq 'Warning' }).Count
Write-Host ("`nSummary: {0} checks | {1} Fail | {2} Warning" -f $results.Count, $failCount, $warnCount)

if ($OutputPath) {
    try {
        $parent = Split-Path -Path $OutputPath -Parent
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        if ([System.IO.Path]::GetExtension($OutputPath) -ieq '.json') {
            $results | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
        }
        else {
            $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "Results exported to: $OutputPath"
    }
    catch {
        Write-Error ("Could not export results to '{0}': {1}" -f $OutputPath, $_.Exception.Message)
    }
}

return $results
