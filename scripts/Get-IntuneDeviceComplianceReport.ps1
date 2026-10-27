<#
.SYNOPSIS
    Exports a Microsoft Intune managed-device compliance report via Microsoft Graph.

.DESCRIPTION
    Get-IntuneDeviceComplianceReport.ps1 retrieves managed devices from
    Microsoft Graph (Get-MgDeviceManagementManagedDevice) and builds a
    compliance-focused report suitable for IT operations reviews, audits
    and help-desk follow-up.

    Each record includes device name, user, operating system, compliance
    state, enrolment type, last sync time, and whether the last sync is
    considered stale for the threshold supplied.

    Results can be filtered by compliance state and operating system, and
    exported to CSV. The script is read-only.

.PARAMETER ComplianceState
    Optional filter. Accepted values: All, Compliant, Noncompliant,
    InGracePeriod, Error, Conflict, ConfigManager.
    Default: All

.PARAMETER OperatingSystem
    Optional operating system filter, e.g. Windows, iOS, Android, macOS.
    Matching is case-insensitive and partial.

.PARAMETER StaleSyncDays
    Number of days after which a device that has not synced is flagged as
    stale. Default: 30

.PARAMETER OutputPath
    Optional path to a CSV file to export the report to.

.PARAMETER DisconnectOnExit
    When specified, Disconnect-MgGraph is called when the script finishes.

.EXAMPLE
    .\Get-IntuneDeviceComplianceReport.ps1

    Lists all Intune managed devices with their compliance state.

.EXAMPLE
    .\Get-IntuneDeviceComplianceReport.ps1 -ComplianceState Noncompliant -OutputPath ".\noncompliant.csv"

    Exports only non-compliant devices to CSV for remediation follow-up.

.EXAMPLE
    .\Get-IntuneDeviceComplianceReport.ps1 -OperatingSystem Windows -StaleSyncDays 14

    Shows Windows devices and flags those not synced in more than 14 days.

.NOTES
    Author      : Zohaib Yousaf (info@zohaibyousaf.com)
    Requires    : Microsoft.Graph.DeviceManagement (v2.x),
                  PowerShell 5.1 or PowerShell 7+
    Permissions : DeviceManagementManagedDevices.Read.All
    Version     : 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'Compliant', 'Noncompliant', 'InGracePeriod', 'Error', 'Conflict', 'ConfigManager')]
    [string]$ComplianceState = 'All',

    [Parameter(Mandatory = $false)]
    [string]$OperatingSystem,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$StaleSyncDays = 30,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$DisconnectOnExit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-GraphConnection {
    $context = $null
    try { $context = Get-MgContext -ErrorAction SilentlyContinue } catch { $context = $null }
    if (-not $context) {
        Write-Verbose "No Microsoft Graph session found. Connecting..."
        Connect-MgGraph -Scopes 'DeviceManagementManagedDevices.Read.All' -NoWelcome
    }
    else {
        Write-Verbose ("Connected as '{0}' (Tenant: {1})" -f $context.Account, $context.TenantId)
    }
}

try {
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Microsoft Graph PowerShell module was not found. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Ensure-GraphConnection

    Write-Verbose "Retrieving Intune managed devices..."
    $devices = Get-MgDeviceManagementManagedDevice -All
    if (-not $devices) {
        Write-Warning "No managed devices were returned."
        return
    }

    $now = Get-Date
    $report = foreach ($device in $devices) {
        $state = [string]$device.ComplianceState
        if ($ComplianceState -ne 'All' -and $state -ine $ComplianceState) { continue }
        if ($OperatingSystem -and ([string]$device.OperatingSystem) -notlike "*$OperatingSystem*") { continue }

        $lastSync = $device.LastSyncDateTime
        $daysSinceSync = $null
        $isStale = $false
        if ($lastSync) {
            $daysSinceSync = [math]::Floor(($now - $lastSync).TotalDays)
            $isStale = $daysSinceSync -gt $StaleSyncDays
        }

        [PSCustomObject]@{
            DeviceName        = $device.DeviceName
            UserPrincipalName = $device.UserPrincipalName
            UserDisplayName   = $device.UserDisplayName
            OperatingSystem   = $device.OperatingSystem
            OsVersion         = $device.OsVersion
            ComplianceState   = $state
            ManagedDeviceType = $device.ManagedDeviceOwnershipType
            EnrolledDateTime  = $device.EnrolledDateTime
            LastSyncDateTime  = $lastSync
            DaysSinceSync     = $daysSinceSync
            IsStale           = $isStale
            SerialNumber      = $device.SerialNumber
            AzureADDeviceId   = $device.AzureAdDeviceId
            Id                = $device.Id
        }
    }

    $report = @($report | Sort-Object -Property ComplianceState, DeviceName)
    if ($report.Count -eq 0) {
        Write-Warning "No devices matched the supplied filters."
        return
    }

    $report | Format-Table DeviceName, UserPrincipalName, OperatingSystem, ComplianceState, LastSyncDateTime, IsStale -AutoSize

    $grouped = $report | Group-Object -Property ComplianceState | Sort-Object -Property Name
    Write-Host "`nCompliance breakdown:"
    foreach ($group in $grouped) {
        Write-Host ("  {0}: {1}" -f $group.Name, $group.Count)
    }
    Write-Host ("  Total: {0} | Stale (>{1} days): {2}" -f $report.Count, $StaleSyncDays, @($report | Where-Object { $_.IsStale }).Count)

    if ($OutputPath) {
        $parent = Split-Path -Path $OutputPath -Parent
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: $OutputPath"
    }

    return $report
}
catch {
    Write-Error ("Get-IntuneDeviceComplianceReport failed: {0}" -f $_.Exception.Message)
    throw
}
finally {
    if ($DisconnectOnExit) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}
