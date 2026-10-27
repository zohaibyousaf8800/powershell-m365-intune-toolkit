<#
.SYNOPSIS
    Finds stale computer accounts in Active Directory or Microsoft Entra ID.

.DESCRIPTION
    Get-StaleADDevices.ps1 reports computer / device accounts that have
    not signed in or synced for longer than the -DaysInactive threshold.

    Two data sources are supported:
      * ActiveDirectory — uses Get-ADComputer (RSAT / ActiveDirectory
        module) and the LastLogonDate property. This is the classic
        on-premises stale-account clean-up report.
      * EntraID — uses Microsoft Graph Get-MgDevice and the
        ApproximateLastSignInDateTime property, for cloud-joined devices.

    The script is read-only by default. It never disables or deletes an
    account. Use the report to review candidates with the asset owner
    before taking any clean-up action in a separate, approved change.

.PARAMETER Source
    Data source to query. Accepted values: ActiveDirectory, EntraID.
    Default: ActiveDirectory

.PARAMETER DaysInactive
    Inactivity threshold in days. Devices with no logon / sign-in at or
    after (today - DaysInactive) are reported as stale.
    Default: 90. Valid range: 1–3650.

.PARAMETER SearchBase
    Optional Active Directory SearchBase (distinguished name) to limit
    the AD query, e.g. "OU=Workstations,DC=contoso,DC=local".
    Ignored when -Source is EntraID.

.PARAMETER IncludeNeverLoggedOn
    When specified, accounts with no recorded logon date are included
    and flagged as NeverLoggedOn. Recommended for clean-up reviews.

.PARAMETER OutputPath
    Optional path to a CSV file to export the report to.

.EXAMPLE
    .\Get-StaleADDevices.ps1

    Lists AD computers with no logon in the last 90 days.

.EXAMPLE
    .\Get-StaleADDevices.ps1 -DaysInactive 180 -IncludeNeverLoggedOn -OutputPath ".\stale-180.csv"

    Exports computers inactive for 180+ days, including never-logged-on accounts.

.EXAMPLE
    .\Get-StaleADDevices.ps1 -Source EntraID -DaysInactive 60

    Lists Entra ID devices with no approximate sign-in in the last 60 days.

.NOTES
    Author      : Zohaib Yousaf (info@zohaibyousaf.com)
    Requires    : Source ActiveDirectory: RSAT ActiveDirectory module and
                  domain read permissions.
                  Source EntraID: Microsoft.Graph.Identity.DirectoryManagement
                  and Device.Read.All.
                  PowerShell 5.1 or PowerShell 7+.
    Version     : 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('ActiveDirectory', 'EntraID')]
    [string]$Source = 'ActiveDirectory',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$DaysInactive = 90,

    [Parameter(Mandatory = $false)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeNeverLoggedOn,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-StaleADComputers {
    param([datetime]$Cutoff, [string]$Base)

    if (-not (Get-Module -ListAvailable -Name 'ActiveDirectory')) {
        throw "ActiveDirectory module (RSAT) was not found. Install RSAT or run this script on a domain-joined management machine."
    }
    Import-Module ActiveDirectory -ErrorAction Stop

    $params = @{
        Filter     = '*'
        Properties = @('LastLogonDate', 'OperatingSystem', 'Enabled', 'Description', 'whenCreated')
    }
    if ($Base) { $params['SearchBase'] = $Base }

    $computers = Get-ADComputer @params
    foreach ($computer in $computers) {
        $last = $computer.LastLogonDate
        $never = ($null -eq $last)
        if ($never -and -not $IncludeNeverLoggedOn) { continue }
        if (-not $never -and $last -ge $Cutoff) { continue }

        [PSCustomObject]@{
            Name              = $computer.Name
            DNSHostName       = $computer.DNSHostName
            OperatingSystem   = $computer.OperatingSystem
            Enabled           = $computer.Enabled
            LastLogonDate     = $last
            DaysInactive      = if ($never) { $null } else { [math]::Floor(((Get-Date) - $last).TotalDays) }
            NeverLoggedOn     = $never
            Created           = $computer.whenCreated
            DistinguishedName = $computer.DistinguishedName
            Source            = 'ActiveDirectory'
        }
    }
}

function Get-StaleEntraDevices {
    param([datetime]$Cutoff)

    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Microsoft Graph PowerShell module was not found. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    $context = $null
    try { $context = Get-MgContext -ErrorAction SilentlyContinue } catch { $context = $null }
    if (-not $context) {
        Connect-MgGraph -Scopes 'Device.Read.All' -NoWelcome
    }

    $devices = Get-MgDevice -All
    foreach ($device in $devices) {
        $last = $device.ApproximateLastSignInDateTime
        $never = ($null -eq $last)
        if ($never -and -not $IncludeNeverLoggedOn) { continue }
        if (-not $never -and $last -ge $Cutoff) { continue }

        [PSCustomObject]@{
            Name              = $device.DisplayName
            DNSHostName       = ''
            OperatingSystem   = $device.OperatingSystem
            Enabled           = $device.AccountEnabled
            LastLogonDate     = $last
            DaysInactive      = if ($never) { $null } else { [math]::Floor(((Get-Date) - $last).TotalDays) }
            NeverLoggedOn     = $never
            Created           = $null
            DistinguishedName = $device.DeviceId
            Source            = 'EntraID'
        }
    }
}

try {
    $cutoff = (Get-Date).AddDays(-$DaysInactive)
    Write-Verbose ("Source: {0} | Cutoff: {1} ({2} days inactive)" -f $Source, $cutoff, $DaysInactive)

    $report = if ($Source -eq 'ActiveDirectory') {
        Get-StaleADComputers -Cutoff $cutoff -Base $SearchBase
    }
    else {
        Get-StaleEntraDevices -Cutoff $cutoff
    }

    $report = @($report | Sort-Object -Property LastLogonDate)
    if ($report.Count -eq 0) {
        Write-Host "No stale devices found for the supplied threshold."
        return
    }

    $report | Format-Table Name, OperatingSystem, Enabled, LastLogonDate, DaysInactive, NeverLoggedOn -AutoSize
    Write-Host ("`nTotal stale devices: {0} (threshold: {1} days, source: {2})" -f $report.Count, $DaysInactive, $Source)
    Write-Warning "Review only — no accounts were disabled or deleted by this script. Confirm ownership before any clean-up change."

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
    Write-Error ("Get-StaleADDevices failed: {0}" -f $_.Exception.Message)
    throw
}
