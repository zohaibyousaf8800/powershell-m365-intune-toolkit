<#
.SYNOPSIS
    Generates a Microsoft 365 licence usage summary for the connected tenant.

.DESCRIPTION
    Get-M365LicenseReport.ps1 connects to Microsoft Graph (using an existing
    session if one is present, otherwise Connect-MgGraph) and retrieves the
    subscribed SKUs for the tenant via Get-MgSubscribedSku.

    For every SKU it calculates:
      * Total licences purchased (PrepaidUnits.Enabled)
      * Licences consumed (ConsumedUnits)
      * Licences available (Total - Consumed)
      * Utilisation percentage
      * Suspended and warning units

    Results are displayed in the console and can optionally be exported to
    CSV for reporting, licence true-ups, or cost optimisation reviews.
    The script is read-only: it never modifies tenant data.

.PARAMETER OutputPath
    Optional path to a CSV file. When supplied, the report is exported to
    that location. Parent folders are created if they do not exist.

.PARAMETER IncludeDisabled
    When specified, SKUs with zero enabled units are included in the output.
    By default such SKUs are filtered out to keep reports focused.

.PARAMETER DisconnectOnExit
    When specified, Disconnect-MgGraph is called when the script finishes.

.EXAMPLE
    .\Get-M365LicenseReport.ps1

    Displays a licence usage table for the currently connected tenant.

.EXAMPLE
    .\Get-M365LicenseReport.ps1 -OutputPath "C:\Reports\M365-Licences.csv"

    Displays the report and exports it to the specified CSV file.

.EXAMPLE
    .\Get-M365LicenseReport.ps1 -IncludeDisabled -OutputPath ".\licences.csv"

    Includes SKUs with no enabled units and exports the full result set.

.NOTES
    Author      : Zohaib Yousaf (info@zohaibyousaf.com)
    Requires    : Microsoft.Graph.Identity.DirectoryManagement or
                  Microsoft.Graph (v2.x), PowerShell 5.1 or PowerShell 7+
    Permissions : Organization.Read.All, LicenseAssignment.Read.All
                  (delegated or application, depending on auth method)
    Version     : 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabled,

    [Parameter(Mandatory = $false)]
    [switch]$DisconnectOnExit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-GraphModule {
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Microsoft Graph PowerShell module was not found. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
}

function Ensure-GraphConnection {
    $context = $null
    try { $context = Get-MgContext -ErrorAction SilentlyContinue } catch { $context = $null }

    if (-not $context) {
        Write-Verbose "No Microsoft Graph session found. Connecting..."
        Connect-MgGraph -Scopes 'Organization.Read.All', 'LicenseAssignment.Read.All' -NoWelcome
    }
    else {
        Write-Verbose ("Connected to Microsoft Graph as '{0}' (Tenant: {1})" -f $context.Account, $context.TenantId)
    }
}

try {
    Test-GraphModule
    Ensure-GraphConnection

    Write-Verbose "Retrieving subscribed SKUs..."
    $skus = Get-MgSubscribedSku -All

    if (-not $skus -or $skus.Count -eq 0) {
        Write-Warning "No subscribed SKUs were returned for this tenant."
        return
    }

    $report = foreach ($sku in $skus) {
        $total     = [int]$sku.PrepaidUnits.Enabled
        $consumed  = [int]$sku.ConsumedUnits
        $suspended = [int]$sku.PrepaidUnits.Suspended
        $warning   = [int]$sku.PrepaidUnits.Warning

        if (-not $IncludeDisabled -and $total -eq 0) { continue }

        $available = $total - $consumed
        $percent   = if ($total -gt 0) { [math]::Round(($consumed / $total) * 100, 2) } else { 0 }

        [PSCustomObject]@{
            SkuPartNumber      = $sku.SkuPartNumber
            SkuId              = $sku.SkuId
            AppliesTo          = $sku.AppliesTo
            CapabilityStatus   = $sku.CapabilityStatus
            TotalLicences      = $total
            ConsumedLicences   = $consumed
            AvailableLicences  = $available
            SuspendedLicences  = $suspended
            WarningLicences    = $warning
            UtilisationPercent = $percent
        }
    }

    $report = @($report | Sort-Object -Property SkuPartNumber)
    $report | Format-Table SkuPartNumber, TotalLicences, ConsumedLicences, AvailableLicences, UtilisationPercent -AutoSize

    $totalPurchased = ($report | Measure-Object -Property TotalLicences -Sum).Sum
    $totalConsumed  = ($report | Measure-Object -Property ConsumedLicences -Sum).Sum
    Write-Host ("`nSummary: {0} SKUs | {1} licences purchased | {2} consumed | {3} available" -f `
        $report.Count, $totalPurchased, $totalConsumed, ($totalPurchased - $totalConsumed))

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
    Write-Error ("Get-M365LicenseReport failed: {0}" -f $_.Exception.Message)
    throw
}
finally {
    if ($DisconnectOnExit) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}
