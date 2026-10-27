<#
.SYNOPSIS
    Builds a structured Microsoft 365 user onboarding report and task checklist.

.DESCRIPTION
    New-M365UserOnboarding.ps1 generates a complete, auditable onboarding
    plan for a new starter — without requiring a live tenant to parse or
    to preview the output.

    It produces:
      * A structured task list covering account creation, licensing,
        MFA / Conditional Access, mailbox, OneDrive, Teams, device
        enrolment (Intune), security groups, and manager sign-off
      * A Markdown report and, optionally, a CSV task list that can be
        imported into a help-desk / ITSM system or Microsoft Planner

    The script deliberately does NOT create a user account by default.
    Account creation is represented as a checklist item so an engineer
    reviews licence, group and policy assignments before any change is
    made in the tenant. When -CreateReportOnly is not used together with
    a Graph session, no tenant calls are made at all.

    Supports -WhatIf (via SupportsShouldProcess): with -WhatIf, it only
    prints what would be generated and writes nothing.

.PARAMETER DisplayName
    Full name of the new starter, e.g. "Ayesha Khan". (Required)

.PARAMETER UserPrincipalName
    Planned User Principal Name, e.g. "ayesha.khan@contoso.com".
    No tenant validation is performed — this is recorded in the report.

.PARAMETER Department
    Department of the new starter, e.g. "Finance".

.PARAMETER JobTitle
    Job title of the new starter, e.g. "Accounts Officer".

.PARAMETER Manager
    Manager name or UPN, recorded as the sign-off owner in the report.

.PARAMETER LicenceSku
    Planned licence SKU part number, e.g. "Microsoft 365 Business Premium"
    or "ENTERPRISEPACK". Recorded only; no licence is assigned by this script.

.PARAMETER DeviceType
    Planned device type. Accepted values: WindowsLaptop, WindowsDesktop,
    Mac, MobileOnly, SharedDevice. Default: WindowsLaptop

.PARAMETER StartDate
    Planned start date. Default: today.

.PARAMETER OutputDirectory
    Folder where the Markdown report (and CSV, if requested) will be
    written. Default: .\output

.PARAMETER ExportCsv
    When specified, also exports the task list as CSV alongside the
    Markdown report.

.EXAMPLE
    .\New-M365UserOnboarding.ps1 -DisplayName "Ayesha Khan" -UserPrincipalName "ayesha.khan@contoso.com" -Department "Finance" -JobTitle "Accounts Officer" -Manager "Finance Manager"

    Generates an onboarding report in .\output for the new starter.

.EXAMPLE
    .\New-M365UserOnboarding.ps1 -DisplayName "Test User" -UserPrincipalName "test.user@contoso.com" -ExportCsv -WhatIf

    Shows exactly what would be created, without writing any files.

.EXAMPLE
    .\New-M365UserOnboarding.ps1 -DisplayName "Bilal Ahmed" -UserPrincipalName "bilal.ahmed@contoso.com" -DeviceType MobileOnly -LicenceSku "Microsoft 365 F3" -OutputDirectory "C:\IT\Onboarding"

    Generates a mobile-only onboarding plan in a custom output folder.

.NOTES
    Author  : Zohaib Yousaf (info@zohaibyousaf.com)
    Requires: PowerShell 5.1 or PowerShell 7+. No tenant or Graph module
              is required to generate the report.
    Version : 1.0.0
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$Department = 'Not specified',

    [Parameter(Mandatory = $false)]
    [string]$JobTitle = 'Not specified',

    [Parameter(Mandatory = $false)]
    [string]$Manager = 'Not specified',

    [Parameter(Mandatory = $false)]
    [string]$LicenceSku = 'To be confirmed by IT',

    [Parameter(Mandatory = $false)]
    [ValidateSet('WindowsLaptop', 'WindowsDesktop', 'Mac', 'MobileOnly', 'SharedDevice')]
    [string]$DeviceType = 'WindowsLaptop',

    [Parameter(Mandatory = $false)]
    [datetime]$StartDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = (Join-Path -Path $PSScriptRoot -ChildPath '..\output'),

    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-OnboardingTasks {
    param(
        [string]$Device,
        [string]$Dept
    )

    $tasks = [System.Collections.Generic.List[object]]::new()
    $add = {
        param($Phase, $Task, $Owner, $Notes)
        $tasks.Add([PSCustomObject]@{
            Id     = $tasks.Count + 1
            Phase  = $Phase
            Task   = $Task
            Owner  = $Owner
            Status = 'Not started'
            Notes  = $Notes
        })
    }

    & $add '1. Account & Licence' 'Create Microsoft 365 user account and verify UPN spelling' 'IT Admin' 'Confirm name spelling with HR before creation'
    & $add '1. Account & Licence' 'Assign licence SKU and verify service plans' 'IT Admin' ("Planned SKU: {0}" -f $LicenceSku)
    & $add '1. Account & Licence' 'Set usage location and department / job title attributes' 'IT Admin' ("Department: {0}" -f $Dept)
    & $add '1. Account & Licence' 'Add user to department security groups (least-privilege)' 'IT Admin' 'Do not copy another user blindly — confirm with manager'

    & $add '2. Security' 'Enforce MFA registration (Authenticator app)' 'IT Admin / User' 'Verify on day 1, do not leave as SMS-only if policy requires app'
    & $add '2. Security' 'Confirm Conditional Access policies apply to the account' 'IT Admin' 'Check sign-in logs after first login'
    & $add '2. Security' 'Set temporary password and require change on first sign-in' 'IT Admin' 'Share via approved secure channel only, never by plain email'

    & $add '3. Mailbox & Collaboration' 'Verify mailbox provisioned and Outlook on the web access' 'IT Admin' 'Allow time for provisioning before testing'
    & $add '3. Mailbox & Collaboration' 'Add to department distribution lists / shared mailbox access' 'IT Admin' 'Only with manager approval'
    & $add '3. Mailbox & Collaboration' 'Provision OneDrive and confirm sync client sign-in' 'IT Admin / User' ''
    & $add '3. Mailbox & Collaboration' 'Add to Microsoft Teams teams and channels for the department' 'Team Owner / IT' ''

    if ($Device -eq 'MobileOnly') {
        & $add '4. Device' 'Enrol mobile device in Intune (App Protection / MDM per policy)' 'IT Admin / User' 'Confirm personal vs corporate-owned handling'
        & $add '4. Device' 'Install Outlook / Teams mobile apps and verify MFA prompts' 'User' ''
    }
    elseif ($Device -eq 'SharedDevice') {
        & $add '4. Device' 'Confirm shared-device sign-in method and kiosk / shared profile' 'IT Admin' 'Do not assign a primary-user device licence pattern'
        & $add '4. Device' 'Test sign-in, printing and line-of-business apps on the shared device' 'IT Admin' ''
    }
    else {
        & $add '4. Device' ("Prepare and enrol {0} in Intune / Autopilot" -f $Device) 'IT Admin' 'Apply device naming standard and asset tag'
        & $add '4. Device' 'Verify device compliance policy evaluates as Compliant' 'IT Admin' 'Check encryption, OS version and endpoint protection'
        & $add '4. Device' 'Install standard software set and verify updates are current' 'IT Admin' ''
    }

    & $add '5. Handover' 'First-login walkthrough with user (Outlook, Teams, OneDrive, MFA)' 'IT Admin' 'Book 30 minutes on start date'
    & $add '5. Handover' 'Manager sign-off on access and equipment' 'Manager' ("Manager: {0}" -f $Manager)
    & $add '5. Handover' 'Close onboarding ticket with report attached' 'IT Admin' 'Attach this report and the CSV task list if exported'

    return $tasks
}

try {
    $safeName = ($UserPrincipalName -replace '[^a-zA-Z0-9._-]', '_')
    $dateStamp = $StartDate.ToString('yyyy-MM-dd')
    $reportName = "Onboarding_{0}_{1}.md" -f $safeName, $dateStamp
    $reportPath = Join-Path -Path $OutputDirectory -ChildPath $reportName
    $csvPath = Join-Path -Path $OutputDirectory -ChildPath ("Onboarding_{0}_{1}.csv" -f $safeName, $dateStamp)

    $tasks = Get-OnboardingTasks -Device $DeviceType -Dept $Department

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# Microsoft 365 User Onboarding Report")
    $lines.Add("")
    $lines.Add("| Field | Value |")
    $lines.Add("|---|---|")
    $lines.Add(("| Display Name | {0} |" -f $DisplayName))
    $lines.Add(("| User Principal Name | {0} |" -f $UserPrincipalName))
    $lines.Add(("| Department | {0} |" -f $Department))
    $lines.Add(("| Job Title | {0} |" -f $JobTitle))
    $lines.Add(("| Manager / Sign-off | {0} |" -f $Manager))
    $lines.Add(("| Planned Licence | {0} |" -f $LicenceSku))
    $lines.Add(("| Device Type | {0} |" -f $DeviceType))
    $lines.Add(("| Start Date | {0} |" -f $dateStamp))
    $lines.Add(("| Generated | {0} |" -f (Get-Date -Format 'yyyy-MM-dd HH:mm')))
    $lines.Add("")
    $lines.Add("## Task Checklist")
    $lines.Add("")
    $lines.Add("| # | Phase | Task | Owner | Status | Notes |")
    $lines.Add("|---|---|---|---|---|---|")
    foreach ($task in $tasks) {
        $lines.Add(("| {0} | {1} | {2} | {3} | {4} | {5} |" -f $task.Id, $task.Phase, $task.Task, $task.Owner, $task.Status, $task.Notes))
    }
    $lines.Add("")
    $lines.Add("## Safety Notes")
    $lines.Add("")
    $lines.Add("- This report is a plan only. No account, licence or group was changed by generating it.")
    $lines.Add("- Verify licence availability and manager approval before creating the account in the tenant.")
    $lines.Add("- Share initial passwords only through an approved secure channel.")
    $lines.Add("")
    $lines.Add("Prepared with the PowerShell Microsoft 365 & Intune Toolkit by Zohaib Yousaf.")

    $content = $lines -join [Environment]::NewLine

    if ($PSCmdlet.ShouldProcess($reportPath, "Create onboarding Markdown report")) {
        if (-not (Test-Path -LiteralPath $OutputDirectory)) {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        }
        Set-Content -LiteralPath $reportPath -Value $content -Encoding UTF8
        Write-Host "Onboarding report created: $reportPath"
    }

    if ($ExportCsv) {
        if ($PSCmdlet.ShouldProcess($csvPath, "Export onboarding task list as CSV")) {
            if (-not (Test-Path -LiteralPath $OutputDirectory)) {
                New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
            }
            $tasks | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "Task list exported: $csvPath"
        }
    }

    $tasks | Format-Table Id, Phase, Task, Owner, Status -AutoSize
    return $tasks
}
catch {
    Write-Error ("New-M365UserOnboarding failed: {0}" -f $_.Exception.Message)
    throw
}
