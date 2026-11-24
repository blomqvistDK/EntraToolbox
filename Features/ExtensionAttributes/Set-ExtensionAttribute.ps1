<#
.SYNOPSIS
    Sets an on-premises extension attribute (1–15) on one or more Entra ID users.

.DESCRIPTION
    Connects to Microsoft Graph and updates the specified onPremisesExtensionAttributes
    property on target users. Accepts a single UPN, a comma-separated list, or a CSV
    file with a UserPrincipalName column.

    For hybrid (AD-synced) users the extension attribute is read-only in Entra ID;
    the update must be made in on-premises AD and synced. This script targets
    cloud-only users or directory-writeable attributes.

.REQUIREMENTS
    Install-Module Microsoft.Graph -Scope CurrentUser
    Required Microsoft Graph permission:
        User.ReadWrite.All

.PARAMETER SkipGraphConnect
    Reuse an existing Graph session with the required delegated scopes.
.PARAMETER NonInteractive
    Reject missing inputs instead of prompting in the console.
.PARAMETER PassThru
    Return structured per-user results in addition to status messages.

.EXAMPLE
    # Set extensionAttribute5 = "ProjectX" on a single user
    .\Set-UserExtensionAttribute.ps1 -UserPrincipalName user@contoso.com -AttributeNumber 5 -Value "ProjectX"

    # Clear extensionAttribute3 for a user
    .\Set-UserExtensionAttribute.ps1 -UserPrincipalName user@contoso.com -AttributeNumber 3 -Clear

    # Bulk-update from CSV (must have a UserPrincipalName column)
    .\Set-UserExtensionAttribute.ps1 -CsvPath .\users.csv -AttributeNumber 10 -Value "Migrated"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(HelpMessage = 'One or more UPNs, comma-separated')]
    [string[]]$UserPrincipalName,

    [Parameter(HelpMessage = 'Path to a CSV with a UserPrincipalName column')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter()]
    [ValidateRange(1, 15)]
    [int]$AttributeNumber,

    [Parameter()]
    [string]$Value,

    [Parameter()]
    [switch]$Clear,
    [switch]$SkipGraphConnect,
    [switch]$NonInteractive,
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '=== Set Entra ID Extension Attribute ===' -ForegroundColor Cyan
Write-Host ''

#region --- Interactive prompts for missing parameters ---
if ($NonInteractive -and ((-not $UserPrincipalName -and -not $CsvPath) -or -not $AttributeNumber -or (-not $Value -and -not $Clear))) {
    throw 'Supply users or a CSV, an attribute number, and a value or -Clear.'
}
if (-not $UserPrincipalName -and -not $CsvPath) {
    $input = Read-Host 'Enter user UPN(s) (comma-separated) or path to a CSV file'
    $input = $input.Trim()
    if ($input -match '\.csv$' -and (Test-Path $input -PathType Leaf)) {
        $CsvPath = $input
    } else {
        $UserPrincipalName = $input -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
}

if (-not $AttributeNumber) {
    do {
        $attrInput = Read-Host 'Which extension attribute? (1-15)'
    } while ($attrInput -notmatch '^\d+$' -or [int]$attrInput -lt 1 -or [int]$attrInput -gt 15)
    $AttributeNumber = [int]$attrInput
}

if (-not $Value -and -not $Clear) {
    $valueInput = Read-Host 'Value to set (leave empty to clear the attribute)'
    if ([string]::IsNullOrWhiteSpace($valueInput)) {
        $Clear = [switch]::new($true)
    } else {
        $Value = $valueInput
    }
}
#endregion

if ($Clear) { $Value = $null }

#region --- Connect to Microsoft Graph ---
$requiredScopes = @('User.ReadWrite.All')
$context = Get-MgContext

if ($SkipGraphConnect) {
    if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {
        throw 'Connect to Microsoft Graph with User.ReadWrite.All before running this operation.'
    }
} elseif (-not $context) {
    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
} else {
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }
    if ($missingScopes) {
        Write-Host "Reconnecting - missing scopes: $($missingScopes -join ', ')" -ForegroundColor Yellow
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    }
}
Write-Host "Connected as: $((Get-MgContext).Account)" -ForegroundColor Green
Write-Host ''
#endregion

#region --- Resolve target users ---
if ($CsvPath) {
    $csv = @(Import-Csv -Path $CsvPath)
    if ($csv.Count -eq 0) { throw 'The user CSV is empty.' }
    if ('UserPrincipalName' -notin $csv[0].PSObject.Properties.Name) {
        throw "CSV must contain a 'UserPrincipalName' column."
    }
    $targetUpns = @($csv | ForEach-Object { $_.UserPrincipalName.Trim() } | Where-Object { $_ })
} else {
    $targetUpns = @($UserPrincipalName | ForEach-Object { $_.Trim() })
}
$attrName     = "extensionAttribute$AttributeNumber"
$displayValue = if ($Clear) { '(cleared)' } else { $Value }

Write-Host "  Attribute : $attrName" -ForegroundColor White
Write-Host "  New value : $displayValue" -ForegroundColor White
Write-Host "  Users     : $($targetUpns.Count)" -ForegroundColor White
Write-Host ''
#endregion

#region --- Update extension attribute ---
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($upn in $targetUpns) {
    $body = @{
        onPremisesExtensionAttributes = @{
            $attrName = $Value
        }
    }

    if ($PSCmdlet.ShouldProcess($upn, "Set $attrName = $displayValue")) {
        try {
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/users/$upn" `
                -Body ($body | ConvertTo-Json -Depth 3) `
                -ContentType 'application/json'

            Write-Host "  [OK]   $upn" -ForegroundColor Green
            $results.Add([PSCustomObject]@{ User = $upn; Status = 'OK'; Detail = $displayValue })
        } catch {
            $errMsg = $_.Exception.Message -replace '\r?\n', ' '
            Write-Host "  [FAIL] $upn - $errMsg" -ForegroundColor Red
            $results.Add([PSCustomObject]@{ User = $upn; Status = 'FAIL'; Detail = $errMsg })
        }
    } elseif ($WhatIfPreference) {
        $results.Add([PSCustomObject]@{ User = $upn; Status = 'WhatIf'; Detail = "$attrName = $displayValue" })
    }
}
#endregion

#region --- Summary ---
$ok   = @($results | Where-Object Status -eq 'OK').Count
$fail = @($results | Where-Object Status -eq 'FAIL').Count

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Cyan
Write-Host "  Succeeded : $ok" -ForegroundColor Green
if ($fail -gt 0) {
    Write-Host "  Failed    : $fail" -ForegroundColor Red
}
Write-Host ''
#endregion
if ($PassThru) { $results.ToArray() }
 