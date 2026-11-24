<#
.SYNOPSIS
    Adds groups, users, or devices as members to a specified Entra ID group.

.DESCRIPTION
    This script looks up the target Entra ID group, then resolves each supplied
    member name against Entra ID (as a group display name, user UPN, or device
    display name depending on -MemberType), checks whether it is already a
    member, and adds it if not.
    Requires Microsoft.Graph PowerShell SDK.

.PARAMETER TargetGroupName
    Display name of the Entra ID group to add members to.

.PARAMETER MemberNames
    Array of names to add. For Groups and Devices supply the display name;
    for Users supply the UserPrincipalName (UPN).

.PARAMETER MemberFile
    Path to a text file with one entry per line. For Groups/Devices each line
    is a display name; for Users each line is a UPN.

.PARAMETER MemberType
    The type of objects to add. Accepted values: Groups, Users, Devices.
    Defaults to Groups.

.PARAMETER SkipGraphConnect
    Require an existing Graph session and throw on integration failures.
.PARAMETER PassThru
    Return structured membership results instead of a formatted table.

.EXAMPLE
    .\Add-MembersToEntraGroup.ps1 `
        -TargetGroupName "SG-All-Devices" `
        -MemberNames "SG-Dept-IT", "SG-Dept-HR"

.EXAMPLE
    .\Add-MembersToEntraGroup.ps1 `
        -TargetGroupName "SG-All-Users" `
        -MemberNames "alice@contoso.com", "bob@contoso.com" `
        -MemberType Users

.EXAMPLE
    .\Add-MembersToEntraGroup.ps1 `
        -TargetGroupName "SG-All-Devices" `
        -MemberNames "DESKTOP-ABC123", "LAPTOP-XYZ456" `
        -MemberType Devices

.EXAMPLE
    # Using a text file — one group display name per line
    .\Add-MembersToEntraGroup.ps1 `
        -TargetGroupName "SG-All-Devices" `
        -MemberFile "C:\groups.txt"

.EXAMPLE
    # Using a text file — one UPN per line
    .\Add-MembersToEntraGroup.ps1 `
        -TargetGroupName "SG-All-Users" `
        -MemberFile "C:\users.txt" `
        -MemberType Users

.EXAMPLE
    # Using a text file — one device display name per line
    .\Add-MembersToEntraGroup.ps1 `
        -TargetGroupName "SG-All-Devices" `
        -MemberFile "C:\devices.txt" `
        -MemberType Devices
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'ByNames')]
param (
    [Parameter(Mandatory)]
    [string]$TargetGroupName,

    [Parameter(Mandatory, ParameterSetName = 'ByNames')]
    [Alias('MemberGroupNames')]
    [string[]]$MemberNames,

    [Parameter(Mandatory, ParameterSetName = 'ByFile')]
    [Alias('MemberGroupsFile')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$MemberFile,

    [Parameter()]
    [ValidateSet('Groups', 'Users', 'Devices')]
    [string]$MemberType = 'Groups',
    [switch]$SkipGraphConnect,
    [switch]$PassThru
)

#region Functions

function Connect-ToGraph {
    $requiredScopes = @(
        "GroupMember.ReadWrite.All",
        "Group.Read.All",
        "User.Read.All",
        "Device.Read.All"
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($SkipGraphConnect) {
        if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {
            throw "An existing Graph session with these scopes is required: $($requiredScopes -join ', ')"
        }
    } elseif ($null -eq $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    } else {
        $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }
        if ($missingScopes) {
            Write-Host "Reconnecting - missing scopes: $($missingScopes -join ', ')" -ForegroundColor Yellow
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        } else {
            Write-Host "Already connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
        }
    }
}

function Get-GroupByDisplayName {
    param ([string]$DisplayName)

    Get-MgGroup -Filter "displayName eq '$DisplayName'" -ConsistencyLevel eventual -CountVariable count -ErrorAction SilentlyContinue |
        Select-Object -First 1
}

function Get-UserByUPN {
    param ([string]$UPN)

    Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction SilentlyContinue |
        Select-Object -First 1
}

function Get-DeviceByDisplayName {
    param ([string]$DisplayName)

    Get-MgDevice -Filter "displayName eq '$DisplayName'" -ConsistencyLevel eventual -CountVariable count -ErrorAction SilentlyContinue |
        Select-Object -First 1
}

function Get-ExistingMemberIds {
    param ([string]$GroupId)

    $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue
    return $members.Id
}

#endregion

#region Resolve input source

if ($PSCmdlet.ParameterSetName -eq 'ByFile') {
    Write-Host "Reading names from file: '$MemberFile'" -ForegroundColor Cyan
    $MemberNames = @(Get-Content -Path $MemberFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($MemberNames.Count -eq 0) {
        if ($SkipGraphConnect) { throw "No entries found in '$MemberFile'." }
        Write-Error "No entries found in '$MemberFile'. Aborting."
        exit 1
    }
    Write-Host "  Found $($MemberNames.Count) entry/entries in file." -ForegroundColor Green
}

#endregion

#region Prerequisites

$requiredModules = @{
    Groups  = 'Microsoft.Graph.Groups'
    Users   = 'Microsoft.Graph.Users'
    Devices = 'Microsoft.Graph.Identity.DirectoryManagement'
}

$modulesToImport = @($requiredModules['Groups'])  # Always needed for target group lookup

if ($MemberType -eq 'Users') {
    $modulesToImport += $requiredModules['Users']
} elseif ($MemberType -eq 'Devices') {
    $modulesToImport += $requiredModules['Devices']
}

foreach ($module in $modulesToImport) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        if ($SkipGraphConnect) { throw "Missing module: $module. Install-Module $module -Scope CurrentUser" }
        Write-Error "$module module is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
        exit 1
    }
    Import-Module $module -ErrorAction Stop
}

Connect-ToGraph

#endregion

#region Resolve target group

Write-Host "`nResolving target group: '$TargetGroupName'" -ForegroundColor Cyan
$targetGroup = Get-GroupByDisplayName -DisplayName $TargetGroupName

if (-not $targetGroup) {
    if ($SkipGraphConnect) { throw "Target group '$TargetGroupName' was not found in Entra ID." }
    Write-Error "Target group '$TargetGroupName' was not found in Entra ID. Aborting."
    exit 1
}

Write-Host "  Target group found: $($targetGroup.DisplayName) [$($targetGroup.Id)]" -ForegroundColor Green

#endregion

#region Get current members

Write-Host "`nRetrieving current members of target group..." -ForegroundColor Cyan
$existingMemberIds = Get-ExistingMemberIds -GroupId $targetGroup.Id

#endregion

#region Validate and add members

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

$memberTypeLabel = switch ($MemberType) {
    'Users'   { 'user' }
    'Devices' { 'device' }
    default   { 'group' }
}

Write-Host "`nValidating $($MemberNames.Count) $($memberTypeLabel)(s)..." -ForegroundColor Cyan

foreach ($memberName in $MemberNames) {
    $memberName = $memberName.Trim()

    if ([string]::IsNullOrWhiteSpace($memberName)) {
        continue
    }

    $result = [PSCustomObject]@{
        MemberName = $memberName
        Found      = $false
        MemberId   = $null
        Status     = $null
    }

    # Lookup the member object based on type
    $memberObject = switch ($MemberType) {
        'Users'   { Get-UserByUPN -UPN $memberName }
        'Devices' { Get-DeviceByDisplayName -DisplayName $memberName }
        default   { Get-GroupByDisplayName -DisplayName $memberName }
    }

    if (-not $memberObject) {
        Write-Warning "  NOT FOUND: '$memberName' - skipping."
        $result.Status = "Not found"
        $results.Add($result)
        continue
    }

    $result.Found    = $true
    $result.MemberId = $memberObject.Id

    # Self-reference check only applies when adding groups
    if ($MemberType -eq 'Groups' -and $memberObject.Id -eq $targetGroup.Id) {
        Write-Warning "  SKIPPED: '$memberName' is the same as the target group."
        $result.Status = "Skipped (same as target)"
        $results.Add($result)
        continue
    }

    if ($existingMemberIds -contains $memberObject.Id) {
        Write-Host "  ALREADY MEMBER: '$memberName'" -ForegroundColor Yellow
        $result.Status = "Already a member"
        $results.Add($result)
        continue
    }

    if ($PSCmdlet.ShouldProcess($targetGroup.DisplayName, "Add $memberTypeLabel '$memberName'")) {
        try {
            $bodyParam = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($memberObject.Id)"
            }
            New-MgGroupMemberByRef -GroupId $targetGroup.Id -BodyParameter $bodyParam -ErrorAction Stop
            Write-Host "  ADDED: '$memberName'" -ForegroundColor Green
            $result.Status = "Added"
        } catch {
            Write-Warning "  FAILED: '$memberName' - $($_.Exception.Message)"
            $result.Status = "Failed: $($_.Exception.Message)"
        }
    } elseif ($WhatIfPreference) {
        $result.Status = 'WhatIf'
    }

    $results.Add($result)
}

#endregion

#region Summary

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
if ($PassThru) { $results.ToArray() } else { $results | Format-Table -Property MemberName, Found, Status -AutoSize }

$added    = ($results | Where-Object Status -eq "Added").Count
$skipped  = ($results | Where-Object Status -eq "Already a member").Count
$notFound = ($results | Where-Object Status -eq "Not found").Count
$failed   = ($results | Where-Object { $_.Status -like "Failed:*" }).Count

Write-Host "Added: $added | Already member: $skipped | Not found: $notFound | Failed: $failed" -ForegroundColor Cyan

#endregion
