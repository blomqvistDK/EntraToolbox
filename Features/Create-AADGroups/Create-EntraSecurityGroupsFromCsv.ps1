# Requires: Microsoft Graph PowerShell module
# Install-Module Microsoft.Graph -Scope CurrentUser

<#
.SYNOPSIS
    Create or reuse security groups and optional nested group memberships from a semicolon CSV.
.PARAMETER SkipGraphConnect
    Require an existing Graph session instead of starting device-code sign-in.
.PARAMETER PassThru
    Return structured group and membership outcomes.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$CsvPath,
    [switch]$SkipGraphConnect,
    [switch]$PassThru
)

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    if ($SkipGraphConnect) { throw 'Supply -CsvPath for noninteractive execution.' }
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV file"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "No CSV file selected."
    }

    $CsvPath = $openFileDialog.FileName
}

<#
$requiredModule = "Microsoft.Graph"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Microsoft.Graph module not found. Installing it..." -ForegroundColor Yellow
    Install-Module $requiredModule -Scope CurrentUser -Force -AllowClobber
}

Import-Module $requiredModule -Force

$requiredCmds = @("Get-MgGroup", "New-MgGroup", "New-MgGroupMemberByRef")
foreach ($cmd in $requiredCmds) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "Required Microsoft Graph command '$cmd' is not available after import."
    }
}
#>
#Disconnect-MgGraph -ErrorAction SilentlyContinue
if ($SkipGraphConnect) {
    $context = Get-MgContext
    $requiredScopes = @('Group.ReadWrite.All', 'Directory.ReadWrite.All')
    if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {
        throw 'An existing Graph session with Group.ReadWrite.All and Directory.ReadWrite.All is required.'
    }
} else { Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All" -UseDeviceCode }
if (-not (Get-MgContext)) {
    throw "Microsoft Graph sign-in did not complete."
}
Write-Host "Signed in to Microsoft Graph."

if (-not (Test-Path $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

$rows = @(Import-Csv $CsvPath -Delimiter ';')
Write-Host "Loaded $($rows.Count) rows from '$CsvPath'."

foreach ($row in $rows) {
    $displayName = $row.GroupName
    $description = $row.Description
    $groupStatus = 'Reused'
    $membershipStatus = 'Not requested'
    Write-Host "Processing group: $displayName"

    if ([string]::IsNullOrWhiteSpace($displayName)) {
        throw "Missing GroupName in CSV row: $($row | ConvertTo-Json -Compress)"
    }

    # Check for an existing group and reuse it if present
    $existing = Get-MgGroup -Filter "displayName eq '$($displayName.Replace("'", "''"))'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Group already exists and will be reused: $displayName"
        $newGroup = $existing
    }
    else {
        # Mail nickname is required by Graph even for security groups
        $mailNickname = ($displayName -replace '[^A-Za-z0-9]', '')
        if ([string]::IsNullOrWhiteSpace($mailNickname)) {
            $mailNickname = "group"
        }

        if ($mailNickname.Length -gt 64) {
            $mailNickname = $mailNickname.Substring(0, 64)
        }

        $params = @{
            DisplayName     = $displayName
            Description     = $description
            MailEnabled     = $false
            SecurityEnabled = $true
            MailNickname    = $mailNickname
        }

        try {
            $newGroup = New-MgGroup @params
            $groupStatus = 'Created'
            Write-Host "Created security group: $($newGroup.DisplayName) (ID: $($newGroup.Id))" -ForegroundColor Green
        }
        catch {
            if ($PassThru) { [pscustomobject]@{ GroupName = $displayName; GroupId = $null; Status = 'Failed'; Membership = 'Not attempted'; Detail = $_.Exception.Message } }
            Write-Error "Failed to create group '$displayName': $($_.Exception.Message)"
            continue
        }
    }

    $memberGroupName = $null
    foreach ($candidate in @("MemberGroup", "MemberGroupName", "GroupMember", "Member")) {
        if ($row.PSObject.Properties.Name -contains $candidate) {
            $memberGroupName = [string]$row.$candidate
            break
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($memberGroupName)) {
        $memberGroupName = $memberGroupName.Trim()
        try {
            $memberGroup = Get-MgGroup -Filter "displayName eq '$($memberGroupName.Replace("'", "''"))'" -ErrorAction Stop
            if ($memberGroup) {
                $groupRef = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($memberGroup.Id)"
                }
                New-MgGroupMemberByRef -GroupId $newGroup.Id -BodyParameter $groupRef -ErrorAction Stop
                $membershipStatus = 'Added'
                Write-Host "Added group $memberGroupName as member to $displayName" -ForegroundColor Green
            }
            else {
                $membershipStatus = 'Not found'
                Write-Warning "Group '$memberGroupName' was not found for membership in '$displayName'."
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($errorMessage -match 'already exist') {
                $membershipStatus = 'Already a member'
                Write-Host "Group '$memberGroupName' is already a member of '$displayName'."
            }
            elseif ($errorMessage -match 'Resource .* does not exist' -or $errorMessage -match 'Request_ResourceNotFound') {
                $membershipStatus = 'Not found'
                Write-Warning "Group '$memberGroupName' was not found for membership in '$displayName'."
            }
            else {
                $membershipStatus = "Failed: $errorMessage"
                Write-Warning "Failed to add group '$memberGroupName' to '$displayName': $errorMessage"
            }
        }
    }
    if ($PassThru) { [pscustomobject]@{ GroupName = $displayName; GroupId = $newGroup.Id; Status = $groupStatus; Membership = $membershipStatus; Detail = $memberGroupName } }
}

Write-Host "Finished processing $($rows.Count) CSV rows."
