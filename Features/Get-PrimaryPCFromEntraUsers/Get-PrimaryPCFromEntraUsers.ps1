<#
.SYNOPSIS
    Finds each user's Windows PCs through Entra device associations or Windows Sign In logs and exports device object IDs.

.DESCRIPTION
    Users come from an Entra group, a CSV file and/or a UPN list. Output CSV matches the Entra bulk-import template header.

.PARAMETER DeviceMode
    With DeviceSource Entra, Latest (default) returns the most recently active Windows Intune-managed device.
    All returns all Windows devices in the user's Entra owned and registered devices,
    including unmanaged devices. Sign-in fallback can return one device when no devices match in either mode.
    With DeviceSource SignInLogs, Most (default) returns all PCs tied for the most successful
    Windows Sign In events. All returns every eligible PC, ordered by count then newest sign-in.
.PARAMETER DeviceSource
    Entra (default) discovers owned and registered devices. SignInLogs discovers devices only from
    successful Windows Sign In events, then resolves them in Entra for Windows validation and export.
.PARAMETER SignInFallback
    With Entra discovery, optionally try the existing Windows Sign In fallback when no device matches.
.PARAMETER SignInDays
    Sign-in lookback in days (default 30). Available history depends on tenant retention, access and licensing.
    Requires AuditLog.Read.All, Microsoft.Graph.Reports and applicable Entra ID P1/P2 licensing.
    Observed Entra events are not a complete history of local Windows logons.
.PARAMETER SkipGraphConnect
    Require an existing Graph session instead of signing in.
.PARAMETER OnlyIntuneDevices
    Return only devices present in the Intune managed-device inventory, including sign-in fallback results.
.PARAMETER NonInteractive
    Disable the optional console prompt for sign-in fallback.
.PARAMETER PassThru
    Return device result objects instead of a formatted console table.

.EXAMPLE
    .\Get-PrimaryPCFromEntraUsers.ps1 -GroupName "Pilot Users"
    .\Get-PrimaryPCFromEntraUsers.ps1 -UserPrincipalName a@contoso.com,b@contoso.com
    .\Get-PrimaryPCFromEntraUsers.ps1 -CsvPath .\users.csv -OutputPath .\devices.csv
    .\Get-PrimaryPCFromEntraUsers.ps1 -GroupName "Pilot Users" -DeviceMode All
    .\Get-PrimaryPCFromEntraUsers.ps1 -UserPrincipalName a@contoso.com -DeviceSource SignInLogs -DeviceMode Most
    .\Get-PrimaryPCFromEntraUsers.ps1 -CsvPath .\users.csv -DeviceSource SignInLogs -DeviceMode All -SignInDays 14
#>
[CmdletBinding()]
param(
    [string]$GroupId,
    [string]$GroupName,
    [string]$CsvPath,
    [string[]]$UserPrincipalName,
    [string]$OutputPath = (Join-Path $PWD 'PrimaryPCs.csv'),
    [switch]$OnlyIntuneDevices,
    [switch]$SignInFallback,
    [ValidateRange(1, 365)]
    [int]$SignInDays = 30,
    [switch]$SkipGraphConnect,
    [switch]$NonInteractive,
    [switch]$PassThru,
    [ValidateSet('Latest', 'Most', 'All')]
    [string]$DeviceMode,
    [ValidateSet('Entra', 'SignInLogs')]
    [string]$DeviceSource = 'Entra'
)

if (-not $PSBoundParameters.ContainsKey('DeviceMode')) {
    $DeviceMode = if ($DeviceSource -eq 'SignInLogs') { 'Most' } else { 'Latest' }
}
if (($DeviceSource -eq 'Entra' -and $DeviceMode -eq 'Most') -or ($DeviceSource -eq 'SignInLogs' -and $DeviceMode -eq 'Latest')) {
    throw "DeviceMode '$DeviceMode' is not supported with DeviceSource '$DeviceSource'. Use Latest/All for Entra or Most/All for SignInLogs."
}
if ($DeviceSource -eq 'SignInLogs' -and $SignInFallback) {
    throw 'SignInFallback is only supported with DeviceSource Entra.'
}

if (-not ($GroupId -or $GroupName -or $CsvPath -or $UserPrincipalName)) {
    throw 'Specify at least one of -GroupId, -GroupName, -CsvPath or -UserPrincipalName.'
}

$scopes = @('User.Read.All', 'GroupMember.Read.All', 'Device.Read.All')
if ($OnlyIntuneDevices) { $scopes += 'DeviceManagementManagedDevices.Read.All' }
if ($SignInFallback -or $DeviceSource -eq 'SignInLogs') { $scopes += 'AuditLog.Read.All' }
if ($SkipGraphConnect) {
    $context = Get-MgContext
    if (-not $context -or @($scopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {
        throw "An existing Graph session with these scopes is required: $($scopes -join ', ')"
    }
} else { Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null }

# Intune MDM app ID, and the Configuration Manager (co-managed) app ID that Entra reports for hybrid co-managed PCs.
$IntuneMdmAppIds = @('0000000a-0000-0000-c000-000000000000', '54b943f8-d761-4f8d-951e-9cea1846db5a')

$intuneDeviceIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if ($OnlyIntuneDevices) {
    $inventoryUri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=azureADDeviceId'
    while ($inventoryUri) {
        $inventoryPage = Invoke-MgGraphRequest -Method GET -Uri $inventoryUri -OutputType PSObject -ErrorAction Stop
        foreach ($managedDevice in $inventoryPage.value) {
            $deviceGuid = [guid]::Empty
            if ([guid]::TryParse([string]$managedDevice.azureADDeviceId,[ref]$deviceGuid) -and $deviceGuid -ne [guid]::Empty) {
                [void]$intuneDeviceIds.Add($deviceGuid.ToString())
            }
        }
        $inventoryUri = [string]$inventoryPage.'@odata.nextLink'
    }
}

function Test-PrimaryPCIntuneDevice {
    param([string]$DeviceId)
    $deviceGuid = [guid]::Empty
    return [guid]::TryParse($DeviceId,[ref]$deviceGuid) -and $intuneDeviceIds.Contains($deviceGuid.ToString())
}

$upns = [System.Collections.Generic.List[string]]::new()

if ($GroupName -and -not $GroupId) {
    $safeName = $GroupName.Replace("'", "''")
    $groups = @(Get-MgGroup -Filter "displayName eq '$safeName'")
    if ($groups.Count -ne 1) { throw "Expected exactly one group named '$GroupName', found $($groups.Count)." }
    $GroupId = $groups[0].Id
}

if ($GroupId) {
    Get-MgGroupTransitiveMember -GroupId $GroupId -All |
        Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
        ForEach-Object { $upns.Add($_.AdditionalProperties['userPrincipalName']) }
}

if ($CsvPath) {
    $rows = Import-Csv -Path $CsvPath
    if (-not $rows) { throw "No rows found in $CsvPath." }
    $col = if ($rows[0].PSObject.Properties.Name -contains 'UserPrincipalName') { 'UserPrincipalName' } else { $rows[0].PSObject.Properties.Name[0] }
    $rows | ForEach-Object { if ($_.$col) { $upns.Add($_.$col.Trim()) } }
}

if ($UserPrincipalName) {
    $UserPrincipalName | ForEach-Object { if ($_) { $upns.Add($_.Trim()) } }
}

$upns = $upns | Where-Object { $_ } | Sort-Object -Unique
Write-Host "Processing $(@($upns).Count) user(s)..."

$signInAccess = @{ Checked = $false; Blocked = $false }

function Get-SignInAccessBlockMessage {
    param([string]$Message)
    if ($Message -match 'Authentication_RequestFromUnsupportedRole') {
        $account = [string](Get-MgContext).Account
        $who = if ($account) { "the signed-in account ($account)" } else { 'the signed-in account' }
        return "Sign-in logs are unavailable: $who has no active Entra role that can read sign-in logs, such as Reports Reader, Global Reader or Security Reader. Activate PIM-eligible roles, then sign in again. Sign-in lookups are skipped."
    }
    if ($Message -match 'Authentication_RequestFromNonPremiumTenantOrB2CTenant') {
        return 'Sign-in logs are unavailable: reading them through Microsoft Graph requires Microsoft Entra ID P1 or P2. Sign-in lookups are skipped.'
    }
}

function Test-SignInLogAccess {
    # One probe turns a missing role or license into a single warning instead of a failure per user.
    if (-not $signInAccess.Checked) {
        $signInAccess.Checked = $true
        $since = (Get-Date).ToUniversalTime().AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
        try { $null = Get-MgAuditLogSignIn -Filter "createdDateTime ge $since" -Top 1 -ErrorAction Stop }
        catch {
            $blockMessage = Get-SignInAccessBlockMessage $_.Exception.Message
            if ($blockMessage) { $signInAccess.Blocked = $true; Write-Warning $blockMessage }
            else { Write-Verbose "Sign-in log access precheck was inconclusive: $($_.Exception.Message)" }
        }
    }
    return -not $signInAccess.Blocked
}

function Get-SignInDevice {
    param($User, [switch]$SignInOnly)
    if (-not (Test-SignInLogAccess)) { return }
    # Sign-in logs identify devices by deviceId (not object ID) and need Entra ID P1/P2.
    $since = (Get-Date).ToUniversalTime().AddDays(-$SignInDays).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
    $lookupName = if ($SignInOnly) { 'Sign-in discovery' } else { 'Sign-in fallback' }
    $lookupLabel = if ($SignInOnly) { 'sign-in discovery' } else { 'fallback' }
    $successFilter = if ($SignInOnly) { ' and status/errorCode eq 0' } else { '' }
    Write-Verbose "$($User.UserPrincipalName) : querying Windows Sign In $lookupLabel events since $since"
    try {
        $signIns = @(Get-MgAuditLogSignIn -Filter "userId eq '$($User.Id)' and appDisplayName eq 'Windows Sign In'$successFilter and createdDateTime ge $since" -All -ErrorAction Stop)
    } catch {
        # The probe can pass when it only sees the operator's own sign-ins, so access may still be denied here.
        $blockMessage = Get-SignInAccessBlockMessage $_.Exception.Message
        if ($blockMessage) { $signInAccess.Blocked = $true; Write-Warning $blockMessage }
        else { Write-Warning "$lookupName query failed for $($User.UserPrincipalName): $($_.Exception.Message)" }
        return
    }
    $deviceSignIns = @($signIns | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.DeviceDetail.DeviceId) -and
        (-not $SignInOnly -or ($null -ne $_.Status.ErrorCode -and $_.Status.ErrorCode -eq 0))
    })
    Write-Verbose "$($User.UserPrincipalName) : $lookupLabel returned $($signIns.Count) event(s); $($deviceSignIns.Count) with a device ID$(if ($SignInOnly) { ' and successful status' })"
    if (-not $deviceSignIns.Count) {
        Write-Warning "$lookupName for $($User.UserPrincipalName): no $(if ($SignInOnly) { 'successful ' })Windows Sign In events with a device ID since $since."
        return
    }
    $candidates = @($deviceSignIns | Group-Object { $_.DeviceDetail.DeviceId } |
        Sort-Object @{ Expression='Count'; Descending=$true }, @{ Expression = { ($_.Group.CreatedDateTime | Measure-Object -Maximum).Maximum }; Descending=$true }, Name)
    $eligible = [System.Collections.Generic.List[object]]::new()
    foreach ($candidate in $candidates) {
        if ($OnlyIntuneDevices -and -not (Test-PrimaryPCIntuneDevice $candidate.Name)) {
            Write-Verbose "$($User.UserPrincipalName) : $lookupLabel device $($candidate.Name) rejected: not present in Intune inventory"
            continue
        }
        $safeDeviceId = $candidate.Name.Replace("'", "''")
        try {
            $deviceMatches = @(Get-MgDevice -Filter "deviceId eq '$safeDeviceId'" -Property 'id,deviceId,displayName,operatingSystem,approximateLastSignInDateTime' -ErrorAction Stop)
        } catch {
            Write-Warning "$lookupName device lookup failed for $($User.UserPrincipalName), device $($candidate.Name): $($_.Exception.Message)"
            continue
        }
        if ($deviceMatches.Count -ne 1) {
            $message = "$($User.UserPrincipalName) : $lookupLabel device $($candidate.Name) rejected: resolved to $($deviceMatches.Count) Entra devices"
            if ($SignInOnly) { Write-Warning $message } else { Write-Verbose $message }
            continue
        }
        $device = $deviceMatches[0]
        if ($SignInOnly -and [string]::IsNullOrWhiteSpace($device.Id)) {
            Write-Warning "$lookupName device $($candidate.Name) for $($User.UserPrincipalName) has no Entra object ID and cannot be exported."
            continue
        }
        if ($device.OperatingSystem -notlike 'Windows*') {
            Write-Verbose "$($User.UserPrincipalName) : $lookupLabel device $($candidate.Name) rejected: Entra OS '$($device.OperatingSystem)' is not Windows"
            continue
        }
        if ($OnlyIntuneDevices -and -not (Test-PrimaryPCIntuneDevice $device.DeviceId)) {
            Write-Verbose "$($User.UserPrincipalName) : $lookupLabel device $($candidate.Name) rejected: resolved Entra device is not present in Intune inventory"
            continue
        }
        Write-Verbose "$($User.UserPrincipalName) : $lookupName accepted $($device.DisplayName) ($($candidate.Count) Windows Sign In events)"
        $result = [pscustomobject]@{
            ObjectId     = $device.Id
            DisplayName  = $device.DisplayName
            LastActivity = if ($SignInOnly) { ([datetimeoffset](($candidate.Group.CreatedDateTime | Measure-Object -Maximum).Maximum)).UtcDateTime } elseif ($device.ApproximateLastSignInDateTime) { $device.ApproximateLastSignInDateTime } else { [datetime]::MinValue }
            Source       = "SignInLogs ($($candidate.Count) events)"
            SignInCount  = $candidate.Count
        }
        if (-not $SignInOnly) { return $result }
        $eligible.Add($result)
    }
    if ($eligible.Count -gt 0) {
        if ($DeviceMode -eq 'Most') { return $eligible | Where-Object { $_.SignInCount -eq $eligible[0].SignInCount } }
        return $eligible.ToArray()
    }
    Write-Warning "$lookupName for $($User.UserPrincipalName): none of $($candidates.Count) device candidate(s) passed the Windows and inventory checks. Enable verbose logging for rejection details."
    return
}

$results = [System.Collections.Generic.List[object]]::new()
$missing = [System.Collections.Generic.List[object]]::new()

foreach ($upn in $upns) {
    $user = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
    if (-not $user) { Write-Warning "User not found: $upn"; continue }

    if ($DeviceSource -eq 'SignInLogs') {
        $selectedDevices = @(Get-SignInDevice -User $user -SignInOnly)
    } else {
        # Joined devices appear under owned devices; registered/hybrid under registered devices. Assess the union.
        $deviceIds = @(Get-MgUserOwnedDevice -UserId $user.Id -All) + @(Get-MgUserRegisteredDevice -UserId $user.Id -All) |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.device' } |
            Select-Object -ExpandProperty Id -Unique

        $all = foreach ($id in $deviceIds) {
            $d = Get-MgDevice -DeviceId $id -Property 'id,deviceId,displayName,operatingSystem,isManaged,mdmAppId,approximateLastSignInDateTime,trustType' -ErrorAction SilentlyContinue
            if (-not $d) { continue }
            [pscustomobject]@{
                ObjectId     = $d.Id
                DeviceId     = $d.DeviceId
                DisplayName  = $d.DisplayName
                OS           = $d.OperatingSystem
                TrustType    = $d.TrustType
                IsManaged    = $d.IsManaged
                MdmAppId     = $d.MdmAppId
                LastActivity = if ($d.ApproximateLastSignInDateTime) { $d.ApproximateLastSignInDateTime } else { [datetime]::MinValue }
            }
        }
        $all | ForEach-Object { Write-Verbose "$upn : $($_.DisplayName) OS=$($_.OS) Trust=$($_.TrustType) Managed=$($_.IsManaged) Mdm=$($_.MdmAppId) Last=$($_.LastActivity)" }

        $devices = $all | Where-Object { $_.OS -like 'Windows*' }
        if ($OnlyIntuneDevices) { $devices = $devices | Where-Object { Test-PrimaryPCIntuneDevice $_.DeviceId } }
        if ($DeviceMode -eq 'Latest') {
            $devices = $devices | Where-Object { $_.IsManaged -and $_.MdmAppId -in $IntuneMdmAppIds }
        }
        $selectedDevices = @($devices | Sort-Object LastActivity -Descending)
        if ($DeviceMode -eq 'Latest') { $selectedDevices = @($selectedDevices | Select-Object -First 1) }
        $source = 'Entra'

        if ($selectedDevices.Count -eq 0 -and $SignInFallback) {
            $fb = Get-SignInDevice -User $user
            if ($fb) { $selectedDevices = @($fb); $source = $fb.Source }
        }

        if ($selectedDevices.Count -eq 0) {
            $deviceType = if ($DeviceMode -eq 'All') { 'Windows device' } else { 'Windows Intune-managed device' }
            if ($OnlyIntuneDevices) { $deviceType += ' present in Intune' }
            Write-Warning "No $deviceType for: $upn ($(@($all).Count) device(s) found; run with -Verbose for details)"
            $missing.Add($user)
            continue
        }
    }

    foreach ($device in $selectedDevices) {
        $results.Add([pscustomobject]@{
            UserPrincipalName = $upn
            DeviceName        = $device.DisplayName
            LastActivity      = $device.LastActivity
            Source            = if ($DeviceSource -eq 'SignInLogs') { $device.Source } else { $source }
            DeviceObjectId    = $device.ObjectId
        })
    }
}

# Offer the sign-in log fallback for users without a device when it wasn't requested up front.
if ($DeviceSource -eq 'Entra' -and $missing.Count -gt 0 -and -not $SignInFallback -and -not $NonInteractive -and [Environment]::UserInteractive) {
    $answer = Read-Host "$($missing.Count) user(s) have no device. Try Windows Sign In logs from the last $SignInDays days for them? (y/N)"
    if ($answer -match '^(y|yes)$') {
        Connect-MgGraph -Scopes ($scopes + 'AuditLog.Read.All') -NoWelcome | Out-Null
        foreach ($user in $missing) {
            if (-not (Test-SignInLogAccess)) { break }
            $fb = Get-SignInDevice -User $user
            if (-not $fb) { Write-Warning "Sign-in logs gave no Windows device for: $($user.UserPrincipalName)"; continue }
            $results.Add([pscustomobject]@{
                UserPrincipalName = $user.UserPrincipalName
                DeviceName        = $fb.DisplayName
                LastActivity      = $fb.LastActivity
                Source            = $fb.Source
                DeviceObjectId    = $fb.ObjectId
            })
        }
    }
}

if ($PassThru) { $results.ToArray() } else { $results | Format-Table -AutoSize }

# Header matches the Entra bulk-import template.
# -NoTypeInformation is required on Windows PowerShell 5.1 but obsolete (warns) on PowerShell 7.
$exportOptions = @{ Encoding = $(if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }) }
if ($PSVersionTable.PSVersion.Major -lt 6) { $exportOptions.NoTypeInformation = $true }
$results | Select-Object -ExpandProperty DeviceObjectId -Unique |
    ForEach-Object { [pscustomobject]@{ 'memberObjectIdOrUpn [memberObjectIdOrUpn] Required' = $_ } } |
    Export-Csv -Path $OutputPath @exportOptions

Write-Host "Exported $(@($results).Count) device(s) to $OutputPath"
