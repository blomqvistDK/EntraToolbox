<#
.SYNOPSIS
Gets selected logged-on users for Intune managed devices.

.DESCRIPTION
By default, the script retrieves all Intune managed devices from Microsoft Graph.
If CsvPath is provided, the script only processes devices listed in the CSV.
If DeviceName is provided, the script only processes the named devices.
If GroupNameOrId is provided, it processes direct device members of that Entra group.
UserMode selects the latest user (default), the most frequent usersLoggedOn userId,
or all distinct logged-on users. Results contain one row per device/selected user.

.PARAMETER CsvPath
Optional path to a CSV file that contains one or both of these headers:
- DeviceId, ManagedDeviceId, Id
- DeviceName, ManagedDeviceName, Name, displayName

An ID can be an Intune managed device ID or an Entra device ID, so Entra group
member exports can be used directly. IDs are preferred over names.

.PARAMETER DeviceName
Optional one or more Intune managed device names to resolve directly from the
command line. This parameter cannot be used together with CsvPath.

.PARAMETER GroupNameOrId
An Entra group display name or object ID. Only direct device members that are
also Intune managed devices are included. Display names must match exactly one group.

.PARAMETER OutputPath
CSV destination. Defaults to the original dated file in the working directory.
.PARAMETER UserMode
Latest selects the newest usable usersLoggedOn entry. Most counts userId
occurrences in that array and returns every user tied for the highest count.
All returns each distinct userId once. Most and All order users by their newest
logon, newest first. Array occurrences are not complete historical sign-in counts.
All modes fall back to the primary user when no usable logged-on userId exists.
.PARAMETER SkipGraphConnect
Require an existing Graph session instead of signing in.
.PARAMETER PassThru
Return the existing result objects without the additional formatted table.

.EXAMPLE
.\Get-DevicesLastUser.ps1

Retrieves all Intune managed devices and exports the result to CSV.

.EXAMPLE
.\Get-DevicesLastUser.ps1 -CsvPath .\Get-DevicesLastUser-Template.csv

Processes only the devices listed in the template CSV file.

.EXAMPLE
.\Get-DevicesLastUser.ps1 -DeviceName PC1

Processes only the managed device named PC1.

.EXAMPLE
.\Get-DevicesLastUser.ps1 -DeviceName PC1,PC2

Processes only the managed devices named PC1 and PC2.

.EXAMPLE
.\Get-DevicesLastUser.ps1 -GroupNameOrId 'Windows Devices'

Processes Intune managed devices that are direct device members of the group.

.EXAMPLE
.\Get-DevicesLastUser.ps1 -DeviceName PC1 -UserMode Most

Returns every user tied for the most userId occurrences in usersLoggedOn.

.EXAMPLE
.\Get-DevicesLastUser.ps1 -DeviceName PC1 -UserMode All

Returns one row per distinct logged-on user, newest logon first.
#>
[CmdletBinding(DefaultParameterSetName = 'AllDevices')]
param(
    [Parameter(Mandatory = $false, ParameterSetName = 'FromCsv')]
    [string]$CsvPath,

    [Parameter(Mandatory = $false, ParameterSetName = 'FromDeviceName')]
    [string[]]$DeviceName,

    [Parameter(Mandatory = $false, ParameterSetName = 'FromGroup')]
    [string]$GroupNameOrId,

    [string]$OutputPath = ".\intune_device_last_logged_on_user_$(Get-Date -Format 'yyyy-MM-dd').csv",
    [ValidateSet('Latest', 'Most', 'All')]
    [string]$UserMode = 'Latest',
    [switch]$SkipGraphConnect,
    [switch]$PassThru
)

function Get-GraphPagedCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$ProgressMessage
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    do {
        $page = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject
        if ($page.value) {
            $results.AddRange($page.value)
        }

        Write-Host ($ProgressMessage -f $results.Count) -ForegroundColor Gray
        $nextUri = $page.'@odata.nextLink'
    } while ($nextUri)

    return $results
}

function Get-ManagedDeviceSelect {
    return 'id,azureADDeviceId,deviceName,manufacturer,model,userId,userDisplayName,userPrincipalName,emailAddress,usersLoggedOn'
}

$script:ManagedDeviceIndex = $null

function Get-ManagedDeviceIndex {
    if (-not $script:ManagedDeviceIndex) {
        $select = Get-ManagedDeviceSelect
        $managedDevices = Get-GraphPagedCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$top=500&`$select=$select" -ProgressMessage 'Loaded {0} Intune devices so far...'
        $byId = @{}
        $byAzureADDeviceId = @{}
        foreach ($managedDevice in $managedDevices) {
            if ($managedDevice.id) { $byId[[string]$managedDevice.id] = $managedDevice }
            if ($managedDevice.azureADDeviceId -and $managedDevice.azureADDeviceId -ne [guid]::Empty.ToString()) {
                $byAzureADDeviceId[[string]$managedDevice.azureADDeviceId] = $managedDevice
            }
        }
        $script:ManagedDeviceIndex = @{ ById = $byId; ByAzureADDeviceId = $byAzureADDeviceId }
    }

    return $script:ManagedDeviceIndex
}

function Resolve-ManagedDeviceFromId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DeviceId
    )

    $normalizedId = $DeviceId.Trim().TrimStart('{').TrimEnd('}')
    if (-not $script:ManagedDeviceIndex) {
        $device = Get-ManagedDeviceById -DeviceId $normalizedId
        if ($device) { return $device }
    }

    $index = Get-ManagedDeviceIndex
    if ($index.ById.ContainsKey($normalizedId)) { return $index.ById[$normalizedId] }
    if ($index.ByAzureADDeviceId.ContainsKey($normalizedId)) { return $index.ByAzureADDeviceId[$normalizedId] }
    return $null
}

function Get-GraphErrorDetails {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $statusCode = $null
    $responseBody = $null
    $exceptionMessage = $ErrorRecord.Exception.Message
    $exceptionResponse = $ErrorRecord.Exception.Response

    if ($exceptionResponse) {
        if ($exceptionResponse.PSObject.Properties.Name -contains 'StatusCode') {
            $statusCode = [int]$exceptionResponse.StatusCode
        }

        if ($exceptionResponse.PSObject.Methods.Name -contains 'GetResponseStream') {
            try {
                $responseStream = $exceptionResponse.GetResponseStream()
                if ($responseStream) {
                    $streamReader = [System.IO.StreamReader]::new($responseStream)
                    try {
                        $responseBody = $streamReader.ReadToEnd()
                    } finally {
                        $streamReader.Dispose()
                        $responseStream.Dispose()
                    }
                }
            } catch {
            }
        }
    }

    return [PSCustomObject]@{
        StatusCode       = $statusCode
        Message          = $exceptionMessage
        ResponseBody     = $responseBody
        IsTransientError = $statusCode -in @(408, 429, 500, 502, 503, 504)
    }
}

function ConvertTo-ODataEscapedString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Replace("'", "''")
}

function Get-ManagedDeviceById {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DeviceId
    )

    $select = Get-ManagedDeviceSelect

    try {
        return Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$DeviceId')?`$select=$select" -OutputType PSObject
    } catch {
        Write-Verbose "Managed device id '$DeviceId' was not found; trying Entra device ID matching."
        return $null
    }
}

function Get-ManagedDevicesByName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DeviceName
    )

    $escapedDeviceName = ConvertTo-ODataEscapedString -Value $DeviceName
    $encodedFilter = [System.Uri]::EscapeDataString("deviceName eq '$escapedDeviceName'")
    $select = Get-ManagedDeviceSelect
    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$encodedFilter&`$select=$select"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
        return @($response.value)
    } catch {
        Write-Warning "Managed device name '$DeviceName' could not be resolved."
        return @()
    }
}

function Get-CsvColumnValue {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Row,

        [Parameter(Mandatory = $true)]
        [string[]]$ColumnNames
    )

    foreach ($columnName in $ColumnNames) {
        $property = $Row.PSObject.Properties | Where-Object { $_.Name -eq $columnName } | Select-Object -First 1
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return $null
}

function Get-RequestedManagedDevicesByName {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$DeviceNames
    )

    $devicesById = @{}

    foreach ($requestedDeviceName in $DeviceNames) {
        $normalizedDeviceName = $requestedDeviceName.Trim()
        if ([string]::IsNullOrWhiteSpace($normalizedDeviceName)) {
            continue
        }

        $matchedDevices = @(Get-ManagedDevicesByName -DeviceName $normalizedDeviceName)
        if ($matchedDevices.Count -eq 0) {
            Write-Warning "Managed device name '$normalizedDeviceName' could not be resolved."
            continue
        }

        if ($matchedDevices.Count -gt 1) {
            Write-Warning "Managed device name '$normalizedDeviceName' matched multiple devices. Skipping this entry because the match is ambiguous."
            continue
        }

        $matchedDevice = $matchedDevices[0]
        if (-not $devicesById.ContainsKey($matchedDevice.id)) {
            $devicesById[$matchedDevice.id] = $matchedDevice
        }
    }

    $resolvedDevices = [System.Collections.Generic.List[object]]::new()
    foreach ($resolvedDevice in $devicesById.Values) {
        $resolvedDevices.Add($resolvedDevice)
    }

    return $resolvedDevices
}

function Get-RequestedManagedDevicesByGroup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupNameOrId
    )

    $normalizedInput = $GroupNameOrId.Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedInput)) {
        throw 'Enter an Entra group name or ID.'
    }

    $parsedGroupId = [guid]::Empty
    if ([guid]::TryParse($normalizedInput, [ref]$parsedGroupId)) {
        $groupUri = "https://graph.microsoft.com/v1.0/groups/$($parsedGroupId.ToString())?`$select=id,displayName"
        try {
            $group = Invoke-MgGraphRequest -Method GET -Uri $groupUri -OutputType PSObject
        } catch {
            throw "Entra group ID '$normalizedInput' could not be resolved: $($_.Exception.Message)"
        }
        if (-not $group.id) { throw "Entra group ID '$normalizedInput' was not found." }
    } else {
        $escapedName = $normalizedInput.Replace("'", "''")
        $encodedFilter = [System.Uri]::EscapeDataString("displayName eq '$escapedName'")
        $groupUri = "https://graph.microsoft.com/v1.0/groups?`$filter=$encodedFilter&`$select=id,displayName&`$top=2"
        try {
            $groupsResponse = Invoke-MgGraphRequest -Method GET -Uri $groupUri -OutputType PSObject
            $groups = @($groupsResponse.value)
        } catch {
            throw "Entra group '$normalizedInput' could not be resolved: $($_.Exception.Message)"
        }
        if ($groups.Count -eq 0) { throw "No Entra group named '$normalizedInput' was found." }
        if ($groups.Count -gt 1) { throw "More than one Entra group is named '$normalizedInput'. Use its group ID instead." }
        $group = $groups[0]
    }

    Write-Host "Fetching direct device members of Entra group '$($group.displayName)'..." -ForegroundColor Cyan
    $membersUri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members?`$select=id,deviceId&`$top=999"
    $members = Get-GraphPagedCollection -Uri $membersUri -ProgressMessage 'Loaded {0} group members so far...'
    $deviceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($member in $members) {
        if ($member.'@odata.type' -eq '#microsoft.graph.device' -and $member.deviceId) {
            [void]$deviceIds.Add([string]$member.deviceId)
        }
    }

    if ($deviceIds.Count -eq 0) {
        Write-Host 'The group has no direct device members.' -ForegroundColor Gray
        return @()
    }

    $index = Get-ManagedDeviceIndex
    $matchedDevices = [System.Collections.Generic.List[object]]::new()
    foreach ($deviceId in $deviceIds) {
        if ($index.ByAzureADDeviceId.ContainsKey($deviceId)) {
            $matchedDevices.Add($index.ByAzureADDeviceId[$deviceId])
        } else {
            Write-Warning "Group device '$deviceId' has no Intune managed device record."
        }
    }

    return $matchedDevices
}

function Get-RequestedManagedDevices {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "CSV file not found: $Path"
    }

    $csvRows = Import-Csv -Path $Path
    if (-not $csvRows) {
        throw "CSV file '$Path' is empty."
    }

    $idHeaders = 'DeviceId', 'ManagedDeviceId', 'Id'
    $nameHeaders = 'DeviceName', 'ManagedDeviceName', 'Name', 'displayName'
    $availableHeaders = $csvRows[0].PSObject.Properties.Name

    $hasIdHeader = $availableHeaders | Where-Object { $_ -in $idHeaders }
    $hasNameHeader = $availableHeaders | Where-Object { $_ -in $nameHeaders }

    if (-not $hasIdHeader -and -not $hasNameHeader) {
        throw "CSV must contain one of these headers: $($idHeaders + $nameHeaders -join ', ')"
    }

    $devicesById = @{}
    $requestedByName = [System.Collections.Generic.List[object]]::new()

    foreach ($row in $csvRows) {
        $deviceId = Get-CsvColumnValue -Row $row -ColumnNames $idHeaders
        $deviceName = Get-CsvColumnValue -Row $row -ColumnNames $nameHeaders

        if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
            $device = Resolve-ManagedDeviceFromId -DeviceId $deviceId
            if ($device) {
                $devicesById[[string]$device.id] = $device
            } else {
                Write-Warning "Device ID '$deviceId' did not match an Intune managed device ID or Entra device ID."
            }

            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($deviceName)) {
            $requestedByName.Add($deviceName)
            continue
        }

        Write-Warning "Skipping CSV row because it does not contain a supported device identifier."
    }

    if ($requestedByName.Count -gt 0) {
        $resolvedByName = Get-RequestedManagedDevicesByName -DeviceNames $requestedByName
        foreach ($matchedDevice in $resolvedByName) {
            if (-not $devicesById.ContainsKey($matchedDevice.id)) {
                $devicesById[$matchedDevice.id] = $matchedDevice
            }
        }
    }

    $resolvedDevices = [System.Collections.Generic.List[object]]::new()
    foreach ($resolvedDevice in $devicesById.Values) {
        $resolvedDevices.Add($resolvedDevice)
    }

    return $resolvedDevices
}

function Get-UserInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Cache,

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 3
    )

    $normalizedUserId = $UserId.Trim().TrimStart('{').TrimEnd('}')

    if ([string]::IsNullOrWhiteSpace($normalizedUserId)) {
        return [PSCustomObject]@{
            UserInfo      = $null
            Error         = [PSCustomObject]@{
                StatusCode       = $null
                Message          = 'UserId is empty or whitespace.'
                ResponseBody     = $null
                IsTransientError = $false
            }
            ExtendedError = $null
            LookupPhase   = 'InvalidInput'
        }
    }

    if ($Cache.ContainsKey($normalizedUserId)) {
        return [PSCustomObject]@{
            UserInfo      = $Cache[$normalizedUserId]
            Error         = $null
            ExtendedError = $null
            LookupPhase   = 'Cached'
        }
    }

    $coreUserUriSelected = "https://graph.microsoft.com/v1.0/users/${normalizedUserId}?`$select=id,displayName,userPrincipalName"
    $coreUserUriFallback = "https://graph.microsoft.com/v1.0/users/$normalizedUserId"
    $extendedUserUri = "https://graph.microsoft.com/v1.0/users/${normalizedUserId}?`$select=jobTitle,companyName,officeLocation"
    $lastErrorDetails = $null

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            try {
                $userInfo = Invoke-MgGraphRequest -Method GET -Uri $coreUserUriSelected -OutputType PSObject
                $lookupPhase = 'CoreResolved'
            } catch {
                Write-Verbose "Core selected query failed for id '$normalizedUserId'. Trying plain user query."
                $userInfo = Invoke-MgGraphRequest -Method GET -Uri $coreUserUriFallback -OutputType PSObject
                $lookupPhase = 'FallbackResolved'
            }

            $extendedErrorDetails = $null

            try {
                $extendedUserInfo = Invoke-MgGraphRequest -Method GET -Uri $extendedUserUri -OutputType PSObject
                # Always add the properties; Graph omits null ones and strict mode fails on missing members.
                foreach ($extendedProperty in 'jobTitle', 'companyName', 'officeLocation') {
                    $extendedPropertyInfo = $extendedUserInfo.PSObject.Properties[$extendedProperty]
                    $extendedValue = if ($extendedPropertyInfo) { $extendedPropertyInfo.Value } else { $null }
                    $userInfo | Add-Member -NotePropertyName $extendedProperty -NotePropertyValue $extendedValue -Force
                }
            } catch {
                $extendedErrorDetails = Get-GraphErrorDetails -ErrorRecord $_
                $extendedErrorSummary = if ($extendedErrorDetails.StatusCode) {
                    "HTTP $($extendedErrorDetails.StatusCode): $($extendedErrorDetails.Message)"
                } else {
                    $extendedErrorDetails.Message
                }

                Write-Verbose "Extended user profile lookup failed for id '$normalizedUserId': $extendedErrorSummary"
            }

            $Cache[$normalizedUserId] = $userInfo

            return [PSCustomObject]@{
                UserInfo      = $userInfo
                Error         = $null
                ExtendedError = $extendedErrorDetails
                LookupPhase   = $lookupPhase
            }
        } catch {
            $lastErrorDetails = Get-GraphErrorDetails -ErrorRecord $_
            $errorSummary = if ($lastErrorDetails.StatusCode) {
                "HTTP $($lastErrorDetails.StatusCode): $($lastErrorDetails.Message)"
            } else {
                $lastErrorDetails.Message
            }

            Write-Verbose "User lookup attempt $attempt/$MaxAttempts failed for id '$normalizedUserId': $errorSummary"

            if (-not $lastErrorDetails.IsTransientError -or $attempt -eq $MaxAttempts) {
                break
            }
        }
    }

    return [PSCustomObject]@{
        UserInfo      = $null
        Error         = $lastErrorDetails
        ExtendedError = $null
        LookupPhase   = 'CoreLookupFailed'
    }
}

$requiredScopes = @('DeviceManagementManagedDevices.Read.All',
    'DeviceManagementManagedDevices.ReadWrite.All',
    'DeviceManagementConfiguration.Read.All',
    'DeviceManagementConfiguration.ReadWrite.All',
    'User.Read.All')
if ($GroupNameOrId) { $requiredScopes += @('Device.Read.All', 'Group.Read.All', 'GroupMember.Read.All') }
if ($SkipGraphConnect) {
    $context = Get-MgContext
    if (-not $context -or @($requiredScopes | Where-Object { $_ -notin $context.Scopes }).Count -gt 0) {
        throw "An existing Graph session with these scopes is required: $($requiredScopes -join ', ')"
    }
} else { Connect-MgGraph -Scopes $requiredScopes }

if ($CsvPath) {
    Write-Host "Fetching managed devices from CSV input..." -ForegroundColor Cyan
    $devices = @(Get-RequestedManagedDevices -Path $CsvPath)
} elseif ($DeviceName) {
    Write-Host "Fetching managed devices from device-name input..." -ForegroundColor Cyan
    $devices = @(Get-RequestedManagedDevicesByName -DeviceNames $DeviceName)
} elseif ($GroupNameOrId) {
    $devices = @(Get-RequestedManagedDevicesByGroup -GroupNameOrId $GroupNameOrId)
} else {
    Write-Host "Fetching managed devices from Microsoft Graph..." -ForegroundColor Cyan
    $select = Get-ManagedDeviceSelect
    $devices = @(Get-GraphPagedCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$top=500&`$select=$select" -ProgressMessage 'Loaded {0} devices so far...')
}

Write-Host "`nTotal devices retrieved: $($devices.Count). Processing user information...`n" -ForegroundColor Green

$userCache = @{}
$counter = 0
$userLookupFailures = @{}

$report = foreach ($device in $devices) {
    $counter++
    $primaryUserId = $device.userId
    $logons = @($device.usersLoggedOn)
    $sortedLogons = @($logons | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_.userId) } |
        Sort-Object -Property @{ Expression = {
            if ($_.lastLogOnDateTime) { [datetimeoffset]$_.lastLogOnDateTime } else { [datetimeoffset]::MinValue }
        }; Descending = $true })
    $logonCounts = @{}
    $selectedUserIds = @(
        foreach ($logon in $sortedLogons) {
            $loggedOnUserId = ([string]$logon.userId).Trim().TrimStart('{').TrimEnd('}')
            if ([string]::IsNullOrWhiteSpace($loggedOnUserId)) { continue }
            if (-not $logonCounts.ContainsKey($loggedOnUserId)) {
                $logonCounts[$loggedOnUserId] = 0
                $loggedOnUserId
            }
            $logonCounts[$loggedOnUserId]++
        }
    )
    if ($UserMode -eq 'Latest') {
        $selectedUserIds = @($selectedUserIds | Select-Object -First 1)
    } elseif ($UserMode -eq 'Most' -and $selectedUserIds.Count -gt 0) {
        $highestCount = ($logonCounts.Values | Measure-Object -Maximum).Maximum
        $selectedUserIds = @($selectedUserIds | Where-Object { $logonCounts[$_] -eq $highestCount })
    }
    $userSource = if ($selectedUserIds.Count -gt 0) { 'UsersLoggedOn' } else { 'PrimaryUser' }
    if ($userSource -eq 'PrimaryUser') { $selectedUserIds = @($primaryUserId) }

    foreach ($userId in $selectedUserIds) {
        $lastLoggedOnUserId = if ($userSource -eq 'UsersLoggedOn') { $userId } else { $null }
        # Device profile fields describe the primary user, not other logged-on users.
        $isPrimaryUser = $userSource -eq 'PrimaryUser' -or ($primaryUserId -and $userId -eq $primaryUserId)
        $userDisplayName = if ($isPrimaryUser) { $device.userDisplayName } else { $null }
        $userUpn = if ($isPrimaryUser) { $device.userPrincipalName } else { $null }
        $jobTitle = $null
        $companyName = $null
        $officeLocation = $null
        $lookupErrorMessage = $null
        $lookupPhase = $null

        if ($userId) {
            $userLookupResult = Get-UserInfo -UserId $userId -Cache $userCache
            $userInfo = $userLookupResult.UserInfo
            $lookupPhase = $userLookupResult.LookupPhase
            if ($userInfo) {
                if ($userInfo.displayName) { $userDisplayName = $userInfo.displayName }
                if ($userInfo.userPrincipalName) { $userUpn = $userInfo.userPrincipalName }
                $jobTitle = $userInfo.jobTitle
                $companyName = $userInfo.companyName
                $officeLocation = $userInfo.officeLocation
                if ($userSource -eq 'UsersLoggedOn') {
                    $resolutionStatus = 'UsersLoggedOnResolved'
                } else {
                    $resolutionStatus = 'PrimaryUserResolved'
                }

                if ($userLookupResult.ExtendedError) {
                    $extendedLookupErrorMessage = if ($userLookupResult.ExtendedError.StatusCode) {
                        "HTTP $($userLookupResult.ExtendedError.StatusCode): $($userLookupResult.ExtendedError.Message)"
                    } else {
                        $userLookupResult.ExtendedError.Message
                    }

                    if ($jobTitle -or $companyName -or $officeLocation) {
                        $lookupErrorMessage = $extendedLookupErrorMessage
                    } else {
                        $lookupErrorMessage = $extendedLookupErrorMessage
                    }
                    $lookupPhase = 'ExtendedLookupFailed'
                }
            } elseif ($userSource -eq 'UsersLoggedOn') {
                $resolutionStatus = 'UsersLoggedOnLookupFailed'
                if ($userLookupResult.Error) {
                    $lookupErrorMessage = if ($userLookupResult.Error.StatusCode) {
                        "HTTP $($userLookupResult.Error.StatusCode): $($userLookupResult.Error.Message)"
                    } else {
                        $userLookupResult.Error.Message
                    }
                    $userLookupFailures[$userId] = $lookupErrorMessage
                    Write-Warning "User lookup failed for '$userId' on device '$($device.deviceName)' (UsersLoggedOn): $lookupErrorMessage"
                }
            } else {
                $resolutionStatus = 'PrimaryUserLookupFailed'
                if ($userLookupResult.Error) {
                    $lookupErrorMessage = if ($userLookupResult.Error.StatusCode) {
                        "HTTP $($userLookupResult.Error.StatusCode): $($userLookupResult.Error.Message)"
                    } else {
                        $userLookupResult.Error.Message
                    }
                    $userLookupFailures[$userId] = $lookupErrorMessage
                    Write-Warning "User lookup failed for '$userId' on device '$($device.deviceName)' (PrimaryUser): $lookupErrorMessage"
                }
            }
        } elseif (@($logons | Where-Object { $_ }).Count -gt 0) {
            $resolutionStatus = 'UsersLoggedOnMissingUserId'
            $lookupErrorMessage = 'UsersLoggedOn entry did not contain a userId.'
            Write-Warning "Device '$($device.deviceName)' has usersLoggedOn entries, but none contained a userId. Falling back to the primary user."
        } else {
            $resolutionStatus = 'PrimaryUserMissing'
            $lookupErrorMessage = 'No usersLoggedOn userId or primary userId was available.'
        }

        Write-Host "[$counter/$($devices.Count)] Device: $($device.deviceName) | Model: $($device.model) | User: $($userUpn) | Display Name: $($userDisplayName) | Source: $userSource | Status: $resolutionStatus | Title: $jobTitle" -ForegroundColor Yellow

        [PSCustomObject]@{
            DeviceId           = $device.id
            DeviceName         = $device.deviceName
            Manufacturer       = $device.manufacturer
            Model              = $device.model
            ResolvedUserId     = $userId
            ResolvedUserSource = $userSource
            ResolutionStatus   = $resolutionStatus
            LookupPhase        = $lookupPhase
            LookupError        = $lookupErrorMessage
            UserDisplayName    = $userDisplayName
            UserPrincipalName  = $userUpn
            JobTitle           = $jobTitle
            CompanyName        = $companyName
            OfficeLocation     = $officeLocation
            UserSource         = $userSource
            LastLoggedOnUserId = $lastLoggedOnUserId
            PrimaryUserId      = $primaryUserId
        }
    }
}

$report

if (-not $PassThru) { $report | Format-Table -AutoSize }

$exportPath = $OutputPath
$exportEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
# -NoTypeInformation is required on Windows PowerShell 5.1 but obsolete (warns) on PowerShell 7.
$exportOptions = @{ Encoding = $exportEncoding }
if ($PSVersionTable.PSVersion.Major -lt 6) { $exportOptions.NoTypeInformation = $true }
$report | Export-Csv -Path $exportPath @exportOptions
Write-Host "Report exported to $exportPath" -ForegroundColor Green
 