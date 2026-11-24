#Requires -Version 5.1
<#
.SYNOPSIS
    Inspect CSV-listed Entra devices or delete explicitly selected device records.
.DESCRIPTION
    DeviceId is the Entra registration ID; Id/ObjectId is the directory object ID.
    CSV accepts displayName/DeviceName, id/ObjectId and deviceId, with comma or
    semicolon delimiters. Group exports are restricted to objectType=device rows.
    Activity is the latest Entra approximate sign-in or linked Intune sync (UTC).
    Deletion removes Intune, then Autopilot, then any remaining Entra object.
    Hybrid and unknown join types are blocked. No wipe or on-premises AD action.
    Delegated deletion requires Directory.AccessAsUser.All, Intune write scopes,
    appropriate administrator roles and an explicit cloud-deletion acknowledgement.
    After all Autopilot deletion requests are accepted, waits ten seconds once
    per device before rechecking Entra. Autopilot completion is not verified.
.PARAMETER RemovalTimeoutSeconds
    Maximum confirmation wait per Intune or Entra record; defaults to 60 seconds.
.PARAMETER AutopilotRemovalTimeoutSeconds
    Deprecated compatibility parameter. Autopilot confirmation polling has been
    replaced by a fixed ten-second delay; this parameter no longer controls it.
.EXAMPLE
    .\Invoke-EntraDeviceCleanup.ps1 -CsvPath .\devices.csv -PassThru
.EXAMPLE
    $devices = .\Invoke-EntraDeviceCleanup.ps1 -CsvPath .\devices.csv -PassThru
    .\Invoke-EntraDeviceCleanup.ps1 -Action Delete -Device $devices[0] -ExpectedTenantId $devices[0].TenantId -AcknowledgeCloudDeletion -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [ValidateSet('Inspect','Delete')][string]$Action = 'Inspect',
    [string]$CsvPath,
    [object[]]$Device,
    [string]$ExpectedTenantId,
    [switch]$AcknowledgeCloudDeletion,
    [switch]$SkipGraphConnect,
    [switch]$PassThru,
    [ValidateRange(1,600)][int]$RemovalTimeoutSeconds = 60,
    [ValidateRange(1,300)][int]$AutopilotRemovalTimeoutSeconds = 300
)

function Get-CleanupProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Get-CleanupHttpStatus {
    param($Exception)
    $response = Get-CleanupProperty $Exception 'Response'
    $status = Get-CleanupProperty $response 'StatusCode'
    if (-not $status) { $status = Get-CleanupProperty $Exception 'ResponseStatusCode' }
    if ($status) { return [int]$status }
    return 0
}

function Invoke-CleanupGraph {
    param([string]$Uri, [string]$Method = 'GET', [switch]$AllowMissing)
    for ($attempt = 0; $attempt -lt 4; $attempt++) {
        try { return Invoke-MgGraphRequest -Uri $Uri -Method $Method -OutputType PSObject -ErrorAction Stop }
        catch {
            $status = Get-CleanupHttpStatus $_.Exception
            if ($AllowMissing -and $status -eq 404) { return $null }
            if ($status -notin @(429,500,502,503,504) -or $attempt -eq 3) { throw }
            $delay = [Math]::Pow(2, $attempt)
            $headers = Get-CleanupProperty (Get-CleanupProperty $_.Exception 'Response') 'Headers'
            $retryAfter = Get-CleanupProperty $headers 'RetryAfter'
            $delta = Get-CleanupProperty $retryAfter 'Delta'
            if ($delta) { $delay = [Math]::Max($delay, $delta.TotalSeconds) }
            if ($delay -gt 60) { throw 'Graph retry delay exceeds the safe operation limit; retry this device later.' }
            Write-Warning "Graph $Method returned HTTP $status. Retrying in $delay second(s) (retry $($attempt + 1) of 3): $Uri"
            [Threading.Thread]::Sleep([int]($delay * 1000))
        }
    }
}

function Get-CleanupCollection {
    param([string]$Uri)
    $items = [Collections.Generic.List[object]]::new()
    while ($Uri) {
        $page = Invoke-CleanupGraph -Uri $Uri
        $hasValue = $null -ne $page -and $(if ($page -is [Collections.IDictionary]) { $page.Contains('value') } else { $null -ne $page.PSObject.Properties['value'] })
        if (-not $hasValue -or $null -eq $page.value) { throw 'Graph returned an incomplete device inventory.' }
        foreach ($item in @(Get-CleanupProperty $page 'value')) { [void]$items.Add($item) }
        $Uri = [string](Get-CleanupProperty $page '@odata.nextLink')
        if ($Uri -and $Uri -notlike 'https://graph.microsoft.com/*') { throw 'Unexpected Graph pagination host.' }
    }
    return $items.ToArray()
}

function Get-CleanupCsvTargets {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Select an existing device CSV.' }
    $aliases = @('displayName','DeviceName','Device Name','Name','deviceId','Device ID','id','ObjectId','Object ID')
    $rows = @()
    foreach ($delimiter in @(',', ';')) {
        $candidateRows = @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -Encoding UTF8 -ErrorAction Stop)
        if ($candidateRows.Count -and @($candidateRows[0].PSObject.Properties.Name | Where-Object { $_ -in $aliases }).Count) {
            $rows = $candidateRows
            break
        }
    }
    if (-not $rows.Count) { throw 'CSV needs displayName/DeviceName, deviceId, or id/ObjectId columns and at least one row.' }
    foreach ($row in $rows) {
        $objectType = [string](Get-CleanupProperty $row 'objectType')
        if ($objectType -and $objectType -ine 'device') { Write-Warning "Skipping non-device CSV row: $objectType"; continue }
        $target = [ordered]@{ Name=''; ObjectId=''; DeviceId='' }
        foreach ($mapping in @(
            @{ Key='Name'; Names=@('displayName','DeviceName','Device Name','Name') },
            @{ Key='ObjectId'; Names=@('id','ObjectId','Object ID') },
            @{ Key='DeviceId'; Names=@('deviceId','Device ID') }
        )) {
            $values = @($mapping.Names | ForEach-Object { [string](Get-CleanupProperty $row $_) } |
                ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
            if ($values.Count -gt 1) { throw "Conflicting CSV aliases for $($mapping.Key)." }
            if ($values.Count) { $target[$mapping.Key] = $values[0] }
        }
        if (-not ($target.Name -or $target.ObjectId -or $target.DeviceId)) { throw 'A device CSV row has no name or ID.' }
        foreach ($key in @('ObjectId','DeviceId')) {
            $guid = [guid]::Empty
            if ($target[$key] -and (-not [guid]::TryParse($target[$key], [ref]$guid) -or $guid -eq [guid]::Empty)) { throw "Invalid $key in CSV for '$($target.Name)'." }
            if ($target[$key]) { $target[$key] = $guid.ToString() }
        }
        [pscustomobject]$target
    }
}

function Resolve-CleanupTarget {
    param($Target)
    $properties = 'id,deviceId,displayName,trustType,approximateLastSignInDateTime,physicalIds'
    if ($Target.ObjectId) {
        $found = Invoke-CleanupGraph -Uri "https://graph.microsoft.com/v1.0/devices/$($Target.ObjectId)?`$select=$properties" -AllowMissing
    } elseif ($Target.DeviceId) {
        $found = Get-CleanupCollection -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$($Target.DeviceId)'&`$select=$properties"
    } else {
        $escaped = [Uri]::EscapeDataString($Target.Name.Replace("'", "''"))
        $found = Get-CleanupCollection -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$escaped'&`$select=$properties"
    }
    foreach ($entry in @($found)) {
        if ($null -eq $entry) { continue }
        if ($Target.DeviceId -and (Get-CleanupProperty $entry 'deviceId') -ine $Target.DeviceId) { throw 'CSV object ID and device ID refer to different devices.' }
        if ($Target.Name -and (Get-CleanupProperty $entry 'displayName') -ine $Target.Name) { throw 'CSV device name differs from the current Entra name. Refresh the CSV before cleanup.' }
        $entry
    }
}

function New-CleanupDeviceRow {
    param($Entra, [array]$Managed, [array]$Autopilot, [string]$Tenant)
    $deviceId = [string](Get-CleanupProperty $Entra 'deviceId')
    if (-not $deviceId -or $deviceId -eq [guid]::Empty.ToString()) { throw 'Entra device has no usable registration device ID.' }
    $managedMatches = @($Managed | Where-Object { (Get-CleanupProperty $_ 'azureADDeviceId') -ieq $deviceId })
    $managedIds = @($managedMatches | ForEach-Object { [string](Get-CleanupProperty $_ 'id') })
    $autopilotMatches = @($Autopilot | Where-Object {
        (Get-CleanupProperty $_ 'azureActiveDirectoryDeviceId') -ieq $deviceId -or
        ([string](Get-CleanupProperty $_ 'managedDeviceId') -in $managedIds)
    })
    $correlationError = ''
    foreach ($registration in $autopilotMatches) {
        $linkedId = [string](Get-CleanupProperty $registration 'azureActiveDirectoryDeviceId')
        if ($linkedId -and $linkedId -ne [guid]::Empty.ToString() -and $linkedId -ine $deviceId) { $correlationError = 'Conflicting Autopilot device linkage.' }
    }
    $markers = @((Get-CleanupProperty $Entra 'physicalIds') | Where-Object { $_ -match '^\[ZTDId\]:' })
    if ($markers.Count -and -not $autopilotMatches.Count) { $correlationError = 'Autopilot marker has no verified registration. Cleanup blocked.' }
    $dates = @()
    foreach ($date in @((Get-CleanupProperty $Entra 'approximateLastSignInDateTime')) + @($managedMatches | ForEach-Object { Get-CleanupProperty $_ 'lastSyncDateTime' })) {
        if (-not $date) { continue }
        $parsed = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse([string]$date, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { throw 'Device activity contains an invalid timestamp.' }
        $dates += $parsed.UtcDateTime
    }
    $trustType = [string](Get-CleanupProperty $Entra 'trustType')
    $joinType = switch ($trustType) { 'AzureAd' { 'Microsoft Entra joined' } 'Workplace' { 'Microsoft Entra registered' } 'ServerAd' { 'Microsoft Entra hybrid joined' } default { 'Unknown' } }
    $blocked = $trustType -notin @('AzureAd','Workplace') -or [bool]$correlationError
    [pscustomobject]@{
        'Device Name' = [string](Get-CleanupProperty $Entra 'displayName')
        'Join Type' = $joinType
        'Intune managed' = if ($managedMatches.Count) { 'Yes' } else { 'No' }
        'Activity date' = if ($dates.Count) { ($dates | Sort-Object -Descending | Select-Object -First 1) } else { $null }
        'Device ID' = $deviceId
        Status = if ($correlationError) { $correlationError } elseif ($blocked) { 'Deletion blocked: hybrid or unknown join type' } else { 'Ready' }
        ObjectId = [string](Get-CleanupProperty $Entra 'id')
        TenantId = $Tenant
        ManagedDeviceIds = $managedIds
        AutopilotIds = @($autopilotMatches | ForEach-Object { [string](Get-CleanupProperty $_ 'id') })
        AutopilotRemovalRequested = $false
        CanDelete = -not $blocked
    }
}

function Wait-CleanupRemoval {
    param([string]$Uri, [int]$TimeoutSeconds, [string]$Stage = 'Device', [string]$RecordId)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $nextProgress = 0
    do {
        if ($null -eq (Invoke-CleanupGraph -Uri $Uri -AllowMissing)) { return }
        $elapsed = $clock.Elapsed.TotalSeconds
        if ($elapsed -ge $TimeoutSeconds) {
            $timeout = [TimeoutException]::new("$Stage removal requested for $RecordId, but absence was not confirmed within $TimeoutSeconds seconds. Later cleanup stages were not started. Reload and verify service-side status before retrying.")
            $timeout.Data['CleanupConfirmationTimeout'] = $true
            throw $timeout
        }
        if ($elapsed -ge $nextProgress) {
            Write-Information "Waiting for $Stage removal confirmation for $RecordId ($([int]$elapsed)/$TimeoutSeconds seconds)."
            $nextProgress = $elapsed + 15
        }
        [Threading.Thread]::Sleep([int][Math]::Max(1, [Math]::Min(2000, ($TimeoutSeconds - $elapsed) * 1000)))
    } while ($true)
}

function Remove-CleanupDeviceRecords {
    param($Row, [string]$Tenant, [int]$TimeoutSeconds)
    foreach ($stage in @(
        @{ Name='Intune'; Collection='managedDevices'; Ids=@($Row.ManagedDeviceIds) },
        @{ Name='Autopilot'; Collection='windowsAutopilotDeviceIdentities'; Ids=@($Row.AutopilotIds) }
    )) {
        foreach ($recordId in $stage.Ids) {
            if (-not $recordId) { throw "Missing $($stage.Name) record ID." }
            if ((Get-MgContext).TenantId -ne $Tenant) { throw 'Graph tenant changed; deletion stopped.' }
            $uri = "https://graph.microsoft.com/v1.0/deviceManagement/$($stage.Collection)/$recordId"
            Write-Information "$($Row.'Device Name'): removing $($stage.Name) record $recordId"
            if ($null -ne (Invoke-CleanupGraph -Uri $uri -AllowMissing)) {
                Invoke-CleanupGraph -Uri $uri -Method DELETE -AllowMissing | Out-Null
                if ($stage.Name -eq 'Autopilot') { $Row.AutopilotRemovalRequested = $true }
                else { Wait-CleanupRemoval -Uri $uri -TimeoutSeconds $TimeoutSeconds -Stage $stage.Name -RecordId $recordId }
            }
        }
        if ($stage.Name -eq 'Intune') { $Row.'Intune managed' = 'No' }
        if ($stage.Name -eq 'Autopilot' -and $Row.AutopilotRemovalRequested) {
            Write-Information "$($Row.'Device Name'): Autopilot deletion requests accepted; waiting 10 seconds before Entra cleanup."
            Start-Sleep -Seconds 10
        }
    }
    if ((Get-MgContext).TenantId -ne $Tenant) { throw 'Graph tenant changed; deletion stopped.' }
    $uri = "https://graph.microsoft.com/v1.0/devices/$($Row.ObjectId)"
    $remaining = Invoke-CleanupGraph -Uri $uri -AllowMissing
    if ($null -ne $remaining) {
        if ((Get-CleanupProperty $remaining 'deviceId') -ine $Row.'Device ID' -or (Get-CleanupProperty $remaining 'trustType') -notin @('AzureAd','Workplace')) { throw 'Entra identity or join type changed; deletion stopped.' }
        Invoke-CleanupGraph -Uri $uri -Method DELETE -AllowMissing | Out-Null
        Wait-CleanupRemoval -Uri $uri -TimeoutSeconds $TimeoutSeconds -Stage 'Entra' -RecordId $Row.ObjectId
    }
}

if ($PSBoundParameters.ContainsKey('AutopilotRemovalTimeoutSeconds')) {
    Write-Verbose 'AutopilotRemovalTimeoutSeconds is deprecated; the Autopilot-to-Entra delay is fixed at ten seconds.'
}
if ($Action -eq 'Inspect' -and (-not $CsvPath -or $Device)) { throw 'Inspection requires CsvPath, not deletion targets.' }
if ($Action -eq 'Delete' -and (-not @($Device).Count -or $CsvPath -or -not $ExpectedTenantId -or -not $AcknowledgeCloudDeletion)) {
    throw 'Deletion requires selected Device records, ExpectedTenantId and AcknowledgeCloudDeletion, not CsvPath.'
}
$targets = if ($Action -eq 'Inspect') { @(Get-CleanupCsvTargets -Path $CsvPath) } else { @($Device) }
if (-not @($targets).Count) { throw 'No device rows were supplied.' }
if (-not $SkipGraphConnect) { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
$scopes = @('Device.Read.All','DeviceManagementManagedDevices.Read.All','DeviceManagementServiceConfig.Read.All')
if ($Action -eq 'Delete' -and -not $WhatIfPreference) {
    $scopes += @('Directory.AccessAsUser.All','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementServiceConfig.ReadWrite.All')
}
$context = Get-MgContext
if ($context -and $ExpectedTenantId -and $context.TenantId -ne $ExpectedTenantId) { throw 'Graph session belongs to a different tenant.' }
if (-not $context -or @($scopes | Where-Object { $_ -notin $context.Scopes }).Count) {
    if ($SkipGraphConnect) { throw 'The existing Graph session lacks the required scopes.' }
    $connect = @{ Scopes=$scopes; ContextScope='Process'; NoWelcome=$true; ErrorAction='Stop' }
    if ($ExpectedTenantId) { $connect.TenantId = $ExpectedTenantId } elseif ($context) { $connect.TenantId = $context.TenantId }
    Connect-MgGraph @connect | Out-Null
    $context = Get-MgContext
}
if (-not $context -or ($ExpectedTenantId -and $context.TenantId -ne $ExpectedTenantId) -or @($scopes | Where-Object { $_ -notin $context.Scopes }).Count) { throw 'Graph tenant or permission verification failed.' }
$tenant = [string]$context.TenantId
$managed = @(Get-CleanupCollection -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=id,azureADDeviceId,lastSyncDateTime')
$autopilot = @(Get-CleanupCollection -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities')
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($target in $targets) {
    $row = $null
    try {
        if ((Get-MgContext).TenantId -ne $tenant) { throw 'Graph tenant changed.' }
        if ($Action -eq 'Delete') {
            if ($target.TenantId -ne $tenant -or -not $target.CanDelete) { throw 'Selected device is blocked or belongs to a different tenant.' }
            $objectId = [guid]::Parse([string]$target.ObjectId).ToString()
            $lookupTarget = [pscustomobject]@{ ObjectId=$objectId; DeviceId=[string]$target.'Device ID'; Name='' }
        } else { $lookupTarget = $target }
        $entries = @(Resolve-CleanupTarget -Target $lookupTarget)
        if (-not $entries.Count) { throw 'Device not found in Entra. No records were deleted.' }
        foreach ($entry in $entries) {
            if (-not $seen.Add([string](Get-CleanupProperty $entry 'id'))) { continue }
            $row = New-CleanupDeviceRow -Entra $entry -Managed $managed -Autopilot $autopilot -Tenant $tenant
            if ($Action -eq 'Delete') {
                if (-not $row.CanDelete) { throw $row.Status }
                if ($PSCmdlet.ShouldProcess("$tenant / $($row.'Device Name') / $($row.ObjectId)", 'Delete Intune records, deregister Autopilot and delete remaining cloud Entra device')) {
                    $row.Status = 'Cleanup started'
                    Remove-CleanupDeviceRecords -Row $row -Tenant $tenant -TimeoutSeconds $RemovalTimeoutSeconds
                    $row.Status = if ($row.AutopilotRemovalRequested) { 'Entra removed; Autopilot removal requested' } else { 'Deleted' }
                    $row.CanDelete = $false; $row.'Intune managed' = 'No'
                } else { $row.Status = 'Preview; no changes' }
            }
            if ($PassThru) { $row } else { $row | Select-Object 'Device Name','Join Type','Intune managed','Activity date','Device ID',Status }
        }
    } catch [System.Management.Automation.PipelineStoppedException] {
        throw
    } catch {
        if (-not $row) {
            $row = [pscustomobject]@{
                'Device Name'=[string](Get-CleanupProperty $target 'Name'); 'Join Type'='Unknown'; 'Intune managed'=''; 'Activity date'=$null
                'Device ID'=[string](Get-CleanupProperty $target 'DeviceId'); Status=''; ObjectId=[string](Get-CleanupProperty $target 'ObjectId')
                TenantId=$tenant; ManagedDeviceIds=@(); AutopilotIds=@(); CanDelete=$false
            }
        }
        if ($_.Exception -is [TimeoutException] -and $_.Exception.Data['CleanupConfirmationTimeout']) {
            $row.Status = "Pending confirmation: $($_.Exception.Message)"
            $row.CanDelete = $false
        } else { $row.Status = "Failed: $($_.Exception.Message)" }
        Write-Warning $row.Status
        $row
    }
}