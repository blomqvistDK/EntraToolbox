param($Path, $Parameters, $Modules, $Scopes, $Kind, $ExpectedTenant, $RequestedTenant, $Queue, $State)
$InformationPreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'
$ConfirmPreference = 'None'
try {
    foreach ($module in $Modules) { Import-Module $module -ErrorAction Stop }
    if ($Kind -eq 'Disconnect') {
        Disconnect-MgGraph -ErrorAction Stop | Out-Null
        $State.Account = ''; $State.Tenant = ''; $State.Connected = $false
        return
    }
    $context = Get-MgContext
    if ($context -and $ExpectedTenant -and $context.TenantId -ne $ExpectedTenant) { throw 'Graph tenant changed. Sign out and reconnect before continuing.' }
    if ($context -and $RequestedTenant -and $context.TenantId -ne $RequestedTenant) { throw 'The Graph session belongs to a different tenant. Sign out first.' }
    if (-not $context -or @($Scopes | Where-Object { $_ -notin $context.Scopes }).Count) {
        if ($Kind -eq 'Run' -and $Parameters['Action'] -in @('Delete','Disable') -and -not $Parameters['WhatIf']) {
            throw 'Sign in on Device Cleanup with Preview only unchecked to grant write permissions before confirming disable or deletion.'
        }
        $connect = @{ Scopes=$Scopes; NoWelcome=$true; ContextScope='Process'; ErrorAction='Stop' }
        if ($RequestedTenant) { $connect.TenantId = $RequestedTenant }
        elseif ($context) { $connect.TenantId = $context.TenantId }
        Connect-MgGraph @connect | Out-Null
        $context = Get-MgContext
    }
    if (-not $context -or @($Scopes | Where-Object { $_ -notin $context.Scopes }).Count) { throw 'Sign-in did not grant the required scopes.' }
    if (($RequestedTenant -and $context.TenantId -ne $RequestedTenant) -or ($ExpectedTenant -and $context.TenantId -ne $ExpectedTenant)) { throw 'Tenant verification failed; no task was started.' }
    $State.Account = [string]$context.Account; $State.Tenant = [string]$context.TenantId; $State.Connected = $true
    if ($Kind -eq 'Run') {
        & $Path @Parameters | ForEach-Object { $Queue.Enqueue($_) }
    }
} catch {
    $State.Fatal = $_.Exception.Message
    Write-Error -ErrorRecord $_ -ErrorAction Continue
}
