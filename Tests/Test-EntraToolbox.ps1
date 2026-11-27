#Requires -Version 5.1
<#
.SYNOPSIS
    Offline integration checks for Entra Toolbox. All Graph commands are mocked.
.EXAMPLE
    powershell.exe -STA -File .\Tests\Test-EntraToolbox.ps1
.EXAMPLE
    pwsh.exe -STA -File .\Tests\Test-EntraToolbox.ps1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$script:CheckCount = 0
function Assert-EntraTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:CheckCount++
    Write-Host "PASS: $Message"
}
$testRoot = Split-Path $PSScriptRoot -Parent
$sourcePaths = @(
    'Invoke-EntraToolbox-GUI.ps1',
    'Shared\Invoke-EntraToolboxOperation.ps1',
    'Features\Add-MembersToAADGroup\Add-MembersToEntraGroup.ps1',
    'Features\Create-AADGroups\Create-EntraSecurityGroupsFromCsv.ps1',
    'Features\Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1',
    'Features\Get-DevicesLastUser.ps1',
    'Features\Invoke-EntraDeviceCleanup.ps1',
    'Features\ExtensionAttributes\Set-ExtensionAttribute.ps1'
)
foreach ($relativePath in $sourcePaths) {
    $parseTokens = $null; $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $testRoot $relativePath),[ref]$parseTokens,[ref]$parseErrors)
    Assert-EntraTest ($parseErrors.Count -eq 0) "Syntax: $relativePath"
}
$portableDirectory = Join-Path ([IO.Path]::GetTempPath()) ('Entra portable tests ' + [guid]::NewGuid().ToString('N'))
$portableRoot = Join-Path $portableDirectory 'Copied Entra Toolbox'
$portableUi = $null
$customUi = $null
try {
    foreach ($relativePath in $sourcePaths) {
        $destination = Join-Path $portableRoot $relativePath
        [void](New-Item -Path (Split-Path $destination -Parent) -ItemType Directory -Force)
        Copy-Item -LiteralPath (Join-Path $testRoot $relativePath) -Destination $destination
    }
    $customRoot = Join-Path $portableDirectory 'Custom feature tree'
    Copy-Item -LiteralPath (Join-Path $portableRoot 'Features') -Destination $customRoot -Recurse
    Push-Location $portableDirectory
    try {
        $portableLauncher = Join-Path $portableRoot 'Invoke-EntraToolbox-GUI.ps1'
        $portableUi = . $portableLauncher -ValidateOnly
        Assert-EntraTest ($portableUi.PageCount -eq 6 -and $portableUi.MissingScripts -eq 0) 'Copied toolbox resolves all six features from another directory containing spaces'
        Assert-EntraTest (@($portableUi.Pages | Where-Object { -not $_.Path.StartsWith((Join-Path $portableRoot 'Features'),[StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) 'Copied toolbox uses only its bundled feature paths'
        $portableUi.Window.Close()
        $portableUi = $null
        $customUi = . $portableLauncher -ScriptsRoot $customRoot -ValidateOnly
        Assert-EntraTest ($customUi.PageCount -eq 6 -and $customUi.MissingScripts -eq 0) 'ScriptsRoot override retains the existing nested feature layout'
        Assert-EntraTest (@($customUi.Pages | Where-Object { -not $_.Path.StartsWith($customRoot,[StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) 'Custom feature root does not change bundled Shared worker resolution'
        $customUi.Window.Close()
        $customUi = $null
    } finally { Pop-Location }
} finally {
    if ($portableUi) { $portableUi.Window.Close() }
    if ($customUi) { $customUi.Window.Close() }
    if (Test-Path -LiteralPath $portableDirectory) { Remove-Item -LiteralPath $portableDirectory -Recurse -Force }
}
$script:StartupProbeCount = 0
function Import-Module { $script:StartupProbeCount++; throw 'Unexpected validation module import' }
function Get-MgContext { $script:StartupProbeCount++; throw 'Unexpected validation context probe' }
try {
    $ui = . (Join-Path $testRoot 'Invoke-EntraToolbox-GUI.ps1') -ValidateOnly
    Assert-EntraTest ($script:StartupProbeCount -eq 0) 'ValidateOnly does not import Graph or probe the connection'
} finally {
    Remove-Item -LiteralPath 'Function:\Import-Module','Function:\Get-MgContext'
}
Assert-EntraTest ($script:WorkerSource -eq (Get-Content -LiteralPath (Join-Path $testRoot 'Shared\Invoke-EntraToolboxOperation.ps1') -Raw)) 'GUI loads the canonical Shared worker without embedding duplicate processing'
Assert-EntraTest ((Get-EntraOutcomeCategory ([pscustomobject]@{ Status='Added' })) -eq 'Success') 'Outcome classifier counts successful results'
Assert-EntraTest ((Get-EntraOutcomeCategory ([pscustomobject]@{ Status='Created'; Membership='Failed: denied' })) -eq 'Failed') 'Outcome classifier gives failure precedence for compound results'
Assert-EntraTest ((Get-EntraOutcomeCategory ([pscustomobject]@{ Status='Already a member' })) -eq 'Skipped') 'Outcome classifier counts no-op results as skipped'
Assert-EntraTest ((Get-EntraOutcomeCategory ([pscustomobject]@{ Status='Preview; no changes' })) -eq 'Preview') 'Outcome classifier keeps preview separate'
Assert-EntraTest ((Get-EntraOutcomeCategory ([pscustomobject]@{ Status='Ready' })) -eq 'Other') 'Outcome classifier preserves neutral feature statuses as other'
$deviceWarning = Format-EntraResultWarning ([pscustomobject]@{ DeviceId='device-id'; DeviceName='D-04527'; Manufacturer='LENOVO'; ResolutionStatus='PrimaryUserMissing'; LookupError='No usersLoggedOn userId or primary userId was available.' })
Assert-EntraTest ($deviceWarning -eq 'DeviceName: D-04527; Error: No usersLoggedOn userId or primary userId was available.') 'Device warning contains only device name and error'
$summaryJob = @{ Kind='Run'; Cancelled=$false; Counts=@{ Processed=4; Success=0; Failed=0; Skipped=0; Preview=0; Other=4 }; ExecutionErrors=0; ActivityEntries=[Collections.Generic.List[object]]::new() }
Add-EntraRunSummaryEntry $summaryJob 'INFO' 'Run started.'
Add-EntraRunSummaryEntry $summaryJob 'INFO' 'Fetching managed devices from device-name input...'
Add-EntraRunSummaryEntry $summaryJob 'WARN' "Managed device name 'D-04514' could not be resolved."
Add-EntraRunSummaryEntry $summaryJob 'WARN' "Managed device name 'D-04753' could not be resolved."
$formattedSummary = Format-EntraRunSummary $summaryJob
Assert-EntraTest ($formattedSummary -like '*Activity:*Run started.*Fetching managed devices from device-name input*') 'Run summary includes informational activity context'
Assert-EntraTest ($formattedSummary -like '*Warnings (2):*D-04514*D-04753*') 'Run summary includes unresolved-device warning details'
$ui.Window.ShowActivated = $false
$ui.Window.WindowStartupLocation = 'Manual'
$ui.Window.Left = -10000; $ui.Window.Top = -10000
$ui.Window.Show()
$statusLight = $ui.Window.FindName('ConnectionStatusLight')
Assert-EntraTest ($statusLight.Fill.Color -eq $ui.Window.Resources['ErrorBrush'].Color -and $ui.Window.FindName('ConnectionText').Text -eq 'Not connected') 'Connection light starts red and disconnected'
& {
    $probe = @{ Context=[pscustomobject]@{ Account='existing@example.test'; TenantId='existing-tenant'; Scopes=@() }; ImportError=''; ContextError=''; Reads=0; UnexpectedCalls=0 }
    function Import-Module {
        [CmdletBinding()]param([string]$Name)
        if ($Name -ne 'Microsoft.Graph.Authentication') { throw "Unexpected module: $Name" }
        if ($probe.ImportError) { throw $probe.ImportError }
    }
    function Get-MgContext {
        [CmdletBinding()]param()
        $probe.Reads++
        if ($probe.ContextError) { throw $probe.ContextError }
        $probe.Context
    }
    function Connect-MgGraph { $probe.UnexpectedCalls++; throw 'Unexpected sign-in' }
    function Disconnect-MgGraph { $probe.UnexpectedCalls++; throw 'Unexpected sign-out' }
    function Invoke-MgGraphRequest { $probe.UnexpectedCalls++; throw 'Unexpected Graph request' }
    $page = $script:Pages[$script:PagesHost.SelectedIndex]
    try {
        Initialize-EntraConnection
        Assert-EntraTest ($ui.Window.FindName('ConnectionText').Text -eq 'Connected' -and $statusLight.Fill.Color -eq $ui.Window.Resources['SuccessBrush'].Color) 'Startup detects an existing context without requiring feature scopes'
        Assert-EntraTest ($script:ConnectedTenant -eq 'existing-tenant' -and $ui.Window.FindName('AccountText').Text -eq 'existing@example.test' -and $ui.Window.FindName('TenantText').Text -eq 'existing-tenant' -and $ui.Window.FindName('DisconnectButton').IsEnabled) 'Startup populates account, actual tenant, and Sign out availability'
        $script:RequestedTenant = 'requested-tenant'
        Initialize-EntraConnection
        Assert-EntraTest ($script:ConnectedTenant -eq 'existing-tenant' -and $ui.Window.FindName('ConnectionText').ToolTip -like '*different tenant*' -and $page.ActivitySummary.Text -like '*[[]WARN[]]*different tenant*') 'Startup shows the actual connection and warns when the requested tenant differs'
        $script:RequestedTenant = ''
        $probe.Context = $null
        & { Set-StrictMode -Version Latest; Initialize-EntraConnection }
        Assert-EntraTest ($ui.Window.FindName('ConnectionText').Text -eq 'Not connected' -and -not $script:ConnectedTenant -and -not $ui.Window.FindName('DisconnectButton').IsEnabled -and -not $ui.Window.FindName('AccountText').Text -and -not $ui.Window.FindName('TenantText').Text -and $statusLight.Fill.Color -eq $ui.Window.Resources['ErrorBrush'].Color) 'No startup context clears identity and leaves the GUI disconnected'
        $probe.ContextError = 'Mock context read failure'
        Initialize-EntraConnection
        Assert-EntraTest ($ui.Window.FindName('ConnectionText').Text -eq 'Connection check failed' -and $page.ActivitySummary.Text -like '*[[]ERROR[]]*Mock context read failure*' -and -not $script:ConnectedTenant) 'Context-read errors are visible rather than reported as a successful disconnected check'
        $probe.ContextError = ''
        $probe.ImportError = 'Mock missing authentication module'
        $readsBefore = $probe.Reads
        Initialize-EntraConnection
        Assert-EntraTest ($probe.Reads -eq $readsBefore -and $ui.Window.FindName('ConnectionText').ToolTip -like '*Mock missing authentication module*' -and $page.ActivitySummary.Text -like '*[[]ERROR[]]*Mock missing authentication module*') 'Module import failure is reported without probing or signing in'
        Assert-EntraTest ($probe.UnexpectedCalls -eq 0 -and -not $script:Job -and -not $script:Runspace) 'Startup detection never authenticates, disconnects, requests Graph data, or starts a feature worker'
    } finally {
        $script:RequestedTenant = ''
        Set-EntraConnectionState -Connected $false
        $page.Log.Document.Blocks.Clear()
        $page.ActivitySummary.Text = ''
        $page.ActivitySummary.ToolTip = ''
        $page.Activity.IsExpanded = $false
    }
}
Assert-EntraTest ($ui.PageCount -eq 6 -and $ui.MissingScripts -eq 0) 'Six pages and companion script paths'
$launcherHelp = Get-Content -LiteralPath (Join-Path $testRoot 'Invoke-EntraToolbox-GUI.ps1') -Raw
$launcherHelp = $launcherHelp.Substring(0,$launcherHelp.IndexOf('#>'))
foreach ($descriptor in $ui.Pages) {
    $relativeScript = $descriptor.Path.Substring($testRoot.Length).TrimStart('\')
    Assert-EntraTest ($launcherHelp.Contains($relativeScript)) "Deployment header lists companion script: $relativeScript"
}
Assert-EntraTest ($ui.Window.FindName('PageTitle').Text -eq 'Find Primary PCs') 'First menu page (Find Primary PCs) header is initialized'
$navLayout = @(foreach ($child in $script:Navigation.Children) {
    if ($child -is [Windows.Controls.TextBlock]) { "[$($child.Text)]" } else { foreach ($item in $child.Children) { $item.Content } }
}) -join '|'
Assert-EntraTest ($navLayout -eq '[USER-DEVICE AFFINITY]|Find Primary PCs|Device Last User(s)|[GROUPS]|Add Group Members|Create Security Groups|[USERS]|User Extension Attributes|[DEVICES]|Device Cleanup') "Menu groups features under section headings: $navLayout"
Assert-EntraTest (@($ui.Pages | Where-Object { $_.Navigation.IsChecked }).Id -join ',' -eq 'PrimaryPC') 'Startup checks only the first menu item'
$themeButton = $ui.Window.FindName('ThemeButton')
$themeButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
Assert-EntraTest ($ui.Window.Resources['BgBrush'].Color.ToString() -eq '#FFF5F6F8') 'Light theme applies through resource references'
Assert-EntraTest ($statusLight.Fill.Color -eq $ui.Window.Resources['ErrorBrush'].Color) 'Disconnected light remains red in light theme'
$themeButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
Assert-EntraTest ($ui.Window.Resources['BgBrush'].Color.ToString() -eq '#FF191C22') 'Dark theme restores through resource references'
$memberPage = $ui.Pages | Where-Object { $_.Id -eq 'Members' }
$groupPage = $ui.Pages | Where-Object { $_.Id -eq 'Groups' }
$primaryPage = $ui.Pages | Where-Object { $_.Id -eq 'PrimaryPC' }
$lastPage = $ui.Pages | Where-Object { $_.Id -eq 'LastUser' }
$attributePage = $ui.Pages | Where-Object { $_.Id -eq 'Attributes' }
$cleanupPage = $ui.Pages | Where-Object { $_.Id -eq 'DeviceCleanup' }
Assert-EntraTest ($cleanupPage.Table.Columns.Count -eq 8 -and $cleanupPage.Controls.Preview.IsChecked -and -not $cleanupPage.DeleteSelected.IsEnabled -and -not $cleanupPage.DisableSelected.IsEnabled) 'Cleanup columns, preview default and disabled deletion'
Assert-EntraTest (-not $cleanupPage.Controls.DeleteAutopilotRegistration.IsChecked) 'Cleanup keeps Autopilot registration by default'
Assert-EntraTest ($cleanupPage.Controls.Source.SelectedItem -eq 'CSV file' -and $cleanupPage.Fields.CsvPath.Visibility -eq 'Visible' -and $cleanupPage.Fields.InactivityDays.Visibility -eq 'Collapsed' -and $cleanupPage.Fields.DeviceNames.Visibility -eq 'Collapsed') 'Cleanup defaults to the CSV source'
$cleanupPage.Controls.Source.SelectedItem = 'Entra lookup'
Assert-EntraTest ($cleanupPage.Fields.CsvPath.Visibility -eq 'Collapsed' -and $cleanupPage.Fields.InactivityDays.Visibility -eq 'Visible' -and $cleanupPage.Controls.InactivityDays.Text -eq '90') 'Entra lookup shows inactivity days defaulting to 90'
$entraInvocation = Get-EntraInvocation $cleanupPage
Assert-EntraTest ($entraInvocation.Parameters.Source -eq 'Entra' -and $entraInvocation.Parameters.InactivityDays -eq 90 -and -not $entraInvocation.Parameters.ContainsKey('CsvPath')) 'Entra lookup builds Source Entra with inactivity days'
$cleanupPage.Controls.InactivityDays.Text = '0'
$daysRejected = $false
try { Get-EntraInvocation $cleanupPage | Out-Null } catch { $daysRejected = $true }
Assert-EntraTest $daysRejected 'Entra lookup rejects invalid inactivity days'
$cleanupPage.Controls.InactivityDays.Text = '90'
$cleanupPage.Controls.Source.SelectedItem = 'Device names'
$cleanupPage.Controls.DeviceNames.Text = " PC-1 `r`n`r`nPC-2"
$namesInvocation = Get-EntraInvocation $cleanupPage
Assert-EntraTest ($cleanupPage.Fields.DeviceNames.Visibility -eq 'Visible' -and $namesInvocation.Parameters.Source -eq 'Names' -and ($namesInvocation.Parameters.DeviceName -join ',') -eq 'PC-1,PC-2') 'Device names source passes trimmed names'
$cleanupPage.Controls.DeviceNames.Text = ''; $cleanupPage.Controls.Source.SelectedItem = 'CSV file'
Assert-EntraTest (-not $primaryPage.Run.IsEnabled) 'Primary-PC Run initially disabled despite output and lookback defaults'
$primarySources = @{ 'Group name'='GroupName'; 'Group ID'='GroupId'; 'User CSV'='CsvPath'; 'User UPNs'='UserPrincipalName' }
foreach ($source in $primarySources.Keys) {
    $key = $primarySources[$source]
    $primaryPage.Controls.Source.SelectedItem = $source
    Assert-EntraTest ($primaryPage.Fields[$key].Visibility -eq 'Visible' -and @($primarySources.Values | Where-Object { $primaryPage.Fields[$_].Visibility -eq 'Visible' }).Count -eq 1) "$source displays only its matching input"
    $primaryPage.Controls[$key].Text = " `t`r`n"
    Assert-EntraTest (-not $primaryPage.Run.IsEnabled) "Whitespace in $key does not enable Primary-PC Run"
    $primaryPage.Controls[$key].Text = 'test-input'
    Assert-EntraTest $primaryPage.Run.IsEnabled "$key independently enables Primary-PC Run"
    $primaryPage.Controls[$key].Text = ''
    Assert-EntraTest (-not $primaryPage.Run.IsEnabled) "Clearing $key disables Primary-PC Run"
}
$primaryPage.Controls.GroupName.Text = 'Test group'
$primaryPage.Controls.UserPrincipalName.Text = 'user@example.test'
$primaryPage.Controls.Source.SelectedItem = 'Group name'
$primaryPage.Controls.GroupName.Text = ''
Assert-EntraTest (-not $primaryPage.Run.IsEnabled) 'Hidden populated UPNs do not enable the selected empty group source'
$primaryPage.Controls.Source.SelectedItem = 'User UPNs'
Assert-EntraTest $primaryPage.Run.IsEnabled 'Switching to populated UPNs enables Run'
Set-EntraBusy $true
$primaryPage.Controls.UserPrincipalName.Text = 'another@example.test'
Assert-EntraTest (-not $primaryPage.Run.IsEnabled) 'Source changes cannot enable Primary-PC Run while busy'
$primaryPage.Controls.UserPrincipalName.Text = ''
Set-EntraBusy $false
Assert-EntraTest (-not $primaryPage.Run.IsEnabled) 'Busy completion without a source leaves Primary-PC Run disabled'
$primaryPage.Controls.GroupName.Text = 'Test group'
$primaryPage.Controls.Source.SelectedItem = 'Group name'
$primaryPage.Available = $false
Update-EntraRunAvailability $primaryPage
Assert-EntraTest (-not $primaryPage.Run.IsEnabled) 'Missing script disables Primary-PC Run despite source input'
$primaryPage.Available = $true
$primaryPage.Reset.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
Assert-EntraTest (-not $primaryPage.Run.IsEnabled) 'Reset disables Primary-PC Run'
Assert-EntraTest ($primaryPage.Controls.Source.SelectedItem -eq 'Choose a source' -and @($primarySources.Values | Where-Object { $primaryPage.Fields[$_].Visibility -eq 'Visible' }).Count -eq 0) 'Reset restores placeholder and hides source inputs'
Assert-EntraTest $memberPage.Run.IsEnabled 'Other pages retain existing Run availability'
$memberPage.Controls.TargetGroupName.Text = "Team O'Brien"
$memberPage.Controls.MemberNames.Text = "Group, One`nGroup Two"
$mapping = Get-EntraInvocation $memberPage
Assert-EntraTest ($mapping.Parameters.MemberNames.Count -eq 2 -and $mapping.Parameters.MemberNames[0] -eq 'Group, One' -and $mapping.Parameters.TargetGroupName -eq "Team O'Brien") 'Typed names preserve apostrophes and commas'
$lastRejected = $false
try { [void](Get-EntraInvocation $lastPage) } catch { $lastRejected = $_.Exception.Message -like 'Choose device names*' }
Assert-EntraTest $lastRejected 'All-device scan requires explicit input mode'
$lastPage.Controls.Source.SelectedItem = 'Entra group'
$lastPage.Controls.GroupNameOrId.Text = 'Device Group'
$mapping = Get-EntraInvocation $lastPage
$requirements = Get-EntraRequirements $lastPage $mapping.Parameters
Assert-EntraTest ($mapping.Parameters.GroupNameOrId -eq 'Device Group' -and $lastPage.Fields.GroupNameOrId.Visibility -eq 'Visible') 'Device Last User accepts an Entra group name in the GUI'
Assert-EntraTest ('Device.Read.All' -in $requirements.Scopes -and 'Group.Read.All' -in $requirements.Scopes -and 'GroupMember.Read.All' -in $requirements.Scopes) 'Device Last User requests group/device scopes only for group input'
Assert-EntraTest ($lastPage.Navigation.Content -eq 'Device Last User(s)' -and $mapping.Parameters.UserMode -eq 'Latest') 'Device Last User(s) defaults to latest user selection'
$attributeChecks = $lastPage.Controls.ExtensionAttributeNumbers.Children
$attributeState = $lastPage.Controls.ExtensionAttributeNumbers.Tag
Assert-EntraTest ($attributeChecks.Count -eq 15 -and ($attributeChecks.Content -join ',') -eq ((1..15 | ForEach-Object { "extensionAttribute$_" }) -join ',') -and -not $mapping.Parameters.ContainsKey('ExtensionAttributeNumbers')) 'All 15 optional attribute checkboxes start unselected and omit the parameter'
Assert-EntraTest ($attributeState.Toggle -is [Windows.Controls.Primitives.ToggleButton] -and -not $attributeState.Popup.IsOpen -and $attributeState.Summary.Text -eq 'Choose attributes') 'Attribute selector is a closed dropdown showing a placeholder'
$attributeChecks[0].IsChecked = $true
$attributeChecks[4].IsChecked = $true
$attributeChecks[14].IsChecked = $true
$mapping = Get-EntraInvocation $lastPage
Assert-EntraTest (($mapping.Parameters.ExtensionAttributeNumbers -join ',') -eq '1,5,15') 'Multiple attribute selections map to numeric script parameters'
Assert-EntraTest ($attributeState.Summary.Text -eq 'extensionAttribute1, extensionAttribute5, extensionAttribute15' -and $attributeState.Toggle.ToolTip -like '*extensionAttribute15*') 'Dropdown summary lists the selected attributes in order'
$lastPage.Navigation.IsChecked = $true
$ui.Window.UpdateLayout()
$attributeState.Toggle.IsChecked = $true
# Popup opening waits for the deferred Loaded event, so let the dispatcher run.
$attributeFrame = [Windows.Threading.DispatcherFrame]::new()
[void]$ui.Window.Dispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background, [Action]{ $attributeFrame.Continue = $false })
[Windows.Threading.Dispatcher]::PushFrame($attributeFrame)
$ui.Window.UpdateLayout()
Assert-EntraTest ($attributeState.Popup.IsOpen) 'Opening the dropdown shows the attribute list'
$attributeChecks[1].IsChecked = $true
Assert-EntraTest ($attributeState.Popup.IsOpen -and $attributeState.Summary.Text -like '*extensionAttribute2*') 'Checking an attribute keeps the dropdown open for further selections'
$attributeChecks[1].IsChecked = $false
$popupScroll = $attributeState.Popup.Child.Child
foreach ($check in $attributeChecks) {
    $check.BringIntoView()
    $popupScroll.UpdateLayout()
    $checkPosition = $check.TranslatePoint([Windows.Point]::new(0,0),$popupScroll)
    Assert-EntraTest ($check.ActualWidth -gt 0 -and $checkPosition.Y -ge -1 -and ($checkPosition.Y + $check.ActualHeight) -le ($popupScroll.ActualHeight + 1)) "$($check.Content) remains reachable in the dropdown list"
}
Assert-EntraTest ($popupScroll.ActualHeight -le 300) 'Dropdown list height is bounded'
Set-EntraBusy $true
Assert-EntraTest (-not $attributeState.Toggle.IsEnabled -and -not $attributeState.Popup.IsOpen) 'Attribute dropdown is disabled and closed while an operation is running'
Set-EntraBusy $false
$attributeState.Toggle.IsChecked = $true
foreach ($mode in @{'Latest logged-on user'='Latest'; 'Most logged-on user(s)'='Most'; 'All logged-on users'='All'}.GetEnumerator()) {
    $lastPage.Controls.UserMode.SelectedItem = $mode.Key
    $mapping = Get-EntraInvocation $lastPage
    Assert-EntraTest ($mapping.Parameters.UserMode -eq $mode.Value -and $mapping.Parameters.GroupNameOrId -eq 'Device Group') "User selection maps $($mode.Key) without changing device source"
}
$lastPage.Reset.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
Assert-EntraTest ($lastPage.Controls.UserMode.SelectedItem -eq 'Latest logged-on user') 'Device-user reset restores Latest selection'
Assert-EntraTest (@($attributeChecks | Where-Object { $_.IsChecked }).Count -eq 0 -and $attributeState.Summary.Text -eq 'Choose attributes' -and -not $attributeState.Popup.IsOpen) 'Device-user reset clears every attribute selection and closes the dropdown'
$lastPage.Navigation.IsChecked = $true
$ui.Window.UpdateLayout()
Assert-EntraTest ($attributeState.Toggle.ActualHeight -le 40 -and $lastPage.Table.ActualHeight -ge $lastPage.Table.MinHeight) 'Attribute dropdown takes a single row while retaining the output area'
$lastPage.FormScroll.ScrollToTop()
$primaryPage.Navigation.IsChecked = $true
$lastPage.Controls.Source.SelectedItem = 'Choose a source'
$attributePage.Controls.UserPrincipalName.Text = 'user@example.test'
$attributePage.Controls.Action.SelectedItem = 'Clear'
$mapping = Get-EntraInvocation $attributePage
Assert-EntraTest ($mapping.Parameters.Clear -and -not $mapping.Parameters.ContainsKey('Value') -and -not $attributePage.Controls.Value.IsEnabled) 'Clear attribute disables and omits value'
$primaryPage.Controls.UserPrincipalName.Text = 'user@example.test'
$primaryPage.Controls.Source.SelectedItem = 'User UPNs'
$primaryPage.Controls.GroupId.Text = 'invalid-hidden-id'
$primaryPage.Controls.CsvPath.Text = 'missing-hidden.csv'
$mapping = Get-EntraInvocation $primaryPage
Assert-EntraTest ($mapping.Parameters.UserPrincipalName.Count -eq 1 -and -not $mapping.Parameters.ContainsKey('GroupId') -and -not $mapping.Parameters.ContainsKey('CsvPath')) 'Only selected UPNs are validated and submitted'
$primaryPage.Controls.Source.SelectedItem = 'Group ID'
$groupIdRejected = $false
try { [void](Get-EntraInvocation $primaryPage) } catch { $groupIdRejected = $_.Exception.Message -eq 'Group ID must be a GUID.' }
Assert-EntraTest $groupIdRejected 'Selected Group ID retains GUID validation'
$primaryPage.Controls.GroupId.Text = '11111111-2222-3333-4444-555555555555'
$groupMapping = Get-EntraInvocation $primaryPage
Assert-EntraTest ($groupMapping.Parameters.GroupId -eq $primaryPage.Controls.GroupId.Text -and -not $groupMapping.Parameters.ContainsKey('UserPrincipalName')) 'Selected Group ID excludes hidden UPNs'
$primaryPage.Controls.Source.SelectedItem = 'Choose a source'
$primarySourceRejected = $false
try { [void](Get-EntraInvocation $primaryPage) } catch { $primarySourceRejected = $_.Exception.Message -like 'Choose a group name*' }
Assert-EntraTest ($primarySourceRejected -and -not $primaryPage.Run.IsEnabled) 'Placeholder rejects hidden populated sources'
$primaryPage.Controls.Source.SelectedItem = 'User UPNs'
Assert-EntraTest ($mapping.Parameters.NonInteractive -and -not $mapping.Parameters.SignInFallback) 'Noninteractive primary-PC lookup does not enable fallback'
Assert-EntraTest (-not $mapping.Parameters.Verbose -and -not $primaryPage.Controls.Verbose.IsChecked) 'Primary-PC verbose logging defaults to off'
Assert-EntraTest (-not $mapping.Parameters.OnlyIntuneDevices) 'Intune inventory filtering defaults to off'
$primaryPage.Controls.OnlyIntuneDevices.IsChecked = $true
$intuneMapping = Get-EntraInvocation $primaryPage
$intuneRequirements = Get-EntraRequirements $primaryPage $intuneMapping.Parameters
Assert-EntraTest ($intuneMapping.Parameters.OnlyIntuneDevices -and 'DeviceManagementManagedDevices.Read.All' -in $intuneRequirements.Scopes) 'Intune checkbox passes filter and adds inventory read scope'
Assert-EntraTest ('DeviceManagementManagedDevices.Read.All' -in (Get-EntraRequirements $primaryPage @{}).Scopes) 'Sign-in requests Intune scope when the checkbox is enabled'
$primaryPage.Controls.OnlyIntuneDevices.IsChecked = $false
Assert-EntraTest ('DeviceManagementManagedDevices.Read.All' -notin (Get-EntraRequirements $primaryPage (Get-EntraInvocation $primaryPage).Parameters).Scopes) 'Intune scope is not requested when the filter is disabled'
$primaryPage.Controls.Verbose.IsChecked = $true
$verboseMapping = Get-EntraInvocation $primaryPage
Assert-EntraTest ($verboseMapping.Parameters.Verbose) 'Primary-PC verbose checkbox passes the common Verbose parameter'
Assert-EntraTest ($mapping.Parameters.DeviceMode -eq 'Latest') 'Primary-PC selection defaults to Latest'
$primaryPage.Controls.DeviceMode.SelectedItem = 'All'
$mapping = Get-EntraInvocation $primaryPage
Assert-EntraTest ($mapping.Parameters.DeviceMode -eq 'All') 'All device selection maps to script parameter'
$primaryPage.Reset.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
Assert-EntraTest ($primaryPage.Controls.DeviceMode.SelectedItem -eq 'Latest') 'Reset restores Latest device selection'
Assert-EntraTest (-not $primaryPage.Controls.Verbose.IsChecked) 'Reset disables Primary-PC verbose logging'
Assert-EntraTest (-not $primaryPage.Controls.OnlyIntuneDevices.IsChecked) 'Reset disables Intune inventory filtering'
$primaryPage.Controls.UserPrincipalName.Text = 'user@example.test'
$primaryPage.Controls.Source.SelectedItem = 'User UPNs'
$primaryPage.Controls.SignInFallback.IsChecked = $true
$mapping = Get-EntraInvocation $primaryPage
$requirements = Get-EntraRequirements $primaryPage $mapping.Parameters
Assert-EntraTest ('AuditLog.Read.All' -in $requirements.Scopes -and 'Microsoft.Graph.Reports' -in $requirements.Modules) 'Fallback scopes and module are opt-in'
$primaryPage.Controls.DeviceSource.SelectedItem = 'Sign-in logs only'
$mapping = Get-EntraInvocation $primaryPage
$requirements = Get-EntraRequirements $primaryPage $mapping.Parameters
Assert-EntraTest ($mapping.Parameters.DeviceSource -eq 'SignInLogs' -and $mapping.Parameters.DeviceMode -eq 'Most' -and -not $mapping.Parameters.SignInFallback) 'Sign-in-only discovery defaults to Most and clears fallback'
Assert-EntraTest ($primaryPage.Controls.SignInDays.IsEnabled -and $primaryPage.Fields.SignInFallback.Visibility -eq 'Collapsed' -and -not $primaryPage.Controls.SignInFallback.IsChecked -and $mapping.Parameters.SignInDays -eq 30) 'Sign-in-only discovery enables lookback and hides fallback'
Assert-EntraTest ('AuditLog.Read.All' -in $requirements.Scopes -and 'Microsoft.Graph.Reports' -in $requirements.Modules) 'Sign-in-only Run requires audit scope and Reports module'
Assert-EntraTest ('AuditLog.Read.All' -in (Get-EntraRequirements $primaryPage @{}).Scopes) 'Sign-in-only connection requirements reflect the selected discovery source'
$primaryPage.Controls.DeviceMode.SelectedItem = 'All logged-on PCs'
Assert-EntraTest ((Get-EntraInvocation $primaryPage).Parameters.DeviceMode -eq 'All') 'All logged-on PCs maps to All in sign-in-only discovery'
foreach ($invalidDays in @('0','366','invalid')) {
    $primaryPage.Controls.SignInDays.Text = $invalidDays
    $daysRejected = $false
    try { [void](Get-EntraInvocation $primaryPage) } catch { $daysRejected = $_.Exception.Message -eq 'Sign-in days must be 1-365.' }
    Assert-EntraTest $daysRejected "Sign-in-only discovery rejects lookback '$invalidDays'"
}
$primaryPage.Reset.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
Assert-EntraTest ($primaryPage.Controls.DeviceSource.SelectedItem -eq 'Entra owned/registered devices' -and $primaryPage.Controls.DeviceMode.SelectedItem -eq 'Latest' -and -not $primaryPage.Controls.SignInFallback.IsChecked) 'Reset restores Entra/Latest discovery after sign-in-only All'
Assert-EntraTest ($primaryPage.Fields.SignInFallback.Visibility -eq 'Visible' -and $primaryPage.Controls.SignInFallback.IsEnabled -and -not $primaryPage.Controls.SignInDays.IsEnabled) 'Reset shows fallback and disables unused lookback'
$primaryPage.Controls.UserPrincipalName.Text = 'user@example.test'
$primaryPage.Controls.Source.SelectedItem = 'User UPNs'
$primaryPage.Controls.SignInFallback.IsChecked = $false

function Get-MgContext {
    [pscustomobject]@{ Account='user@example.test'; TenantId='11111111-1111-1111-1111-111111111111'; Scopes=@('GroupMember.ReadWrite.All','Group.Read.All','Group.ReadWrite.All','Directory.ReadWrite.All','Directory.AccessAsUser.All','User.Read.All','User.ReadWrite.All','GroupMember.Read.All','Device.Read.All','DeviceManagementManagedDevices.Read.All','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementServiceConfig.Read.All','DeviceManagementServiceConfig.ReadWrite.All','DeviceManagementConfiguration.Read.All','DeviceManagementConfiguration.ReadWrite.All','AuditLog.Read.All') }
}
function Connect-MgGraph { throw 'Unexpected real sign-in' }
function Read-Host { throw 'Unexpected console prompt' }
function Get-Module { [pscustomobject]@{ Name='Mock module' } }
function Import-Module { }
function Get-MgGroup { [CmdletBinding()] param([string]$Filter,[string]$ConsistencyLevel,[string]$CountVariable) [pscustomobject]@{ Id=$(if ($Filter -like "*'Target'*") { 'target-group' } else { 'group-1' }); DisplayName='Test Group' } }
function Get-MgGroupMember { [CmdletBinding()] param([string]$GroupId,[switch]$All) @() }
function New-MgGroupMemberByRef { [CmdletBinding()] param($GroupId,$BodyParameter) $global:EntraMockState.Writes++ }
function New-MgGroup { [CmdletBinding()] param($DisplayName,$Description,$MailEnabled,$SecurityEnabled,$MailNickname) $global:EntraMockState.Writes++; [pscustomobject]@{ Id='new-group'; DisplayName=$DisplayName } }
function Get-MgUser { [CmdletBinding()] param($UserId,$Filter) [pscustomobject]@{ Id='user-1'; UserPrincipalName='user@example.test'; DisplayName='Test User' } }
function Get-MgUserOwnedDevice {
    [CmdletBinding()] param($UserId,[switch]$All)
    if ($global:EntraMockState.NoDevices) { return }
    $ownedIds = if ($global:EntraMockState.Devices) { @('device-1','device-2','device-3') } else { @('device-1') }
    foreach ($deviceId in $ownedIds) { [pscustomobject]@{ Id=$deviceId; AdditionalProperties=@{ '@odata.type'='#microsoft.graph.device' } } }
}
function Get-MgUserRegisteredDevice {
    [CmdletBinding()] param($UserId,[switch]$All)
    if ($global:EntraMockState.Devices -and -not $global:EntraMockState.NoDevices) {
        foreach ($deviceId in @('device-1','device-4','device-5')) { [pscustomobject]@{ Id=$deviceId; AdditionalProperties=@{ '@odata.type'='#microsoft.graph.device' } } }
    }
}
function Get-MgDevice {
    [CmdletBinding()] param($DeviceId,$Property,$Filter)
    if ($DeviceId -and $global:EntraMockState.Devices) { return $global:EntraMockState.Devices[$DeviceId] }
    [pscustomobject]@{ Id='device-1'; DisplayName='TEST-PC'; OperatingSystem='Windows'; IsManaged=$true; MdmAppId='0000000a-0000-0000-c000-000000000000'; ApproximateLastSignInDateTime=[datetime]'2026-10-01'; TrustType='AzureAd' }
}
function Get-MgAuditLogSignIn {
    [CmdletBinding()] param($Filter,[switch]$All)
    $global:EntraMockState.SignInCalls++
    [pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId='entra-device-1' }; CreatedDateTime=[datetime]'2026-10-20' }
}
function Invoke-MgGraphRequest {
    [CmdletBinding()] param($Method,$Uri,$OutputType,$Body,$ContentType)
    if ($Method -eq 'PATCH') { $global:EntraMockState.Writes++; return }
    if ($Uri -match '/v1\.0/groups\?') {
        return [pscustomobject]@{ value=@($global:EntraMockState.GroupNameMatches); '@odata.nextLink'=$null }
    }
    if ($Uri -match '/v1\.0/groups/[^/]+/members') {
        return [pscustomobject]@{ value=@($global:EntraMockState.GroupMembers); '@odata.nextLink'=$null }
    }
    if ($Uri -match '/v1\.0/groups/([^/?]+)\?') {
        return [pscustomobject]@{ id=$Matches[1]; displayName='Device Group' }
    }
    if ($Uri -like '*managedDevices*') {
        $device = [pscustomobject]@{ id='managed-1'; deviceName='TEST-PC'; manufacturer='Test'; model='Model'; azureADDeviceId='aad-device-1'; userId='user-1'; userDisplayName='Test User'; userPrincipalName='user@example.test'; emailAddress='user@example.test'; usersLoggedOn=@() }
        if ($global:EntraMockState.ContainsKey('LastUserDevice')) { $device = $global:EntraMockState.LastUserDevice }
        if ($Uri -match '\$select=([^&]+)') {
            $selected = $Matches[1] -split ','
            $device = $device | Select-Object -Property ($device.PSObject.Properties.Name | Where-Object { $_ -in $selected })
        }
        if ($Uri -match "managedDevices\('([^']+)'\)") {
            if ($Matches[1] -ne 'managed-1') { throw 'Response status code does not indicate success: NotFound (Not Found).' }
            return $device
        }
        $mockDevices = @($device)
        if ($global:EntraMockState['RepeatLastUserDevice']) {
            $secondDevice = $device | Select-Object *
            $secondDevice.id = 'managed-2'
            $secondDevice.deviceName = 'TEST-PC-2'
            $mockDevices += $secondDevice
        }
        return [pscustomobject]@{ value=$mockDevices; '@odata.nextLink'=$null }
    }
    if ($Uri -match '/v1\.0/users/([^?]+)') {
        $requestedUserId = $Matches[1]
        if ($global:EntraMockState.ContainsKey('LastUserRequests')) { $global:EntraMockState.LastUserRequests.Add([string]$Uri) }
        if ($requestedUserId -in $global:EntraMockState.FailedUserIds) { throw "User lookup denied for $requestedUserId" }
        if ($Uri -like '*onPremisesExtensionAttributes*') {
            if ($requestedUserId -in $global:EntraMockState['FailedAttributeUserIds']) { throw "Attribute lookup denied for $requestedUserId" }
            if ($global:EntraMockState.LastUserAttributes.ContainsKey($requestedUserId)) {
                return [pscustomobject]@{ onPremisesExtensionAttributes=$global:EntraMockState.LastUserAttributes[$requestedUserId] }
            }
            return [pscustomobject]@{}
        }
        if ($Uri -notlike '*jobTitle*' -and $global:EntraMockState.ContainsKey('LastUserProfiles') -and $global:EntraMockState.LastUserProfiles.ContainsKey($requestedUserId)) {
            return $global:EntraMockState.LastUserProfiles[$requestedUserId]
        }
    }
    if ($Uri -like '*jobTitle*') { return [pscustomobject]@{} }
    return [pscustomobject]@{ id='user-1'; displayName='Test User'; userPrincipalName='user@example.test' }
}
$global:EntraMockState = @{
    Writes=0
    NoDevices=$false
    GroupNameMatches=@([pscustomobject]@{ id='group-1'; displayName='Device Group' })
    GroupMembers=@(
        [pscustomobject]@{ '@odata.type'='#microsoft.graph.device'; id='entra-object-1'; deviceId='aad-device-1' },
        [pscustomobject]@{ '@odata.type'='#microsoft.graph.device'; id='entra-object-1-duplicate'; deviceId='AAD-DEVICE-1' },
        [pscustomobject]@{ '@odata.type'='#microsoft.graph.device'; id='entra-object-2'; deviceId='aad-device-not-managed' },
        [pscustomobject]@{ '@odata.type'='#microsoft.graph.user'; id='user-1'; deviceId='aad-device-1' },
        [pscustomobject]@{ '@odata.type'='#microsoft.graph.group'; id='nested-group'; deviceId='aad-device-1' }
    )
}
$tempDirectory = Join-Path ([IO.Path]::GetTempPath()) ('EntraToolboxTests-' + [guid]::NewGuid())
[void](New-Item -Path $tempDirectory -ItemType Directory)
try {
    $memberRows = @(& $memberPage.Path -TargetGroupName 'Target' -MemberNames 'Member' -MemberType Groups -SkipGraphConnect -PassThru -WhatIf)
    Assert-EntraTest ($memberRows.Count -eq 1 -and $memberRows[0].Status -eq 'WhatIf' -and $global:EntraMockState.Writes -eq 0) 'Members WhatIf returns objects and never writes'
    $attributeRows = @(& $attributePage.Path -UserPrincipalName 'user@example.test' -AttributeNumber 5 -Value 'Project' -SkipGraphConnect -NonInteractive -PassThru -WhatIf)
    Assert-EntraTest ($attributeRows.Count -eq 1 -and $attributeRows[0].Status -eq 'WhatIf' -and $global:EntraMockState.Writes -eq 0) 'Attributes WhatIf returns objects and never writes'
    $attributeRows = @(& $attributePage.Path -UserPrincipalName 'user@example.test' -AttributeNumber 5 -Clear -SkipGraphConnect -NonInteractive -PassThru -Confirm:$false)
    Assert-EntraTest ($attributeRows[0].Status -eq 'OK' -and $global:EntraMockState.Writes -eq 1) 'Attribute clear invokes only mocked PATCH'
    $primaryCsv = Join-Path $tempDirectory 'PrimaryPC.csv'
    $deviceRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
    $importRows = @(Import-Csv -LiteralPath $primaryCsv)
    Assert-EntraTest ($deviceRows.Count -eq 1 -and $deviceRows[0].DeviceObjectId -eq 'device-1' -and $importRows[0].PSObject.Properties.Name -contains 'memberObjectIdOrUpn [memberObjectIdOrUpn] Required') 'Primary-PC objects and unchanged Entra import header'
    $global:EntraMockState.Devices = @{}
    foreach ($number in 1..5) {
        $deviceId = "device-$number"
        $global:EntraMockState.Devices[$deviceId] = [pscustomobject]@{
            Id=$deviceId; DisplayName="TEST-PC-$number"; OperatingSystem='Windows'; IsManaged=$true
            MdmAppId='0000000a-0000-0000-c000-000000000000'; ApproximateLastSignInDateTime=([datetime]'2026-10-01').AddDays($number); TrustType='AzureAd'
        }
    }
    $global:EntraMockState.Devices['device-3'].IsManaged = $false
    $global:EntraMockState.Devices['device-3'].MdmAppId = $null
    $global:EntraMockState.Devices['device-4'].MdmAppId = 'other-mdm'
    $global:EntraMockState.Devices['device-4'].ApproximateLastSignInDateTime = $null
    $global:EntraMockState.Devices['device-5'].OperatingSystem = 'Linux'
    $deviceRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
    Assert-EntraTest ($deviceRows.Count -eq 1 -and $deviceRows[0].DeviceObjectId -eq 'device-2') 'Latest preserves management filter and selects most recent eligible Windows device'
    $global:EntraMockState.SignInCalls = 0
    $deviceRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -DeviceMode All -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
    $importRows = @(Import-Csv -LiteralPath $primaryCsv)
    Assert-EntraTest ($deviceRows.Count -eq 4 -and ($deviceRows.DeviceObjectId -join ',') -eq 'device-3,device-2,device-1,device-4') 'All combines owned and registered Windows devices without duplicates, includes unmanaged devices, and sorts by activity'
    Assert-EntraTest ($importRows.Count -eq 4 -and $importRows[0].PSObject.Properties.Name -contains 'memberObjectIdOrUpn [memberObjectIdOrUpn] Required') 'All exports unique device IDs with unchanged Entra import header'
    Assert-EntraTest ($global:EntraMockState.SignInCalls -eq 0) 'All does not use sign-in fallback when associated Windows devices exist'
    & {
        $signInState = @{ Filter=''; All=$false; Mode='Events'; IntuneEmpty=$false; Scopes=@('User.Read.All','GroupMember.Read.All','Device.Read.All','AuditLog.Read.All','DeviceManagementManagedDevices.Read.All') }
        $signInEvents = [Collections.Generic.List[object]]::new()
        $eventTime = [datetimeoffset]::UtcNow.AddHours(-1)
        foreach ($spec in @(@(1,2),@(2,2),@(3,1),@(4,5),@(5,4),@(6,3),@(7,3),@(8,3))) {
            $deviceGuid = 'aaaaaaaa-0000-0000-0000-{0:D12}' -f $spec[0]
            foreach ($eventNumber in 1..$spec[1]) {
                $signInEvents.Add([pscustomobject]@{
                    DeviceDetail=[pscustomobject]@{ DeviceId=$(if ($eventNumber -eq 2) { $deviceGuid.ToUpperInvariant() } else { $deviceGuid }) }
                    Status=[pscustomobject]@{ ErrorCode=0 }; CreatedDateTime=$eventTime.AddSeconds($spec[0] * 10 + $eventNumber)
                })
            }
        }
        foreach ($number in 1..6) {
            $signInEvents.Add([pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId='aaaaaaaa-0000-0000-0000-000000000003' }; Status=[pscustomobject]@{ ErrorCode=50074 }; CreatedDateTime=$eventTime.AddMinutes(1) })
        }
        $signInEvents.Add([pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId=' ' }; Status=[pscustomobject]@{ ErrorCode=0 }; CreatedDateTime=$eventTime })
        $signInEvents.Add([pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId='aaaaaaaa-0000-0000-0000-000000000003' }; Status=$null; CreatedDateTime=$eventTime })
        function Get-MgContext { [pscustomobject]@{ Scopes=$signInState.Scopes } }
        function Get-MgUserOwnedDevice { throw 'Sign-in-only mode queried owned devices' }
        function Get-MgUserRegisteredDevice { throw 'Sign-in-only mode queried registered devices' }
        function Get-MgAuditLogSignIn {
            [CmdletBinding()]param($Filter,[switch]$All,[int]$Top)
            $roleError = '[Authentication_RequestFromUnsupportedRole] : User is not in the allowed roles'
            if ($Top) {
                $signInState.Probes++
                if ($signInState.Mode -eq 'NoRole') { throw $roleError }
                return
            }
            $signInState.Queries++
            $signInState.Filter = $Filter; $signInState.All = [bool]$All
            if ($signInState.Mode -in @('NoRole','NoRoleAfterProbe')) { throw $roleError }
            if ($signInState.Mode -eq 'QueryError') { throw 'Audit access denied' }
            if ($signInState.Mode -eq 'Empty') { return }
            return $signInEvents.ToArray()
        }
        function Get-MgDevice {
            [CmdletBinding()]param($DeviceId,$Property,$Filter)
            if ($DeviceId) { throw 'Sign-in-only mode queried an associated device' }
            if ($Filter -notmatch "deviceId eq '(aaaaaaaa-0000-0000-0000-(\d{12}))'") { throw "Unexpected device filter: $Filter" }
            $deviceGuid = $Matches[1]; $number = [int]$Matches[2]
            if ($number -eq 5) { return }
            if ($number -eq 6) { throw 'Device resolution denied' }
            $device = [pscustomobject]@{
                Id=$(if ($number -eq 8) { '' } else { "object-$number" }); DeviceId=$deviceGuid; DisplayName="LOG-PC-$number"
                OperatingSystem=$(if ($number -eq 4) { 'Linux' } else { 'Windows' }); ApproximateLastSignInDateTime=[datetime]'2030-01-01'
            }
            if ($number -eq 7) { return @($device,$device) }
            return $device
        }
        function Invoke-MgGraphRequest {
            [CmdletBinding()]param($Method,$Uri,$OutputType)
            if ($signInState.IntuneEmpty) { return [pscustomobject]@{ value=@(); '@odata.nextLink'=$null } }
            [pscustomobject]@{ value=@([pscustomobject]@{ azureADDeviceId='aaaaaaaa-0000-0000-0000-000000000002' },[pscustomobject]@{ azureADDeviceId='aaaaaaaa-0000-0000-0000-000000000003' }); '@odata.nextLink'=$null }
        }
        $logParameters = @{ UserPrincipalName='user@example.test'; DeviceSource='SignInLogs'; OutputPath=$primaryCsv; SkipGraphConnect=$true; PassThru=$true; WarningAction='SilentlyContinue' }
        $logWarnings = @()
        $mostRows = @(& $primaryPage.Path @logParameters -SignInDays 14 -WarningVariable logWarnings)
        Assert-EntraTest (($mostRows.DeviceObjectId -join ',') -eq 'object-2,object-1' -and @($mostRows | Where-Object Source -ne 'SignInLogs (2 events)').Count -eq 0) 'Sign-in-only Most returns tied eligible Windows PCs, ignoring failed events and case-varied duplicate IDs'
        Assert-EntraTest ($mostRows[0].LastActivity -eq $eventTime.AddSeconds(22).UtcDateTime -and $mostRows[1].LastActivity -eq $eventTime.AddSeconds(12).UtcDateTime) 'Sign-in-only activity uses each users newest successful log event, including DateTimeOffset timestamps'
        Assert-EntraTest ($signInState.All -and $signInState.Filter -like "*userId eq 'user-1' and appDisplayName eq 'Windows Sign In' and status/errorCode eq 0 and createdDateTime ge *") 'Sign-in-only query filters successful Windows Sign In events per user and requests all pages'
        Assert-EntraTest (($logWarnings -join ' ') -like '*resolved to 0 Entra devices*' -and ($logWarnings -join ' ') -like '*Device resolution denied*' -and ($logWarnings -join ' ') -like '*resolved to 2 Entra devices*' -and ($logWarnings -join ' ') -like '*has no Entra object ID*') 'Unresolvable, ambiguous and non-exportable devices produce explicit warnings'
        $allLogRows = @(& $primaryPage.Path @logParameters -DeviceMode All)
        Assert-EntraTest (@($logWarnings | Where-Object { "$_" -like '*obsolete*' }).Count -eq 0 -and (Get-Content -LiteralPath $primaryCsv -TotalCount 1) -notlike '#TYPE*') 'Primary PC CSV export emits no obsolete-parameter warning or type header'
        Assert-EntraTest (($allLogRows.DeviceObjectId -join ',') -eq 'object-2,object-1,object-3' -and ($allLogRows[0].PSObject.Properties.Name -join ',') -eq 'UserPrincipalName,DeviceName,LastActivity,Source,DeviceObjectId') 'Sign-in-only All returns distinct eligible PCs ordered by count then newest event with unchanged result columns'
        $intuneLogRows = @(& $primaryPage.Path @logParameters -DeviceMode Most -OnlyIntuneDevices)
        Assert-EntraTest (($intuneLogRows.DeviceObjectId -join ',') -eq 'object-2') 'Sign-in-only Most selects winners after applying optional Intune inventory filtering'
        $signInState.IntuneEmpty = $true
        $intuneLogRows = @(& $primaryPage.Path @logParameters -DeviceMode All -OnlyIntuneDevices)
        Assert-EntraTest ($intuneLogRows.Count -eq 0) 'Sign-in-only All cannot bypass an empty Intune inventory'
        $signInState.IntuneEmpty = $false
        $userCsv = Join-Path $tempDirectory 'SignInUsers.csv'
        @([pscustomobject]@{ UserPrincipalName='user@example.test' },[pscustomobject]@{ UserPrincipalName='other@example.test' }) | Export-Csv -LiteralPath $userCsv -NoTypeInformation
        $csvParameters = $logParameters.Clone(); $csvParameters.Remove('UserPrincipalName'); $csvParameters.CsvPath=$userCsv
        $csvLogRows = @(& $primaryPage.Path @csvParameters -DeviceMode All)
        $logImport = @(Import-Csv -LiteralPath $primaryCsv)
        Assert-EntraTest ($csvLogRows.Count -eq 6 -and @($csvLogRows.UserPrincipalName | Sort-Object -Unique).Count -eq 2 -and $logImport.Count -eq 3 -and ($logImport.'memberObjectIdOrUpn [memberObjectIdOrUpn] Required' -join ',') -eq 'object-2,object-1,object-3') 'User CSV remains user input and shared log devices export as unique object IDs with the original import header'
        $savedCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            $culture = [Globalization.CultureInfo]::new('da-DK'); $culture.DateTimeFormat.TimeSeparator='.'
            [Threading.Thread]::CurrentThread.CurrentCulture = $culture
            $null = @(& $primaryPage.Path @logParameters -SignInDays 14)
            Assert-EntraTest ($signInState.Filter -match 'createdDateTime ge (\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)$') 'Sign-in-only lookback uses culture-invariant Graph UTC timestamps'
            $since = [datetime]::ParseExact($Matches[1],"yyyy-MM-dd'T'HH:mm:ss'Z'",[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
            Assert-EntraTest ([Math]::Abs(($since - [datetime]::UtcNow.AddDays(-14)).TotalSeconds) -lt 10) 'Sign-in-only query honors the selected lookback rather than the default'
        } finally { [Threading.Thread]::CurrentThread.CurrentCulture = $savedCulture }
        foreach ($mode in @('QueryError','Empty')) {
            $signInState.Mode = $mode; $logWarnings = @()
            $emptyLogRows = @(& $primaryPage.Path @logParameters -WarningVariable logWarnings)
            $expectedWarning = if ($mode -eq 'QueryError') { '*query failed*Audit access denied*' } else { '*no successful Windows Sign In events with a device ID*' }
            Assert-EntraTest ($emptyLogRows.Count -eq 0 -and ($logWarnings -join ' ') -like $expectedWarning) "Sign-in-only $mode warns without prompting or querying owned/registered devices"
        }
        foreach ($mode in @('NoRole','NoRoleAfterProbe')) {
            $signInState.Mode = $mode; $signInState.Probes = 0; $signInState.Queries = 0; $logWarnings = @()
            $deniedRows = @(& $primaryPage.Path @csvParameters -WarningVariable logWarnings)
            $expectedQueries = if ($mode -eq 'NoRole') { 0 } else { 1 }
            Assert-EntraTest ($deniedRows.Count -eq 0 -and @($logWarnings).Count -eq 1 -and "$($logWarnings[0])" -like '*no active Entra role*Reports Reader*' -and $signInState.Probes -eq 1 -and $signInState.Queries -eq $expectedQueries) "Missing sign-in role ($mode) warns once and skips remaining users' lookups"
        }
        $signInState.Mode = 'Events'
        foreach ($invalidOptions in @(@{ DeviceMode='Latest' },@{ SignInFallback=$true })) {
            $invalidRejected = $false
            try { & $primaryPage.Path @logParameters @invalidOptions | Out-Null } catch { $invalidRejected = $_.Exception.Message -match 'not supported|only supported' }
            Assert-EntraTest $invalidRejected 'Sign-in-only rejects incompatible legacy selection/fallback options'
        }
        $signInState.Scopes = @('User.Read.All','GroupMember.Read.All','Device.Read.All')
        $scopeRejected = $false
        try { & $primaryPage.Path @logParameters | Out-Null } catch { $scopeRejected = $_.Exception.Message -like '*existing Graph session*AuditLog.Read.All*' }
        Assert-EntraTest $scopeRejected 'Standalone sign-in-only discovery requires audit permission even without fallback'
    }
    & {
        $intuneTestState = @{ Calls=0; Empty=$false; Fail=$false }
        foreach ($number in 1..5) {
            $global:EntraMockState.Devices["device-$number"] | Add-Member -NotePropertyName DeviceId -NotePropertyValue ("AAAAAAAA-0000-0000-0000-{0:D12}" -f $number) -Force
        }
        function Get-MgContext { [pscustomobject]@{ Scopes=@('User.Read.All','GroupMember.Read.All','Device.Read.All','DeviceManagementManagedDevices.Read.All','AuditLog.Read.All') } }
        function Invoke-MgGraphRequest {
            [CmdletBinding()]param($Method,$Uri,$OutputType)
            $intuneTestState.Calls++
            if ($intuneTestState.Fail) { throw 'Intune inventory unavailable' }
            if ($intuneTestState.Empty) { return [pscustomobject]@{ value=@(); '@odata.nextLink'=$null } }
            if ($Uri -like '*skiptoken*') { return [pscustomobject]@{ value=@([pscustomobject]@{ azureADDeviceId='aaaaaaaa-0000-0000-0000-000000000003' }); '@odata.nextLink'=$null } }
            [pscustomobject]@{ value=@([pscustomobject]@{ azureADDeviceId='aaaaaaaa-0000-0000-0000-000000000001' },[pscustomobject]@{ azureADDeviceId=[guid]::Empty.ToString() }); '@odata.nextLink'='https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$skiptoken=next' }
        }
        function Get-MgAuditLogSignIn {
            [CmdletBinding()]param($Filter,[switch]$All)
            [pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId='aaaaaaaa-0000-0000-0000-000000000001' }; CreatedDateTime=[datetime]'2026-10-20' }
        }
        function Get-MgDevice {
            [CmdletBinding()]param($DeviceId,$Property,$Filter)
            if ($DeviceId) { return $global:EntraMockState.Devices[$DeviceId] }
            return $global:EntraMockState.Devices['device-1']
        }
        $intuneRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -OnlyIntuneDevices -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
        Assert-EntraTest ($intuneRows.Count -eq 1 -and $intuneRows[0].DeviceObjectId -eq 'device-1' -and $intuneTestState.Calls -eq 2) 'Latest excludes stale managed devices and reads all Intune inventory pages'
        $intuneRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -DeviceMode All -OnlyIntuneDevices -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
        Assert-EntraTest (($intuneRows.DeviceObjectId -join ',') -eq 'device-3,device-1') 'All filters by actual Intune presence using case-insensitive Entra device IDs'
        $global:EntraMockState.NoDevices = $true
        foreach ($deviceMode in @('Latest','All')) {
            $intuneRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -DeviceMode $deviceMode -OnlyIntuneDevices -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
            Assert-EntraTest ($intuneRows.Count -eq 1 -and $intuneRows[0].Source -like 'SignInLogs*') "$deviceMode fallback accepts devices confirmed in Intune"
        }
        $intuneTestState.Empty = $true
        $intuneRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -DeviceMode All -OnlyIntuneDevices -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru 3>$null)
        Assert-EntraTest ($intuneRows.Count -eq 0) 'Sign-in fallback cannot bypass an empty Intune inventory'
        $intuneTestState.Fail = $true
        $inventoryRejected = $false
        try { & $primaryPage.Path -UserPrincipalName 'user@example.test' -OnlyIntuneDevices -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru | Out-Null } catch { $inventoryRejected = $_.Exception.Message -eq 'Intune inventory unavailable' }
        Assert-EntraTest $inventoryRejected 'Intune inventory failure stops the report instead of returning unverified devices'
        $global:EntraMockState.NoDevices = $false
    }
    $global:EntraMockState.Devices = $null
    $global:EntraMockState.NoDevices = $true
    $deviceRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
    Assert-EntraTest ($deviceRows.Count -eq 0) 'Missing primary device does not prompt or force fallback'
    $deviceRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -DeviceMode All -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
    Assert-EntraTest ($deviceRows.Count -eq 0 -and $global:EntraMockState.SignInCalls -eq 0) 'All with no devices does not prompt or force fallback'
    foreach ($deviceMode in @('Latest','All')) {
        $deviceRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -DeviceMode $deviceMode -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru)
        Assert-EntraTest ($deviceRows.Count -eq 1 -and $deviceRows[0].DeviceObjectId -eq 'device-1' -and $deviceRows[0].Source -like 'SignInLogs*') "$deviceMode uses optional sign-in fallback when no devices match"
    }
    & {
        $fallbackState = @{ Mode='Candidates'; Lookups=[Collections.Generic.List[string]]::new(); SignInFilter=''; ValidateTimestamp=$false }
        function Get-MgAuditLogSignIn {
            [CmdletBinding()]param($Filter,[switch]$All)
            $fallbackState.SignInFilter = $Filter
            if ($fallbackState.ValidateTimestamp -and $Filter -notmatch 'createdDateTime ge \d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$') { throw 'Invalid Graph timestamp' }
            if ($fallbackState.Mode -eq 'QueryError') { throw 'Sign-in access denied' }
            if ($fallbackState.Mode -eq 'NoDeviceId') { return [pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId='' }; CreatedDateTime=[datetime]'2026-11-03' } }
            foreach ($candidateSpec in @(@('missing-device',4),@('linux-device',3),@('lookup-error',2),@('valid-device',1))) {
                foreach ($eventNumber in 1..$candidateSpec[1]) {
                    [pscustomobject]@{ DeviceDetail=[pscustomobject]@{ DeviceId=$candidateSpec[0] }; CreatedDateTime=([datetime]'2026-11-03').AddMinutes($eventNumber) }
                }
            }
        }
        function Get-MgDevice {
            [CmdletBinding()]param($DeviceId,$Property,$Filter)
            [void]$fallbackState.Lookups.Add($Filter)
            if ($Filter -like '*missing-device*') { return }
            if ($Filter -like '*lookup-error*') { throw 'Device lookup denied' }
            [pscustomobject]@{ Id='fallback-object'; DeviceId='valid-device'; DisplayName='FALLBACK-PC'; OperatingSystem=$(if ($Filter -like '*linux-device*') { 'Linux' } else { 'Windows' }); ApproximateLastSignInDateTime=[datetime]'2026-11-03' }
        }
        $fallbackWarnings = @()
        $fallbackOutput = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -SignInFallback -Verbose -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru -WarningVariable fallbackWarnings -WarningAction SilentlyContinue 4>&1)
        $fallbackRows = @($fallbackOutput | Where-Object { $_ -isnot [Management.Automation.VerboseRecord] })
        $fallbackDetails = ($fallbackOutput | Where-Object { $_ -is [Management.Automation.VerboseRecord] }) -join "`n"
        Assert-EntraTest ($fallbackRows.Count -eq 1 -and $fallbackRows[0].DeviceObjectId -eq 'fallback-object' -and $fallbackState.Lookups.Count -eq 4) 'Fallback continues past missing, non-Windows and lookup-error candidates to a valid device'
        Assert-EntraTest ($fallbackDetails -like '*10 event(s); 10 with a device ID*' -and $fallbackDetails -like '*resolved to 0 Entra devices*' -and $fallbackDetails -like "*Entra OS 'Linux' is not Windows*") 'Verbose fallback diagnostics report event counts and candidate rejection reasons'
        Assert-EntraTest (($fallbackWarnings -join ' ') -like '*device lookup failed*Device lookup denied*') 'Fallback device lookup errors are visible without verbose logging'
        $savedCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            $danishCulture = [Globalization.CultureInfo]::new('da-DK')
            $danishCulture.DateTimeFormat.TimeSeparator = '.'
            [Threading.Thread]::CurrentThread.CurrentCulture = $danishCulture
            $fallbackState.ValidateTimestamp = $true
            $cultureRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru -WarningAction SilentlyContinue)
            Assert-EntraTest ($fallbackState.SignInFilter -match 'createdDateTime ge (\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)$') 'Danish culture preserves Graph UTC timestamp colons in the actual sign-in filter'
            $filterTimestamp = [datetime]::ParseExact($Matches[1],"yyyy-MM-dd'T'HH:mm:ss'Z'",[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
            Assert-EntraTest ($filterTimestamp.Kind -eq [DateTimeKind]::Utc -and $cultureRows.Count -eq 1) 'Danish-locale fallback submits a valid UTC filter and returns an eligible device'
        } finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $savedCulture
            $fallbackState.ValidateTimestamp = $false
        }
        $fallbackState.Mode = 'QueryError'
        $fallbackWarnings = @()
        $fallbackRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru -WarningVariable fallbackWarnings -WarningAction SilentlyContinue)
        Assert-EntraTest ($fallbackRows.Count -eq 0 -and ($fallbackWarnings -join ' ') -like '*fallback query failed*Sign-in access denied*') 'Fallback query failure is reported rather than silently treated as no events'
        $fallbackState.Mode = 'NoDeviceId'
        $fallbackWarnings = @()
        $fallbackRows = @(& $primaryPage.Path -UserPrincipalName 'user@example.test' -SignInFallback -OutputPath $primaryCsv -SkipGraphConnect -NonInteractive -PassThru -WarningVariable fallbackWarnings -WarningAction SilentlyContinue)
        Assert-EntraTest ($fallbackRows.Count -eq 0 -and ($fallbackWarnings -join ' ') -like '*no Windows Sign In events with a device ID*') 'Fallback explains when retrieved events have no device ID'
    }
    $lastCsv = Join-Path $tempDirectory 'LastUser.csv'
    $lastRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -OutputPath $lastCsv -SkipGraphConnect -PassThru)
    $bytes = [IO.File]::ReadAllBytes($lastCsv)
    Assert-EntraTest ($lastRows.Count -eq 1 -and $lastRows[0].ResolutionStatus -eq 'PrimaryUserResolved') 'Last-user typed result and absent extended properties'
    Assert-EntraTest ($bytes.Length -gt 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) 'Last-user CSV preserves UTF8 BOM'
    $baseLastUserColumns = 'DeviceId,DeviceName,Manufacturer,Model,ResolvedUserId,ResolvedUserSource,ResolutionStatus,LookupPhase,LookupError,UserDisplayName,UserPrincipalName,JobTitle,CompanyName,OfficeLocation,UserSource,LastLoggedOnUserId,PrimaryUserId'
    Assert-EntraTest (($lastRows[0].PSObject.Properties.Name -join ',') -eq $baseLastUserColumns) 'No selection preserves the exact existing last-user schema'
    try {
        $global:EntraMockState.LastUserRequests = [Collections.Generic.List[string]]::new()
        $global:EntraMockState.LastUserAttributes = @{
            'user-1'=[pscustomobject]@{ extensionAttribute1='Primary'; extensionAttribute5=''; extensionAttribute15='Fifteen' }
            'user-2'=[pscustomobject]@{ extensionAttribute1='Second'; extensionAttribute15='Other' }
            'user-3'=$null
        }
        $defaultRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        Assert-EntraTest (@($global:EntraMockState.LastUserRequests | Where-Object { $_ -like '*onPremisesExtensionAttributes*' }).Count -eq 0) 'No selection does not request extension attributes'
        $selectedRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -ExtensionAttributeNumbers 15,1,5,1 -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        $selectedColumns = "$baseLastUserColumns,extensionAttribute1,extensionAttribute5,extensionAttribute15"
        $selectedCsv = @(Import-Csv -LiteralPath $lastCsv)
        Assert-EntraTest (($selectedRows[0].PSObject.Properties.Name -join ',') -eq $selectedColumns -and $selectedRows[0].extensionAttribute1 -eq 'Primary' -and -not $selectedRows[0].extensionAttribute5 -and $selectedRows[0].extensionAttribute15 -eq 'Fifteen') 'Primary-user fallback appends selected attributes once in numeric order'
        Assert-EntraTest (($selectedCsv[0].PSObject.Properties.Name -join ',') -eq $selectedColumns -and $selectedCsv[0].extensionAttribute15 -eq 'Fifteen') 'Automatic CSV contains the selected attribute headers and values'
        $lastPage.Navigation.IsChecked = $true
        $lastPage.Results.Clear(); $lastPage.Results.Add($selectedRows[0]); Update-EntraResults $lastPage
        $ui.Window.UpdateLayout()
        Assert-EntraTest (($lastPage.Table.Columns.Header -join ',') -eq $selectedColumns) 'Output grid generates the selected attribute columns'
        $attributeExport = Join-Path $tempDirectory 'SelectedAttributes.csv'
        Export-EntraResults $lastPage $attributeExport
        $exportedAttributes = @(Import-Csv -LiteralPath $attributeExport)
        Assert-EntraTest (($exportedAttributes[0].PSObject.Properties.Name -join ',') -eq $selectedColumns -and $exportedAttributes[0].extensionAttribute1 -eq 'Primary') 'Export results retains the selected attribute schema and values'
        $changedRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -ExtensionAttributeNumbers 5 -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        $lastPage.Results.Clear(); $lastPage.Results.Add($changedRows[0]); Update-EntraResults $lastPage
        $ui.Window.UpdateLayout()
        Assert-EntraTest (($lastPage.Table.Columns.Header -join ',') -eq "$baseLastUserColumns,extensionAttribute5" -and ((Import-Csv -LiteralPath $lastCsv).PSObject.Properties.Name -join ',') -eq "$baseLastUserColumns,extensionAttribute5") 'Rerunning with changed selection removes stale grid and CSV columns'
        $global:EntraMockState.LastUserDevice = [pscustomobject]@{
            id='managed-1'; deviceName='TEST-PC'; manufacturer='Test'; model='Model'; azureADDeviceId='aad-device-1'
            userId='user-1'; userDisplayName='Test User'; userPrincipalName='user@example.test'; emailAddress='user@example.test'
            usersLoggedOn=@(
                [pscustomobject]@{ userId='user-2'; lastLogOnDateTime='2026-10-02T10:00:00Z' },
                [pscustomobject]@{ userId='user-2'; lastLogOnDateTime='2026-10-01T10:00:00Z' },
                [pscustomobject]@{ userId='user-3'; lastLogOnDateTime='2026-09-01T10:00:00Z' },
                [pscustomobject]@{ userId='user-4'; lastLogOnDateTime='2026-08-01T10:00:00Z' }
            )
        }
        foreach ($mode in @('Latest','Most','All')) {
            $modeRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode $mode -ExtensionAttributeNumbers 1,5,15 -OutputPath $lastCsv -SkipGraphConnect -PassThru)
            Assert-EntraTest ($modeRows[0].ResolvedUserId -eq 'user-2' -and $modeRows[0].extensionAttribute1 -eq 'Second' -and -not $modeRows[0].extensionAttribute5) "$mode retrieves attributes from the selected logged-on user, not the primary user"
            if ($mode -eq 'All') {
                Assert-EntraTest ($modeRows.Count -eq 3 -and -not $modeRows[1].extensionAttribute1 -and -not $modeRows[2].extensionAttribute1 -and @($modeRows | Where-Object { ($_.PSObject.Properties.Name -join ',') -ne $selectedColumns }).Count -eq 0) 'Null and omitted attribute objects retain blank cells and a uniform row schema'
            }
        }
        $global:EntraMockState.RepeatLastUserDevice = $true
        $global:EntraMockState.LastUserRequests.Clear()
        $cachedRows = @(& $lastPage.Path -ExtensionAttributeNumbers 1 -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        Assert-EntraTest ($cachedRows.Count -eq 2 -and ($cachedRows.extensionAttribute1 -join ',') -eq 'Second,Second' -and @($global:EntraMockState.LastUserRequests | Where-Object { $_ -like '*onPremisesExtensionAttributes*' }).Count -eq 1) 'Repeated users reuse a single extended-profile request across devices'
        $global:EntraMockState.FailedAttributeUserIds = @('user-2')
        $attributeWarnings = @()
        $failedAttributeRows = @(& $lastPage.Path -ExtensionAttributeNumbers 1 -OutputPath $lastCsv -SkipGraphConnect -PassThru -WarningVariable attributeWarnings -WarningAction SilentlyContinue)
        Assert-EntraTest ($failedAttributeRows.Count -eq 2 -and @($failedAttributeRows | Where-Object { $_.LookupPhase -ne 'ExtendedLookupFailed' -or $_.LookupError -notlike '*Attribute lookup denied*' -or $_.extensionAttribute1 -or -not $_.UserPrincipalName }).Count -eq 0 -and $attributeWarnings.Count -eq 2) 'Attribute lookup failures preserve identity, blank cells and explicit diagnostics even for cached users'
        $global:EntraMockState.FailedAttributeUserIds = @()
        $global:EntraMockState.FailedUserIds = @('user-2')
        $failedCoreRows = @(& $lastPage.Path -ExtensionAttributeNumbers 1 -OutputPath $lastCsv -SkipGraphConnect -PassThru -WarningAction SilentlyContinue)
        Assert-EntraTest ($failedCoreRows.Count -eq 2 -and $failedCoreRows[0].ResolutionStatus -eq 'UsersLoggedOnLookupFailed' -and -not $failedCoreRows[0].extensionAttribute1) 'Core lookup failure retains selected attribute columns without leaking prior values'
        $global:EntraMockState.LastUserDevice.usersLoggedOn = @()
        $global:EntraMockState.LastUserDevice.userId = $null
        $missingRows = @(& $lastPage.Path -ExtensionAttributeNumbers 1 -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        Assert-EntraTest ($missingRows[0].ResolutionStatus -eq 'PrimaryUserMissing' -and $missingRows[0].PSObject.Properties['extensionAttribute1'] -and -not $missingRows[0].extensionAttribute1) 'Missing user retains the selected blank attribute column'
        foreach ($invalidNumber in @(0,16)) {
            $invalidRejected = $false
            try { & $lastPage.Path -DeviceName 'TEST-PC' -ExtensionAttributeNumbers 1,$invalidNumber -OutputPath $lastCsv -SkipGraphConnect -PassThru }
            catch { $invalidRejected = $_.FullyQualifiedErrorId -like '*ParameterArgumentValidationError*' }
            Assert-EntraTest $invalidRejected "Attribute selection rejects out-of-range entry $invalidNumber"
        }
    } finally {
        foreach ($key in @('LastUserRequests','LastUserAttributes','LastUserDevice','RepeatLastUserDevice','FailedAttributeUserIds','FailedUserIds')) { $global:EntraMockState.Remove($key) }
        $lastPage.Results.Clear(); Update-EntraResults $lastPage
        $primaryPage.Navigation.IsChecked = $true
    }
    try {
        $global:EntraMockState.LastUserDevice = [pscustomobject]@{
            id='managed-1'; deviceName='TEST-PC'; manufacturer='Test'; model='Model'; azureADDeviceId='aad-device-1'
            userId='user-1'; userDisplayName='Test User'; userPrincipalName='user@example.test'; emailAddress='user@example.test'
            usersLoggedOn=@(
                [pscustomobject]@{ userId='user-2'; lastLogOnDateTime='2026-10-01T10:00:00Z' },
                [pscustomobject]@{ userId='user-3'; lastLogOnDateTime='2026-10-10T09:00:00Z' },
                [pscustomobject]@{ userId='user-2'; lastLogOnDateTime='2026-10-02T10:00:00Z' },
                [pscustomobject]@{ userId='user-4'; lastLogOnDateTime='2026-10-10T10:00:00+02:00' },
                [pscustomobject]@{ userId='user-2'; lastLogOnDateTime='2026-10-02T10:00:00Z' },
                [pscustomobject]@{ userId=' '; lastLogOnDateTime='2026-10-11T10:00:00Z' }
            )
        }
        $global:EntraMockState.LastUserProfiles = @{
            'user-2'=[pscustomobject]@{ id='user-2'; displayName='User Two'; userPrincipalName='two@example.test' }
            'user-3'=[pscustomobject]@{ id='user-3'; displayName='User Three'; userPrincipalName='three@example.test' }
            'user-4'=[pscustomobject]@{ id='user-4'; displayName=$null; userPrincipalName=$null }
        }
        $latestRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        Assert-EntraTest ($latestRows.Count -eq 1 -and $latestRows[0].ResolvedUserId -eq 'user-3' -and $latestRows[0].UserPrincipalName -eq 'three@example.test') 'Default latest user uses newest timestamp across offsets, not array order or blank IDs'
        $mostRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode Most -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        Assert-EntraTest ($mostRows.Count -eq 1 -and $mostRows[0].ResolvedUserId -eq 'user-2') 'Most counts repeated userIds including identical timestamps rather than selecting latest'
        $allRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode All -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        $allCsvRows = @(Import-Csv -LiteralPath $lastCsv)
        Assert-EntraTest ($allRows.Count -eq 3 -and ($allRows.ResolvedUserId -join ',') -eq 'user-3,user-4,user-2') 'All returns unique users in newest-logon order'
        Assert-EntraTest ($allCsvRows.Count -eq 3 -and ($allCsvRows.ResolvedUserId -join ',') -eq ($allRows.ResolvedUserId -join ',') -and ($allCsvRows[0].PSObject.Properties.Name -join ',') -eq ($lastRows[0].PSObject.Properties.Name -join ',')) 'All CSV retains columns and exports one row per selected device/user'
        Assert-EntraTest (@($allRows | Where-Object { $_.DeviceId -ne 'managed-1' -or $_.LastLoggedOnUserId -ne $_.ResolvedUserId -or $_.PrimaryUserId -ne 'user-1' -or $_.ResolutionStatus -ne 'UsersLoggedOnResolved' }).Count -eq 0) 'Each selected user row retains correct device, selected-user IDs and source status'
        Assert-EntraTest (-not $allRows[1].UserDisplayName -and -not $allRows[1].UserPrincipalName) 'Missing selected-user profile fields do not inherit primary-user details'
        $global:EntraMockState.LastUserDevice.usersLoggedOn += @(
            [pscustomobject]@{ userId='user-4'; lastLogOnDateTime='2026-10-03T10:00:00Z' },
            [pscustomobject]@{ userId='user-4'; lastLogOnDateTime='2026-10-04T10:00:00Z' }
        )
        $tiedRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode Most -OutputPath $lastCsv -SkipGraphConnect -PassThru)
        $tiedCsvRows = @(Import-Csv -LiteralPath $lastCsv)
        Assert-EntraTest ($tiedRows.Count -eq 2 -and ($tiedRows.ResolvedUserId -join ',') -eq 'user-4,user-2' -and ($tiedCsvRows.ResolvedUserId -join ',') -eq 'user-4,user-2') 'Most returns and exports every tied winner in newest-logon order'
        $global:EntraMockState.FailedUserIds = @('user-2')
        $lookupWarnings = @()
        $failedRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode All -OutputPath $lastCsv -SkipGraphConnect -PassThru -WarningVariable lookupWarnings -WarningAction SilentlyContinue)
        $failedUser = $failedRows | Where-Object { $_.ResolvedUserId -eq 'user-2' }
        Assert-EntraTest ($failedRows.Count -eq 3 -and $failedUser.ResolutionStatus -eq 'UsersLoggedOnLookupFailed' -and $failedUser.LookupError -like '*User lookup denied*' -and ($lookupWarnings -join ' ') -like '*user-2*') 'Selected-user lookup failure remains explicit without losing other user rows'
        Assert-EntraTest (-not $failedUser.UserDisplayName -and -not $failedUser.UserPrincipalName -and $failedRows[0].UserPrincipalName -eq 'three@example.test') 'Failed selected-user lookup cannot display primary or previous selected-user identity'
        $global:EntraMockState.FailedUserIds = @()
        $singleLogon = @([pscustomobject]@{ userId='user-2'; lastLogOnDateTime='2026-10-01T10:00:00Z' })
        foreach ($mode in @('Latest','Most','All')) {
            $global:EntraMockState.LastUserDevice.usersLoggedOn = $singleLogon
            $singleRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode $mode -OutputPath $lastCsv -SkipGraphConnect -PassThru)
            Assert-EntraTest ($singleRows.Count -eq 1 -and $singleRows[0].ResolvedUserId -eq 'user-2' -and $singleRows[0].UserPrincipalName -eq 'two@example.test') "$mode returns the same single logged-on user"
            foreach ($fallbackCase in @(
                @{ Logons=@(); PrimaryId='user-1'; Status='PrimaryUserResolved' },
                @{ Logons=@([pscustomobject]@{ userId=' '; lastLogOnDateTime='2026-10-01T10:00:00Z' }); PrimaryId='user-1'; Status='PrimaryUserResolved' },
                @{ Logons=@([pscustomobject]@{ lastLogOnDateTime='2026-10-01T10:00:00Z' }); PrimaryId=$null; Status='UsersLoggedOnMissingUserId' },
                @{ Logons=@(); PrimaryId=$null; Status='PrimaryUserMissing' }
            )) {
                $global:EntraMockState.LastUserDevice.usersLoggedOn = $fallbackCase.Logons
                $global:EntraMockState.LastUserDevice.userId = $fallbackCase.PrimaryId
                $fallbackRows = @(& $lastPage.Path -DeviceName 'TEST-PC' -UserMode $mode -OutputPath $lastCsv -SkipGraphConnect -PassThru -WarningAction SilentlyContinue)
                Assert-EntraTest ($fallbackRows.Count -eq 1 -and $fallbackRows[0].ResolvedUserId -eq $fallbackCase.PrimaryId -and $fallbackRows[0].ResolutionStatus -eq $fallbackCase.Status -and $fallbackRows[0].UserSource -eq 'PrimaryUser' -and -not $fallbackRows[0].LastLoggedOnUserId) "$mode preserves primary-user fallback/missing status $($fallbackCase.Status)"
            }
            $global:EntraMockState.LastUserDevice.userId = 'user-1'
        }
    } finally {
        foreach ($key in @('LastUserDevice','LastUserProfiles','FailedUserIds')) { $global:EntraMockState.Remove($key) }
    }
    $groupRows = @(& $lastPage.Path -GroupNameOrId 'Device Group' -OutputPath $lastCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($groupRows.Count -eq 1 -and $groupRows[0].DeviceId -eq 'managed-1') 'Group input includes matching Intune devices and ignores non-device or unmanaged members'
    $groupIdRows = @(& $lastPage.Path -GroupNameOrId '11111111-2222-3333-4444-555555555555' -OutputPath $lastCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($groupIdRows.Count -eq 1 -and $groupIdRows[0].DeviceId -eq 'managed-1') 'Group input resolves a group by object ID'
    $entraExportCsv = Join-Path $tempDirectory 'EntraGroupExport.csv'
    @(
        [pscustomobject]@{ displayName='TEST-PC'; objectType='Device'; id='entra-object-1'; deviceId='aad-device-1' },
        [pscustomobject]@{ displayName='OTHER-PC'; objectType='Device'; id='entra-object-2'; deviceId='aad-device-not-managed' }
    ) | Export-Csv -LiteralPath $entraExportCsv -NoTypeInformation
    $exportRows = @(& $lastPage.Path -CsvPath $entraExportCsv -OutputPath $lastCsv -SkipGraphConnect -PassThru 3>$null)
    Assert-EntraTest ($exportRows.Count -eq 1 -and $exportRows[0].DeviceId -eq 'managed-1') 'CSV input matches Entra device IDs from a group member export'
    $managedIdCsv = Join-Path $tempDirectory 'ManagedIds.csv'
    @([pscustomobject]@{ ManagedDeviceId='managed-1' }) | Export-Csv -LiteralPath $managedIdCsv -NoTypeInformation
    $managedIdRows = @(& $lastPage.Path -CsvPath $managedIdCsv -OutputPath $lastCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($managedIdRows.Count -eq 1 -and $managedIdRows[0].DeviceId -eq 'managed-1') 'CSV input still resolves Intune managed device IDs'
    $nameOnlyCsv = Join-Path $tempDirectory 'DisplayNames.csv'
    @([pscustomobject]@{ displayName='TEST-PC' }) | Export-Csv -LiteralPath $nameOnlyCsv -NoTypeInformation
    $nameOnlyRows = @(& $lastPage.Path -CsvPath $nameOnlyCsv -OutputPath $lastCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($nameOnlyRows.Count -eq 1 -and $nameOnlyRows[0].DeviceId -eq 'managed-1') 'CSV input accepts a displayName column'
    $global:EntraMockState.GroupNameMatches = @()
    $missingGroupRejected = $false
    try { & $lastPage.Path -GroupNameOrId 'Missing Group' -OutputPath $lastCsv -SkipGraphConnect -PassThru | Out-Null } catch { $missingGroupRejected = $_.Exception.Message -like "No Entra group named*" }
    Assert-EntraTest $missingGroupRejected 'Group input rejects a missing display name'
    $global:EntraMockState.GroupNameMatches = @(
        [pscustomobject]@{ id='group-1'; displayName='Device Group' },
        [pscustomobject]@{ id='group-2'; displayName='Device Group' }
    )
    $ambiguousGroupRejected = $false
    try { & $lastPage.Path -GroupNameOrId 'Device Group' -OutputPath $lastCsv -SkipGraphConnect -PassThru | Out-Null } catch { $ambiguousGroupRejected = $_.Exception.Message -like 'More than one Entra group*' }
    Assert-EntraTest $ambiguousGroupRejected 'Group input rejects ambiguous display names'
    $global:EntraMockState.GroupNameMatches = @([pscustomobject]@{ id='group-1'; displayName='Device Group' })
    $groupsCsv = Join-Path $tempDirectory 'Groups.csv'
    [pscustomobject]@{ GroupName='Existing Group'; Description='Description'; MemberGroup='Nested Group' } | Export-Csv -LiteralPath $groupsCsv -Delimiter ';' -NoTypeInformation
    $groupRows = @(& $groupPage.Path -CsvPath $groupsCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($groupRows.Count -eq 1 -and $groupRows[0].Status -eq 'Reused' -and $groupRows[0].Membership -eq 'Added') 'Security group CSV returns group and membership outcomes'
    $groupPage.Results.Clear()
    foreach ($row in $groupRows) { [void]$groupPage.Results.Add($row) }
    Update-EntraResults $groupPage
    $groupPage.Table.UpdateLayout()
    Assert-EntraTest ($groupPage.Table.Items.Count -eq 1) 'Result objects bind to the WPF table'
    $groupPage.Filter.Text = 'does-not-match'
    Assert-EntraTest ($groupPage.Table.Items.Count -eq 0) 'Result filter excludes nonmatches'
    $groupPage.Filter.Text = ''
    $danishTitle = 'Socialr' + [char]0xE5 + 'dgiver'
    $resultsExportCsv = Join-Path $tempDirectory 'ResultsExport.csv'
    $lastPage.Results.Clear(); [void]$lastPage.Results.Add([pscustomobject]@{ DeviceName='TEST-PC'; JobTitle=$danishTitle })
    Export-EntraResults $lastPage $resultsExportCsv
    $exportBytes = [IO.File]::ReadAllBytes($resultsExportCsv)
    $exportedRows = @(Import-Csv -LiteralPath $resultsExportCsv -Encoding UTF8)
    Assert-EntraTest ($exportBytes[0] -eq 0xEF -and $exportBytes[1] -eq 0xBB -and $exportBytes[2] -eq 0xBF -and $exportedRows[0].JobTitle -eq $danishTitle) 'Results export writes UTF-8 with BOM so Excel shows Danish characters'
    $lastPage.Results.Clear()

    Add-Type -TypeDefinition 'public class EntraCleanupTestException : System.Exception { public int ResponseStatusCode { get; set; } public EntraCleanupTestException(int status) : base("Mock Graph failure") { ResponseStatusCode = status; } }'
    $cleanupState = @{
        Entra=[pscustomobject]@{ id='cccccccc-1111-1111-1111-111111111111'; deviceId='dddddddd-1111-1111-1111-111111111111'; displayName='TEST-CLEANUP'; trustType='AzureAd'; physicalIds=@('[ZTDId]:autopilot-1'); approximateLastSignInDateTime='2026-10-20T08:00:00Z'; accountEnabled=$true }
        ExtraDevices=@()
        Managed=@([pscustomobject]@{ id='managed-cleanup'; azureADDeviceId='dddddddd-1111-1111-1111-111111111111'; lastSyncDateTime='2026-10-21T10:00:00+02:00' })
        Autopilot=@([pscustomobject]@{ id='autopilot-1'; azureActiveDirectoryDeviceId='dddddddd-1111-1111-1111-111111111111'; managedDeviceId='managed-cleanup' })
        Removed=@{}; Writes=[Collections.Generic.List[string]]::new(); FailStage=''; InventoryFailure=$false; Sticky=$false; DuplicateName=$false; AutoRemoveEntra=$false
        AutopilotRequests=0; AutopilotFailures=0; AutopilotPaging=$false; AutopilotFailureStatus=500
        AutopilotDelayPolls=0; AutopilotPolls=0; StickyAutopilotId=''; AutopilotPollFailure=0; PendingAutopilot=@{}
        UsePostDeleteInventory=$false; PostDeleteInventory=@(); PostDeleteEntraMode=''; PostDeleteInventoryStatus=0; PostDeleteMalformedInventory=$false; PostDeletePageFailure=$false; PostDeleteSnapshots=0
        Delays=[Collections.Generic.List[int]]::new(); Events=[Collections.Generic.List[string]]::new(); DelayAction=''; TenantChanged=$false
        WithoutAutopilotWrite=$false; ChangeTenantAfterIntune=$false
    }
    $cleanupContextProvider = ${function:Get-MgContext}
    function Get-MgContext {
        $context = & $cleanupContextProvider
        if ($cleanupState.TenantChanged) { $context.TenantId = 'different-tenant' }
        if ($cleanupState.WithoutAutopilotWrite) { $context.Scopes = @($context.Scopes | Where-Object { $_ -ne 'DeviceManagementServiceConfig.ReadWrite.All' }) }
        return $context
    }
    function Start-Sleep {
        [CmdletBinding()]param([int]$Seconds)
        [void]$cleanupState.Delays.Add($Seconds)
        [void]$cleanupState.Events.Add("Delay:$Seconds")
        switch ($cleanupState.DelayAction) {
            'Cancel' { throw [System.Management.Automation.PipelineStoppedException]::new() }
            'Tenant' { $cleanupState.TenantChanged = $true }
        }
    }
    function Invoke-MgGraphRequest {
        [CmdletBinding()]param($Method,$Uri,$OutputType,$Body,$ContentType)
        [void]$cleanupState.Events.Add("${Method}:$($Uri.Split('/')[-1])")
        if ($Method -eq 'PATCH') {
            [void]$cleanupState.Writes.Add("PATCH $Uri $Body")
            if ($Uri -like "*/$($cleanupState.Entra.id)") { $cleanupState.Entra.accountEnabled = $false }
            return
        }
        if ($Method -eq 'DELETE') {
            if ($cleanupState.FailStage -and $Uri -like "*$($cleanupState.FailStage)*") { throw [EntraCleanupTestException]::new(403) }
            [void]$cleanupState.Writes.Add($Uri)
            if ($Uri -like '*windowsAutopilotDeviceIdentities/*' -and $cleanupState.AutopilotDelayPolls -gt 0) {
                $cleanupState.PendingAutopilot[$Uri] = $cleanupState.AutopilotDelayPolls
            } elseif (-not $cleanupState.Sticky -and (-not $cleanupState.StickyAutopilotId -or $Uri -notlike "*/$($cleanupState.StickyAutopilotId)")) {
                $cleanupState.Removed[$Uri] = $true
            }
            if ($cleanupState.AutoRemoveEntra -and $Uri -like '*windowsAutopilotDeviceIdentities/*') {
                $cleanupState.Removed['https://graph.microsoft.com/v1.0/devices/' + $cleanupState.Entra.id] = $true
            }
            if ($cleanupState.ChangeTenantAfterIntune -and $Uri -like '*managedDevices/*') { $cleanupState.TenantChanged = $true }
            return
        }
        if ($Uri -like '*windowsAutopilotDeviceIdentities/*' -and $cleanupState.Writes.Contains($Uri)) {
            $cleanupState.AutopilotPolls++
            if ($cleanupState.AutopilotPollFailure) { throw [EntraCleanupTestException]::new($cleanupState.AutopilotPollFailure) }
            if ($cleanupState.PendingAutopilot.ContainsKey($Uri)) {
                if ($cleanupState.PendingAutopilot[$Uri] -le 0) { $cleanupState.Removed[$Uri] = $true }
                else { $cleanupState.PendingAutopilot[$Uri]-- }
            }
        }
        if ($cleanupState.Removed.ContainsKey(($Uri -split '\?',2)[0])) { throw [EntraCleanupTestException]::new(404) }
        if ($Uri -like '*managedDevices?*') {
            if ($cleanupState.InventoryFailure) { throw [EntraCleanupTestException]::new(403) }
            return [pscustomobject]@{ value=@($cleanupState.Managed); '@odata.nextLink'=$null }
        }
        if ($Uri -match '/windowsAutopilotDeviceIdentities(?:\?|$)') {
            $cleanupState.AutopilotRequests++
            if ($Uri -match '(?i)[?&]\$select=') { throw [EntraCleanupTestException]::new(500) }
            if ($cleanupState.AutopilotFailures -gt 0) {
                $cleanupState.AutopilotFailures--
                throw [EntraCleanupTestException]::new($cleanupState.AutopilotFailureStatus)
            }
            if ($cleanupState.UsePostDeleteInventory -and @($cleanupState.Writes | Where-Object { $_ -like '*windowsAutopilotDeviceIdentities/*' }).Count -gt 0) {
                $cleanupState.PostDeleteSnapshots++
                if ($cleanupState.PostDeleteInventoryStatus) { throw [EntraCleanupTestException]::new($cleanupState.PostDeleteInventoryStatus) }
                if ($cleanupState.PostDeleteMalformedInventory) { return [pscustomobject]@{ '@odata.nextLink'=$null } }
                if ($cleanupState.PostDeletePageFailure) {
                    if ($Uri -match '\$skiptoken=') { throw [EntraCleanupTestException]::new(403) }
                    return [pscustomobject]@{ value=@(); '@odata.nextLink'='https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$skiptoken=verification-test' }
                }
                return [pscustomobject]@{ value=@($cleanupState.PostDeleteInventory); '@odata.nextLink'=$null }
            }
            if ($cleanupState.AutopilotPaging -and $Uri -notmatch '\$skiptoken=') {
                return [pscustomobject]@{ value=@(); '@odata.nextLink'='https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$skiptoken=cleanup-test' }
            }
            return [pscustomobject]@{ value=@($cleanupState.Autopilot); '@odata.nextLink'=$null }
        }
        if ($Uri -like '*managedDevices/managed-cleanup') { return $cleanupState.Managed[0] }
        if ($Uri -match '/windowsAutopilotDeviceIdentities/([^/?]+)$') {
            $registration = @($cleanupState.Autopilot | Where-Object { $_.id -eq $Matches[1] })
            if ($registration.Count) { return $registration[0] }
            throw [EntraCleanupTestException]::new(404)
        }
        if ($Uri -like '*devices/cccccccc-1111-1111-1111-111111111111*') {
            if ($cleanupState.UsePostDeleteInventory -and @($cleanupState.Writes | Where-Object { $_ -like '*windowsAutopilotDeviceIdentities/*' }).Count -gt 0) {
                $entry = $cleanupState.Entra.PSObject.Copy()
                switch ($cleanupState.PostDeleteEntraMode) {
                    'Clear' { $entry.physicalIds = @() }
                    'Absent' { throw [EntraCleanupTestException]::new(404) }
                    'Missing' { $entry.PSObject.Properties.Remove('physicalIds') }
                    'Null' { $entry.physicalIds = $null }
                    'Invalid' { $entry.physicalIds = 'not-an-array' }
                    'WrongDevice' { $entry.deviceId = 'dddddddd-2222-2222-2222-222222222222' }
                    'Hybrid' { $entry.trustType = 'ServerAd' }
                }
                return $entry
            }
            return $cleanupState.Entra
        }
        if ($Uri -like '*devices?*') {
            $values = @($cleanupState.Entra)
            if ($cleanupState.DuplicateName) { $values += [pscustomobject]@{ id='cccccccc-2222-2222-2222-222222222222'; deviceId='dddddddd-2222-2222-2222-222222222222'; displayName='TEST-CLEANUP'; trustType='Workplace'; physicalIds=@(); approximateLastSignInDateTime=$null; accountEnabled=$true } }
            if ($Uri -match "operatingSystem eq 'Windows'") { $values += @($cleanupState.ExtraDevices) }
            if ($Uri -match "displayName eq '([^']*)'") {
                $requestedName = [Uri]::UnescapeDataString($Matches[1])
                $values = @($values | Where-Object { $_.displayName -eq $requestedName })
            }
            return [pscustomobject]@{ value=$values; '@odata.nextLink'=$null }
        }
        throw "Unexpected cleanup mock URI: $Uri"
    }
    $cleanupCsv = Join-Path $tempDirectory 'Cleanup.csv'
    @([pscustomobject]@{ displayName='TEST-CLEANUP'; objectType='device'; id=$cleanupState.Entra.id; deviceId=$cleanupState.Entra.deviceId },
      [pscustomobject]@{ displayName='Ignored user'; objectType='user'; id='eeeeeeee-1111-1111-1111-111111111111'; deviceId='' }) |
        Export-Csv -LiteralPath $cleanupCsv -NoTypeInformation -Encoding UTF8
    $cleanupRows = @(& $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($cleanupRows.Count -eq 1 -and $cleanupRows[0].ObjectId -eq $cleanupState.Entra.id -and $cleanupRows[0].'Device ID' -eq $cleanupState.Entra.deviceId) 'Cleanup imports Entra group export and separates object ID from device ID'
    Assert-EntraTest ($cleanupRows[0].'Intune managed' -eq 'Yes' -and $cleanupRows[0].'Activity date' -eq ([DateTimeOffset]'2026-10-21T08:00:00Z').UtcDateTime) 'Cleanup correlates Intune by registration ID and chooses latest UTC activity'
    Assert-EntraTest ($cleanupRows[0].AutopilotIds.Count -eq 1 -and $cleanupRows[0].CanDelete -and $cleanupState.Writes.Count -eq 0) 'Inspection resolves Autopilot without writes'
    Assert-EntraTest ($cleanupState.AutopilotRequests -eq 1) 'Autopilot inventory uses one plain request without select projection'
    $cleanupState.AutopilotPaging = $true; $cleanupState.AutopilotFailures = 1
    $beforeRequests = $cleanupState.AutopilotRequests
    $retryRows = @(& $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($retryRows[0].AutopilotIds.Count -eq 1 -and $cleanupState.AutopilotRequests - $beforeRequests -eq 3) 'Transient Autopilot 500 retries and follows the service nextLink before correlating devices'
    $cleanupState.AutopilotPaging = $false
    $requirements = Get-EntraRequirements $cleanupPage @{ Action='Inspect' }
    Assert-EntraTest ('Directory.AccessAsUser.All' -notin $requirements.Scopes) 'Inspection requests read permissions only'
    $requirements = Get-EntraRequirements $cleanupPage @{ Action='Delete'; WhatIf=$false }
    Assert-EntraTest ('Directory.AccessAsUser.All' -in $requirements.Scopes -and 'DeviceManagementManagedDevices.ReadWrite.All' -in $requirements.Scopes -and 'DeviceManagementServiceConfig.Read.All' -in $requirements.Scopes -and 'DeviceManagementServiceConfig.ReadWrite.All' -notin $requirements.Scopes) 'Keep-mode deletion needs directory and Intune write permissions, but only Autopilot read'
    $requirements = Get-EntraRequirements $cleanupPage @{ Action='Delete'; WhatIf=$false; DeleteAutopilotRegistration=$true }
    Assert-EntraTest ('DeviceManagementServiceConfig.ReadWrite.All' -in $requirements.Scopes) 'Selected deregistration adds Autopilot write permission'
    $requirements = Get-EntraRequirements $cleanupPage @{ Action='Delete'; WhatIf=$true; DeleteAutopilotRegistration=$true }
    Assert-EntraTest ('Directory.AccessAsUser.All' -notin $requirements.Scopes -and 'DeviceManagementServiceConfig.ReadWrite.All' -notin $requirements.Scopes) 'Deregistration preview still needs only read permissions'
    $cleanupPage.Controls.Preview.IsChecked = $false
    $requirements = Get-EntraRequirements $cleanupPage @{}
    Assert-EntraTest ('Directory.AccessAsUser.All' -in $requirements.Scopes -and 'DeviceManagementServiceConfig.ReadWrite.All' -notin $requirements.Scopes) 'Cleanup sign-in reflects the unchecked deregistration option'
    $cleanupPage.Controls.DeleteAutopilotRegistration.IsChecked = $true
    $requirements = Get-EntraRequirements $cleanupPage @{}
    Assert-EntraTest ('DeviceManagementServiceConfig.ReadWrite.All' -in $requirements.Scopes) 'Cleanup sign-in reflects the checked deregistration option'
    $cleanupPage.Controls.DeleteAutopilotRegistration.IsChecked = $false
    $cleanupPage.Controls.Preview.IsChecked = $true
    $requirements = Get-EntraRequirements $cleanupPage @{ Action='Disable'; WhatIf=$false }
    Assert-EntraTest ('Directory.AccessAsUser.All' -in $requirements.Scopes -and 'DeviceManagementServiceConfig.ReadWrite.All' -notin $requirements.Scopes) 'Disable requests only the directory write permission'
    $requirements = Get-EntraRequirements $cleanupPage @{ Action='Disable'; WhatIf=$true }
    Assert-EntraTest ('Directory.AccessAsUser.All' -notin $requirements.Scopes) 'Disable preview requests read permissions only'

    $recentSync = [DateTime]::UtcNow.AddDays(-2).ToString('o')
    $cleanupState.ExtraDevices = @(
        [pscustomobject]@{ id='cccccccc-3333-3333-3333-333333333333'; deviceId='dddddddd-3333-3333-3333-333333333333'; displayName='STALE-PC'; trustType='AzureAd'; physicalIds=@(); approximateLastSignInDateTime='2020-01-01T00:00:00Z'; accountEnabled=$true },
        [pscustomobject]@{ id='cccccccc-4444-4444-4444-444444444444'; deviceId='dddddddd-4444-4444-4444-444444444444'; displayName='NEVER-PC'; trustType='AzureAd'; physicalIds=@(); approximateLastSignInDateTime=$null; accountEnabled=$false },
        [pscustomobject]@{ id='cccccccc-5555-5555-5555-555555555555'; deviceId='dddddddd-5555-5555-5555-555555555555'; displayName='RECENT-PC'; trustType='AzureAd'; physicalIds=@(); approximateLastSignInDateTime=$recentSync; accountEnabled=$true },
        [pscustomobject]@{ id='cccccccc-6666-6666-6666-666666666666'; deviceId='dddddddd-6666-6666-6666-666666666666'; displayName='SYNCED-PC'; trustType='AzureAd'; physicalIds=@(); approximateLastSignInDateTime='2020-01-01T00:00:00Z'; accountEnabled=$true })
    $savedManaged = $cleanupState.Managed
    $cleanupState.Managed = @($savedManaged) + [pscustomobject]@{ id='managed-synced'; azureADDeviceId='dddddddd-6666-6666-6666-666666666666'; lastSyncDateTime=$recentSync }
    $staleRows = @(& $cleanupPage.Path -Source Entra -InactivityDays 30 -SkipGraphConnect -PassThru)
    $staleNames = @($staleRows | Where-Object { $_.'Device Name' -ne 'TEST-CLEANUP' } | ForEach-Object { $_.'Device Name' } | Sort-Object)
    Assert-EntraTest (($staleNames -join ',') -eq 'NEVER-PC,STALE-PC' -and $cleanupState.Writes.Count -eq 0) 'Entra lookup lists stale and never-active Windows devices but not recent Entra or Intune activity'
    $neverRow = @($staleRows | Where-Object { $_.'Device Name' -eq 'NEVER-PC' })[0]
    Assert-EntraTest ($neverRow.Enabled -eq 'No' -and $neverRow.Autopilot -eq 'No' -and $null -eq $neverRow.'Activity date' -and @($staleRows | Where-Object { $_.'Device Name' -eq 'STALE-PC' })[0].Enabled -eq 'Yes') 'Entra lookup rows show enabled state and Autopilot registration'
    Assert-EntraTest (@($cleanupState.Events | Where-Object { $_ -like "GET:devices?*operatingSystem eq 'Windows'*" }).Count -eq 1) 'Entra lookup queries Windows devices once'
    $cleanupState.Managed = $savedManaged; $cleanupState.ExtraDevices = @()
    $cleanupState.DuplicateName = $true
    $nameRows = @(& $cleanupPage.Path -Source Names -DeviceName ' test-cleanup ','TEST-CLEANUP','MISSING-PC' -SkipGraphConnect -PassThru 3>$null)
    $cleanupState.DuplicateName = $false
    Assert-EntraTest ($nameRows.Count -eq 3 -and @($nameRows | Where-Object { $_.'Device Name' -eq 'TEST-CLEANUP' }).Count -eq 2) 'Device names match exactly once per name and list every duplicate'
    Assert-EntraTest (@($nameRows | Where-Object { $_.'Device Name' -eq 'MISSING-PC' -and $_.Status -like 'Failed:*not found*' -and -not $_.CanDelete }).Count -eq 1) 'Unknown device names are reported as not found'
    $namesRejected = $false
    try { & $cleanupPage.Path -Source Names -DeviceName ' ' -SkipGraphConnect -PassThru | Out-Null } catch { $namesRejected = $true }
    Assert-EntraTest $namesRejected 'Device names source requires at least one name'

    $disableArguments = @{ Action='Disable'; Device=@($cleanupRows[0]); ExpectedTenantId=$cleanupRows[0].TenantId; SkipGraphConnect=$true; PassThru=$true; Confirm=$false }
    $cleanupState.Writes.Clear()
    $disablePreview = @(& $cleanupPage.Path @disableArguments -WhatIf)
    Assert-EntraTest ($disablePreview[0].Status -eq 'Preview; no changes' -and $cleanupState.Writes.Count -eq 0 -and $cleanupState.Entra.accountEnabled) 'Disable WhatIf does not write'
    $disabledRows = @(& $cleanupPage.Path @disableArguments)
    Assert-EntraTest ($disabledRows[0].Status -eq 'Disabled' -and $disabledRows[0].Enabled -eq 'No' -and $disabledRows[0].CanDelete -and $cleanupState.Writes.Count -eq 1 -and $cleanupState.Writes[0] -like "PATCH */devices/$($cleanupState.Entra.id) *accountEnabled*false*") 'Disable patches only accountEnabled and confirms the result'
    Assert-EntraTest (@($cleanupState.Writes | Where-Object { $_ -notlike 'PATCH*' }).Count -eq 0) 'Disable never deletes Intune, Autopilot or Entra records'
    $alreadyRows = @(& $cleanupPage.Path @disableArguments)
    Assert-EntraTest ($alreadyRows[0].Status -eq 'Already disabled' -and $cleanupState.Writes.Count -eq 1) 'Already disabled device is skipped without another write'
    Assert-EntraTest ((Get-EntraOutcomeCategory $disabledRows[0]) -eq 'Success' -and (Get-EntraOutcomeCategory $alreadyRows[0]) -eq 'Skipped') 'Disable outcomes count as success or skipped'
    $cleanupState.Entra.accountEnabled = $true; $cleanupState.Entra.trustType = 'ServerAd'
    $hybridDisable = @(& $cleanupPage.Path @disableArguments 3>$null)
    Assert-EntraTest ($hybridDisable[0].Status -like 'Failed:*blocked*' -and $cleanupState.Writes.Count -eq 1) 'Refreshed hybrid identity blocks disable'
    $cleanupState.Entra.trustType = 'AzureAd'
    $disableRejected = $false
    try { & $cleanupPage.Path -Action Disable -Device $cleanupRows[0] -SkipGraphConnect -PassThru -Confirm:$false | Out-Null } catch { $disableRejected = $true }
    Assert-EntraTest ($disableRejected -and $cleanupState.Writes.Count -eq 1) 'Disable requires the expected tenant'
    $cleanupState.Writes.Clear(); $cleanupState.Events.Clear()
    $script:ConnectedTenant = $cleanupRows[0].TenantId
    $cleanupPage.Results.Clear(); [void]$cleanupPage.Results.Add($cleanupRows[0]); Update-EntraResults $cleanupPage
    $cleanupPage.Table.SelectedItem = $cleanupRows[0]
    $deletion = Get-EntraCleanupDeletionInvocation $cleanupPage
    Assert-EntraTest ($deletion.Preview -and $deletion.Parameters.Device.Count -eq 1 -and $cleanupPage.DeleteSelected.IsEnabled) 'Cleanup selected preview captures only eligible same-tenant IDs'
    $cleanupPage.Controls.DeleteAutopilotRegistration.IsChecked = $true
    $deregistration = Get-EntraCleanupDeletionInvocation $cleanupPage
    Assert-EntraTest (-not $deletion.Parameters.DeleteAutopilotRegistration -and $deregistration.Parameters.DeleteAutopilotRegistration -and $deregistration.Summary -like '*Deregister linked Autopilot registrations.*' -and $deregistration.Summary -notlike '*enrollment failures*') 'Deletion captures the chosen mode without changing an already prepared invocation'
    $cleanupPage.Controls.DeleteAutopilotRegistration.IsChecked = $false
    $disableInvocation = Get-EntraCleanupDisableInvocation $cleanupPage
    Assert-EntraTest ($cleanupPage.DisableSelected.IsEnabled -and $disableInvocation.Preview -and $disableInvocation.Parameters.Action -eq 'Disable' -and $disableInvocation.Parameters.Device.Count -eq 1 -and -not $disableInvocation.Parameters.ContainsKey('AcknowledgeCloudDeletion')) 'Disable selected previews enabled eligible devices without the deletion acknowledgement'
    $cleanupPage.Controls.Preview.IsChecked = $false
    $disableInvocation = Get-EntraCleanupDisableInvocation $cleanupPage
    Assert-EntraTest (-not $disableInvocation.Preview -and $disableInvocation.Summary -like 'Disable 1 selected device*' -and (Get-EntraConfirmationMessage $cleanupPage $disableInvocation) -like "*Tenant: $script:ConnectedTenant*") 'Live disable requires a tenant confirmation but no acknowledgement'
    $cleanupPage.Controls.Preview.IsChecked = $true
    $disabledCopy = $cleanupRows[0].PSObject.Copy(); $disabledCopy.Enabled = 'No'
    $cleanupPage.Results.Clear(); [void]$cleanupPage.Results.Add($disabledCopy); Update-EntraResults $cleanupPage
    $cleanupPage.Table.SelectedItem = $disabledCopy; Update-EntraCleanupSelection $cleanupPage
    Assert-EntraTest (-not $cleanupPage.DisableSelected.IsEnabled -and $cleanupPage.DeleteSelected.IsEnabled) 'Disable selected ignores devices that are already disabled'
    $cleanupPage.Results.Clear(); [void]$cleanupPage.Results.Add($cleanupRows[0]); Update-EntraResults $cleanupPage
    $cleanupPage.Table.SelectedItem = $cleanupRows[0]
    $keepConfirmation = Get-EntraConfirmationMessage -Page $cleanupPage -Invocation $deletion
    Assert-EntraTest ($keepConfirmation -like 'Delete 1 selected device?*Keep Autopilot registrations.*enrollment failures*' -and $keepConfirmation -notlike '*TEST-CLEANUP*') 'Cleanup confirmation includes the keep-mode warning and count without device names'
    $cleanupPage.Results.Clear()
    foreach ($number in 1..200) {
        $copy = $cleanupRows[0].PSObject.Copy()
        $copy.'Device Name' = "Confirmation device $number"; $copy.ObjectId = [guid]::NewGuid().ToString()
        [void]$cleanupPage.Results.Add($copy)
    }
    Update-EntraResults $cleanupPage
    $cleanupPage.SelectAll.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    $bulkDeletion = Get-EntraCleanupDeletionInvocation $cleanupPage
    $bulkConfirmation = Get-EntraConfirmationMessage -Page $cleanupPage -Invocation $bulkDeletion
    Assert-EntraTest ($bulkConfirmation -like 'Delete 200 selected devices?*Keep Autopilot registrations.*' -and $bulkConfirmation -notlike '*Confirmation device*') 'Bulk cleanup confirmation stays concise and mode-specific for 200 devices'
    Assert-EntraTest ($bulkDeletion.Parameters.Device.Count -eq 200 -and @($bulkDeletion.Parameters.Device.ObjectId | Select-Object -Unique).Count -eq 200 -and $bulkDeletion.Parameters.ExpectedTenantId -eq $script:ConnectedTenant) 'Concise confirmation preserves every selected identity and tenant binding'
    $otherInvocation = @{ Summary='Update selected users.'; Parameters=@{} }
    Assert-EntraTest ((Get-EntraConfirmationMessage -Page $attributePage -Invocation $otherInvocation) -eq "Update selected users.`n`nTenant: $script:ConnectedTenant`nChanges take effect immediately. Continue?") 'Other pages retain their existing confirmation footer'
    Assert-EntraTest ($cleanupState.Writes.Count -eq 0) 'Confirmation formatting performs no Graph writes'
    $cleanupPage.Results.Clear(); [void]$cleanupPage.Results.Add($cleanupRows[0]); Update-EntraResults $cleanupPage
    $cleanupPage.Controls.NameFilter.Text = 'absent'
    Assert-EntraTest ($cleanupPage.Table.Items.Count -eq 0 -and -not $cleanupPage.DeleteSelected.IsEnabled) 'Filtering hides and excludes selected cleanup targets'
    $cleanupPage.Controls.NameFilter.Text = ''
    $selectionFixtures = @($cleanupRows[0])
    foreach ($name in @('TEST-SECOND','HIDDEN-DEVICE','TEST-BLOCKED','TEST-OTHER-TENANT','TEST-DELETED')) {
        $copy = $cleanupRows[0].PSObject.Copy()
        $copy.'Device Name' = $name; $copy.ObjectId = $name
        switch ($name) {
            'TEST-BLOCKED' { $copy.CanDelete = $false; $copy.'Join Type' = 'Microsoft Entra hybrid joined' }
            'TEST-OTHER-TENANT' { $copy.TenantId = 'different-tenant' }
            'TEST-DELETED' { $copy.Status = 'Deleted' }
        }
        $selectionFixtures += $copy
    }
    $cleanupPage.Results.Clear()
    foreach ($row in $selectionFixtures) { [void]$cleanupPage.Results.Add($row) }
    $cleanupPage.Controls.NameFilter.Text = 'TEST-'
    $cleanupPage.SelectAll.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    Assert-EntraTest ($cleanupPage.Table.SelectedItems.Count -eq 2 -and @(Get-EntraCleanupSelection $cleanupPage).Count -eq 2) 'Select all excludes filtered, hybrid, other-tenant and deleted devices'
    $cleanupPage.SelectAll.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    Assert-EntraTest ($cleanupPage.Table.SelectedItems.Count -eq 2 -and -not $cleanupPage.SelectAll.IsEnabled -and $cleanupPage.ClearSelection.IsEnabled) 'Repeated select all is idempotent and updates button states'
    $cleanupPage.ClearSelection.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    Assert-EntraTest ($cleanupPage.Table.SelectedItems.Count -eq 0 -and -not $cleanupPage.DeleteSelected.IsEnabled -and $cleanupPage.SelectAll.IsEnabled) 'Clear selection deselects devices and disables deletion'
    [void]$cleanupPage.Table.SelectedItems.Add($selectionFixtures[3])
    $cleanupPage.ClearSelection.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    Assert-EntraTest ($cleanupPage.Table.SelectedItems.Count -eq 0) 'Clear selection also deselects manually selected blocked rows'
    $script:Job = @{ Page=$cleanupPage }
    try {
        Set-EntraBusy $true
        $cleanupPage.SelectAll.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
        Assert-EntraTest (-not $cleanupPage.SelectAll.IsEnabled -and -not $cleanupPage.ClearSelection.IsEnabled -and $cleanupPage.Table.SelectedItems.Count -eq 0) 'Busy selection actions are disabled and cannot select rows even through direct events'
        [void]$cleanupPage.Table.SelectedItems.Add($selectionFixtures[0])
        $cleanupPage.ClearSelection.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
        Assert-EntraTest ($cleanupPage.Table.SelectedItems.Count -eq 1) 'Busy clear-selection event cannot change the active selection'
    } finally { $script:Job = $null; Set-EntraBusy $false }
    $cleanupPage.Controls.NameFilter.Text = 'absent'
    Assert-EntraTest (-not $cleanupPage.SelectAll.IsEnabled -and -not $cleanupPage.ClearSelection.IsEnabled) 'Empty filtered list disables both selection actions'
    Assert-EntraTest ($cleanupState.Writes.Count -eq 0 -and $cleanupPage.Controls.NameFilter.Text -eq 'absent') 'Selection actions do not call Graph writes or change filters'
    $cleanupPage.Results.Clear(); [void]$cleanupPage.Results.Add($cleanupRows[0]); $cleanupPage.Controls.NameFilter.Text = ''; Update-EntraResults $cleanupPage
    $filterRows = @(Select-EntraCleanupRows -Rows $cleanupRows -Name 'CLEANUP' -Managed 'Yes' -JoinType 'Microsoft Entra joined' -From '2026-10-21' -To '2026-10-21')
    Assert-EntraTest ($filterRows.Count -eq 1) 'Cleanup column filters combine and use inclusive UTC date boundaries'
    $cleanupPage.Table.SelectedItem = $cleanupRows[0]
    $cleanupPage.Controls.Preview.IsChecked = $false
    $ackRejected = $false
    try { Get-EntraCleanupDeletionInvocation $cleanupPage | Out-Null } catch { $ackRejected = $true }
    Assert-EntraTest $ackRejected 'Cleanup real deletion requires warning acknowledgement'
    $cleanupPage.Controls.Preview.IsChecked = $true
    $deleteArguments = @{ Action='Delete'; Device=$cleanupRows; ExpectedTenantId=$cleanupRows[0].TenantId; AcknowledgeCloudDeletion=$true; SkipGraphConnect=$true; PassThru=$true; Confirm=$false }
    $keepArguments = $deleteArguments.Clone(); $keepArguments.DeleteAutopilotRegistration = $false
    $previewRows = @(& $cleanupPage.Path @deleteArguments -WhatIf)
    $keepPreview = @(& $cleanupPage.Path @keepArguments -WhatIf 3>$null)
    Assert-EntraTest ($cleanupState.Writes.Count -eq 0 -and $cleanupState.Delays.Count -eq 0 -and $previewRows[0].Status -like 'Preview*deregister*' -and $keepPreview[0].Status -like 'Preview*keep registration*') 'CLI deregisters Autopilot by default; explicit false previews keep mode, without writes or delay'
    $cleanupState.WithoutAutopilotWrite = $true; $cleanupState.Events.Clear()
    try {
        $keepWarnings = @()
        $keptRows = @(& $cleanupPage.Path @keepArguments -WarningVariable keepWarnings 3>$null)
        $keepEvents = @($cleanupState.Events | Where-Object { $_ -like 'DELETE:*' -or $_ -like 'Delay:*' })
        Assert-EntraTest (($keepEvents -join ',') -eq 'DELETE:managed-cleanup,DELETE:cccccccc-1111-1111-1111-111111111111' -and $cleanupState.Delays.Count -eq 0 -and $keptRows[0].Status -eq 'Entra removed; Autopilot registration kept' -and -not $keptRows[0].AutopilotRemovalRequested -and $keptRows[0].Autopilot -eq 'Yes' -and -not $keptRows[0].CanDelete) 'CLI keep mode deletes Intune then Entra without deregistration, delay or Autopilot write permission'
        Assert-EntraTest (($keepWarnings -join ' ') -like '*enrollment failures*' -and (Get-EntraOutcomeCategory $keptRows[0]) -eq 'Success') 'Keep mode warns about enrollment risk and reports confirmed Entra removal as success'
        $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear()
        $permissionRejected = $false
        try { & $cleanupPage.Path @deleteArguments | Out-Null } catch { $permissionRejected = $_.Exception.Message -like '*lacks the required scopes*' }
        Assert-EntraTest ($permissionRejected -and $cleanupState.Writes.Count -eq 0) 'Default CLI deregistration requires Autopilot write permission before any deletion'
    } finally { $cleanupState.WithoutAutopilotWrite = $false }
    $cleanupState.FailStage = $cleanupState.Entra.id; $cleanupState.Events.Clear()
    $keepFailedRows = @(& $cleanupPage.Path @keepArguments 3>$null)
    Assert-EntraTest ($keepFailedRows[0].Status -like 'Failed:*' -and $cleanupState.Writes.Count -eq 1 -and @($cleanupState.Events | Where-Object { $_ -like 'DELETE:autopilot-*' }).Count -eq 0 -and $cleanupState.Delays.Count -eq 0) 'Keep-mode Entra rejection is a failure, with no fallback Autopilot deregistration'
    $cleanupState.FailStage = ''; $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Sticky = $true
    $keepPendingRows = @(& $cleanupPage.Path @keepArguments -RemovalTimeoutSeconds 1 3>$null)
    Assert-EntraTest ($keepPendingRows[0].Status -like 'Pending confirmation:*Intune*' -and -not $keepPendingRows[0].CanDelete -and $cleanupState.Writes.Count -eq 1) 'Keep mode still blocks Entra deletion when Intune removal is unconfirmed'
    $cleanupState.Sticky = $false; $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.ChangeTenantAfterIntune = $true
    try {
        $keepChangedRows = @(& $cleanupPage.Path @keepArguments 3>$null)
        Assert-EntraTest ($keepChangedRows[0].Status -like 'Failed:*tenant changed*' -and $cleanupState.Writes.Count -eq 1) 'Keep mode rechecks the tenant before deleting Entra'
    } finally { $cleanupState.ChangeTenantAfterIntune = $false; $cleanupState.TenantChanged = $false }
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Events.Clear()
    $cleanupState.AutopilotFailures = 4; $beforeRequests = $cleanupState.AutopilotRequests
    $autopilotRejected = $false
    try { & $cleanupPage.Path @deleteArguments | Out-Null } catch { $autopilotRejected = $true }
    Assert-EntraTest ($autopilotRejected -and $cleanupState.AutopilotRequests - $beforeRequests -eq 4 -and $cleanupState.Writes.Count -eq 0) 'Persistent Autopilot 500 stops after bounded retries and blocks all deletion stages'
    $cleanupState.AutopilotFailures = 1; $cleanupState.AutopilotFailureStatus = 403; $beforeRequests = $cleanupState.AutopilotRequests
    $autopilotRejected = $false
    try { & $cleanupPage.Path @deleteArguments | Out-Null } catch { $autopilotRejected = $true }
    Assert-EntraTest ($autopilotRejected -and $cleanupState.AutopilotRequests - $beforeRequests -eq 1 -and $cleanupState.Writes.Count -eq 0) 'Autopilot permission failure is not retried or treated as an empty inventory'
    $cleanupState.AutopilotFailureStatus = 500
    $cleanupState.FailStage = 'windowsAutopilotDeviceIdentities'
    $failedRows = @(& $cleanupPage.Path @deleteArguments)
    Assert-EntraTest ($cleanupState.Writes.Count -eq 1 -and $cleanupState.Writes[0] -like '*managedDevices/*' -and $failedRows[0].Status -like 'Failed:*' -and $cleanupState.Delays.Count -eq 0) 'Autopilot failure prevents the delay and Entra deletion after Intune removal'
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.FailStage = 'managedDevices'
    $failedRows = @(& $cleanupPage.Path @deleteArguments)
    Assert-EntraTest ($cleanupState.Writes.Count -eq 0 -and $failedRows[0].Status -like 'Failed:*') 'Intune removal failure prevents Autopilot and Entra writes'
    $cleanupState.FailStage = ''; $cleanupState.Sticky = $true
    $failedRows = @(& $cleanupPage.Path @deleteArguments -RemovalTimeoutSeconds 1)
    Assert-EntraTest ($cleanupState.Writes.Count -eq 1 -and $failedRows[0].Status -like 'Pending confirmation:*Intune*not confirmed*' -and -not $failedRows[0].CanDelete) 'Unconfirmed Intune removal is pending and blocks later stages and stale retries'
    $cleanupState.Sticky = $false
    $cleanupState.FailStage = ''; $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear()
    $cleanupState.Events.Clear(); $cleanupState.Delays.Clear()
    $deletedRows = @(& $cleanupPage.Path @deleteArguments)
    $expectedEvents = @('DELETE:managed-cleanup','DELETE:autopilot-1','Delay:10','DELETE:cccccccc-1111-1111-1111-111111111111')
    $actualEvents = @($cleanupState.Events | Where-Object { $_ -like 'DELETE:*' -or $_ -like 'Delay:*' })
    Assert-EntraTest (($actualEvents -join ',') -eq ($expectedEvents -join ',') -and $deletedRows[0].Status -eq 'Entra removed; Autopilot removal requested' -and -not $deletedRows[0].CanDelete) 'Cleanup confirms Intune, submits Autopilot deletion, pauses ten seconds once, then confirms Entra removal'
    Assert-EntraTest ($cleanupState.Delays.Count -eq 1 -and $cleanupState.Delays[0] -eq 10 -and $cleanupState.AutopilotPolls -eq 0) 'Autopilot handoff does not poll after the accepted delete request'
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.AutoRemoveEntra = $true
    $deletedRows = @(& $cleanupPage.Path @deleteArguments)
    Assert-EntraTest ($deletedRows[0].Status -eq 'Entra removed; Autopilot removal requested' -and $cleanupState.Writes.Count -eq 2 -and $cleanupState.Delays.Count -eq 1) 'Autopilot auto-removal of Entra avoids unnecessary directory deletion after the single pause'
    $cleanupState.AutoRemoveEntra = $false
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.StickyAutopilotId = 'autopilot-1'
    $beforeRequests = $cleanupState.AutopilotRequests
    $unconfirmedRows = @(& $cleanupPage.Path @deleteArguments -AutopilotRemovalTimeoutSeconds 1)
    Assert-EntraTest ($unconfirmedRows[0].Status -eq 'Entra removed; Autopilot removal requested' -and $cleanupState.Delays.Count -eq 1 -and $cleanupState.Delays[0] -eq 10 -and $cleanupState.AutopilotRequests -eq $beforeRequests + 1 -and $cleanupState.AutopilotPolls -eq 0) 'Accepted Autopilot deletion proceeds after the fixed pause even when the direct record remains visible'
    $cleanupState.Autopilot += [pscustomobject]@{ id='autopilot-2'; azureActiveDirectoryDeviceId=$cleanupState.Entra.deviceId; managedDeviceId='managed-cleanup' }
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.Events.Clear()
    $multipleRows = @(& $cleanupPage.Path @deleteArguments)
    $actualEvents = @($cleanupState.Events | Where-Object { $_ -like 'DELETE:*' -or $_ -like 'Delay:*' })
    Assert-EntraTest (($actualEvents -join ',') -eq 'DELETE:managed-cleanup,DELETE:autopilot-1,DELETE:autopilot-2,Delay:10,DELETE:cccccccc-1111-1111-1111-111111111111' -and $cleanupState.Delays.Count -eq 1 -and $multipleRows[0].AutopilotRemovalRequested) 'Multiple Autopilot registrations share exactly one delay after the final accepted request'
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.FailStage = 'autopilot-2'
    $failedRows = @(& $cleanupPage.Path @deleteArguments)
    Assert-EntraTest ($failedRows[0].Status -like 'Failed:*' -and $failedRows[0].AutopilotRemovalRequested -and $cleanupState.Writes.Count -eq 2 -and $cleanupState.Delays.Count -eq 0) 'A later Autopilot request failure prevents the pause and Entra deletion despite an earlier accepted request'
    $cleanupState.FailStage = ''
    $cleanupState.Autopilot = @($cleanupState.Autopilot[0]); $cleanupState.StickyAutopilotId = ''
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear()
    $cleanupState.Removed['https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/autopilot-1'] = $true
    $absentRows = @(& $cleanupPage.Path @deleteArguments)
    Assert-EntraTest ($absentRows[0].Status -eq 'Deleted' -and -not $absentRows[0].AutopilotRemovalRequested -and $cleanupState.Delays.Count -eq 0 -and $cleanupState.Writes.Count -eq 2) 'Already absent Autopilot records require neither another delete nor a delay'
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.DelayAction = 'Tenant'
    try {
        $changedRows = @(& $cleanupPage.Path @deleteArguments)
        Assert-EntraTest ($changedRows[0].Status -like 'Failed:*tenant changed*' -and $cleanupState.Writes.Count -eq 2 -and $cleanupState.Delays.Count -eq 1) 'Tenant is revalidated after the ten-second pause before Entra writes'
    } finally { $cleanupState.TenantChanged = $false; $cleanupState.DelayAction = '' }
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.DelayAction = 'Cancel'
    $cancelWorker = [Management.Automation.PowerShell]::Create()
    try {
        [void]$cancelWorker.AddScript(@'
param($Path,$Arguments,$State,$ContextProvider,$GraphSource)
$cleanupState = $State
function Get-MgContext { & $ContextProvider }
Invoke-Expression ("function Invoke-MgGraphRequest { " + $GraphSource + " }")
function Start-Sleep {
    param([int]$Seconds)
    [void]$cleanupState.Delays.Add($Seconds)
    throw [System.Management.Automation.PipelineStoppedException]::new()
}
& $Path @Arguments
'@).AddArgument($cleanupPage.Path).AddArgument($deleteArguments).AddArgument($cleanupState).AddArgument($cleanupContextProvider).AddArgument(${function:Invoke-MgGraphRequest}.ToString())
        $cancelled = $false
        try { [void]$cancelWorker.Invoke() } catch { $cancelled = $true }
        $cancelled = $cancelled -or $cancelWorker.InvocationStateInfo.State -in @('Stopped','Failed')
    } finally { $cancelWorker.Dispose(); $cleanupState.DelayAction = '' }
    Assert-EntraTest ($cancelled -and $cleanupState.Writes.Count -eq 2 -and $cleanupState.Delays.Count -eq 1) 'Cancellation during the pause propagates and prevents Entra deletion'
    $cleanupState.UsePostDeleteInventory = $true
    foreach ($mode in @('WrongDevice','Hybrid')) {
        $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.PostDeleteEntraMode = $mode
        $changedRows = @(& $cleanupPage.Path @deleteArguments)
        Assert-EntraTest ($changedRows[0].Status -like 'Failed:*identity or join type changed*' -and $cleanupState.Writes.Count -eq 2 -and $cleanupState.Delays.Count -eq 1) "Entra '$mode' after the pause blocks directory deletion"
    }
    $cleanupState.UsePostDeleteInventory = $false; $cleanupState.PostDeleteEntraMode = ''
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear(); $cleanupState.StickyAutopilotId = $cleanupState.Entra.id
    $pendingRows = @(& $cleanupPage.Path @deleteArguments -RemovalTimeoutSeconds 1)
    Assert-EntraTest ($pendingRows[0].Status -like 'Pending confirmation:*Entra*' -and -not $pendingRows[0].CanDelete -and $pendingRows[0].AutopilotRemovalRequested -and $cleanupState.Writes.Count -eq 3) 'Entra confirmation still reports pending when the final deletion is not verified'
    $cleanupState.StickyAutopilotId = ''
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Entra.trustType = 'ServerAd'
    $hybridRows = @(& $cleanupPage.Path @deleteArguments)
    Assert-EntraTest ($cleanupState.Writes.Count -eq 0 -and $hybridRows[0].Status -like '*blocked*') 'Refreshed hybrid identity blocks every deletion stage'
    $cleanupState.Entra.trustType = 'AzureAd'
    $cleanupState.InventoryFailure = $true; $inventoryRejected = $false
    try { & $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru | Out-Null } catch { $inventoryRejected = $true }
    Assert-EntraTest ($inventoryRejected -and $cleanupState.Writes.Count -eq 0) 'Failed Intune inventory never becomes unmanaged No'
    $cleanupState.InventoryFailure = $false
    $badTenantArguments = $deleteArguments.Clone(); $badTenantArguments.ExpectedTenantId = 'different-tenant'
    $tenantRejected = $false
    try { & $cleanupPage.Path @badTenantArguments | Out-Null } catch { $tenantRejected = $true }
    Assert-EntraTest ($tenantRejected -and $cleanupState.Writes.Count -eq 0) 'Cleanup rejects tenant change without writes'
    $cleanupState.Autopilot = @()
    $markerRows = @(& $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest (-not $markerRows[0].CanDelete -and $markerRows[0].Status -like '*marker*') 'Unresolved Autopilot marker blocks cleanup'
    $cleanupState.InventoryFailure = $false; $cleanupState.Managed = @(); $cleanupState.Autopilot = @(); $cleanupState.Entra.physicalIds = @(); $cleanupState.Entra.approximateLastSignInDateTime = $null
    [pscustomobject]@{ DeviceName='TEST-CLEANUP' } | Export-Csv -LiteralPath $cleanupCsv -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    $unmanagedRows = @(& $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($unmanagedRows[0].'Intune managed' -eq 'No' -and $null -eq $unmanagedRows[0].'Activity date' -and $unmanagedRows[0].CanDelete) 'Simple semicolon name CSV and empty inventories retain unknown date and unmanaged No'
    Assert-EntraTest (@(Select-EntraCleanupRows -Rows $unmanagedRows -DateMode 'Unknown dates').Count -eq 1 -and @(Select-EntraCleanupRows -Rows $unmanagedRows -DateMode 'Known dates').Count -eq 0) 'Unknown date filtering is explicit'
    $cleanupState.DuplicateName = $true
    $duplicateRows = @(& $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($duplicateRows.Count -eq 2 -and @($duplicateRows.ObjectId | Select-Object -Unique).Count -eq 2) 'Duplicate device names retain separate Entra object IDs'
    $cleanupState.DuplicateName = $false
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear(); $cleanupState.Delays.Clear()
    $unmanagedDelete = $deleteArguments.Clone(); $unmanagedDelete.Device = $unmanagedRows
    $unmanagedDeleted = @(& $cleanupPage.Path @unmanagedDelete)
    Assert-EntraTest ($unmanagedDeleted[0].Status -eq 'Deleted' -and $cleanupState.Writes.Count -eq 1 -and $cleanupState.Writes[0] -like '*/devices/*' -and $cleanupState.Delays.Count -eq 0) 'Unmanaged non-Autopilot device only deletes Entra record without delay'
    $cleanupState.Writes.Clear(); $cleanupState.Removed.Clear()
    [pscustomobject]@{ DeviceName='TEST-CLEANUP'; id=$cleanupState.Entra.id; deviceId='dddddddd-2222-2222-2222-222222222222' } | Export-Csv -LiteralPath $cleanupCsv -NoTypeInformation
    $mismatchRows = @(& $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru)
    Assert-EntraTest ($mismatchRows[0].Status -like 'Failed:*different devices*' -and -not $mismatchRows[0].CanDelete) 'Mismatched CSV IDs do not fall back to a name match'
    [pscustomobject]@{ DeviceName='TEST-CLEANUP'; deviceId='invalid' } | Export-Csv -LiteralPath $cleanupCsv -NoTypeInformation
    $invalidRejected = $false
    try { & $cleanupPage.Path -CsvPath $cleanupCsv -SkipGraphConnect -PassThru | Out-Null } catch { $invalidRejected = $true }
    Assert-EntraTest $invalidRejected 'Invalid CSV GUID rejected before Graph lookup'
    $script:ConnectedTenant = ''

    $workerBootstrap = @'
function Import-Module { }
function Get-MgContext { [pscustomobject]@{ Account='mock@example.test'; TenantId='test-tenant'; Scopes=@('User.Read.All') } }
function Connect-MgGraph { throw 'Unexpected authentication' }
'@
    $worker = [Management.Automation.PowerShell]::Create()
    [void]$worker.AddScript($workerBootstrap).AddStatement()
    $workerQueue = [Collections.Concurrent.ConcurrentQueue[object]]::new()
    $workerState = [hashtable]::Synchronized(@{ Fatal=''; Connected=$null; Account=''; Tenant='' })
    [void]$worker.AddScript($script:WorkerSource).AddArgument($primaryPage.Path).AddArgument(@{}).AddArgument(@()).AddArgument(@('User.Read.All')).AddArgument('Connect').AddArgument('different-tenant').AddArgument('').AddArgument($workerQueue).AddArgument($workerState)
    [void]$worker.Invoke()
    Assert-EntraTest ($workerState.Fatal -like '*tenant changed*' -and $workerQueue.IsEmpty) 'Runner rejects changed tenant before executing task'
    $worker.Dispose()

    function Wait-EntraTestJob {
        param([switch]$StrictChecks)
        $frame = [Windows.Threading.DispatcherFrame]::new()
        $waitTimer = [Windows.Threading.DispatcherTimer]::new()
        $waitTimer.Interval = [TimeSpan]::FromMilliseconds(50)
        $waitState = @{ Frame=$frame; Timer=$waitTimer; Clock=[Diagnostics.Stopwatch]::StartNew(); TimedOut=$false; StrictChecks=[bool]$StrictChecks; Failure='' }
        $waitTimer.Tag = $waitState
        $waitTimer.Add_Tick({
            param($sender,$eventArgs)
            if ($sender.Tag.StrictChecks -and $script:Job) {
                try { & { Set-StrictMode -Version Latest; Receive-EntraOperation } }
                catch { $sender.Tag.Failure = $_.Exception.Message; $sender.Stop(); $sender.Tag.Frame.Continue=$false; return }
            }
            if (-not $script:Job -or $sender.Tag.Clock.Elapsed.TotalSeconds -gt 15) {
                $sender.Tag.TimedOut = [bool]$script:Job
                $sender.Stop(); $sender.Tag.Frame.Continue=$false
            }
        })
        $waitTimer.Start()
        [Windows.Threading.Dispatcher]::PushFrame($frame)
        if ($waitState.Failure) { Stop-EntraOperation; throw "StrictMode completion failed: $($waitState.Failure)" }
        if ($waitState.TimedOut) { Stop-EntraOperation; throw 'Timed out waiting for mocked background operation.' }
    }
    $script:Runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:Runspace.ApartmentState = 'STA'; $script:Runspace.ThreadOptions='UseNewThread'; $script:Runspace.Open()
    $bootstrap = [Management.Automation.PowerShell]::Create(); $bootstrap.Runspace=$script:Runspace
    [void]$bootstrap.AddScript(@'
function Import-Module { Write-Verbose "Importing cmdlet 'Mock-GraphCommand'." }
function Get-MgContext { [pscustomobject]@{ Account='mock@example.test'; TenantId='mock-tenant'; Scopes=@('User.ReadWrite.All','User.Read.All','GroupMember.Read.All','Device.Read.All','DeviceManagementManagedDevices.Read.All','DeviceManagementServiceConfig.Read.All') } }
function Connect-MgGraph { throw 'Unexpected sign-in' }
function Disconnect-MgGraph { }
function Get-MgUser { [pscustomobject]@{ Id='user-1'; UserPrincipalName='mock@example.test' } }
function Get-MgUserOwnedDevice { while ($true) { Get-Date | Out-Null } }
function Get-MgUserRegisteredDevice { @() }
function Invoke-MgGraphRequest {
    [CmdletBinding()]param($Method,$Uri,$OutputType)
    if ($Method -ne 'GET') { throw 'Unexpected write in cleanup preview worker' }
    if ($Uri -like '*deviceManagement/*') { return [pscustomobject]@{ value=@(); '@odata.nextLink'=$null } }
    $entry = [pscustomobject]@{ id='cccccccc-1111-1111-1111-111111111111'; deviceId='dddddddd-1111-1111-1111-111111111111'; displayName='TEST-CLEANUP'; trustType='AzureAd'; physicalIds=@(); approximateLastSignInDateTime='2026-10-21T08:00:00Z' }
    if ($Uri -like '*/devices/*') { return $entry }
    return [pscustomobject]@{ value=@($entry); '@odata.nextLink'=$null }
}
'@)
    [void]$bootstrap.Invoke(); $bootstrap.Dispose()
    $originalWorkerSource = $script:WorkerSource
    try {
        $script:WorkerSource = @'
param($Path, $Parameters, $Modules, $Scopes, $Kind, $ExpectedTenant, $RequestedTenant, $Queue, $State)
if ('Microsoft.Graph.Reports' -notin $Modules -or 'AuditLog.Read.All' -notin $Scopes -or 'DeviceManagementManagedDevices.Read.All' -notin $Scopes) { throw 'Sign-in-only worker requirements missing' }
if ($Parameters.DeviceSource -ne 'SignInLogs' -or $Parameters.SignInFallback) { throw 'Sign-in-only worker parameters incorrect' }
$State.Account='mock@example.test'; $State.Tenant='mock-tenant'; $State.Connected=$true
if ($Kind -eq 'Run') {
    if ($Parameters.DeviceMode -ne 'Most' -or $Parameters.SignInDays -ne 14) { throw 'Sign-in-only selection not forwarded to worker' }
    $Queue.Enqueue([pscustomobject]@{ DeviceName='LOG-PC'; DeviceObjectId='log-object'; Source='SignInLogs (2 events)' })
}
'@
        $primaryPage.Controls.DeviceSource.SelectedItem = 'Sign-in logs only'
        $primaryPage.Controls.OnlyIntuneDevices.IsChecked = $true
        $primaryPage.Controls.SignInDays.Text = '14'
        $primaryPage.Controls.OutputPath.Text = Join-Path $tempDirectory 'SignInWorker.csv'
        Start-EntraOperation -Page $primaryPage -Kind 'Connect'
        Wait-EntraTestJob
        Assert-EntraTest ($primaryPage.Status.Text -eq 'Completed') 'GUI Sign in forwards sign-in-only audit/module and optional Intune requirements'
        Start-EntraOperation $primaryPage
        Wait-EntraTestJob
        Assert-EntraTest ($primaryPage.Status.Text -eq 'Completed' -and $primaryPage.Results.Count -eq 1 -and $primaryPage.Results[0].DeviceObjectId -eq 'log-object') 'GUI Run forwards sign-in-only selection/lookback and receives typed results'
    } finally {
        $script:WorkerSource = $originalWorkerSource
        $primaryPage.Controls.DeviceSource.SelectedItem = 'Entra owned/registered devices'
        $primaryPage.Controls.OnlyIntuneDevices.IsChecked = $false
        $primaryPage.Controls.SignInDays.Text = '30'
    }
    $verboseBootstrap = [Management.Automation.PowerShell]::Create(); $verboseBootstrap.Runspace=$script:Runspace
    [void]$verboseBootstrap.AddScript(@'
function Get-MgUserOwnedDevice { [pscustomobject]@{ Id='verbose-device'; AdditionalProperties=@{ '@odata.type'='#microsoft.graph.device' } } }
function Get-MgDevice { [pscustomobject]@{ Id='verbose-device'; DisplayName='TEST-VERBOSE'; OperatingSystem='Linux'; TrustType='AzureAd'; IsManaged=$false; MdmAppId=''; ApproximateLastSignInDateTime=[datetime]'2026-11-01' } }
'@)
    [void]$verboseBootstrap.Invoke(); $verboseBootstrap.Dispose()
    $primaryPage.Controls.Verbose.IsChecked = $true
    $primaryPage.Controls.OutputPath.Text = Join-Path $tempDirectory 'PrimaryPC-verbose.csv'
    Start-EntraOperation $primaryPage
    Wait-EntraTestJob
    $verboseActivity = [Windows.Documents.TextRange]::new($primaryPage.Log.Document.ContentStart,$primaryPage.Log.Document.ContentEnd).Text
    Assert-EntraTest ($verboseActivity -like '*[[]DETAIL[]]*TEST-VERBOSE OS=Linux*') 'Primary-PC verbose device details reach the GUI activity log'
    Assert-EntraTest ($verboseActivity -notlike '*Importing cmdlet*') 'Enabling feature verbose output does not expose module import chatter'
    $primaryPage.Log.Document.Blocks.Clear()
    $primaryPage.Controls.Verbose.IsChecked = $false
    $primaryPage.Controls.OutputPath.Text = Join-Path $tempDirectory 'PrimaryPC-quiet.csv'
    Start-EntraOperation $primaryPage
    Wait-EntraTestJob
    $quietActivity = [Windows.Documents.TextRange]::new($primaryPage.Log.Document.ContentStart,$primaryPage.Log.Document.ContentEnd).Text
    Assert-EntraTest ($quietActivity -notlike '*TEST-VERBOSE OS=Linux*') 'Disabling verbose logging suppresses device details on the next run'
    $verboseBootstrap = [Management.Automation.PowerShell]::Create(); $verboseBootstrap.Runspace=$script:Runspace
    [void]$verboseBootstrap.AddScript('function Get-MgUserOwnedDevice { while ($true) { Get-Date | Out-Null } }')
    [void]$verboseBootstrap.Invoke(); $verboseBootstrap.Dispose()
    $attributePage.Controls.WhatIf.IsChecked = $true
    Start-EntraOperation $attributePage
    Assert-EntraTest ([bool]$script:Job -and -not $attributePage.Form.IsEnabled) 'Background runner starts and locks parameter edits'
    Wait-EntraTestJob
    Assert-EntraTest ($attributePage.Results.Count -eq 1 -and $attributePage.Results[0].Status -eq 'WhatIf' -and $attributePage.Form.IsEnabled) 'Background runner delivers typed results and restores form'
    Assert-EntraTest ($statusLight.Fill.Color -eq $ui.Window.Resources['SuccessBrush'].Color -and $ui.Window.FindName('ConnectionText').Text -eq 'Connected') 'Confirmed connection turns status light green'
    Assert-EntraTest ($ui.Window.FindName('AccountText').Text -eq 'mock@example.test' -and $ui.Window.FindName('TenantText').Text -eq 'mock-tenant') 'Sidebar shows connected account and tenant'
    $themeButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    Assert-EntraTest ($statusLight.Fill.Color -eq $ui.Window.Resources['SuccessBrush'].Color) 'Connected light remains green in light theme'
    $themeButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    $activityText = [Windows.Documents.TextRange]::new($attributePage.Log.Document.ContentStart,$attributePage.Log.Document.ContentEnd).Text
    Assert-EntraTest ($activityText -notlike '*Importing cmdlet*') 'Module import verbose output does not flood activity log'
    Assert-EntraTest ($activityText -like '*Connected as:*' -and $activityText -like '*Completed*') 'Write-Host stream and completion reach activity log'
    Assert-EntraTest ($activityText -like '*Run summary: Processed 1; Success 0; Failed 0; Skipped 0; Preview 1; Other 0*') 'Activity log summarizes emitted WhatIf results'
    Assert-EntraTest ($activityText -like '*Activity:*Run started.*Connected as:*') 'Activity summary includes run and worker information messages'
    $lastActivityBlock = $attributePage.Log.Document.Blocks.LastBlock
    $lastActivityText = [Windows.Documents.TextRange]::new($lastActivityBlock.ContentStart,$lastActivityBlock.ContentEnd).Text
    Assert-EntraTest ($lastActivityText -like '*Run summary: Processed 1; Success 0; Failed 0; Skipped 0; Preview 1; Other 0*') 'Run summary remains the final activity-log entry'
    $originalWorkerSource = $script:WorkerSource
    try {
        $script:WorkerSource = @'
param($Path, $Parameters, $Modules, $Scopes, $Kind, $ExpectedTenant, $RequestedTenant, $Queue, $State)
Write-Information 'Fetching managed devices from device-name input...' -InformationAction Continue
Write-Warning "Managed device name 'D-04514' could not be resolved."
$rows = @(
    [pscustomobject]@{ DeviceName='D-04527'; ResolutionStatus='PrimaryUserResolved'; LookupError=''; Manufacturer='LENOVO' },
    [pscustomobject]@{ DeviceName='D-04504'; DeviceId='d77f7175-1833-46e8-9d4c-7b8789273e7c'; ResolutionStatus='PrimaryUserMissing'; LookupError='No usersLoggedOn userId or primary userId was available.'; Manufacturer='LENOVO' }
)
$rows | Export-Csv -LiteralPath $Parameters['OutputPath'] -NoTypeInformation
Write-Information "Report exported to $($Parameters['OutputPath'])" -InformationAction Continue
foreach ($row in $rows) { $Queue.Enqueue($row) }
'@
        $lastUserOutput = Join-Path $tempDirectory 'LastUser-summary.csv'
        $lastPage.Controls.Source.SelectedItem = 'Device names'
        $lastPage.Controls.DeviceName.Text = 'D-04504'
        $lastPage.Controls.OutputPath.Text = $lastUserOutput
        Start-EntraOperation -Page $lastPage
        Assert-EntraTest ([bool]$script:Job) 'LastUser regression starts through the real form invocation'
        $script:Timer.Stop()
        Wait-EntraTestJob -StrictChecks
        $lastUserLog = [Windows.Documents.TextRange]::new($lastPage.Log.Document.ContentStart,$lastPage.Log.Document.ContentEnd).Text
        Assert-EntraTest ($lastUserLog -like '*DeviceName: D-04504; Error: No usersLoggedOn userId or primary userId was available.*' -and $lastUserLog -notmatch 'Manufacturer\s*:|DeviceId\s*:') 'LastUser runner logs concise device warnings without result dumps'
        Assert-EntraTest ($lastUserLog.TrimEnd().EndsWith('Run summary: Processed 2; Success 1; Failed 1; Skipped 0; Preview 0; Other 0')) 'LastUser runner ends with success and failure totals after warning details'
        Assert-EntraTest ($lastPage.Results.Count -eq 2 -and $lastPage.Results[1].Manufacturer -eq 'LENOVO' -and $lastPage.OpenOutput.IsEnabled -and -not $script:Job) 'LastUser summary preserves full results and export while releasing the job'
    } finally { $script:WorkerSource = $originalWorkerSource }
    [pscustomobject]@{ DeviceName='TEST-CLEANUP' } | Export-Csv -LiteralPath $cleanupCsv -NoTypeInformation
    $cleanupPage.Controls.CsvPath.Text = $cleanupCsv
    Start-EntraOperation $cleanupPage
    Wait-EntraTestJob
    Assert-EntraTest ($cleanupPage.Results.Count -eq 1 -and $cleanupPage.Results[0].TenantId -eq 'mock-tenant') 'Background cleanup inventory returns typed same-tenant rows'
    $cleanupPage.Table.SelectedItem = $cleanupPage.Results[0]
    $previewInvocation = Get-EntraCleanupDeletionInvocation $cleanupPage
    Start-EntraOperation -Page $cleanupPage -PreparedInvocation $previewInvocation
    Wait-EntraTestJob
    Assert-EntraTest ($cleanupPage.Results.Count -eq 1 -and $cleanupPage.Results[0].Status -like 'Preview*') 'Background cleanup preview replaces the existing row without clearing or duplicating results'
    & {
        Set-StrictMode -Version Latest
        $requirements = Get-EntraRequirements $cleanupPage @{}
        Assert-EntraTest ('Device.Read.All' -in $requirements.Scopes) 'StrictMode cleanup connection accepts absent Action and WhatIf keys'
        Start-EntraOperation $cleanupPage
        Assert-EntraTest ([bool]$script:Job) 'StrictMode cleanup load starts without OutputPath'
        $script:Timer.Stop()
        Wait-EntraTestJob -StrictChecks
        Assert-EntraTest ($cleanupPage.Results.Count -eq 1 -and $cleanupPage.Status.Text -eq 'Completed' -and -not $cleanupPage.OpenOutput.IsEnabled -and -not $script:Job) 'StrictMode cleanup load completes and releases job without OutputPath'
        $cleanupPage.Table.SelectedItem = $cleanupPage.Results[0]
        $invocation = Get-EntraCleanupDeletionInvocation $cleanupPage
        Start-EntraOperation -Page $cleanupPage -PreparedInvocation $invocation
        Assert-EntraTest ([bool]$script:Job) 'StrictMode cleanup preview starts without OutputPath'
        $script:Timer.Stop()
        Wait-EntraTestJob -StrictChecks
        Assert-EntraTest ($cleanupPage.Results.Count -eq 1 -and $cleanupPage.Results[0].Status -like 'Preview*' -and -not $script:Job) 'StrictMode preview completion preserves keyed results without OutputPath'
        Start-EntraOperation $attributePage
        Assert-EntraTest ([bool]$script:Job -and $attributePage.Results.Count -eq 0) 'StrictMode ordinary run still clears results when Action is absent'
        $script:Timer.Stop()
        Wait-EntraTestJob -StrictChecks
        Assert-EntraTest ($attributePage.Status.Text -eq 'Completed' -and $attributePage.Results.Count -eq 1 -and -not $script:Job) 'StrictMode ordinary run completes without optional Action or OutputPath'
    }
    $primaryPage.Controls.OutputPath.Text = Join-Path $tempDirectory 'Cancelled.csv'
    Start-EntraOperation $primaryPage
    Stop-EntraOperation
    Wait-EntraTestJob
    Assert-EntraTest ($primaryPage.Status.Text -like 'Stopped*' -and $primaryPage.Form.IsEnabled) 'Stop cancels background pipeline without locking interface'
    Start-EntraOperation -Page $groupPage -Kind 'Connect'
    Wait-EntraTestJob
    Assert-EntraTest ($groupPage.Status.Text -eq 'Finished with errors' -and $groupPage.Form.IsEnabled) 'Authentication failure is visible and restores controls'
    Assert-EntraTest ($statusLight.Fill.Color -eq $ui.Window.Resources['SuccessBrush'].Color) 'Failed scope request does not falsely disconnect existing session'
    Start-EntraOperation -Page $groupPage -Kind 'Disconnect'
    Wait-EntraTestJob
    Assert-EntraTest ($statusLight.Fill.Color -eq $ui.Window.Resources['ErrorBrush'].Color -and $ui.Window.FindName('ConnectionText').Text -eq 'Not connected') 'Sign-out turns status light red'
    Assert-EntraTest (-not $ui.Window.FindName('AccountText').Text -and -not $ui.Window.FindName('TenantText').Text -and -not $ui.Window.FindName('DisconnectButton').IsEnabled) 'Sign-out clears identity and disables sign-out action'
    $ui.Window.FindName('AccountText').Text = 'a-very-long-account-name-that-needs-trimming@example.test'
    $ui.Window.FindName('TenantText').Text = '11111111-2222-3333-4444-555555555555'

    foreach ($size in @(@(1050,720),@(1280,900),@(1600,1000))) {
        $ui.Window.Width=$size[0]; $ui.Window.Height=$size[1]
        foreach ($testPage in $ui.Pages) {
            $script:PagesHost.SelectedIndex = [int]$testPage.Navigation.Tag
            $testPage.Activity.IsExpanded = $true
            $ui.Window.Measure([Windows.Size]::new($size[0],$size[1]))
            $ui.Window.Arrange([Windows.Rect]::new(0,0,$size[0],$size[1]))
            $ui.Window.UpdateLayout()
            Assert-EntraTest ($testPage.Table.ActualWidth -gt 100 -and $testPage.Table.ActualHeight -gt 0) "Layout $($size[0])x$($size[1]): $($testPage.Id)"
            if ($testPage.CompactForm) {
                Assert-EntraTest ($testPage.FormScroll.ScrollableHeight -le 1) "No parameter scrollbar at $($size[0])x$($size[1]): $($testPage.Id)"
                $footerBottom = $testPage.Footer.TranslatePoint([Windows.Point]::new(0,$testPage.Footer.ActualHeight),$testPage.Layout).Y
                Assert-EntraTest ($footerBottom -le ($testPage.Layout.ActualHeight + 1)) "Compact form and expanded log fit page: $($testPage.Id)"
                $variants = switch ($testPage.Id) {
                    'Members' { @(@{ Source='Names' },@{ Source='Text file' }) }
                    'PrimaryPC' { @(@{ Source='Choose a source'; DeviceSource='Entra owned/registered devices' },@{ Source='Group name' },@{ Source='Group ID' },@{ Source='User CSV' },@{ Source='User UPNs' },@{ Source='User UPNs'; DeviceSource='Sign-in logs only' },@{ Source='Group name'; DeviceSource='Sign-in logs only' }) }
                    'Attributes' { @(@{ Source='UPN list'; Action='Set' },@{ Source='CSV file'; Action='Clear' }) }
                    default { @(@{}) }
                }
                foreach ($variant in $variants) {
                    foreach ($key in $variant.Keys) { $testPage.Controls[$key].SelectedItem = $variant[$key] }
                    foreach ($expanded in @($false,$true)) {
                        $testPage.Activity.IsExpanded = $expanded
                        foreach ($lightTheme in @($false,$true)) {
                            if ($script:IsDark -eq $lightTheme) { $themeButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent)) }
                            $ui.Window.UpdateLayout()
                            $fieldsFit = $true
                            foreach ($field in $testPage.Form.Children) {
                                if ($field.Visibility -eq 'Collapsed') { continue }
                                $top = $field.TranslatePoint([Windows.Point]::new(0,0),$testPage.FormScroll)
                                if ($top.Y -lt -1 -or ($top.Y + $field.ActualHeight) -gt ($testPage.FormScroll.ActualHeight + 1)) { $fieldsFit=$false }
                            }
                            $footerBottom = $testPage.Footer.TranslatePoint([Windows.Point]::new(0,$testPage.Footer.ActualHeight),$testPage.Layout).Y
                            $scenario = "$($testPage.Id) $($size[0])x$($size[1]) $($variant.Source) expanded=$expanded light=$lightTheme"
                            Assert-EntraTest ($fieldsFit -and $testPage.FormScroll.ScrollableHeight -le 1) "All parameter fields visible: $scenario"
                            Assert-EntraTest ($footerBottom -le ($testPage.Layout.ActualHeight + 1) -and $testPage.Table.ActualHeight -ge ($testPage.Table.MinHeight - 1)) "Results and log remain within page: $scenario"
                        }
                    }
                }
                if (-not $script:IsDark) { $themeButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent)) }
            }
        }
        $connectionPanel = $ui.Window.FindName('ConnectionPanel')
        $panelPosition = $connectionPanel.TranslatePoint([Windows.Point]::new(0,0),$ui.Window)
        Assert-EntraTest ($panelPosition.X -lt 30 -and $panelPosition.Y -gt ($ui.Window.ActualHeight / 2) -and ($panelPosition.Y + $connectionPanel.ActualHeight) -le $ui.Window.ActualHeight) "Connection panel stays in lower-left corner at $($size[0])x$($size[1])"
        $navBottom = $script:Navigation.TranslatePoint([Windows.Point]::new(0,$script:Navigation.ActualHeight),$ui.Window).Y
        Assert-EntraTest ($navBottom -le $panelPosition.Y) "Navigation does not overlap connection footer at $($size[0])x$($size[1])"
    }
    $ui.Window.FindName('AccountText').Text = ''
    $ui.Window.FindName('TenantText').Text = ''
    $ui.Window.Width=1280; $ui.Window.Height=900
    $memberPage.Controls.Source.SelectedItem='Names'
    $attributePage.Controls.Source.SelectedItem='UPN list'
    $attributePage.Controls.Action.SelectedItem='Set'
    foreach ($compactPage in @($memberPage,$primaryPage,$attributePage,$cleanupPage)) {
        $script:PagesHost.SelectedIndex=[int]$compactPage.Navigation.Tag
        $compactPage.Activity.IsExpanded=$false
        $ui.Window.UpdateLayout()
        $pageBitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new(1280,900,96,96,[Windows.Media.PixelFormats]::Pbgra32)
        $pageBitmap.Render($ui.Window)
        $pageEncoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
        $pageEncoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($pageBitmap))
        $pageImagePath = Join-Path ([IO.Path]::GetTempPath()) "EntraToolbox-$($compactPage.Id)-$($PSVersionTable.PSVersion.Major).png"
        $pageImageStream = [IO.File]::Create($pageImagePath)
        try { $pageEncoder.Save($pageImageStream) } finally { $pageImageStream.Dispose() }
        Write-Host "Screenshot: $pageImagePath"
    }
    $script:PagesHost.SelectedIndex=[int]$attributePage.Navigation.Tag
    $ui.Window.UpdateLayout()
    $render = [Windows.Media.Imaging.RenderTargetBitmap]::new(1280,900,96,96,[Windows.Media.PixelFormats]::Pbgra32)
    $render.Render($ui.Window)
    $imagePath = Join-Path ([IO.Path]::GetTempPath()) "EntraToolbox-$($PSVersionTable.PSVersion.Major).png"
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($render))
    $imageStream = [IO.File]::OpenWrite($imagePath)
    try { $encoder.Save($imageStream) } finally { $imageStream.Dispose() }
    Write-Host "Screenshot: $imagePath"
    $ui.Window.Close()
    Write-Host "PASS: $script:CheckCount offline checks on PowerShell $($PSVersionTable.PSVersion)"
} finally { Remove-Item -LiteralPath $tempDirectory -Recurse -Force; Remove-Variable -Name EntraMockState -Scope Global -ErrorAction SilentlyContinue }